#!/usr/bin/env python3
"""
extract_contacts_no_llm.py
============================

Extract a unique, deduplicated list of external business contacts (clients /
customers / vendors) from Outlook (.msg / .eml) or Gmail (.mbox, from Google
Takeout) email archives - PURE PYTHON, NO LLM, NO API CALLS, NO COST.

This is the rule-based counterpart to email_contact_extractor.py. It runs
the same parse -> filter -> aggregate -> dedupe -> export pipeline, but
instead of sending unique candidates to Claude for cleanup, it extracts
name and company directly from the email headers and a regex-scanned
signature block. Nothing ever leaves your machine.

Trade-off, stated plainly: rule-based extraction is less accurate than the
LLM version on messy or unusual signatures (foreign name orders, signatures
with no company suffix like "Inc"/"LLC", contacts on personal Gmail used for
business, etc). It also cannot judge "is this actually a business contact"
the way a model can - so by default nothing is silently dropped on that
basis; every non-internal, non-automated address is included, and it's on
you to skim the output. Use --exclude-public-domains if you'd rather only
keep contacts on a company domain.

Pipeline
--------
1. PARSE     - Walk the input directory and parse messages from .eml, .msg,
               or .mbox files into a lightweight, uniform representation.
2. FILTER    - Drop internal-company addresses and automated/system senders
               *per address*, not per message (a message from a colleague
               that CC's a client should still yield the client).
3. AGGREGATE - Collapse to one candidate row per unique email address,
               keeping the best available display name and a short
               signature snippet.
4. ENRICH    - Rule-based only: trust the header display name when present;
               otherwise look for a name-shaped line in the signature.
               Look for a line containing a company suffix (Inc/LLC/Ltd/
               GmbH/...) in the signature; otherwise guess from the domain.
5. MERGE     - Deduplicate on normalized email, keeping the richer fields.
6. EXPORT    - Write out contacts.csv and contacts.xlsx.

Quick start
-----------
    pip install -r requirements_no_llm.txt

    # Gmail (Google Takeout .mbox export)
    python extract_contacts_no_llm.py \\
        --mode gmail \\
        --input-dir ./takeout_export \\
        --internal-domains "mycompany.com,mycompany.co.uk" \\
        --self-emails "me@mycompany.com,me@gmail.com"

    # Outlook (.msg or .eml export)
    python extract_contacts_no_llm.py \\
        --mode outlook \\
        --input-dir ./outlook_export \\
        --internal-domains "mycompany.com"

See README.md ("Handling .pst files") if you're starting from a .pst.
"""

from __future__ import annotations

import argparse
import csv
import html as html_lib
import json
import logging
import mailbox
import os
import re
import sys
from dataclasses import dataclass, asdict
from email import policy
from email.message import EmailMessage
from email.parser import BytesParser
from email.utils import getaddresses
from pathlib import Path
from typing import Dict, Iterable, Iterator, List, Optional, Set, Tuple

# --------------------------------------------------------------------------
# Dependencies - only pandas/openpyxl are hard requirements. extract_msg is
# optional (only needed for .msg files). No LLM SDK, no pydantic.
# --------------------------------------------------------------------------
try:
    import pandas as pd
except ImportError:
    print("Missing dependency 'pandas'. Run: pip install -r requirements_no_llm.txt", file=sys.stderr)
    raise

try:
    import extract_msg  # type: ignore
    HAS_EXTRACT_MSG = True
except ImportError:
    HAS_EXTRACT_MSG = False


# ==========================================================================
# Logging
# ==========================================================================

def setup_logging(verbose: bool, log_file: Optional[str]) -> logging.Logger:
    logger = logging.getLogger("email_extractor_no_llm")
    logger.setLevel(logging.DEBUG if verbose else logging.INFO)
    logger.handlers.clear()

    fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", "%H:%M:%S")

    console = logging.StreamHandler(sys.stdout)
    console.setFormatter(fmt)
    logger.addHandler(console)

    if log_file:
        Path(log_file).parent.mkdir(parents=True, exist_ok=True)
        fh = logging.FileHandler(log_file, encoding="utf-8")
        fh.setFormatter(fmt)
        fh.setLevel(logging.DEBUG)
        logger.addHandler(fh)

    return logger


log = logging.getLogger("email_extractor_no_llm")


# ==========================================================================
# Constants
# ==========================================================================

DEFAULT_AUTOMATED_LOCAL_PARTS = [
    r"no[-_.]?reply", r"do[-_.]?not[-_.]?reply", r"notifications?", r"notify",
    r"alerts?", r"mailer[-_.]?daemon", r"postmaster", r"newsletter", r"news",
    r"marketing", r"campaign", r"bounce", r"automated", r"system", r"admin",
    r"info", r"support", r"help(desk)?", r"billing", r"subscriptions?",
    r"unsubscribe", r"feedback", r"survey", r"webmaster", r"digest",
]

PUBLIC_EMAIL_DOMAINS = {
    "gmail.com", "googlemail.com", "yahoo.com", "ymail.com", "hotmail.com",
    "outlook.com", "live.com", "msn.com", "icloud.com", "me.com", "mac.com",
    "aol.com", "protonmail.com", "proton.me", "gmx.com", "zoho.com",
    "yandex.com", "mail.com", "hey.com", "fastmail.com",
}

TWO_LABEL_SUFFIXES = {
    "co.uk", "org.uk", "ac.uk", "gov.uk", "ltd.uk", "plc.uk",
    "co.jp", "co.nz", "co.za", "co.in", "co.kr", "co.il", "co.th",
    "com.au", "com.br", "com.mx", "com.sg", "com.hk", "com.cn", "com.tw",
    "net.au", "org.au", "net.nz", "org.nz",
}

SIGNATURE_MARKERS = [
    r"^--\s*$",
    r"^best regards,?\s*$", r"^regards,?\s*$", r"^best,?\s*$",
    r"^sincerely,?\s*$", r"^thanks,?\s*$", r"^thank you,?\s*$",
    r"^cheers,?\s*$", r"^warm regards,?\s*$", r"^kind regards,?\s*$",
    r"^many thanks,?\s*$", r"^respectfully,?\s*$",
    r"sent from my iphone", r"sent from my android", r"sent from my mobile",
    r"get outlook for",
]

EMAIL_REGEX = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")

COMPANY_SUFFIX_RE = re.compile(
    r"\b(Inc\.?|L\.?L\.?C\.?|Ltd\.?|Limited|Corp\.?|Corporation|Co\.|Company|"
    r"Group|GmbH|LLP|PLC|Pty\.?(\s?Ltd\.?)?|S\.?A\.?|S\.?L\.?|S\.?R\.?L\.?|AG|"
    r"Solutions|Technologies|Enterprises|Partners|Associates|Consulting|"
    r"Studio|Agency|Labs|Ventures)\b",
    re.IGNORECASE,
)

# Words that mean "this line is not a person's name" even if it superficially
# looks like one (2-4 capitalized words).
NON_NAME_WORDS = {
    "sent", "from", "mobile", "phone", "tel", "fax", "website", "www",
    "http", "https", "email", "e-mail", "confidential", "disclaimer",
    "unsubscribe", "regards", "best", "sincerely", "thanks", "thank",
    "you", "cheers", "team", "sales", "support", "info", "admin", "office",
    "department", "dept", "customer", "service", "notice", "please",
    "click", "view", "download", "attachment", "attached",
}


@dataclass
class Config:
    mode: str
    input_dir: Path
    internal_domains: Set[str]
    internal_domain_regexes: List[re.Pattern]
    self_emails: Set[str]
    automated_patterns: List[re.Pattern]
    output_csv: Path
    output_xlsx: Path
    cache_file: Optional[Path]
    use_cache: bool
    exclude_public_domains: bool
    max_snippet_chars: int
    include_evidence: bool


# ==========================================================================
# Data model
# ==========================================================================

@dataclass
class RawCandidate:
    email: str
    domain: str
    header_name: str = ""
    snippet: str = ""
    source_file: str = ""
    direction: str = ""


@dataclass
class AggregatedCandidate:
    id: int
    email: str
    domain: str
    header_name: str = ""
    guessed_company: str = ""
    snippet: str = ""
    occurrences: int = 0


@dataclass
class FinalContact:
    full_name: str = ""
    email_address: str = ""
    company_name: str = ""
    signature_snippet: str = ""  # only populated/exported if --include-evidence


# ==========================================================================
# Filtering helpers
# ==========================================================================

def build_internal_matcher(domains: Set[str], regexes: List[re.Pattern]):
    domains = {d.lower().strip().lstrip("@") for d in domains if d.strip()}

    def is_internal(email_addr: str) -> bool:
        email_addr = email_addr.lower().strip()
        if "@" not in email_addr:
            return True
        domain = email_addr.split("@", 1)[1]
        if domain in domains:
            return True
        for d in domains:
            if domain.endswith("." + d):
                return True
        for rx in regexes:
            if rx.search(email_addr):
                return True
        return False

    return is_internal


def build_automated_matcher(patterns: List[re.Pattern]):
    def is_automated(email_addr: str, headers: Optional[dict] = None) -> bool:
        local_part = email_addr.split("@", 1)[0].lower() if "@" in email_addr else email_addr.lower()
        for rx in patterns:
            if rx.search(local_part):
                return True
        if headers:
            if headers.get("List-Unsubscribe") or headers.get("List-Id"):
                return True
            precedence = (headers.get("Precedence") or "").lower()
            if precedence in ("bulk", "list", "junk"):
                return True
            auto_submitted = (headers.get("Auto-Submitted") or "").lower()
            if auto_submitted and auto_submitted != "no":
                return True
        return False

    return is_automated


def guess_company_from_domain(domain: str) -> str:
    """Anchored from the right so subdomains (mail.acme.co.uk) don't get
    mistaken for the company name - the registrable label right before the
    public suffix is what we want (acme), not the first label (mail)."""
    if not domain:
        return ""
    domain = domain.lower().strip()
    if domain in PUBLIC_EMAIL_DOMAINS:
        return ""
    labels = domain.split(".")
    if len(labels) < 2:
        core = labels[0]
    else:
        last_two = ".".join(labels[-2:])
        if last_two in TWO_LABEL_SUFFIXES and len(labels) >= 3:
            core = labels[-3]
        else:
            core = labels[-2]
    core = re.sub(r"[-_]+", " ", core).strip()
    return core.title()


# ==========================================================================
# Body / signature extraction
# ==========================================================================

def html_to_text(html: str) -> str:
    text = re.sub(r"<(script|style)[^>]*>.*?</\1>", " ", html, flags=re.S | re.I)
    text = re.sub(r"<br\s*/?>", "\n", text, flags=re.I)
    text = re.sub(r"</p>", "\n", text, flags=re.I)
    text = re.sub(r"<[^>]+>", " ", text)
    text = html_lib.unescape(text)
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n\s*\n+", "\n", text)
    return text.strip()


def extract_signature_snippet(body: str, max_lines: int = 20, max_chars: int = 1500) -> str:
    if not body:
        return ""
    lines = [l.rstrip() for l in body.splitlines()]
    while lines and not lines[-1].strip():
        lines.pop()
    if not lines:
        return ""

    search_start = max(0, len(lines) - 40)
    marker_idx = None
    for i in range(len(lines) - 1, search_start - 1, -1):
        candidate = lines[i].strip().lower()
        if not candidate:
            continue
        for pattern in SIGNATURE_MARKERS:
            if re.match(pattern, candidate):
                marker_idx = i
                break
        if marker_idx is not None:
            break

    if marker_idx is not None:
        snippet_lines = lines[marker_idx: marker_idx + max_lines + 1]
    else:
        snippet_lines = lines[-max_lines:]

    snippet = "\n".join(snippet_lines).strip()
    return snippet[:max_chars]


def get_body_text(msg: EmailMessage) -> str:
    try:
        body_part = msg.get_body(preferencelist=("plain", "html"))
        if body_part is None:
            return ""
        content = body_part.get_content()
        if body_part.get_content_type() == "text/html":
            return html_to_text(content)
        return content
    except Exception:
        try:
            payload = msg.get_payload(decode=True)
            if payload:
                charset = msg.get_content_charset() or "utf-8"
                return payload.decode(charset, errors="replace")
        except Exception:
            pass
        return ""


def parse_address_list(header_value: Optional[str]) -> List[Tuple[str, str]]:
    if not header_value:
        return []
    normalized = header_value.replace(";", ",")
    pairs = getaddresses([normalized])
    out = []
    for name, addr in pairs:
        addr = addr.strip().lower()
        if EMAIL_REGEX.fullmatch(addr):
            out.append((name.strip(), addr))
    return out


def parse_loose_address_list(raw: Optional[str]) -> List[Tuple[str, str]]:
    if not raw:
        return []
    strict = parse_address_list(raw)
    if strict:
        return strict
    out = []
    for m in EMAIL_REGEX.finditer(raw):
        out.append(("", m.group(0).lower()))
    return out


# ==========================================================================
# Rule-based name / company extraction (this replaces the LLM step)
# ==========================================================================

def looks_like_person_name(s: str) -> bool:
    """Unicode-aware heuristic: 2-4 capitalized words, no digits, no email,
    no company suffix, none of them a known non-name word. Deliberately
    conservative - false negatives (rejecting a real name) are safer here
    than false positives (accepting a role/company as a name)."""
    s = s.strip().strip(",.;:")
    if not s or len(s) > 50:
        return False
    if EMAIL_REGEX.search(s):
        return False
    if any(ch.isdigit() for ch in s):
        return False
    if COMPANY_SUFFIX_RE.search(s):
        return False
    words = s.split()
    if not (2 <= len(words) <= 4):
        return False
    lowered = {w.lower().strip(".,'-") for w in words}
    if lowered & NON_NAME_WORDS:
        return False
    for w in words:
        core = w.strip(".,'-")
        if not core:
            return False
        if not core[0].isupper():
            return False
        if not all(c.isalpha() or c in "'-." for c in core):
            return False
    return True


def clean_header_name(name: str) -> str:
    """Trust the header display name as-is (it's real signal), but tidy up
    the common case of an all-caps or all-lowercase mail client export."""
    name = name.strip().strip('"').strip()
    if not name:
        return ""
    if name.isupper() or name.islower():
        return name.title()
    return name


def extract_name_from_signature(snippet: str, max_check_lines: int = 4) -> str:
    if not snippet:
        return ""
    lines = [l.strip() for l in snippet.splitlines() if l.strip()]
    checked = 0
    for line in lines:
        if any(re.match(pat, line.lower()) for pat in SIGNATURE_MARKERS):
            continue
        if looks_like_person_name(line):
            return line.strip(" ,")
        checked += 1
        if checked >= max_check_lines:
            break
    return ""


def extract_company_from_signature(snippet: str, exclude_line: str = "") -> str:
    if not snippet:
        return ""
    exclude_line = exclude_line.strip()
    for line in snippet.splitlines():
        line = line.strip()
        if not line or line == exclude_line:
            continue
        if EMAIL_REGEX.search(line):
            continue
        if re.match(r"^\+?[\d\s().\-]{7,}$", line):  # phone-number-only line
            continue
        if COMPANY_SUFFIX_RE.search(line) and 2 <= len(line) <= 70:
            return line.strip(" \t,.;:-|")
    return ""


def enrich_rule_based(agg: AggregatedCandidate) -> Tuple[str, str]:
    """Returns (full_name, company_name) using headers + signature only."""
    full_name = clean_header_name(agg.header_name)
    if not full_name:
        full_name = extract_name_from_signature(agg.snippet)

    company = extract_company_from_signature(agg.snippet, exclude_line=full_name)
    if not company:
        company = agg.guessed_company

    return full_name, company


# ==========================================================================
# File-format parsers -> yield (headers: dict, body_text: str, source_file: str)
# ==========================================================================

def iter_eml_files(root: Path) -> Iterator[Tuple[dict, str, str]]:
    for path in root.rglob("*.eml"):
        try:
            with open(path, "rb") as f:
                msg = BytesParser(policy=policy.default).parse(f)
            headers = {
                "From": msg.get("From", ""),
                "To": msg.get("To", ""),
                "Cc": msg.get("Cc", ""),
                "List-Unsubscribe": msg.get("List-Unsubscribe", ""),
                "List-Id": msg.get("List-Id", ""),
                "Precedence": msg.get("Precedence", ""),
                "Auto-Submitted": msg.get("Auto-Submitted", ""),
                "Message-ID": msg.get("Message-ID", ""),
            }
            body = get_body_text(msg)
            yield headers, body, str(path)
        except Exception as e:
            log.warning("Skipping corrupted .eml file %s (%s)", path, e)
            continue


def iter_msg_files(root: Path) -> Iterator[Tuple[dict, str, str]]:
    if not HAS_EXTRACT_MSG:
        log.error("extract_msg is not installed - cannot parse .msg files. "
                   "Run: pip install extract-msg")
        return
    for path in root.rglob("*.msg"):
        m = None
        try:
            m = extract_msg.Message(str(path))
            sender = getattr(m, "sender", "") or ""
            to = getattr(m, "to", "") or ""
            cc = getattr(m, "cc", "") or ""
            headers = {"From": sender, "To": to, "Cc": cc}
            body = getattr(m, "body", "") or ""
            yield headers, body, str(path)
        except Exception as e:
            log.warning("Skipping corrupted .msg file %s (%s)", path, e)
            continue
        finally:
            if m is not None:
                try:
                    m.close()
                except Exception:
                    pass


def iter_mbox_messages(root: Path) -> Iterator[Tuple[dict, str, str]]:
    seen_message_ids: Set[str] = set()
    mbox_paths = list(root.rglob("*.mbox"))
    if not mbox_paths:
        log.warning("No .mbox files found under %s", root)
    for mbox_path in mbox_paths:
        try:
            box = mailbox.mbox(
                str(mbox_path),
                factory=lambda f: BytesParser(policy=policy.default).parse(f),
            )
        except Exception as e:
            log.warning("Could not open mbox file %s (%s)", mbox_path, e)
            continue

        for key in box.keys():
            try:
                msg = box.get_message(key)
            except Exception as e:
                log.warning("Skipping corrupted message in %s (%s)", mbox_path, e)
                continue
            try:
                msg_id = msg.get("Message-ID", "")
                if msg_id:
                    if msg_id in seen_message_ids:
                        continue
                    seen_message_ids.add(msg_id)

                headers = {
                    "From": msg.get("From", ""),
                    "To": msg.get("To", ""),
                    "Cc": msg.get("Cc", ""),
                    "List-Unsubscribe": msg.get("List-Unsubscribe", ""),
                    "List-Id": msg.get("List-Id", ""),
                    "Precedence": msg.get("Precedence", ""),
                    "Auto-Submitted": msg.get("Auto-Submitted", ""),
                    "Message-ID": msg_id,
                }
                body = get_body_text(msg)
                yield headers, body, f"{mbox_path}#{key}"
            except Exception as e:
                log.warning("Skipping unparseable message in %s (%s)", mbox_path, e)
                continue
        try:
            box.close()
        except Exception:
            pass


def check_for_pst_files(root: Path) -> None:
    pst_files = list(root.rglob("*.pst"))
    if pst_files:
        log.error(
            "%d .pst file(s) found, but this script does not read .pst directly. "
            "Convert first, e.g.:\n"
            "    readpst -r -o ./converted_eml <file>.pst\n"
            "then re-run with --mode outlook --input-dir ./converted_eml",
            len(pst_files),
        )


# ==========================================================================
# Candidate extraction / aggregation
# ==========================================================================

def messages_to_candidates(
    messages: Iterable[Tuple[dict, str, str]],
    is_internal, is_automated, self_emails: Set[str],
    max_snippet_chars: int,
) -> Tuple[List[RawCandidate], dict]:
    stats = {"messages_scanned": 0, "addresses_seen": 0,
              "filtered_internal": 0, "filtered_automated": 0, "filtered_self": 0}
    raw_candidates: List[RawCandidate] = []

    for headers, body, source_file in messages:
        stats["messages_scanned"] += 1
        from_pairs = parse_address_list(headers.get("From", "")) or parse_loose_address_list(headers.get("From", ""))
        to_pairs = parse_address_list(headers.get("To", "")) or parse_loose_address_list(headers.get("To", ""))
        cc_pairs = parse_address_list(headers.get("Cc", "")) or parse_loose_address_list(headers.get("Cc", ""))

        snippet = ""
        if from_pairs:
            snippet = extract_signature_snippet(body, max_chars=max_snippet_chars)

        for direction, pairs in (("from", from_pairs), ("to", to_pairs), ("cc", cc_pairs)):
            for name, addr in pairs:
                stats["addresses_seen"] += 1
                addr = addr.lower().strip()
                if addr in self_emails:
                    stats["filtered_self"] += 1
                    continue
                if is_internal(addr):
                    stats["filtered_internal"] += 1
                    continue
                if is_automated(addr, headers if direction == "from" else None):
                    stats["filtered_automated"] += 1
                    continue
                domain = addr.split("@", 1)[1] if "@" in addr else ""
                raw_candidates.append(RawCandidate(
                    email=addr,
                    domain=domain,
                    header_name=name.strip(),
                    snippet=snippet if direction == "from" else "",
                    source_file=source_file,
                    direction=direction,
                ))

    return raw_candidates, stats


def aggregate_candidates(raw_candidates: List[RawCandidate]) -> List[AggregatedCandidate]:
    by_email: Dict[str, AggregatedCandidate] = {}
    next_id = 0
    for rc in raw_candidates:
        agg = by_email.get(rc.email)
        if agg is None:
            next_id += 1
            agg = AggregatedCandidate(id=next_id, email=rc.email, domain=rc.domain)
            agg.guessed_company = guess_company_from_domain(rc.domain)
            by_email[rc.email] = agg
        agg.occurrences += 1
        if rc.header_name and (not agg.header_name or len(rc.header_name) > len(agg.header_name)):
            agg.header_name = rc.header_name
        if rc.snippet and len(rc.snippet) > len(agg.snippet):
            agg.snippet = rc.snippet
    return list(by_email.values())


# ==========================================================================
# Merge + export
# ==========================================================================

def build_final_contacts(
    aggregated: List[AggregatedCandidate], exclude_public_domains: bool,
) -> List[FinalContact]:
    by_email: Dict[str, FinalContact] = {}

    for agg in aggregated:
        if exclude_public_domains and agg.domain in PUBLIC_EMAIL_DOMAINS:
            continue

        full_name, company_name = enrich_rule_based(agg)
        norm_email = agg.email.lower().strip()

        existing = by_email.get(norm_email)
        if existing is None:
            by_email[norm_email] = FinalContact(
                full_name=full_name, email_address=norm_email,
                company_name=company_name, signature_snippet=agg.snippet,
            )
        else:
            if len(full_name) > len(existing.full_name):
                existing.full_name = full_name
            if len(company_name) > len(existing.company_name):
                existing.company_name = company_name
            if len(agg.snippet) > len(existing.signature_snippet):
                existing.signature_snippet = agg.snippet

    return sorted(by_email.values(), key=lambda c: (c.company_name.lower(), c.full_name.lower()))


def export_contacts(contacts: List[FinalContact], csv_path: Path, xlsx_path: Path, include_evidence: bool) -> None:
    rows = []
    for c in contacts:
        row = {"Full Name": c.full_name, "Email Address": c.email_address, "Company Name": c.company_name}
        if include_evidence:
            row["Signature Snippet"] = c.signature_snippet
        rows.append(row)
    df = pd.DataFrame(rows)

    csv_path.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(csv_path, index=False, quoting=csv.QUOTE_MINIMAL, encoding="utf-8")
    try:
        df.to_excel(xlsx_path, index=False, engine="openpyxl")
    except Exception as e:
        log.error("Could not write .xlsx (%s). CSV was still written. "
                   "Install openpyxl: pip install openpyxl", e)


# ==========================================================================
# Cache (skip re-parsing a huge archive on rerun)
# ==========================================================================

def save_cache(path: Path, aggregated: List[AggregatedCandidate]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump([asdict(a) for a in aggregated], f, ensure_ascii=False, indent=2)
    log.info("Cached %d candidates to %s", len(aggregated), path)


def load_cache(path: Path) -> List[AggregatedCandidate]:
    with open(path, "r", encoding="utf-8") as f:
        raw = json.load(f)
    return [AggregatedCandidate(**item) for item in raw]


# ==========================================================================
# CLI
# ==========================================================================

def parse_args(argv=None):
    p = argparse.ArgumentParser(
        description="Extract unique external business contacts from Outlook/Gmail archives - "
                    "pure Python, no LLM, no API cost.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    p.add_argument("--mode", choices=["outlook", "gmail"], required=True,
                   help="outlook: read .eml/.msg files. gmail: read .mbox files (Google Takeout).")
    p.add_argument("--input-dir", required=True, type=Path,
                   help="Directory to scan recursively for source files.")
    p.add_argument("--internal-domains", default="",
                   help="Comma-separated list of your own past company domains to exclude, "
                        "e.g. 'acme.com,acme.co.uk'.")
    p.add_argument("--internal-domains-regex", default="",
                   help="Optional extra regex (applied to the full address) for matching "
                        "internal addresses that don't fit a simple domain list.")
    p.add_argument("--self-emails", default="",
                   help="Comma-separated list of your own email addresses to exclude.")
    p.add_argument("--automated-patterns", default="",
                   help="Comma-separated list of EXTRA regex patterns (matched against the "
                        "local part before the @) to treat as automated senders.")
    p.add_argument("--exclude-public-domains", action="store_true",
                   help="Drop contacts on personal webmail domains (gmail.com, yahoo.com, "
                        "etc). Off by default, since many real business contacts (freelancers, "
                        "small vendors) legitimately use personal email.")
    p.add_argument("--output-csv", default="contacts.csv", type=Path)
    p.add_argument("--output-xlsx", default="contacts.xlsx", type=Path)
    p.add_argument("--include-evidence", action="store_true",
                   help="Add a 'Signature Snippet' column to the export, so you can quickly "
                        "eyeball what the name/company guess was based on.")
    p.add_argument("--cache-file", default=None, type=Path,
                   help="Where to save/load the parsed-candidate cache (JSON), so you can "
                        "re-run enrichment/export without re-parsing a huge archive.")
    p.add_argument("--use-cache", action="store_true",
                   help="Load candidates from --cache-file instead of re-parsing --input-dir.")
    p.add_argument("--max-snippet-chars", type=int, default=1500,
                   help="Max characters of signature snippet scanned per sender.")
    p.add_argument("--verbose", action="store_true")
    p.add_argument("--log-file", default="extraction.log")

    args = p.parse_args(argv)

    internal_domains = {d.strip().lower() for d in args.internal_domains.split(",") if d.strip()}
    internal_regexes = [re.compile(args.internal_domains_regex, re.I)] if args.internal_domains_regex else []
    self_emails = {e.strip().lower() for e in args.self_emails.split(",") if e.strip()}

    automated_patterns = [re.compile(pat, re.I) for pat in DEFAULT_AUTOMATED_LOCAL_PARTS]
    extra_patterns = [p_.strip() for p_ in args.automated_patterns.split(",") if p_.strip()]
    automated_patterns += [re.compile(pat, re.I) for pat in extra_patterns]

    cfg = Config(
        mode=args.mode,
        input_dir=args.input_dir,
        internal_domains=internal_domains,
        internal_domain_regexes=internal_regexes,
        self_emails=self_emails,
        automated_patterns=automated_patterns,
        output_csv=args.output_csv,
        output_xlsx=args.output_xlsx,
        cache_file=args.cache_file,
        use_cache=args.use_cache,
        exclude_public_domains=args.exclude_public_domains,
        max_snippet_chars=args.max_snippet_chars,
        include_evidence=args.include_evidence,
    )
    return cfg, args.verbose, args.log_file


# ==========================================================================
# Main
# ==========================================================================

def main(argv=None) -> int:
    cfg, verbose, log_file = parse_args(argv)
    global log
    log = setup_logging(verbose, log_file)

    if cfg.use_cache:
        if not cfg.cache_file or not cfg.cache_file.exists():
            log.error("--use-cache was passed but --cache-file does not exist.")
            return 1
        log.info("Loading candidates from cache %s", cfg.cache_file)
        aggregated = load_cache(cfg.cache_file)
    else:
        if not cfg.input_dir.exists():
            log.error("Input directory does not exist: %s", cfg.input_dir)
            return 1

        is_internal = build_internal_matcher(cfg.internal_domains, cfg.internal_domain_regexes)
        is_automated = build_automated_matcher(cfg.automated_patterns)

        if cfg.mode == "outlook":
            check_for_pst_files(cfg.input_dir)
            messages = list(iter_eml_files(cfg.input_dir)) + list(iter_msg_files(cfg.input_dir))
        else:
            messages = list(iter_mbox_messages(cfg.input_dir))

        log.info("Parsed %d messages from %s", len(messages), cfg.input_dir)

        raw_candidates, stats = messages_to_candidates(
            messages, is_internal, is_automated, cfg.self_emails, cfg.max_snippet_chars,
        )
        log.info(
            "Messages: %d | addresses seen: %d | filtered internal: %d | "
            "filtered automated: %d | filtered self: %d",
            stats["messages_scanned"], stats["addresses_seen"],
            stats["filtered_internal"], stats["filtered_automated"], stats["filtered_self"],
        )

        aggregated = aggregate_candidates(raw_candidates)
        log.info("Aggregated to %d unique external candidate addresses", len(aggregated))

        if cfg.cache_file:
            save_cache(cfg.cache_file, aggregated)

    final_contacts = build_final_contacts(aggregated, cfg.exclude_public_domains)
    log.info("Final unique contact count: %d", len(final_contacts))

    export_contacts(final_contacts, cfg.output_csv, cfg.output_xlsx, cfg.include_evidence)
    log.info("Wrote %s and %s", cfg.output_csv, cfg.output_xlsx)
    return 0


if __name__ == "__main__":
    sys.exit(main())
