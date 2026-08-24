# Email Contact Extractor - No-LLM Edition

Same goal as `email_contact_extractor.py` - a unique, deduplicated list of
external business contacts from Outlook or Gmail archives - but the entire
pipeline is pure Python. No API key, no network calls, no per-run cost.

## What's different from the LLM version

| | LLM version | This version |
|---|---|---|
| Cost | Anthropic API usage (small, since only unique contacts + short snippets are sent) | $0 |
| Name/company source | Header, then Claude cleans it up using the signature | Header (trusted as-is), falling back to a regex-matched name-shaped line in the signature when there's no header name at all |
| Company detection | Claude reads the signature in context | Regex match for a company-suffix line (Inc/LLC/Ltd/GmbH/Group/...) in the signature; falls back to a domain-name guess |
| "Is this actually a business contact?" | Claude judges and flags it | No judgment call is made - nothing is dropped on that basis. Use `--exclude-public-domains` for a blunter, rule-based version of the same idea |
| Dependencies | pandas, openpyxl, pydantic, anthropic, (extract-msg) | pandas, openpyxl, (extract-msg) |

## Install

```bash
pip install -r requirements_no_llm.txt
```

That's the whole setup - no API key needed.

## Usage

```bash
# Gmail (Google Takeout .mbox export)
python extract_contacts_no_llm.py \
    --mode gmail \
    --input-dir ./takeout_export \
    --internal-domains "mycompany.com,mycompany.co.uk" \
    --self-emails "me@mycompany.com,myname@gmail.com"

# Outlook (.eml or .msg export)
python extract_contacts_no_llm.py \
    --mode outlook \
    --input-dir ./outlook_export \
    --internal-domains "mycompany.com"
```

For `.pst` files, convert to `.eml` first with `readpst` - see the main
README's "Handling .pst files" section, it applies identically here.

### Recommended: start with `--include-evidence`

Add `--include-evidence` on your first run - it appends a "Signature
Snippet" column so you can see exactly what text each name/company guess
came from, which is the fastest way to judge how well the heuristics are
doing on your particular archive before you trust the output as-is:

```bash
python extract_contacts_no_llm.py --mode gmail --input-dir ./takeout_export \
    --internal-domains "mycompany.com" --include-evidence
```

### Flags

| Flag | Purpose |
|---|---|
| `--internal-domains` | Comma-separated list of your own company's domains to exclude |
| `--internal-domains-regex` | Extra regex for internal addresses that don't fit a simple domain list |
| `--self-emails` | Your own address(es), including personal ones |
| `--automated-patterns` | Extra regex patterns on top of the built-in automated-sender list |
| `--exclude-public-domains` | Drop contacts on gmail.com/yahoo.com/etc entirely |
| `--include-evidence` | Add the signature snippet each guess was based on, for manual QA |
| `--cache-file` / `--use-cache` | Save/reload the parsed candidate list as JSON, so a large archive isn't re-parsed on every rerun |

Run `python extract_contacts_no_llm.py --help` for the full list.

## Accuracy: what to expect, honestly

This will get straightforward cases right - a normal `Name <email>` header,
or a signature block that ends in "Best regards, / Full Name / Title /
Company Inc." But it has no semantic understanding, so:

- A header display name is always trusted as typed (just re-cased if it was
  ALL CAPS or all lowercase) - if someone's mail client shows "Sales Team"
  or "iPhone User" as their name, that's what you'll get, verbatim.
- Company names only get pulled from a signature line if it contains a
  recognizable suffix (Inc, LLC, Ltd, GmbH, Group, Solutions, ...) - a
  signature that just says "Acme" with no suffix will fall back to the
  domain-based guess instead.
- Nothing is dropped as "not a real business contact" - a personal friend
  on Gmail who happens to survive the internal/automated filters will
  appear in the output. Skim the CSV, or use `--exclude-public-domains` for
  a stricter (but blunter) cut.

`--include-evidence` is there specifically so you can spot-check these
cases quickly rather than trusting the output blind. If accuracy matters
more than the API cost for your use case, `email_contact_extractor.py` (the
LLM version) will do meaningfully better on messy, inconsistent signatures.

## Output

`contacts.csv` and `contacts.xlsx` with columns `Full Name`,
`Email Address`, `Company Name` (plus `Signature Snippet` if
`--include-evidence` is set), sorted by company then name.
