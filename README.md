# Email Contact Extractor
A unique, deduplicated list of
external business contacts from Outlook or Gmail archives and the entire
pipeline is pure Python.

## Install

```bash
pip install -r requirements.txt
```

That's the whole setup.

## Usage

```bash
# Gmail (Google Takeout .mbox export)
python extract_contacts.py \
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
python extract_contacts.py --mode gmail --input-dir ./takeout_export \
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

Run `python extract_contacts.py --help` for the full list.

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
cases quickly rather than trusting the output blind.
## Output

`contacts.csv` and `contacts.xlsx` with columns `Full Name`,
`Email Address`, `Company Name` (plus `Signature Snippet` if
`--include-evidence` is set), sorted by company then name.
