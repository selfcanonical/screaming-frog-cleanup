---
name: screaming-frog-cleanup
description: Clean a Screaming Frog Bulk Export > Issues > All export (zip of CSVs or merged xlsx) into a clean audit workbook, a raw merged workbook, and a plain-language executive SEO audit summary PDF.
---

# Screaming Frog Audit Cleanup

Turns a noisy Screaming Frog "Bulk Export > Issues > All" export into:

1. **A clean workbook** (`<client>_seo_audit_clean.xlsx`): Executive Summary tab (color-coded priority, linked to detail tabs), All Issues tab (every SF issue with SF vs clean counts and SF fix notes), and one trimmed tab per kept issue.
2. **A raw merged workbook** (`sf_raw_merged.xlsx`): every CSV from the export as its own tab, untouched, with real headers in row 1 and a linked Index tab. Made for manual work in Excel or Google Sheets. Only created when the input is a zip or folder of CSVs.
3. **An executive audit summary PDF** (`<client>_executive_summary.pdf`) written for a client or stakeholder.

This skill ships two scripts in its `scripts/` folder:
- `scripts/sf_clean.py` builds the two workbooks and a JSON summary (needs `pandas` and `openpyxl`).
- `scripts/sf_pdf.py` turns the Markdown summary into the PDF (needs `reportlab`).

## Inputs accepted
- A `.zip` of the SF CSVs (the raw export, e.g. `All_Issues.zip`)
- A folder of the CSVs
- A merged `.xlsx` where row 1 of each tab is `File Name: x.csv` (for example, from https://selfcanonical.com)

The export must include `issues_overview_report.csv`. If it is missing, ask the user to re-export with Bulk Export > Issues > All. The set of CSVs changes with each site, since SF only exports checks that found something. The scripts handle any subset.

## Steps

1. Find the uploaded file. Copy it into its own empty folder (e.g. `~/sf_in/`). Ask for the client name only if it is not clear from the request; otherwise the script uses the domain.
2. Locate this skill's `scripts/` folder (`<skill dir>/scripts`). Check the libraries: `python3 -c "import pandas, openpyxl, reportlab"`. Install anything missing with `pip install pandas openpyxl reportlab` (add `--break-system-packages` if pip asks). If installs are blocked and `reportlab` is missing, say so and deliver the summary as a .docx or Markdown instead.
3. Run: `cd /tmp && python3 -I <skill dir>/scripts/sf_clean.py <input> ~/sf_out "<Client Name>"`. Small exports take seconds. Large ones (hundreds of thousands of rows) can take about a minute because of the raw merged workbook, so allow a long enough timeout.
4. Read `~/sf_out/<client>_summary.json`. If `unmapped` is not empty, SF has added checks this skill does not know yet. Tell the user and suggest a rule for each one. Check `raw_merged.fits_google_sheets`: if it is false, the file is over Google Sheets' 10 million cell limit, so tell the user it will only open in Excel, or offer to leave out the largest `_inlinks` tabs.
5. Spot-check two or three detail tabs. Confirm the Critical items are real and not site-specific false positives (for example, noindex pages that are meant to be noindex). Look at example URLs for patterns (one template, one folder, PDFs, a subdomain), because patterns make the summary useful.
6. Write the executive summary (format below) as Markdown to `~/sf_out/<client>_executive_summary.md`. Use only numbers from the JSON and the detail tabs. Never invent figures, traffic estimates, or a health score. Check every pattern you name (URL formats, languages, folders) against the data before writing it.
7. Build the PDF: `cd /tmp && python3 -I <skill dir>/scripts/sf_pdf.py ~/sf_out/<client>_executive_summary.md ~/sf_out/<client>_executive_summary.pdf`. Then look at it: render page 1 to an image (for example `pdftoppm -r 60 -f 1 -l 1 -png <pdf> /tmp/prev`) and view it. Fix anything that reads badly (a cramped table, a wrapped word) by shortening the Markdown and rebuilding.
8. Deliver the summary PDF, the clean workbook, and the raw merged workbook. In two or three lines, say what was kept, what was cut, and any judgment calls.

## Executive summary format

Writing rules: plain language at about an 8th-grade reading level, no em dashes, short and structured. Write for a marketing manager, not a developer. Name the pattern behind an issue ("most broken links come from one old landing page template"), not just the count.

1. **Title**: `# <Client> Technical SEO Audit: Executive Summary`, then a blank line, then one italic subtitle line such as `_Prepared <date> from a Screaming Frog crawl (Bulk Export > Issues > All)_`. The PDF shows it under the title.
2. **The short version** (3 to 5 bullets): the biggest risks and the single most valuable fix.
3. **Priority issues** table: Priority, Issue (plain-English name), Pages Affected, Why It Matters, Effort. All Critical items plus the Important items that matter most. About 12 rows max.
4. **Issue details by area**: for each issue in the table, what is wrong, the pattern found, 1 to 3 example URLs, and the fix in one or two sentences.
5. **Quick wins**: small fixes that clear many URLs.
6. **Recommended next steps**: a 30/60/90-day order.
7. **Appendix**: one line saying the full data is in the clean workbook, plus the checks left out on purpose and why (low value, or a duplicate of another check).

Markdown the PDF script supports: `#`, `##`, `###` headings, paragraphs, `-` bullets, `1.` numbered lists, pipe tables, `**bold**`, `*italic*`, `` `code` ``, `[links](url)`. A table whose first column holds Critical, Important or Minor gets that cell color-coded. Keep tables to 5 columns or fewer and cells short.

## Tuning the rules

All judgment lives at the top of `scripts/sf_clean.py`. It covers all 351 checks in Screaming Frog v23:
- `RULES[name] = K(tier, area, effort, why, indexable_only)`. Tier 1 is Critical, 2 is Important, 3 is Minor. `indexable_only=True` drops non-indexable and non-HTML URLs (such as PDFs) from that issue.
- `DROP` and `DROP_PREFIXES` send checks to the appendix only. By default that covers Accessibility, Social, H2, URL formatting, security headers, readability, and duplicate length checks.
- Names must match the `Issue Name` column in `issues_overview_report.csv` exactly.
- Any check not covered falls back on SF priority (High becomes Important, Medium becomes Minor, Low goes to the appendix). It is flagged "(no rule)" in the All Issues tab and listed under `unmapped`.

If the user says something like "always drop X", "make Y critical", or "never show Z", apply it for this run by editing a copy of the script. Then offer to update the skill so the change sticks.
