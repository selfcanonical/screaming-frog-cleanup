# Screaming Frog Cleanup: a Claude skill

Turn a messy Screaming Frog export into a client-ready SEO audit in about a minute.

Screaming Frog's **Bulk Export > Issues > All** gives you 100+ CSV files and hundreds of thousands of rows. Most of it is noise for a client report. This skill sorts it for you and gives back three files:

| File | What it is |
| --- | --- |
| `<client>_executive_summary.pdf` | A plain-language summary for a marketing manager: the short version, a priority table, the pattern behind each issue, quick wins, and a 30/60/90-day plan. |
| `<client>_seo_audit_clean.xlsx` | The issues that matter, ranked Critical, Important or Minor, one tab per issue, with links to the detail tabs. |
| `sf_raw_merged.xlsx` | Every CSV from the export as its own tab, with an index. Opens in Excel or Google Sheets for your own manual work. |

## What it cleans

- **Covers all 351 Screaming Frog checks** (v23). Each one is either kept with a priority, an area, an effort rating and a "why it matters" line, or moved to the appendix.
- **Drops the noise:** security headers, social and X card tags, H2 checks, readability, image titles, URL formatting, and Accessibility checks go to the appendix.
- **Merges duplicate checks**, such as title length in characters vs. pixels.
- **Uses the inlinks files instead of dumping them.** Each broken or redirected URL gets an "Internal Links Pointing Here" count and an example linking page, so the biggest problems sort to the top.
- **Filters content checks** to indexable HTML pages, so PDFs and noindexed pages don't inflate counts.
- **Flags new checks.** If a future Screaming Frog version adds a check the skill doesn't know, it says so instead of guessing.

## Install

**Claude app (claude.ai or desktop)**
1. Turn on code execution: **Settings > Capabilities**.
2. Download `screaming-frog-cleanup.zip` from the [Releases](../../releases) page (or zip the `screaming-frog-cleanup` folder yourself).
3. Go to **Customize > Skills**, click **+**, and upload the zip.

**Claude Code**
```bash
git clone https://github.com/<your-username>/screaming-frog-cleanup.git
cp -r screaming-frog-cleanup/screaming-frog-cleanup ~/.claude/skills/
```

## Use

1. In Screaming Frog, run your crawl, then **Bulk Export > Issues > All**. Zip the folder it creates.
2. Upload the zip to Claude and say: *"Clean up this Screaming Frog audit for [Client Name]."*

It also accepts a folder of the CSVs, or a single merged `.xlsx` (for example from [selfcanonical.com](https://selfcanonical.com)).

## Run the scripts without Claude

The workbooks come from a plain Python script, so you can run it yourself:

```bash
pip install pandas openpyxl reportlab
python3 screaming-frog-cleanup/scripts/sf_clean.py All_Issues.zip out "Client Name"
python3 screaming-frog-cleanup/scripts/sf_pdf.py summary.md summary.pdf
```

`sf_clean.py` writes the two workbooks plus `<client>_summary.json`. Claude uses that JSON to write the executive summary, and `sf_pdf.py` turns any Markdown summary into a styled PDF.

## Customize the priorities

All the judgment lives at the top of `scripts/sf_clean.py`:

- `RULES`: checks that are kept, with their tier (1 Critical, 2 Important, 3 Minor), area, effort and plain-English "why it matters" line.
- `DROP` and `DROP_PREFIXES`: checks that only go to the appendix.

Move a check between them to match how you audit. Pull requests with better rules are welcome.

## Repo layout

```
screaming-frog-cleanup/        <- the skill (this folder is what you install)
  SKILL.md                     <- instructions Claude follows
  scripts/
    sf_clean.py                <- builds the clean and raw merged workbooks
    sf_pdf.py                  <- turns the Markdown summary into a PDF
README.md
LICENSE
```

## License

MIT. See [LICENSE](LICENSE).

Not affiliated with Screaming Frog Ltd. "Screaming Frog" is a trademark of its owner.
