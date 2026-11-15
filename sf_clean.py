#!/usr/bin/env python3
"""Clean a Screaming Frog "Bulk Export > Issues > All" export into:
  1. <out>/<name>_seo_audit_clean.xlsx - Executive Summary, All Issues, one tab per kept issue
  2. <out>/<name>_summary.json         - data for writing the executive report
  3. <out>/sf_raw_merged.xlsx          - every CSV as its own tab (zip or folder input only)
Input: .zip of the CSVs, a folder of CSVs, or a merged .xlsx (selfcsv format).
Usage: python3 -I sf_clean.py <export> <out_dir> [client name]
"""
import sys, re, json, os
from collections import Counter
from urllib.parse import urlparse
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------- RULES
# Every check Screaming Frog can report (351 as of v23) is either in RULES (kept)
# or caught by DROP / DROP_PREFIXES (appendix only). Anything else is "unmapped":
# it falls back on SF priority and is listed in the summary so a rule can be added.
# K(tier, area, effort, why, indexable_only=False). Tier 1 Critical, 2 Important, 3 Minor.
IDX, BRK, INTL, HTML, ONP, CQ, SPD, LNK, JS, SD, MOB, SM, GSC, AMP = (
    "Indexing & Crawling", "Broken Links & Redirects", "International (Hreflang)", "HTML Structure",
    "On-Page SEO", "Content Quality", "Speed & Images", "Internal Links", "JavaScript Rendering",
    "Structured Data", "Mobile", "XML Sitemaps", "Search Console & Analytics", "AMP")
def K(tier, area, effort, why, idx=False):
    return ("keep", tier, area, effort, idx, why)
RULES = {
 # Response codes
 "Response Codes: Internal Client Error (4xx)": K(1,BRK,"Low","Broken pages that users and Google reach through the site's own links."),
 "Response Codes: Internal Server Error (5xx)": K(1,BRK,"Medium","Pages that fail with a server error. Users and Google see an error instead of the page."),
 "Response Codes: Internal No Response": K(1,BRK,"Medium","Internal URLs that did not respond at all."),
 "Response Codes: Internal Redirect Loop": K(1,BRK,"Low","Redirects that loop forever, so the page never loads."),
 "Response Codes: Internal Redirect Chain": K(2,BRK,"Low","Links pass through several redirects before reaching the page. This slows users and wastes crawl budget."),
 "Response Codes: Internal Redirection (3xx)": K(2,BRK,"Low","Internal links point to URLs that redirect. Update the links to the final URL."),
 "Response Codes: Internal Redirection (JavaScript)": K(3,BRK,"Low","Pages redirect with JavaScript, which Google handles less reliably than a server redirect."),
 "Response Codes: Internal Redirection (Meta Refresh)": K(3,BRK,"Low","Pages redirect with a meta refresh instead of a proper server redirect."),
 "Response Codes: Internal Redirection (HTTP Refresh)": K(3,BRK,"Low","Pages redirect with an HTTP refresh instead of a proper server redirect."),
 "Response Codes: Internal Blocked by Robots.txt": K(2,IDX,"Low","Internal URLs blocked by robots.txt. Google cannot crawl them even though the site links to them."),
 "Response Codes: Internal Blocked Resource": K(2,JS,"Low","Files Google needs to render the page (CSS, JS, images) are blocked."),
 "Response Codes: External Client Error (4xx)": K(3,BRK,"Low","Links to other sites that are broken. Hurts user trust."),
 "Security: HTTP URLs": K(1,BRK,"Low","Insecure HTTP pages are still linked internally."),
 "Security: Mixed Content": K(1,BRK,"Low","HTTPS pages load HTTP files, which browsers may block or flag as not secure."),
 "Security: Form URL Insecure": K(1,BRK,"Low","Forms send data over insecure HTTP."),
 "Security: Form on HTTP URL": K(1,BRK,"Low","Forms sit on insecure HTTP pages, so anything typed in is not protected."),
 # Indexing
 "Directives: Noindex": K(1,IDX,"Low","These pages tell Google not to index them. Confirm none of them should rank."),
 "Directives: None": K(1,IDX,"Low","These pages tell Google not to index them or follow their links."),
 "Directives: Outside <head>": K(1,IDX,"Low","Robots tags sit outside the head, so Google may ignore them."),
 "Directives: Nofollow": K(2,IDX,"Low","These pages tell Google not to follow their links, which stops link value from flowing."),
 "Directives: Unavailable_After": K(3,IDX,"Low","Pages are set to drop out of Google after a date. Check the date is right."),
 "Directives: NoSnippet": K(3,IDX,"Low","Google is told not to show a text snippet for these pages."),
 "Directives: NoImageIndex": K(3,IDX,"Low","Images on these pages are blocked from Google Images."),
 "Canonicals: Non-Indexable Canonical": K(1,IDX,"Low","The canonical points to a page Google cannot index, so neither page may rank."),
 "Canonicals: Multiple Conflicting": K(1,IDX,"Low","Pages have more than one canonical, each pointing somewhere different."),
 "Canonicals: Invalid Attribute In Annotation": K(1,IDX,"Low","Canonical tags have extra attributes, so Google ignores them."),
 "Canonicals: Contains Fragment URL": K(1,IDX,"Low","Canonicals include a # fragment, so Google ignores them."),
 "Canonicals: Outside <head>": K(1,IDX,"Low","Google ignores a canonical tag that sits outside the head."),
 "Canonicals: Canonicalised": K(2,IDX,"Medium","These pages point Google to a different URL as the main version. Fine if intended, a problem if they should rank."),
 "Canonicals: Missing": K(2,IDX,"Low","No canonical tag, so Google has to guess which URL is the main version.",True),
 "Canonicals: Multiple": K(3,IDX,"Low","Pages set the canonical more than once."),
 "Canonicals: Unlinked": K(3,IDX,"Low","Canonical URLs that the site never links to."),
 "Pagination: Non-Indexable": K(2,IDX,"Medium","Paginated pages are not indexable, so products or posts deeper in the list may be missed."),
 "Pagination: Pagination URL Not in Anchor Tag": K(2,IDX,"Medium","Next/previous pages are not linked with normal links, so Google may not reach them."),
 "Pagination: Non-200 Pagination URLs": K(2,IDX,"Low","Next/previous links point to broken or redirected pages."),
 "Pagination: Unlinked Pagination URLs": K(3,IDX,"Low","Paginated pages are not linked anywhere on the site."),
 "Pagination: Multiple Pagination URLs": K(3,IDX,"Low","Pages have more than one next/prev tag."),
 "Pagination: Pagination Loop": K(3,IDX,"Low","Pagination loops back on itself."),
 "Pagination: Sequence Error": K(3,IDX,"Low","Next/prev tags do not match up between pages."),
 "Links: Pages With High Crawl Depth": K(2,IDX,"Medium","These pages are many clicks from the homepage. Google crawls them less and they rank weaker.",True),
 "Links: Pages Without Internal Outlinks": K(3,IDX,"Low","Dead-end pages with no internal links out."),
 "URL: GA Tracking Parameters": K(3,IDX,"Low","Internal links carry tracking parameters, which creates duplicate URLs and skews analytics."),
 # Hreflang
 "Hreflang: Missing Return Links": K(1,INTL,"Medium","Language versions do not link back to each other, so Google may ignore the hreflang setup."),
 "Hreflang: Non-200 hreflang URLs": K(1,INTL,"Low","Hreflang points to pages that redirect or are broken."),
 "Hreflang: Noindex Return Links": K(1,INTL,"Low","Hreflang points to pages that are set to noindex."),
 "Hreflang: Inconsistent Language & Region Return Links": K(1,INTL,"Medium","Language pages disagree on which language or region each one is."),
 "Hreflang: Non-Canonical Return Links": K(1,INTL,"Low","Hreflang points to non-canonical versions of pages."),
 "Hreflang: Not Using Canonical": K(1,INTL,"Low","Pages use a non-canonical URL in their own hreflang."),
 "Hreflang: Incorrect Language & Region Codes": K(1,INTL,"Low","Hreflang uses invalid language or country codes, so Google ignores it."),
 "Hreflang: Outside <head>": K(1,INTL,"Low","Hreflang tags sit outside the head, so Google ignores them."),
 "Hreflang: Self Reference Only": K(2,INTL,"Medium","Pages list only themselves in hreflang, so other language versions are not connected."),
 "Hreflang: Multiple Entries": K(2,INTL,"Low","The same language code is used more than once."),
 "Hreflang: Mismatch with HTML Lang Attribute": K(2,INTL,"Low","Hreflang and the page's lang attribute disagree on the language."),
 "Hreflang: Missing Self Reference": K(2,INTL,"Low","Pages do not include themselves in their hreflang set."),
 "Hreflang: Missing X-Default": K(3,INTL,"Low","No default version for users whose language is not covered."),
 "Hreflang: Unlinked hreflang URLs": K(3,INTL,"Low","Hreflang pages that are not linked anywhere on the site."),
 # HTML structure
 "Validation: Invalid HTML Elements in <head>": K(1,HTML,"Medium","Invalid tags in the head make Google stop reading it early, so tags after them are missed."),
 "Validation: <head> Not First In <html> Element": K(1,HTML,"Medium","Something comes before the head, so Google may treat key tags as body content."),
 "Validation: <body> Element Preceding <html>": K(1,HTML,"Medium","Body content comes before the html tag, so Google may ignore the head."),
 "Validation: Missing <head> Tag": K(1,HTML,"Medium","Pages have no head section."),
 "Validation: Multiple <head> Tags": K(1,HTML,"Medium","Broken HTML can make Google ignore key tags like the title and canonical."),
 "Validation: Multiple <body> Tags": K(1,HTML,"Medium","Broken HTML can make Google misread the page."),
 "Validation: HTML Document Over 2MB": K(1,HTML,"Medium","Google only reads the first 2MB of a page, so content after that is lost."),
 "Validation: Missing <body> Tag": K(2,HTML,"Medium","Pages have no body tag."),
 "Validation: Resource Over 2MB": K(2,HTML,"Medium","CSS or JS files over 2MB may be cut off by Google."),
 "Page Titles: Outside <head>": K(1,HTML,"Low","Title tag is outside the head, so Google may not use it."),
 "Meta Description: Outside <head>": K(2,HTML,"Low","Meta description is outside the head, so Google may not use it."),
 "Page Titles: Multiple": K(2,HTML,"Low","Pages have more than one title tag."),
 "Meta Description: Multiple": K(3,HTML,"Low","Pages have more than one meta description."),
 # On-page
 "Page Titles: Missing": K(1,ONP,"Low","Pages with no title tag. Titles are one of the strongest ranking signals.",True),
 "Page Titles: Duplicate": K(2,ONP,"Medium","Several pages share the same title, so they compete with each other in search.",True),
 "H1: Missing": K(2,ONP,"Low","Pages have no main heading.",True),
 "Page Titles: Over 561 Pixels": K(3,ONP,"Medium","Titles too long for search results, so they get cut off.",True),
 "Page Titles: Below 30 Characters": K(3,ONP,"Low","Very short titles waste space to describe the page.",True),
 "H1: Duplicate": K(3,ONP,"Low","Several pages share the same H1.",True),
 "H1: Multiple": K(3,ONP,"Low","Pages have more than one H1.",True),
 "Meta Description: Missing": K(3,ONP,"Low","No meta description, so Google writes its own snippet.",True),
 "Meta Description: Duplicate": K(3,ONP,"Low","Several pages share the same meta description.",True),
 # Content quality
 "Content: Exact Duplicates": K(1,CQ,"Medium","Identical pages on different URLs split ranking signals.",True),
 "Content: Soft 404 Pages": K(1,CQ,"Low","Error pages that return a 200 OK, so Google may index them."),
 "Content: Lorem Ipsum Placeholder": K(1,CQ,"Low","Placeholder text left on live pages."),
 "Content: Near Duplicates": K(2,CQ,"High","Pages that are almost the same compete with each other and look low quality.",True),
 "Content: Low Content Pages": K(2,CQ,"High","Thin pages rarely rank and can drag down site quality.",True),
 "Content: Semantically Similar": K(3,CQ,"High","Pages cover the same topic and may compete for the same searches.",True),
 "Content: Spelling Errors": K(3,CQ,"Low","Spelling mistakes are a sign of low quality to users and Google's raters.",True),
 # Internal links
 "Links: Outlinks To Localhost": K(1,LNK,"Low","Links point to localhost, a development address that does not work for users."),
 "Links: Non-Indexable Page Inlinks Only": K(2,LNK,"Medium","Pages only linked from non-indexable pages, so Google may not find or value them."),
 "Links: Pages With Uncrawlable Internal Outlinks": K(2,LNK,"Medium","Links not built as normal <a href> links, so Google cannot follow them."),
 "Links: Internal Nofollow Inlinks Only": K(2,LNK,"Low","Pages only reached through nofollow links, so they may not be indexed."),
 "Links: Pages With Empty Anchor Tags": K(3,LNK,"Low","Links with no text give Google no clue about the target page."),
 "Links: Internal Nofollow Outlinks": K(3,LNK,"Low","Internal links marked nofollow block link value inside the site."),
 # Speed & images
 "Images: Over 100 kB": K(2,SPD,"Medium","Heavy images slow pages down and hurt Core Web Vitals."),
 "PageSpeed: Improve Image Delivery": K(2,SPD,"Medium","Images could load faster with compression or modern formats. Affects LCP."),
 "PageSpeed: Render Blocking Requests": K(2,SPD,"Medium","Files that stop the page from showing until they load."),
 "PageSpeed: LCP Request Discovery": K(2,SPD,"Low","The main image is found late or lazy-loaded, which slows LCP."),
 "PageSpeed: Reduce Unused JavaScript": K(3,SPD,"High","Pages load JavaScript they do not use."),
 "PageSpeed: Reduce Unused CSS": K(3,SPD,"Medium","Pages load CSS they do not use."),
 "PageSpeed: Reduce JavaScript Execution Time": K(3,SPD,"High","JavaScript takes over 2 seconds to run."),
 "PageSpeed: Minimize Main-Thread Work": K(3,SPD,"High","The browser is busy too long, so the page feels slow to respond."),
 "PageSpeed: Layout Shift Culprits": K(3,SPD,"Medium","Elements move around while loading (CLS)."),
 "PageSpeed: Document Request Latency": K(3,SPD,"Medium","Slow server response, redirects, or no compression."),
 "PageSpeed: Use Efficient Cache Lifetimes": K(3,SPD,"Low","Files are not cached, so repeat visits are slower."),
 "PageSpeed: Avoid Enormous Network Payloads": K(3,SPD,"Medium","Pages are very heavy to download."),
 "PageSpeed: Minify CSS": K(3,SPD,"Low","CSS files are larger than needed."),
 "PageSpeed: Minify JavaScript": K(3,SPD,"Low","JavaScript files are larger than needed."),
 "Images: Missing Size Attributes": K(3,SPD,"Low","Images without width/height make the layout jump while loading (CLS)."),
 "Images: Incorrectly Sized Images": K(3,SPD,"Low","Images are much bigger than their display size."),
 "Images: Missing Alt Text": K(3,SPD,"Low","Images without alt text are hard for Google and screen readers to understand."),
 "Images: Missing Alt Attribute": K(3,SPD,"Low","Images have no alt attribute at all."),
 # JavaScript
 "JavaScript: Canonical Mismatch": K(1,JS,"Medium","The canonical changes after JavaScript runs, which sends Google mixed signals."),
 "JavaScript: Noindex Only in Original HTML": K(1,JS,"Low","Noindex in the raw HTML stops Google before JavaScript can remove it."),
 "JavaScript: Nofollow Only in Original HTML": K(1,JS,"Low","Nofollow in the raw HTML stops Google from following links."),
 "JavaScript: Pages with Blocked Resources": K(2,JS,"Low","Google cannot load files it needs to render these pages."),
 "JavaScript: Page Title Only in Rendered HTML": K(3,JS,"Medium","Title only appears after JavaScript runs."),
 "JavaScript: H1 Only in Rendered HTML": K(3,JS,"Medium","H1 only appears after JavaScript runs."),
 "JavaScript: Meta Description Only in Rendered HTML": K(3,JS,"Medium","Meta description only appears after JavaScript runs."),
 "JavaScript: Canonical Only in Rendered HTML": K(3,JS,"Medium","Canonical only appears after JavaScript runs."),
 "JavaScript: Contains JavaScript Links": K(3,JS,"Medium","Some links only appear after JavaScript runs."),
 "JavaScript: Uses Old AJAX Crawling Scheme URLs": K(3,JS,"Medium","Uses an AJAX crawling scheme Google no longer supports."),
 "JavaScript: Uses Old AJAX Crawling Scheme Meta Fragment Tag": K(3,JS,"Low","Old AJAX meta fragment tag left on pages."),
 # Structured data
 "Structured Data: Parse Errors": K(2,SD,"Low","Schema code is broken, so Google cannot read it."),
 "Structured Data: Validation Errors": K(2,SD,"Low","Schema has errors against Schema.org."),
 "Structured Data: Rich Result Validation Errors": K(2,SD,"Low","Schema is missing required fields, so pages cannot get rich results."),
 "Structured Data: Validation Warnings": K(3,SD,"Low","Schema uses old or deprecated markup."),
 "Structured Data: Missing": K(3,SD,"Medium","Pages have no structured data. Missed chance for rich results.",True),
 # Mobile
 "Mobile: Viewport Not Set": K(1,MOB,"Low","No viewport tag, so mobile shows a shrunken desktop page."),
 "Mobile: Content Not Sized Correctly": K(2,MOB,"Medium","Content is wider than the screen on mobile."),
 "Mobile: Illegible Font Size": K(2,MOB,"Low","Text is too small to read on mobile."),
 "Mobile: Contains Unsupported Plugins": K(2,MOB,"High","Uses Flash or similar plugins that mobile cannot run."),
 "Mobile: Target Size": K(3,MOB,"Low","Buttons and links are too small or too close to tap."),
 # Sitemaps
 "Sitemaps: XML Sitemap with over 50k URLs": K(1,SM,"Low","Sitemap is over the 50,000 URL limit."),
 "Sitemaps: XML Sitemap over 50MB": K(1,SM,"Low","Sitemap is over the 50MB limit."),
 "Sitemaps: Non-Indexable URLs in Sitemap": K(2,SM,"Low","The sitemap lists pages Google cannot index."),
 "Sitemaps: Orphan URLs": K(2,SM,"Medium","Pages in the sitemap that the site never links to."),
 "Sitemaps: URLs not in Sitemap": K(3,SM,"Low","Indexable pages missing from the sitemap."),
 # Search Console / Analytics (only when connected)
 "Search Console: Indexable URL Not Indexed": K(1,GSC,"Medium","Pages that can be indexed but Google has not indexed them."),
 "Search Console: User-Declared Canonical Not Selected": K(2,GSC,"Medium","Google picked a different main URL than the canonical says."),
 "Search Console: Orphan URLs": K(2,GSC,"Medium","Pages getting search impressions that the site does not link to."),
 "Search Console: Rich Result Invalid": K(2,GSC,"Low","Rich result errors in Search Console."),
 "Search Console: Non-Indexable with Search Analytics Data": K(3,GSC,"Low","Non-indexable pages still getting impressions. Check if Google ignores the canonical."),
 "Search Console: URL is on Google But Has Issues": K(3,GSC,"Low","Indexed pages with rich result or AMP issues."),
 "Analytics: Orphan URLs": K(2,GSC,"Medium","Pages getting traffic that the site does not link to."),
 "Analytics: Non-Indexable with GA Data": K(3,GSC,"Low","Non-indexable pages still getting traffic."),
 # AMP
 "AMP: Non-200 Response": K(2,AMP,"Low","AMP pages are broken or redirect."),
 "AMP: Non-Indexable Canonical": K(2,AMP,"Low","AMP canonical points to a page Google cannot index."),
 "AMP: Missing Canonical to Non-AMP": K(2,AMP,"Low","AMP canonical points to another AMP page."),
 "AMP: Missing Non-AMP Return Link": K(2,AMP,"Low","Regular pages do not link to their AMP version."),
 "AMP: Missing Canonical": K(2,AMP,"Low","AMP pages have no canonical."),
 "AMP: Indexable": K(3,AMP,"Low","AMP pages are indexable. Check they are standalone pages."),
}
# Appendix only: duplicates of kept checks, or low SEO value
DROP = {
 "Page Titles: Over 60 Characters","Page Titles: Below 200 Pixels","Page Titles: Same as H1",
 "Meta Description: Over 155 Characters","Meta Description: Over 985 Pixels","Meta Description: Below 70 Characters",
 "Meta Description: Below 400 Pixels","H1: Over 70 Characters","H1: Non-Sequential","H1: Alt Text in H1",
 "Images: Missing Image Title","Images: Alt Text Over 100 Characters","Images: Background Images",
 "Links: Internal Outlinks With No Anchor Text","Links: Non-Descriptive Anchor Text In Internal Outlinks",
 "Links: Follow & Nofollow Internal Inlinks To Page","Links: Pages With High External Outlinks",
 "Links: Pages With High Internal Outlinks","Response Codes: External No Response",
 "Response Codes: External Server Error (5xx)","Response Codes: External Blocked Resource",
 "Directives: NoODP","Directives: NoYDIR","Directives: NoTranslate","Canonicals: Canonical Is Relative",
 "JavaScript: Page Title Updated by JavaScript","JavaScript: H1 Updated by JavaScript",
 "JavaScript: Meta Description Updated by JavaScript","JavaScript: Contains JavaScript Content",
 "JavaScript: Pages with JavaScript Errors","Content: Grammar Errors","Content: Low Relevance Content",
 "Structured Data: Rich Result Validation Warnings","Sitemaps: URLs in Multiple Sitemaps",
 "Mobile: Mobile Alternate Link","Validation: High Carbon Rating","PageSpeed: Optimize DOM Size",
 "PageSpeed: Font Display","PageSpeed: Legacy JavaScript","PageSpeed: Forced Reflow",
 "PageSpeed: Network Dependency Tree","PageSpeed: Duplicated JavaScript",
 "Search Console: No Search Analytics Data","Search Console: URL is Not on Google",
 "Search Console: Page is Not Mobile Friendly","Search Console: AMP URL Invalid",
 "Analytics: Bounce Rate Above 70%","Analytics: No GA Data",
}
DROP_PREFIXES = ("H2:", "Social:", "Accessibility:", "URL:", "Content: Readability", "AMP: Missing",
 "AMP: Contains", "AMP: Other", "Security: Missing", "Security: Unsafe", "Security: Bad", "Security: Protocol")
NOISE_COLS = {"Size (Bytes)","Path Type","Link Path","Link Position","Link Origin","Target","Rel","Follow",
 "Crawlability","Type","Occurrences","Hash","Response Time","Last Modified","Language","HTTP Version"}
TIER_NAME = {1:"Critical",2:"Important",3:"Minor"}
SF_TIER = {"High":2,"Medium":3,"Low":None}
# ----------------------------------------------------------------------

def norm(s):
    """Match issue names to CSV file names: keep only letters and digits."""
    return re.sub(r"[^a-z0-9]", "", str(s).lower())

def rule_for(name, sf_priority):
    if name in RULES:
        return RULES[name], False
    if name in DROP or name.startswith(DROP_PREFIXES):
        return ("appendix",None,"Appendix","",False,""), False
    t = SF_TIER.get(sf_priority)  # unmapped: fall back on SF priority and flag it
    return (("keep",t,"Other","Medium",False,"") if t else ("appendix",None,"Appendix","",False,"")), True

def is_html(row):
    ct = str(row.get("Content Type", "")) if "Content Type" in row else ""
    if ct and ct != "nan":
        return "html" in ct.lower()
    return not re.search(r"\.(pdf|jpe?g|png|gif|webp|svg|docx?|xlsx?|zip|mp4)(\?|$)", str(row.get("Address","")), re.I)

def tab_name(name, used):
    base = re.sub(r"[\[\]\*\?/\\:]", "", name.replace(": ", " - "))[:31].strip()
    n, i = base, 2
    while n in used:
        n = base[:28] + f" {i}"; i += 1
    used.add(n); return n

def style(ws, widths=None, wrap_cols=()):
    hdr = PatternFill("solid", fgColor="1F3864")
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF"); c.fill = hdr
        c.alignment = Alignment(vertical="center", wrap_text=True)
    ws.freeze_panes = "A2"
    if ws.max_row > 1:
        ws.auto_filter.ref = ws.dimensions
    for i, col in enumerate(ws.iter_cols(min_row=1, max_row=min(ws.max_row, 200)), 1):
        L = get_column_letter(i)
        w = (widths or {}).get(col[0].value) or min(max(len(str(c.value or "")) for c in col) + 2, 70)
        ws.column_dimensions[L].width = max(w, 10)
        if col[0].value in wrap_cols:
            for c in ws[L][1:]:
                c.alignment = Alignment(wrap_text=True, vertical="top")

GSHEETS_CELL_LIMIT = 10_000_000

def write_raw_merged(file_to_sheet, read, out_path):
    """Merge every CSV into one workbook: Index tab first, then issues_overview_report,
    then one tab per CSV (A to Z). Real headers in row 1 so it filters cleanly in Excel
    and Google Sheets. Returns (tabs, total_cells)."""
    names = sorted(file_to_sheet, key=lambda k: (k != "issues_overview_report", k))
    used, index, cells = {"Index"}, [], 0
    with pd.ExcelWriter(out_path, engine="openpyxl") as xw:
        pd.DataFrame(columns=["Tab", "Source CSV", "Rows", "Columns"]).to_excel(xw, "Index", index=False)
        for stem in names:
            df = read(file_to_sheet[stem])
            tab = tab_name(stem, used)
            df.to_excel(xw, tab, index=False)
            style(xw.book[tab])
            index.append({"Tab": tab, "Source CSV": stem + ".csv", "Rows": len(df), "Columns": df.shape[1]})
            cells += (len(df) + 1) * max(df.shape[1], 1)
        pd.DataFrame(index).to_excel(xw, "Index", index=False)
        ws = xw.book["Index"]
        for row in ws.iter_rows(min_row=2):
            row[0].hyperlink = f"#'{row[0].value}'!A1"; row[0].font = Font(color="0563C1", underline="single")
        style(ws)
    return len(index), cells

def open_export(src):
    """Return ({csv_stem: handle}, reader). Accepts a .zip of CSVs, a folder of CSVs,
    or a merged .xlsx where row 1 of each tab is 'File Name: x.csv'."""
    if os.path.isdir(src):
        m = {f[:-4]: os.path.join(src, f) for f in os.listdir(src) if f.lower().endswith(".csv")}
        return m, lambda p: pd.read_csv(p, encoding="utf-8-sig", low_memory=False)
    if src.lower().endswith(".zip"):
        import zipfile
        z = zipfile.ZipFile(src)
        m = {os.path.basename(n)[:-4]: n for n in z.namelist() if n.lower().endswith(".csv")
             and not os.path.basename(n).startswith("._")}
        return m, lambda n: pd.read_csv(z.open(n), encoding="utf-8-sig", low_memory=False)
    wb = openpyxl.load_workbook(src, read_only=True)
    m, hdr = {}, {}
    for ws in wb.worksheets:
        first = next(ws.iter_rows(max_row=1, values_only=True), (None,))[0]
        if isinstance(first, str) and first.startswith("File Name:"):
            m[first.split(":", 1)[1].strip().removesuffix(".csv")] = ws.title
            hdr[ws.title] = 1
        else:
            m[ws.title] = ws.title
            hdr[ws.title] = 0
    wb.close()
    return m, lambda sh: pd.read_excel(src, sheet_name=sh, header=hdr[sh])

def main(src, out_dir, client=None):
    os.makedirs(out_dir, exist_ok=True)
    file_to_sheet, read = open_export(src)
    if "issues_overview_report" not in file_to_sheet:
        sys.exit("No issues_overview_report tab found. Export Bulk Export > Issues > All from Screaming Frog.")
    ov = read(file_to_sheet["issues_overview_report"])
    ov = ov[pd.to_numeric(ov["URLs"], errors="coerce").fillna(0) > 0]  # skip checks with no hits
    by_norm = {norm(k): v for k, v in file_to_sheet.items()}

    raw_info = None
    if os.path.isdir(src) or src.lower().endswith(".zip"):  # a merged .xlsx input is already merged
        raw_path = os.path.join(out_dir, "sf_raw_merged.xlsx")
        tabs, cells = write_raw_merged(file_to_sheet, read, raw_path)
        raw_info = dict(path=raw_path, tabs=tabs, cells=cells, fits_google_sheets=cells <= GSHEETS_CELL_LIMIT)

    issues, details, hosts = [], {}, Counter()
    for _, r in ov.iterrows():
        name = str(r["Issue Name"]).strip()
        (action, tier, theme, effort, idx_only, why), unmapped = rule_for(name, str(r.get("Issue Priority", "")))
        rec = dict(issue=name, sf_type=r.get("Issue Type"), sf_priority=r.get("Issue Priority"),
                   sf_urls=int(r.get("URLs") or 0), action=action, tier=tier, theme=theme, effort=effort,
                   why=why or str(r.get("Description", ""))[:300], how_to_fix=str(r.get("How To Fix", "")),
                   clean_urls=None, inlinks=None, examples=[], tab=None, unmapped=unmapped)
        s = norm(name)
        sheet = by_norm.get(s)
        if action == "keep" and sheet:
            df = read(sheet).dropna(how="all")
            if idx_only and "Indexability" in df.columns:
                df = df[df["Indexability"].astype(str).str.strip().eq("Indexable")]
            if idx_only and ("Address" in df.columns):
                df = df[df.apply(is_html, axis=1)]
            key = "Address" if "Address" in df.columns else ("Destination" if "Destination" in df.columns else df.columns[0])
            inl = by_norm.get(s + "inlinks")
            if inl and key == "Address":
                il = read(inl)
                if "Destination" in il.columns and "Source" in il.columns:
                    g = il.groupby("Destination")["Source"].agg(["count", "first"])
                    df = df.merge(g.rename(columns={"count": "Internal Links Pointing Here", "first": "Example Linking Page"}),
                                  left_on="Address", right_index=True, how="left")
                    rec["inlinks"] = int(il["Destination"].isin(df["Address"]).sum())
            df = df[[c for c in df.columns if c not in NOISE_COLS and df[c].notna().any()]]
            if "Internal Links Pointing Here" in df.columns:
                df = df.sort_values("Internal Links Pointing Here", ascending=False)
            rec["clean_urls"] = int(df[key].nunique())
            rec["examples"] = df[key].astype(str).head(3).tolist()
            for u in df[key].astype(str).head(500):
                hosts[urlparse(u).netloc] += 1
            if rec["clean_urls"]:
                details[name] = df
        elif action == "keep":
            rec["clean_urls"] = rec["sf_urls"]
        issues.append(rec)

    domain = hosts.most_common(1)[0][0] if hosts else ""
    kept = [i for i in issues if i["action"] == "keep" and i["clean_urls"]]
    kept.sort(key=lambda i: (i["tier"] or 9, -(i["clean_urls"] or 0)))
    appendix = sorted([i for i in issues if i not in kept], key=lambda i: -i["sf_urls"])

    base = re.sub(r"[^A-Za-z0-9]+", "-", (client or domain or "site")).strip("-").lower()
    xlsx = os.path.join(out_dir, f"{base}_seo_audit_clean.xlsx")
    used = {"Executive Summary", "All Issues"}
    for i in kept:
        if i["issue"] in details:
            i["tab"] = tab_name(i["issue"], used)

    with pd.ExcelWriter(xlsx, engine="openpyxl") as xw:
        es = pd.DataFrame([{"Priority": TIER_NAME[i["tier"]], "Area": i["theme"], "Issue": i["issue"],
              "Pages Affected": i["clean_urls"], "Internal Links Affected": i["inlinks"],
              "Why It Matters": i["why"], "Effort": i["effort"], "Details Tab": i["tab"] or ""} for i in kept])
        es.to_excel(xw, "Executive Summary", index=False)
        al = pd.DataFrame([{"Issue": i["issue"], "Included": ("Yes" if i in kept else "Appendix only") + (" (no rule)" if i["unmapped"] else ""),
              "SF Priority": i["sf_priority"], "SF Type": i["sf_type"], "SF URL Count": i["sf_urls"],
              "Clean URL Count": i["clean_urls"], "How To Fix (Screaming Frog)": i["how_to_fix"]}
              for i in kept + appendix])
        al.to_excel(xw, "All Issues", index=False)
        for i in kept:
            if i["tab"]:
                details[i["issue"]].to_excel(xw, i["tab"], index=False)
        bk = xw.book
        style(bk["Executive Summary"], {"Why It Matters": 70, "Issue": 45}, ("Why It Matters",))
        style(bk["All Issues"], {"How To Fix (Screaming Frog)": 80, "Issue": 50}, ("How To Fix (Screaming Frog)",))
        fills = {"Critical": "F8CBAD", "Important": "FFE699", "Minor": "E2EFDA"}
        for row in bk["Executive Summary"].iter_rows(min_row=2):
            row[0].fill = PatternFill("solid", fgColor=fills.get(row[0].value, "FFFFFF"))
            tab = row[7].value
            if tab:
                row[7].hyperlink = f"#'{tab}'!A1"; row[7].font = Font(color="0563C1", underline="single")
        for i in kept:
            if i["tab"]:
                style(bk[i["tab"]])

    summary = dict(client=client, domain=domain, source_file=os.path.basename(src), clean_workbook=xlsx, raw_merged=raw_info,
        counts=dict(sf_issue_types=len(issues), kept=len(kept), appendix=len(appendix),
                    critical=sum(i["tier"] == 1 for i in kept), important=sum(i["tier"] == 2 for i in kept),
                    minor=sum(i["tier"] == 3 for i in kept)),
        issues=[{k: v for k, v in i.items() if k not in ("action",)} for i in kept],
        appendix=[{"issue": i["issue"], "sf_urls": i["sf_urls"]} for i in appendix],
        unmapped=[i["issue"] for i in issues if i["unmapped"]])
    js = os.path.join(out_dir, f"{base}_summary.json")
    with open(js, "w") as f:
        json.dump(summary, f, indent=1, default=str)
    print(f"Domain: {domain}\nKept {len(kept)} of {len(issues)} issue types "
          f"({summary['counts']['critical']} critical, {summary['counts']['important']} important, "
          f"{summary['counts']['minor']} minor)\nWorkbook: {xlsx}\nSummary: {js}")
    if raw_info:
        print(f"Raw merged: {raw_info['path']} ({raw_info['tabs']} tabs, {raw_info['cells']:,} cells"
              + ("" if raw_info["fits_google_sheets"] else ", OVER the 10M-cell Google Sheets limit") + ")")
    for i in kept:
        print(f"  [{TIER_NAME[i['tier']]}] {i['issue']}: {i['clean_urls']} pages (SF {i['sf_urls']})")
    if summary["unmapped"]:
        print("UNMAPPED (no rule, used SF priority): " + "; ".join(summary["unmapped"]))

if __name__ == "__main__":
    if len(sys.argv) < 3:
        sys.exit(__doc__)
    main(sys.argv[1], sys.argv[2], sys.argv[3] if len(sys.argv) > 3 else None)
