# SEO Toolkit — Beta

**Status: Beta.** This is an internal tool under active development. Features, output format, and behavior may change without notice. Not for external/client distribution yet.

---

## What this does

Two tools, one app:

### 1. Internal Link Implementation Checker
Checks whether a suggested internal link has actually been implemented on a live page.

You give it a spreadsheet where each row says: "on this page, in this paragraph, this anchor text should link to this URL." The tool visits the page and confirms whether that exact anchor is linked to that exact URL inside that specific paragraph — not just linked *somewhere* on the page.

A row only counts as **Implemented** if all three are true:
- The specified paragraph exists on the page
- The exact anchor text is a link inside that paragraph
- That link points to the suggested URL (relative links / trailing slashes count as the same URL)

**Output:** your original sheet, plus columns for Implementation Status, the actual hyperlink found (if any), a short explanation, and the page's HTTP status — so you can tell "not implemented" apart from "page is broken/blocked."

### 2. HTML Sitemap Audit
Compares your master list of URLs against what's actually live on a site's HTML sitemap page.

You upload a sheet of URLs + status codes. You give it the sitemap page URL. It collects every link that page actually shows to a visitor (not just what's in the raw page source — some sites build their sitemap with JavaScript, so a plain fetch would miss links). If the site blocks automated browsing, there's a manual fallback: open the sitemap in Chrome, copy the links using a short console command, paste them in.

**Output:** a 3-sheets one excel file —
- **Summary** — counts at a glance
- **Existing URLs in HTML Sitemap** — every URL that's both on the sitemap and in your uploaded list
- **URLs Not in HTML Sitemap** — URLs from your list (status 200 only) that are missing from the sitemap

Before it builds the file, it shows you a sample of what it collected so you can confirm the count looks right — this stops it from silently generating a wrong report off an incomplete page load.

---

## Access

Access the tool here: https://basiccheck.streamlit.app/

## Rules of use (internal)
- Don't share this app or its output files outside the team without checking first.
- Don't run it against a client's live site without their knowledge if it involves repeated automated visits — treat it like any crawl and be considerate of load.
- Report bad/incorrect output immediately rather than trusting a run blindly — it's beta, verify before you send anything to a client.

## Development progress
- [x] Internal Link Checker — working
- [x] HTML Sitemap Audit — working, includes manual fallback for blocked sites
- [ ] AI-suggested anchor text for missing sitemap URLs — planned, not built yet
- [ ] Known limitation: heavily bot-protected sites may still need the manual paste method
