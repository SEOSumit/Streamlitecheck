# SEO Toolkit

The Streamlit app currently includes:

- Internal Link Implementation Checker
- HTML Sitemap Audit

## Run on Windows

1. Install Python 3.10 or newer if it is not already installed.
2. Double-click `run_tool.bat`.
3. A browser window will open.
4. Upload the `.xlsx` sheet, click **Check internal links**, and download the result.

## Result rule

`Implemented` requires all three conditions:

1. The specified paragraph is present on the source page.
2. The exact suggested anchor text is an `<a>` link inside that paragraph.
3. The anchor destination matches the suggested hyperlink.

Relative links and trailing-slash differences are treated as equivalent. A link elsewhere on the page does not count.

## Output

The original workbook is preserved and the tool adds:

- Source Page URL
- Implementation Status
- Actual Hyperlink Found
- Check Details
- Page HTTP Status
- Final Page URL
- Checked At

A **Link Check Summary** sheet is also added.

## HTML Sitemap Audit

Upload one file with URLs in column A, status codes in column B, and final redirect URLs in column C. The tool fetches the HTML sitemap once by default and does not visit every URL. If a sitemap link is 3xx and column C contains its final destination, `Suggested Link` is populated automatically. Saved HTML upload and pasted source remain available only as fallbacks.

The user can audit the complete sitemap or only URLs containing a folder pattern such as `/servers-storage/`. Clean and parameterized versions are treated as the same URL. The output matches the three-sheet audit format: Summary, Existing Pages in HTML Sitemap, and URLs Not in HTML Sitemap.
