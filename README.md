# Internal Link Implementation Checker

This tool checks whether each suggested internal link is implemented in the specified paragraph.

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
