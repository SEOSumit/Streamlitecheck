import sys
from xml_sitemap import generate_xml_export_workbook

# Sample URLs to test the export logic
test_urls = [
    "https://example.com/page1",
    "https://example.com/page2",
    "https://example.com/page3",
]

try:
    print("Testing generate_xml_export_workbook...")
    output_bytes = generate_xml_export_workbook(test_urls)
    print(f"Success! Output size: {len(output_bytes)} bytes")
except Exception as e:
    print(f"Error: {e}")
    sys.exit(1)
