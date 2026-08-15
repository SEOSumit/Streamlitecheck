import json
import os
import sys

from dlp_content_audit import process_single_url

test_urls = [
    "https://www.lenovo.com/jp/ja/d/15-inch-laptops/",
    "https://www.lenovo.com/jp/ja/d/gaming-laptops/",
    "https://www.lenovo.com/us/en/c/laptops/thinkpad/thinkpad-t-series/",
    "https://www.lenovo.com/de/de/c/laptops/yoga/yoga-2-in-1-series/",
    "https://www.lenovo.com/au/en/c/tablets/android-tablets/"
]

def run_tests():
    results = []
    for url in test_urls:
        print(f"Testing URL: {url}")
        try:
            result = process_single_url(url)
            results.append(result)
            print(f"  Page Status: {result['Page Status']}")
            print(f"  Reported Product Count: {result['Reported Product Count']}")
            print(f"  Rendered Product Cards: {result['Rendered Product Cards']}")
            print(f"  BOPC Available: {result['Is BOPC Available?']}")
            print(f"  Q&A Available: {result['Is Q&A Available?']}")
        except Exception as e:
            print(f"  Error: {e}")
    
    with open("bulk_test_results.json", "w", encoding="utf-8") as f:
        json.dump(results, f, indent=4, ensure_ascii=False)
        print("Wrote detailed results to bulk_test_results.json")

if __name__ == "__main__":
    run_tests()
