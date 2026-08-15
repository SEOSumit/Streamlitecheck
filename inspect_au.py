from playwright.sync_api import sync_playwright

def inspect_au():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        print("Navigating to AU tablets...")
        page.goto("https://www.lenovo.com/au/en/c/tablets/android-tablets/", wait_until="domcontentloaded")
        page.wait_for_timeout(3000)
        
        for _ in range(10):
            page.evaluate("window.scrollBy(0, 500)")
            page.wait_for_timeout(500)
        
        # Look for common product wrappers
        print("Product grid count:", page.locator('.product-grid, [data-product-list], ul.products').count())
        print("Items with 'product' in class:", page.locator('[class*="product"]').count())
        
        # See what the structure is
        html = page.content()
        with open("au_html_snippet.html", "w", encoding="utf-8") as f:
            f.write(html)
        print("Wrote full HTML to au_html_snippet.html")
        browser.close()

if __name__ == "__main__":
    inspect_au()
