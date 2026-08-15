from playwright.sync_api import sync_playwright

def dump_page():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        page.goto("https://www.lenovo.com/jp/ja/d/15-inch-laptops/", wait_until="networkidle", timeout=60000)
        
        # Scroll down multiple times
        for _ in range(15):
            page.evaluate("window.scrollBy(0, 800)")
            page.wait_for_timeout(1000)
            
        html = page.content()
        with open("lenovo_dump.html", "w", encoding="utf-8") as f:
            f.write(html)
        browser.close()

if __name__ == "__main__":
    dump_page()
