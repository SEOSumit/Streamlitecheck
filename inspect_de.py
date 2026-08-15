from playwright.sync_api import sync_playwright
import time

def inspect_de():
    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page()
        page.goto('https://www.lenovo.com/de/de/c/laptops/yoga/yoga-2-in-1-series/', wait_until='domcontentloaded')
        time.sleep(3)
        for _ in range(5):
            page.evaluate('window.scrollBy(0, 500)')
            time.sleep(0.5)
        
        print("Product Grid count:", page.locator('.product-grid, .product-list, .product_list, .productList_container').count())
        html = page.content()
        with open('de_html_2.html', 'w', encoding='utf-8') as f:
            f.write(html)
        browser.close()

if __name__ == "__main__":
    inspect_de()
