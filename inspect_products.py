from playwright.sync_api import sync_playwright

def inspect_products():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        page.goto("https://www.lenovo.com/jp/ja/d/15-inch-laptops/", wait_until="domcontentloaded")
        page.wait_for_timeout(3000)
        
        # Check if skeleton exists
        print("Skeleton count:", page.locator('.skeleton, .skeleton_product, .pc_dlp_skeleton_box').count())
        
        # Scroll down
        for _ in range(10):
            page.evaluate("window.scrollBy(0, 500)")
            page.wait_for_timeout(500)
            
        print("Skeleton count after scroll:", page.locator('.skeleton, .skeleton_product, .pc_dlp_skeleton_box').count())
        
        dlp = page.locator('[componentname="ofp-2c-mobile-new-dlp"]')
        if dlp.count() > 0:
            print("DLP HTML:")
            print(dlp.inner_html()[:2000])
            
        browser.close()

if __name__ == "__main__":
    inspect_products()
