import time
from playwright.sync_api import sync_playwright

with sync_playwright() as p:
    browser = p.chromium.launch(headless=True)
    page = browser.new_page()
    page.goto('https://www.lenovo.com/sg/en/d/gaming-monitor/', wait_until='domcontentloaded', timeout=15000)
    page.wait_for_timeout(2000)
    
    # 1. Click all 'read more' buttons
    read_more_btns = page.locator('div.readMore[role="button"]')
    count = read_more_btns.count()
    for i in range(count):
        btn = read_more_btns.nth(i)
        if btn.is_visible():
            try:
                btn.click()
                time.sleep(0.5)
            except:
                pass

    # 2. Click all FAQ accordions
    faq_accordions = page.locator('.faq button, .q-a button, [class*="faq" i] button, [class*="question" i] button, details summary')
    count = faq_accordions.count()
    print("Found", count, "faq buttons")
    for i in range(count):
        btn = faq_accordions.nth(i)
        if btn.is_visible():
            try:
                # scroll into view and click
                btn.scroll_into_view_if_needed()
                btn.click()
                time.sleep(0.5)
            except:
                pass
                
    time.sleep(1)
    
    # Check what is now in the body
    extract_script = """
    () => {
        const clone = document.body.cloneNode(true);
        const elementsToRemove = clone.querySelectorAll(
            "script, style, noscript, iframe, link, meta, svg, " +
            "header, footer, nav, aside, " +
            "[role='navigation'], [role='banner'], [role='contentinfo'], " +
            "#header, #footer"
        );
        for (let el of elementsToRemove) {
            el.remove();
        }
        return clone.textContent;
    }
    """
    inner_text = page.evaluate(extract_script)
    cleaned_text = " ".join(inner_text.split())
    print("Extracted text snippet:")
    print(cleaned_text[-1000:]) # Last 1000 chars should be FAQ + BOPC
    browser.close()
