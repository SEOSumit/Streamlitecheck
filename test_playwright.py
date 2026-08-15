import asyncio
from playwright.async_api import async_playwright
import markdownify

async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36"
        )
        page = await context.new_page()
        # trying a sample Lenovo page that was erroring
        print("Loading page...")
        await page.goto("https://www.lenovo.com/sg/en/d/8-inch-tablets/", wait_until='domcontentloaded', timeout=30000)
        await page.wait_for_timeout(2000)
        
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
            return clone.innerHTML;
        }
        """
        print("Evaluating script...")
        raw_html = await page.evaluate(extract_script)
        print("HTML LENGTH:", len(raw_html) if raw_html else 0)
        
        if raw_html:
            md_text = markdownify.markdownify(raw_html, heading_style="ATX", strip=['img', 'svg'])
            print("MD LENGTH:", len(md_text))
            print("MD PREVIEW:", md_text[:500].strip())
            if len(md_text.strip()) == 0:
                print("Wait, markdownify stripped everything!")
                print("RAW HTML Preview:", raw_html[:1000])
        else:
            print("raw_html is None!")
            
        await browser.close()

asyncio.run(main())
