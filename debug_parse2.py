from bs4 import BeautifulSoup
soup = BeautifulSoup(open('us_thinkpad.html', encoding='utf-8'), 'html.parser')
main = soup.find('main') or soup.find(role='main') or soup.find(id='main-content')
for c in main.find_all('div', recursive=False):
    is_product = c.select('.product-list, .product-item, .product-grid, [componentname*=\"ofp-2c-mobile-new-dlp\"], .product_card')
    bopc_text = c.get_text(strip=True)
    if is_product: 
        print(c.get('class'), 'CONTAINS PRODUCTS. Chars:', len(bopc_text))
        print('  Contains container9999?', len(c.select('.container9999')))
    else:
        print(c.get('class'), 'Does NOT contain products. Chars:', len(bopc_text))
        print('  Contains container9999?', len(c.select('.container9999')))
