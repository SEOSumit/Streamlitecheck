from bs4 import BeautifulSoup
soup = BeautifulSoup(open('us_thinkpad.html', encoding='utf-8'), 'html.parser')

print('--- FAQ ---')
faqs = soup.select('.faqs .sectionLi')
print('Lenovo faqs:', len(faqs))
accs = soup.select('details, [aria-expanded], [data-accordion]')
print('Accordions:', len(accs))

print('--- BOPC ---')
main = soup.find('main') or soup.find(role='main') or soup.find(id='main-content')
if main:
    for c in main.find_all(recursive=False):
        text = c.get_text(separator=' ', strip=True)
        h = c.find_all(['h1','h2','h3','h4','h5'])
        p = c.find_all('p')
        if len(text) > 200:
            print(f'TAG: {c.name}, CLASS: {c.get("class")}, CHARS: {len(text)}, HEADINGS: {len(h)}, PARAGRAPHS: {len(p)}')
            if len(h) > 0:
                print('   Headings:', [x.get_text(strip=True)[:50] for x in h[:3]])
