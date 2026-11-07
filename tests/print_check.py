"""Print / PDF check: long Word document -> every page printed with its text (Chrome print engine)."""
import base64, io, os, sys
from playwright.sync_api import sync_playwright
from pypdf import PdfReader

BASE = sys.argv[1] if len(sys.argv) > 1 else 'http://127.0.0.1:8090/something/'
HERE = os.path.dirname(os.path.abspath(__file__))
DROP_JS = open(os.path.join(HERE, 'e2e.py')).read().split('DROP_JS = """')[1].split('"""')[0]
ok = True
with sync_playwright() as p:
    b = p.chromium.launch(args=['--no-sandbox'])
    pg = b.new_page(viewport={'width': 1440, 'height': 900})
    pg.add_init_script("window.__printed = 0; window.print = () => { window.__printed++; };")
    pg.goto(BASE, wait_until='networkidle')
    pg.evaluate(DROP_JS, ['long.docx', base64.b64encode(open(os.path.join(HERE, 'fixtures/long.docx'), 'rb').read()).decode()])
    pg.locator('text=Long document').first.wait_for(timeout=20000)
    pg.wait_for_timeout(2500)
    screen_pages = pg.locator('.docx-editor .docx-paginated-surface .docx-page').count()
    pg.get_by_text('File', exact=True).first.click()
    pg.wait_for_timeout(400)
    menu = pg.locator('[role=menu]').first.inner_text()
    hidden_ok = 'Print' not in menu and 'Export' not in menu
    print(('PASS' if hidden_ok else 'FAIL'), 'File menu has no Print/Export rows:', menu.replace('\n', ' | '))
    ok &= hidden_ok
    pg.keyboard.press('Escape')
    pg.get_by_role('button', name='Print', exact=True).click()
    # (wait_for_function would need eval, which the app's CSP correctly blocks: poll instead)
    for _ in range(600):
        if pg.evaluate('window.__printed') > 0:
            break
        pg.wait_for_timeout(100)
    pdf = pg.pdf(prefer_css_page_size=True, print_background=True)
    open(os.path.join(HERE, 'out', 'printed.pdf'), 'wb').write(pdf)
    reader = PdfReader(io.BytesIO(pdf))
    n = len(reader.pages)
    texts = [pg_.extract_text() or '' for pg_ in reader.pages]
    print(('PASS' if n == screen_pages else 'FAIL'), f'printed {n} pages; editor shows {screen_pages}')
    ok &= n == screen_pages
    empty = [i + 1 for i, t in enumerate(texts) if 'Paragraph' not in t and 'Long document' not in t]
    print(('PASS' if not empty else 'FAIL'), f'every printed page has text (empty: {empty})')
    ok &= not empty
    last = texts[-1]
    print(('PASS' if 'Paragraph 120' in last else 'FAIL'), 'last page carries the last paragraph')
    ok &= 'Paragraph 120' in last
    box = reader.pages[0].mediabox
    print('   page size (pt):', float(box.width), 'x', float(box.height))
    b.close()
print('ALL PASSED' if ok else 'FAILED')
sys.exit(0 if ok else 1)
