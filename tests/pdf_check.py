"""
Word -> PDF checks (Papyros "Download PDF"), in real Chrome.

  python3 tests/pdf_check.py http://127.0.0.1:8090/something/

Checks: the PDF has as many pages as the editor shows, every page carries its text, fonts are
embedded (subset), the picture and the table come through, and no CSP violation happens.
Page 1 of each PDF is rendered to tests/out/*.png (needs pdftoppm) for a visual comparison.
"""
import base64, io, os, shutil, subprocess, sys
from playwright.sync_api import sync_playwright
from pypdf import PdfReader

BASE = sys.argv[1] if len(sys.argv) > 1 else 'http://127.0.0.1:8090/something/'
HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, 'out')
os.makedirs(OUT, exist_ok=True)
DROP_JS = open(os.path.join(HERE, 'e2e.py')).read().split('DROP_JS = """')[1].split('"""')[0]
failures = []


def check(cond, msg):
    print(('PASS ' if cond else 'FAIL ') + msg)
    if not cond:
        failures.append(msg)


def fonts_of(reader):
    names, embedded = set(), True
    for page in reader.pages:
        for f in (page.get('/Resources') or {}).get('/Font', {}).values():
            f = f.get_object()
            for d in f.get('/DescendantFonts', [f]):
                d = d.get_object()
                fd = d.get('/FontDescriptor')
                if fd is None:
                    continue
                fd = fd.get_object()
                names.add(str(fd.get('/FontName')))
                embedded &= any(k in fd for k in ('/FontFile', '/FontFile2', '/FontFile3'))
    return names, embedded


def images_of(reader):
    n = 0
    for page in reader.pages:
        xo = (page.get('/Resources') or {}).get('/XObject', {})
        n += sum(1 for v in xo.values() if v.get_object().get('/Subtype') == '/Image')
    return n


def render(path, name):
    if shutil.which('pdftoppm'):
        subprocess.run(['pdftoppm', '-r', '60', '-f', '1', '-l', '1', '-png', path, os.path.join(OUT, name)], check=False)


with sync_playwright() as p:
    b = p.chromium.launch(args=['--no-sandbox'])
    ctx = b.new_context(accept_downloads=True, viewport={'width': 1440, 'height': 900})
    ctx.add_init_script("delete window.showOpenFilePicker; delete window.showSaveFilePicker;")
    pg = ctx.new_page()
    errors = []
    pg.on('console', lambda m: (m.type == 'error') and errors.append(m.text))
    pg.add_init_script("document.addEventListener('securitypolicyviolation', e => console.error('CSP-VIOLATION ' + e.violatedDirective + ' ' + e.blockedURI));")
    pg.goto(BASE, wait_until='networkidle')

    # 1. Long document: every page, every page's text.
    pg.evaluate(DROP_JS, ['long.docx', base64.b64encode(open(os.path.join(HERE, 'fixtures/long.docx'), 'rb').read()).decode()])
    pg.locator('text=Long document').first.wait_for(timeout=30000)
    pg.wait_for_timeout(2500)
    screen = pg.locator('.docx-editor .docx-paginated-surface .docx-page').count()
    with pg.expect_download(timeout=120000) as dl:
        pg.get_by_role('button', name='Download PDF').click()
    path = os.path.join(OUT, 'long.pdf')
    dl.value.save_as(path)
    r = PdfReader(path)
    texts = [x.extract_text() or '' for x in r.pages]
    check(dl.value.suggested_filename == 'long.pdf', f'PDF named after the document ({dl.value.suggested_filename})')
    check(len(r.pages) == screen, f'PDF has {len(r.pages)} pages; the editor shows {screen}')
    check(all('Paragraph' in t or 'Long document' in t for t in texts), 'every PDF page carries its text')
    check('Paragraph 120' in texts[-1], 'the last page carries the last paragraph')
    names, embedded = fonts_of(r)
    check(embedded and names and all('+' in n for n in names), f'fonts embedded as subsets {sorted(names)}')
    render(path, 'long-p1')
    pg.wait_for_timeout(300)
    check(any('PDF saved' in t for t in pg.locator('.toast').all_inner_texts()), 'Papyros reports the PDF was saved')

    # 2. Rich document: table text, picture, page numbers, footnote, Greek.
    pg.get_by_role('button', name='Back to the Greek Office hub').click()
    pg.evaluate(DROP_JS, ['rich.docx', base64.b64encode(open(os.path.join(HERE, 'fixtures/rich.docx'), 'rb').read()).decode()])
    pg.locator('text=Rich document').first.wait_for(timeout=30000)
    pg.wait_for_timeout(2500)
    with pg.expect_download(timeout=120000) as dl:
        pg.get_by_role('button', name='Download PDF').click()
    path = os.path.join(OUT, 'rich.pdf')
    dl.value.save_as(path)
    r = PdfReader(path)
    t0 = r.pages[0].extract_text() or ''
    check('R2C3' in t0 and 'Mixed formatting' in t0, 'table and body text in the PDF')
    check('Page 1' in t0 and 'Page 2' in (r.pages[-1].extract_text() or ''), 'footer page numbers resolved per page')
    check('footnote sits' in t0, 'footnote text in the PDF')
    check('Ἀθῆναι' in t0.replace('\n', ''), 'polytonic Greek survives (fallback font)')
    check('Text inside a shape' in t0, 'text-box shape text in the PDF')
    check(images_of(r) >= 1, 'the picture is embedded')
    render(path, 'rich-p1')
    # Chrome's own print engine on the same document, for a side-by-side look.
    # (A function wrapper: an expression that evaluates to a function would be called by Playwright.)
    pg.evaluate("() => { window.__printed = 0; window.print = () => { window.__printed += 1; }; }")
    pg.get_by_role('button', name='Print', exact=True).click()
    for _ in range(300):
        if pg.evaluate('window.__printed || 0') > 0:
            break
        pg.wait_for_timeout(100)
    printed = pg.pdf(prefer_css_page_size=True, print_background=True)
    open(os.path.join(OUT, 'rich-print.pdf'), 'wb').write(printed)
    check(len(PdfReader(io.BytesIO(printed)).pages) == len(r.pages), f'same page count as Print ({len(r.pages)})')
    render(os.path.join(OUT, 'rich-print.pdf'), 'rich-print-p1')

    csp = [e for e in errors if 'CSP-VIOLATION' in e]
    check(not csp, f'no CSP violations {csp[:3]}')
    b.close()

print()
print('ALL PASSED' if not failures else f'{len(failures)} FAILED:\n  - ' + '\n  - '.join(failures))
sys.exit(1 if failures else 0)
