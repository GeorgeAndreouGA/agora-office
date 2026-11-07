"""
End-to-end checks in real Chrome (Playwright). Run with the app served at BASE.

  python3 tests/e2e.py http://127.0.0.1:8090/something/

Checks: hub loads with no CSP violations or console errors, the gate refuses hostile files,
Word, Excel and PowerPoint files open, edit and save back (download fallback), the saved files
are valid and keep their content, nothing is written to browser storage. Pinax (PowerPoint) is
also checked for PDF export, the slideshow and new presentations; saved decks are re-opened with
python-pptx and, when it is installed, converted by LibreOffice as an independent validity check.
"""
import base64, os, shutil, subprocess, sys, time, zipfile, io
from playwright.sync_api import sync_playwright

BASE = sys.argv[1] if len(sys.argv) > 1 else 'http://127.0.0.1:8090/something/'
HERE = os.path.dirname(os.path.abspath(__file__))
FIX = os.path.join(HERE, 'fixtures')
OUT = os.path.join(HERE, 'out')
os.makedirs(OUT, exist_ok=True)

failures = []
def check(cond, msg):
    print(('PASS ' if cond else 'FAIL ') + msg)
    if not cond:
        failures.append(msg)

DROP_JS = """
async ([name, b64]) => {
  const bin = Uint8Array.from(atob(b64), c => c.charCodeAt(0));
  const file = new File([bin], name);
  const dt = new DataTransfer();
  dt.items.add(file);
  window.dispatchEvent(new DragEvent('dragenter', { dataTransfer: dt, bubbles: true }));
  window.dispatchEvent(new DragEvent('drop', { dataTransfer: dt, bubbles: true, cancelable: true }));
}
"""

def drop(page, name):
    data = base64.b64encode(open(os.path.join(FIX, name), 'rb').read()).decode()
    page.evaluate(DROP_JS, [name, data])

with sync_playwright() as p:
    browser = p.chromium.launch(args=['--no-sandbox'])

    def new_page(disable_fsa=False):
        ctx = browser.new_context(accept_downloads=True, viewport={'width': 1440, 'height': 900})
        if disable_fsa:
            # Exercise the download fallback (what Firefox/Safari users get).
            ctx.add_init_script("delete window.showOpenFilePicker; delete window.showSaveFilePicker;")
        page = ctx.new_page()
        page.errors = []
        page.csp = []
        page.on('console', lambda m: (m.type == 'error') and page.errors.append(m.text))
        page.on('pageerror', lambda e: page.errors.append(str(e)))
        page.add_init_script("""
          document.addEventListener('securitypolicyviolation', e => {
            console.error('CSP-VIOLATION ' + e.violatedDirective + ' ' + e.blockedURI);
          });
        """)
        return ctx, page

    # ---------- hub ----------
    ctx, page = new_page()
    t0 = time.time()
    page.goto(BASE, wait_until='networkidle')
    check(page.get_by_role('heading', name='Enter the Greek Office').is_visible(), f'hub renders ({time.time()-t0:.2f}s)')
    check(page.evaluate('crossOriginIsolated') is True, 'page is cross-origin isolated (COOP+COEP)')
    cards = page.locator('.card .card__title').all_inner_texts()
    check(cards == ['Papyros', 'Abakion', 'Pinax'], f'hub shows the three editors {cards}')
    page.screenshot(path=os.path.join(OUT, 'hub.png'), full_page=True)

    # ---------- gate ----------
    cases = {
        'macro.docx': 'macros',
        'xxe.docx': 'unsafe XML',
        'late_doctype.docx': 'unsafe XML',
        'remote_template.docx': 'links to external content',
        'mismatch.docx': 'contains a Excel workbook',
        'fake.xlsx': 'not a valid',
        'bomb.xlsx': 'zip bomb',
        'protected.xlsx': 'password-protected',
        'evil.xlsm': 'Macro-enabled Excel',
        'vba.pptx': 'macros',
        'xxe.pptx': 'unsafe XML',
        'ole_link.pptx': 'links to external content',
        'macro.pptm': 'Macro-enabled PowerPoint',
        'mismatch_pptx.docx': 'contains a PowerPoint presentation',
        'legacy.ppt': 'Old PowerPoint',
    }
    for name, expect in cases.items():
        drop(page, name)
        try:
            page.locator('.toast--error', has_text=expect).first.wait_for(timeout=5000)
            check(True, f'gate refuses {name}')
        except Exception:
            toasts = page.locator('.toast').all_inner_texts()
            check(False, f'gate refuses {name} (toasts: {toasts})')
    check(page.get_by_role('heading', name='Enter the Greek Office').is_visible(), 'still on hub after refusals')
    ctx.close()

    # ---------- Papyros (Word) ----------
    ctx, page = new_page(disable_fsa=True)
    page.goto(BASE, wait_until='networkidle')
    drop(page, 'sample.docx')
    page.locator('.hdr__app', has_text='Papyros').wait_for(timeout=20000)
    page.wait_for_timeout(4000)
    body_text = page.locator('.editor-body').inner_text()
    check('Agora of Athens' in body_text or page.locator('text=Agora of Athens').count() > 0, 'Word document content renders')
    page.screenshot(path=os.path.join(OUT, 'papyros.png'))
    # Type at the start of the document
    page.locator('text=Agora of Athens').first.click()
    page.keyboard.press('End')
    page.keyboard.type(' ZETA-TEST')
    page.wait_for_timeout(800)
    check(page.locator('.hdr__pill--dirty').count() == 1, 'Word edit marks the document unsaved')
    with page.expect_download(timeout=20000) as dl:
        page.get_by_role('button', name='Save', exact=True).click()
    path = os.path.join(OUT, 'saved.docx')
    dl.value.save_as(path)
    from docx import Document
    doc = Document(path)
    text = '\n'.join(par.text for par in doc.paragraphs)
    check('ZETA-TEST' in text, 'saved .docx contains the typed text')
    check(len(doc.tables) == 1 and doc.tables[0].cell(1, 0).text == 'Stoa of Attalos', 'saved .docx keeps the table')
    page.wait_for_timeout(500)
    check(page.locator('.hdr__pill--dirty').count() == 0, 'unsaved marker clears after save')
    storage = page.evaluate("async () => ({ ls: Object.keys(localStorage), ss: Object.keys(sessionStorage), idb: (await indexedDB.databases()).map(d => d.name) })")
    print('   storage after Word edit:', storage)
    check(not any('docx' in k.lower() or 'doc' == k.lower() for k in storage['ls']), 'no document content in localStorage')
    csp = [e for e in page.errors if 'CSP-VIOLATION' in e]
    check(not csp, f'no CSP violations in Papyros {csp[:3]}')
    errs = [e for e in page.errors if 'CSP-VIOLATION' not in e]
    print('   console errors (Papyros):', errs[:5])
    ctx.close()

    # ---------- Papyros: Delete page ----------
    import re as _re
    ctx, page = new_page(disable_fsa=True)
    page.goto(BASE, wait_until='networkidle')
    def delete_page(n):
        page.get_by_role('button', name='Delete page').click()
        page.locator('dialog.confirm[open] input[type=number]').fill(str(n))
        page.locator('dialog.confirm[open]').get_by_role('button', name=_re.compile('^Delete page')).click()
        page.locator('.toast').first.wait_for(timeout=15000)
        page.wait_for_timeout(1500)
        return page.locator('.toast').all_inner_texts()
    pages = lambda: page.locator('.docx-editor .docx-paginated-surface .docx-page').count()
    drop(page, 'long.docx')
    page.locator('text=Long document').first.wait_for(timeout=30000)
    page.wait_for_timeout(2500)
    n0 = pages()
    toast = delete_page(5)
    n1 = pages()
    check(n1 == n0 - 1 and any('deleted' in t for t in toast), f'Delete page removes a page of text ({n0} -> {n1})')
    page.keyboard.press('Control+z')
    page.wait_for_timeout(1500)
    check(pages() == n0, f'Ctrl+Z brings the page back ({pages()})')
    page.get_by_role('button', name='Back to the Greek Office hub').click()
    if page.get_by_role('button', name='Discard and leave').count():
        page.get_by_role('button', name='Discard and leave').click()
    drop(page, 'blank_page.docx')
    page.locator('text=Alpha on page one').first.wait_for(timeout=30000)
    page.wait_for_timeout(2500)
    toast = delete_page(2)
    check(pages() == 2, f'Delete page removes a blank page ({pages()} pages, {toast})')
    with page.expect_download(timeout=20000) as dl:
        page.get_by_role('button', name='Save', exact=True).click()
    bpath = os.path.join(OUT, 'saved-blank.docx')
    dl.value.save_as(bpath)
    bdoc = Document(bpath)
    check([p_.text for p_ in bdoc.paragraphs if p_.text] == ['Alpha on page one.', 'Beta after a blank page.'] and bdoc.element.xml.count('w:type="page"') == 1,
          'the saved file keeps the text and one page break')
    page.get_by_role('button', name='Show paragraph marks and breaks').click()
    check(page.get_by_role('button', name='Show paragraph marks and breaks').get_attribute('aria-pressed') == 'true', 'the ¶ toggle switches on')
    ctx.close()

    # ---------- Abakion (Excel) ----------
    ctx, page = new_page(disable_fsa=True)
    page.goto(BASE, wait_until='networkidle')
    drop(page, 'sample.xlsx')
    page.locator('.hdr__app', has_text='Abakion').wait_for(timeout=20000)
    page.locator('.abakion__grid canvas').first.wait_for(timeout=20000)
    page.wait_for_timeout(3000)
    page.screenshot(path=os.path.join(OUT, 'abakion.png'))
    check(page.locator('.hdr__pill--dirty').count() == 0, 'opening a workbook does not mark it unsaved')
    # Edit: click cell A1 on the main grid canvas (the largest one), then move down to A6.
    boxes = [c.bounding_box() for c in page.locator('.abakion__grid canvas').all()]
    box = max((b for b in boxes if b), key=lambda b: b['width'] * b['height'])
    page.mouse.click(box['x'] + 46 + 40, box['y'] + 20 + 10)
    page.wait_for_timeout(200)
    for _ in range(5):
        page.keyboard.press('ArrowDown')
    page.keyboard.type('Honey')
    page.keyboard.press('Enter')
    page.wait_for_timeout(800)
    check(page.locator('.hdr__pill--dirty').count() == 1, 'Excel edit marks the workbook unsaved')
    with page.expect_download(timeout=20000) as dl:
        page.get_by_role('button', name='Save', exact=True).click()
    path = os.path.join(OUT, 'saved.xlsx')
    dl.value.save_as(path)
    from openpyxl import load_workbook
    wb = load_workbook(path)
    ws = wb['Ledger']
    check(ws['A6'].value == 'Honey', f"saved .xlsx contains the typed value (A6={ws['A6'].value!r})")
    check(str(ws['D5'].value).upper() == '=SUM(D2:D4)', f"formula kept (D5={ws['D5'].value!r})")
    check(str(ws['D2'].value).upper() == '=B2*C2', f"cell formula kept (D2={ws['D2'].value!r})")
    check('A7:D7' in [str(r) for r in ws.merged_cells.ranges], 'merged range kept')
    check(ws.freeze_panes == 'A2', f'frozen pane kept ({ws.freeze_panes})')
    check(abs((ws.column_dimensions['A'].width or 0) - 22) < 1.5, f"column width kept ({ws.column_dimensions['A'].width})")
    check(ws['A1'].font.bold and (ws['A1'].fill.fgColor.rgb or '').upper().endswith('4E5B23'), 'cell style kept (bold + fill)')
    check(ws['D5'].border.top.style == 'double', f'border kept ({ws["D5"].border.top.style})')
    ws2 = wb['Suppliers']
    check(wb.sheetnames == ['Ledger', 'Suppliers'], f'sheet order kept {wb.sheetnames}')
    check(str(ws2['C2'].value).startswith('=_xlfn.CONCAT('), f"new-function prefix restored ({ws2['C2'].value!r})")
    b2 = ws2['B2'].value
    check(hasattr(b2, 'year') and b2.year == 2024, f'date kept ({b2!r}, fmt {ws2["B2"].number_format})')
    # The file must also be readable by a strict zip reader
    zipfile.ZipFile(path).testzip()
    csp = [e for e in page.errors if 'CSP-VIOLATION' in e]
    check(not csp, f'no CSP violations in Abakion {csp[:3]}')
    errs = [e for e in page.errors if 'CSP-VIOLATION' not in e]
    print('   console errors (Abakion):', errs[:5])
    storage = page.evaluate("async () => ({ ls: Object.keys(localStorage), idb: (await indexedDB.databases()).map(d => d.name) })")
    print('   storage after Excel edit:', storage)
    univer_rows = page.evaluate("""() => new Promise(res => {
      const r = indexedDB.open('UniverLocalStorage');
      r.onsuccess = () => { const db = r.result;
        if (!db.objectStoreNames.contains('kv')) { db.close(); return res(0); }
        const q = db.transaction('kv').objectStore('kv').count();
        q.onsuccess = () => { db.close(); res(q.result); }; };
      r.onerror = () => res(-1);
    })""")
    check(univer_rows == 0 and not storage['ls'], f'nothing persisted to browser storage (Univer store rows: {univer_rows})')

    # Leaving with unsaved changes asks first (focus the grid again: Save took focus)
    page.mouse.click(box['x'] + 46 + 40, box['y'] + 20 + 10 + 20 * 8)
    page.keyboard.type('x')
    page.keyboard.press('Enter')
    page.wait_for_timeout(500)
    page.get_by_role('button', name='Back to the Greek Office hub').click()
    check(page.get_by_role('heading', name='Leave without saving?').is_visible(), 'leaving unsaved work asks for confirmation')
    page.get_by_role('button', name='Discard and leave').click()
    page.get_by_role('heading', name='Enter the Greek Office').wait_for(timeout=5000)
    check(True, 'returns to hub after confirming')
    ctx.close()

    # ---------- workbook with parts that cannot be kept ----------
    ctx, page = new_page()
    page.goto(BASE, wait_until='networkidle')
    drop(page, 'chart.xlsx')
    page.locator('.abakion__grid canvas').first.wait_for(timeout=20000)
    page.wait_for_timeout(1500)
    note = page.locator('.editor-notice').inner_text() if page.locator('.editor-notice').count() else ''
    check('charts' in note, f'chart workbook shows a notice ({note[:70]!r}...)')
    check(page.locator('.hdr__chip-btn', has_text='Choose where to autosave').count() == 1, 'chart workbook never autosaves over the original')
    check('(edited)' in page.locator('.hdr__file').inner_text(), 'chart workbook suggests a new file name')
    ctx.close()
    ctx, page = new_page(disable_fsa=True)
    page.goto(BASE, wait_until='networkidle')
    drop(page, 'chart.xlsx')
    page.locator('.abakion__grid canvas').first.wait_for(timeout=20000)
    page.wait_for_timeout(1000)
    with page.expect_download(timeout=20000) as dl:
        page.get_by_role('button', name='Save', exact=True).click()
    cpath = os.path.join(OUT, 'saved-chart.xlsx')
    dl.value.save_as(cpath)
    cws = load_workbook(cpath).active
    check(cws['A2'].value == 'Oil' and cws['B4'].value == 60, 'chart workbook saves its cells to a valid new file')
    check(dl.value.suggested_filename.endswith('(edited).xlsx'), f'saved under a new name ({dl.value.suggested_filename})')
    ctx.close()

    # ---------- hyperlink, hidden empty row, tab colour ----------
    ctx, page = new_page(disable_fsa=True)
    page.goto(BASE, wait_until='networkidle')
    drop(page, 'link.xlsx')
    page.locator('.abakion__grid canvas').first.wait_for(timeout=20000)
    page.wait_for_timeout(1000)
    with page.expect_download(timeout=20000) as dl:
        page.get_by_role('button', name='Save', exact=True).click()
    lpath = os.path.join(OUT, 'saved-link.xlsx')
    dl.value.save_as(lpath)
    lws = load_workbook(lpath).active
    check(lws['A1'].hyperlink is not None and lws['A1'].hyperlink.target == 'https://example.org/agora', 'hyperlink kept')
    check(lws.row_dimensions[3].hidden is True, 'hidden empty row kept')
    check((lws.sheet_properties.tabColor.rgb or '').endswith('A4472A'), 'tab colour kept')
    ctx.close()

    # ---------- Pinax (PowerPoint) ----------
    ctx, page = new_page(disable_fsa=True)
    page.goto(BASE, wait_until='networkidle')
    drop(page, 'sample.pptx')
    page.locator('.hdr__app', has_text='Pinax').wait_for(timeout=20000)
    page.locator('.px-canvas .px-slide').wait_for(timeout=20000)
    page.wait_for_timeout(800)
    check(page.locator('.px-thumb').count() == 5, f"all 5 slides listed ({page.locator('.px-thumb').count()})")
    check(page.locator('.px-thumb--hidden').count() == 1, 'hidden slide is marked')
    check(page.locator('.hdr__pill--dirty').count() == 0, 'opening a presentation does not mark it unsaved')
    page.screenshot(path=os.path.join(OUT, 'pinax.png'))
    # retitle slide 1
    page.locator('.px-canvas [data-top]').first.dblclick()
    page.locator('.px-text-editor .ProseMirror').wait_for(timeout=5000)
    page.keyboard.press('Control+a')
    page.keyboard.type('The Agora, edited')
    page.keyboard.press('Escape')
    page.wait_for_timeout(200)
    check(page.locator('.hdr__pill--dirty').count() == 1, 'text edit marks the presentation unsaved')
    # move the rounded rectangle on slide 3 to the right
    page.locator('.px-thumb__btn').nth(2).click()
    page.wait_for_timeout(300)
    rect = page.locator('.px-canvas [data-top]').nth(1).bounding_box()
    cx, cy = rect['x'] + rect['width'] / 2, rect['y'] + rect['height'] / 2
    page.mouse.move(cx, cy)
    page.mouse.down()
    page.mouse.move(cx + 80, cy + 2, steps=6)
    page.mouse.move(cx + 160, cy + 4, steps=6)
    page.mouse.up()
    # toolbar formatting on the selected shape, then a picture
    page.locator('.px-canvas [data-top]').nth(1).click()
    page.get_by_role('button', name='Bold (Ctrl+B)').click()
    page.get_by_role('button', name='Fill').click()
    page.get_by_role('button', name='#FFC000').click()
    from PIL import Image as _Image
    _Image.new('RGB', (200, 100), (46, 90, 124)).save(os.path.join(OUT, 'insert.png'))
    with page.expect_file_chooser() as fc:
        page.get_by_role('button', name='Picture', exact=True).click()
    fc.value.set_files(os.path.join(OUT, 'insert.png'))
    page.wait_for_timeout(400)
    # a table cell on slide 4
    page.locator('.px-thumb__btn').nth(3).click()
    page.wait_for_timeout(300)
    page.locator('.px-canvas td[data-cell="2,0"]').dblclick()
    page.locator('.px-text-editor .ProseMirror').wait_for(timeout=5000)
    page.keyboard.press('Control+a')
    page.keyboard.type('Honey')
    page.keyboard.press('Escape')
    page.locator('.px-thumb__btn').nth(2).click()
    page.wait_for_timeout(200)
    # a new slide from a layout, with a title
    page.get_by_role('button', name='New slide ▾').click()
    page.get_by_role('button', name='Title and Content').click()
    page.wait_for_timeout(300)
    page.locator('.px-canvas [data-top]').first.dblclick()
    page.locator('.px-text-editor .ProseMirror').wait_for(timeout=5000)
    page.keyboard.type('A brand new slide')
    page.keyboard.press('Escape')
    page.locator('#px-notes-text').fill('Notes on the new slide.')
    # duplicate slide 2, delete the hidden slide (last), notes on slide 1
    page.locator('.px-thumb').nth(1).hover()
    page.get_by_role('button', name='Duplicate slide 2').click()
    n = page.locator('.px-thumb').count()
    page.locator('.px-thumb').nth(n - 1).hover()
    page.get_by_role('button', name=f'Delete slide {n}').click()
    page.locator('.px-thumb__btn').nth(0).click()
    page.wait_for_timeout(200)
    page.locator('#px-notes-text').fill('Opening remarks.')
    # undo / redo round trip on the last notes edit
    page.locator('.px-canvas').click(position={'x': 5, 'y': 5})
    page.keyboard.press('Control+z')
    page.wait_for_timeout(150)
    undone = page.locator('#px-notes-text').input_value()
    page.keyboard.press('Control+y')
    page.wait_for_timeout(150)
    check(undone == '' and page.locator('#px-notes-text').input_value() == 'Opening remarks.', f'undo/redo ({undone!r})')
    with page.expect_download(timeout=30000) as dl:
        page.get_by_role('button', name='Save', exact=True).click()
    ppath = os.path.join(OUT, 'saved.pptx')
    dl.value.save_as(ppath)
    from pptx import Presentation
    prs = Presentation(ppath)
    titles = [s.shapes.title.text if s.shapes.title is not None else '' for s in prs.slides]
    check(titles == ['The Agora, edited', 'What stood there', 'What stood there', 'Shapes and a picture', 'A brand new slide', 'A table and a chart'], f'saved slides in order {titles}')
    notes = [s.notes_slide.notes_text_frame.text if s.has_notes_slide else None for s in prs.slides]
    check(notes[0] == 'Opening remarks.' and notes[1] == notes[2] == 'Speaker notes for slide two.' and notes[4] == 'Notes on the new slide.', f'notes saved {notes}')
    shapes = prs.slides[3]
    rr = next(sh for sh in shapes.shapes if sh.name.startswith('Rounded'))
    check(rr.left > 731520 + 600000, f'moved shape saved at its new position ({rr.left} EMU)')
    check(rr.fill.type == 1 and str(rr.fill.fore_color.rgb) == 'FFC000' and all(r.font.bold for pp in rr.text_frame.paragraphs for r in pp.runs), 'fill colour and bold saved')
    check(sum(1 for sh in shapes.shapes if sh.shape_type == 13) == 2, 'inserted picture saved')
    tbl = next(sh for sh in prs.slides[5].shapes if sh.has_table).table
    check(tbl.cell(2, 0).text == 'Honey' and tbl.cell(2, 1).text == '25', f'table cell edit saved ({tbl.cell(2, 0).text!r})')
    arrow = next(sh for sh in shapes.shapes if sh.name.startswith('Right Arrow'))
    check(abs(arrow.rotation - 15) < 0.01, 'untouched rotated shape keeps its rotation')
    names = zipfile.ZipFile(ppath).namelist()
    check(any(n.startswith('ppt/charts/chart') for n in names), 'chart part kept')
    check(not any(n.endswith('slide5.xml') for n in names if n.startswith('ppt/slides/')), 'deleted slide removed from the package')
    zipfile.ZipFile(ppath).testzip()
    if shutil.which('soffice'):
        r = subprocess.run(['soffice', '--headless', '--convert-to', 'pdf', '--outdir', OUT, ppath], capture_output=True, text=True, timeout=240)
        lo = os.path.join(OUT, 'saved.pdf')
        from pypdf import PdfReader
        check(r.returncode == 0 and os.path.exists(lo) and len(PdfReader(lo).pages) == 6, 'LibreOffice opens the saved presentation (6 slides)')
    else:
        print('   (LibreOffice not installed: skipping the independent validity check)')
    check(page.locator('.hdr__pill--dirty').count() == 0, 'unsaved marker clears after save')
    # PDF export: visible slides only, real text, slide-sized pages
    with page.expect_download(timeout=120000) as dl:
        page.get_by_role('button', name='Download PDF').click()
    pdfpath = os.path.join(OUT, 'pinax.pdf')
    dl.value.save_as(pdfpath)
    from pypdf import PdfReader
    rd = PdfReader(pdfpath)
    texts = [pg_.extract_text() or '' for pg_ in rd.pages]
    check(len(rd.pages) == 6, f'Pinax PDF has one page per slide ({len(rd.pages)})')
    check('The Agora, edited' in texts[0] and 'Goods' in texts[-1], 'Pinax PDF carries real text')
    box = rd.pages[0].mediabox
    check(abs(float(box.width) - 960) < 1 and abs(float(box.height) - 540) < 1, f'PDF page is the slide size ({float(box.width)} x {float(box.height)})')
    # slideshow
    page.get_by_role('button', name='Slideshow').click()
    page.locator('.px-show').wait_for(timeout=5000)
    page.keyboard.press('ArrowRight')
    page.wait_for_timeout(150)
    label = page.locator('.px-show').get_attribute('aria-label') or ''
    page.keyboard.press('Escape')
    page.wait_for_timeout(300)
    check('slide 2 of' in label and page.locator('.px-show').count() == 0, f'slideshow moves and closes ({label})')
    csp = [e for e in page.errors if 'CSP-VIOLATION' in e]
    check(not csp, f'no CSP violations in Pinax {csp[:3]}')
    errs = [e for e in page.errors if 'CSP-VIOLATION' not in e]
    check(not errs, f'no console errors in Pinax {errs[:3]}')
    storage = page.evaluate("async () => ({ ls: Object.keys(localStorage), ss: Object.keys(sessionStorage) })")
    check(not storage['ls'] and not storage['ss'], f'Pinax stores nothing in the browser {storage}')
    ctx.close()

    # a new presentation from the built-in template
    ctx, page = new_page(disable_fsa=True)
    page.goto(BASE, wait_until='networkidle')
    page.get_by_role('button', name='New presentation').click()
    page.locator('.px-canvas .px-slide').wait_for(timeout=20000)
    page.wait_for_timeout(500)
    page.locator('.px-canvas [data-top]').first.dblclick()
    page.locator('.px-text-editor .ProseMirror').wait_for(timeout=5000)
    page.keyboard.type('Pinax works')
    page.keyboard.press('Escape')
    page.get_by_role('button', name='Shape ▾').click()
    page.get_by_role('button', name='Ellipse').click()
    page.get_by_role('button', name='New slide ▾').click()
    page.get_by_role('button', name='Two Content').click()
    page.locator('#px-notes-text').fill('Notes on slide two.')
    with page.expect_download(timeout=30000) as dl:
        page.get_by_role('button', name='Save', exact=True).click()
    npath = os.path.join(OUT, 'new.pptx')
    dl.value.save_as(npath)
    nprs = Presentation(npath)
    check(len(nprs.slides) == 2 and nprs.slides[0].shapes.title.text == 'Pinax works', 'new presentation saves its slides')
    check(nprs.slides[1].slide_layout.name == 'Two Content' and nprs.slides[1].notes_slide.notes_text_frame.text == 'Notes on slide two.', 'new slide keeps its layout and notes')
    if shutil.which('soffice'):
        r = subprocess.run(['soffice', '--headless', '--convert-to', 'pdf', '--outdir', OUT, npath], capture_output=True, text=True, timeout=240)
        check(r.returncode == 0 and os.path.exists(os.path.join(OUT, 'new.pdf')), 'LibreOffice opens the new presentation')
    csp = [e for e in page.errors if 'CSP-VIOLATION' in e]
    check(not csp, f'no CSP violations in a new presentation {csp[:3]}')
    ctx.close()

    # ---------- new blank documents ----------
    ctx, page = new_page()
    page.goto(BASE, wait_until='networkidle')
    page.get_by_role('button', name='New document').click()
    page.locator('.hdr__app', has_text='Papyros').wait_for(timeout=20000)
    page.wait_for_timeout(2500)
    check(page.locator('.hdr__chip-btn', has_text='Choose where to autosave').count() == 1, 'new doc offers to choose an autosave location')
    page.screenshot(path=os.path.join(OUT, 'papyros-new.png'))
    page.get_by_role('button', name='Back to the Greek Office hub').click()
    page.get_by_role('button', name='New workbook').click()
    page.locator('.abakion__grid canvas').first.wait_for(timeout=20000)
    page.wait_for_timeout(1500)
    check(page.locator('.hdr__pill--dirty').count() == 0, 'new workbook starts clean')
    csp = [e for e in page.errors if 'CSP-VIOLATION' in e]
    check(not csp, f'no CSP violations on new docs {csp[:3]}')
    ctx.close()
    browser.close()

print()
print('ALL PASSED' if not failures else f'{len(failures)} FAILED:\n  - ' + '\n  - '.join(failures))
sys.exit(1 if failures else 0)
