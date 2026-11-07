"""Autosave check: open via the File System Access path (faked handle), edit, fast-forward 5 minutes,
confirm the file was written back silently, and that nothing is written when there are no changes."""
import base64, io, os, sys, zipfile
from playwright.sync_api import sync_playwright

BASE = sys.argv[1] if len(sys.argv) > 1 else 'http://127.0.0.1:8090/something/'
HERE = os.path.dirname(os.path.abspath(__file__))
xlsx_b64 = base64.b64encode(open(os.path.join(HERE, 'fixtures/sample.xlsx'), 'rb').read()).decode()

FAKE_FSA = """
(b64) => {
  window.__writes = [];
  const bytes = Uint8Array.from(atob(b64), c => c.charCodeAt(0));
  const handle = {
    kind: 'file', name: 'ledger.xlsx',
    async getFile() { return new File([bytes], 'ledger.xlsx'); },
    async queryPermission() { return 'granted'; },
    async requestPermission() { return 'granted'; },
    async createWritable() {
      const chunks = [];
      return {
        async write(blob) { chunks.push(new Uint8Array(await blob.arrayBuffer())); },
        async close() { const n = chunks.reduce((a, c) => a + c.length, 0); window.__writes.push({ at: Date.now(), size: n,
          head: Array.from(chunks[0].slice(0, 4)) }); },
        async abort() {},
      };
    },
  };
  window.showOpenFilePicker = async () => [handle];
  window.showSaveFilePicker = async () => handle;
}
"""
ok = True
def check(c, m):
    global ok
    print(('PASS ' if c else 'FAIL ') + m)
    ok &= bool(c)

with sync_playwright() as p:
    b = p.chromium.launch(args=['--no-sandbox'])
    pg = b.new_page(viewport={'width': 1440, 'height': 900})
    pg.clock.install()  # before load, so the app's autosave timer is on the controllable clock
    pg.goto(BASE, wait_until='networkidle')
    pg.evaluate(FAKE_FSA, xlsx_b64)
    pg.get_by_role('button', name='Open .xlsx').click()
    pg.locator('.abakion__grid canvas').first.wait_for(timeout=30000)
    pg.wait_for_timeout(2500)
    print('   loaded', flush=True)
    check(pg.locator('.hdr__chip-btn').count() == 0, 'autosave is armed (permission granted, no prompt chip)')
    # Nothing to save yet: the header says autosave is on, with no countdown to a write that will not happen.
    check(pg.locator('.hdr__autosave').inner_text().strip() == 'Autosave on',
          f"clean document shows 'Autosave on', no countdown: {pg.locator('.hdr__autosave').inner_text()!r}")
    # No changes: a full interval must not write
    pg.clock.fast_forward(5 * 60 * 1000 + 1000)
    check(pg.evaluate('window.__writes.length') == 0, 'no write when nothing changed')
    # Edit a cell
    boxes = [c.bounding_box() for c in pg.locator('.abakion__grid canvas').all()]
    box = max((x for x in boxes if x), key=lambda x: x['width'] * x['height'])
    pg.mouse.click(box['x'] + 46 + 40, box['y'] + 20 + 10 + 20 * 9)
    pg.keyboard.type('autosaved')
    pg.keyboard.press('Enter')
    pg.wait_for_timeout(1000)
    check(pg.locator('.hdr__pill--dirty').count() == 1, 'edit marks unsaved')
    check('autosave in' in pg.locator('.hdr__autosave').inner_text(),
          f"countdown shown once there are unsaved changes: {pg.locator('.hdr__autosave').inner_text()!r}")
    pg.clock.fast_forward(5 * 60 * 1000 + 1000)
    for _ in range(50):
        if pg.evaluate('window.__writes.length') > 0:
            break
        pg.wait_for_timeout(200)
    writes = pg.evaluate('window.__writes')
    check(len(writes) == 1 and writes[0]['head'] == [0x50, 0x4B, 0x03, 0x04], f'autosave wrote a valid .xlsx back to the file once ({writes})')
    pg.wait_for_timeout(1000)
    check(pg.locator('.hdr__pill--dirty').count() == 0, 'unsaved marker cleared by autosave')
    check('All changes saved' in pg.locator('.hdr__autosave').inner_text(), f"status shows last save: {pg.locator('.hdr__autosave').inner_text()!r}")
    b.close()
print('ALL PASSED' if ok else 'FAILED')
sys.exit(0 if ok else 1)
