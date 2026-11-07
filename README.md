# Greek Office · ΕΛΛΗΝΙΚΟ ΓΡΑΦΕΙΟ

Word (`.docx`), Excel (`.xlsx`) and PowerPoint (`.pptx`) editors that run entirely in the browser, with an
Ancient Greek look.

| | |
|---|---|
| **Papyros** (Πάπυρος) | Word editor: styles, tables, images, headers/footers, page setup, one-click **Download PDF**, Print |
| **Abakion** (Ἀβάκιον) | Excel editor: formulas (worker-thread engine), multiple sheets, formatting, merges, frozen panes, CSV export (displayed values; text that looks like a formula is prefixed with `'` so spreadsheets do not run it) |
| **Pinax** (Πίναξ) | PowerPoint editor: opens and saves real `.pptx`; text, shapes, pictures, table text, slide sorter, speaker notes, full-screen slideshow, **Download PDF** |
| **Hub** | Pick an editor, open a file, or drop one anywhere on the page |

**Documents never leave the browser.** The container is a static file server; files are opened, edited and
saved inside the tab. They live in RAM only: reload or close the tab and they are gone unless saved (you are
warned first). In Chrome/Edge, **autosave writes back to the original file every 5 minutes** when there are
unsaved changes.

---

## Deploy on the Raspberry Pi

Requirements: Docker Engine with Compose v2 (arm64 or armv7 Raspberry Pi OS), and the host nginx that already
holds your TLS certificate.

```bash
# 1. Say who may use it: set GREEK_ALLOWED_IPS in .env (see "IP allow-list" under Security)
cd greek-office
cp .env.example .env

# 2. Build and start (bound to 127.0.0.1:8088 only)
docker compose up -d --build

# 3. Check it
docker compose ps                      # STATUS should become "healthy"
curl -s http://127.0.0.1:8088/healthz  # -> ok
```

```bash
# 4. Host nginx: paste deploy/host-nginx-subpath.conf into your HTTPS server block,
#    change /something/ to the path you want, then:
sudo nginx -t && sudo systemctl reload nginx
```

Open `https://your-domain/something/`. The same build works at any subpath or at the domain root without
rebuilding: all asset URLs are relative, and the host nginx strips the prefix. Requests to `/something` without
the trailing slash are redirected.

Change the local port with `GREEK_PORT=9000 docker compose up -d` (and in the host nginx `proxy_pass`).

> **Upgrading from "Agora Office":** the compose project, image and container are now `greek-office`, and the
> port variable is `GREEK_PORT` (was `AGORA_PORT`). On the Pi, stop the old one first with
> `docker compose down` run from the *old* checkout (or `docker rm -f agora-office`), then
> `docker compose up -d --build`. The default port 8088 and the host nginx snippet are unchanged unless you
> set a custom port.

### Building

- **On the Pi:** `docker compose build` takes a few minutes on a Pi 4. The build stage uses up to ~3 GB of RAM,
  which fits in 8 GB.
- **On a faster PC, for the Pi:**
  ```bash
  docker buildx build --platform linux/arm64 -t greek-office:1.0.0 --load .
  docker save greek-office:1.0.0 | gzip | ssh pi 'gunzip | docker load'
  # then on the Pi: docker compose up -d   (no --build)
  ```
  The Node build stage runs natively on the PC (its output is plain static files), so there is no slow
  emulation.

### Resource limits (docker-compose.yml)

| Limit | Value | Why |
|---|---|---|
| CPU | 1.0 of 4 cores | static files need very little |
| Memory | 128 MB hard, 32 MB reserved | nginx uses roughly 10–20 MB |
| Swap | none (`memswap_limit` = memory) | protects the SD card |
| Processes | 64 | stops fork bombs |
| Open files | 4096 | |
| Logs | json-file, 5 MB × 2 | caps SD-card writes |
| /tmp | 16 MB tmpfs, `noexec,nosuid,nodev` | the only writable path |

> **Raspberry Pi OS:** if `docker info` prints `WARNING: No memory limit support`, the memory cgroup is off.
> Add `cgroup_enable=memory cgroup_memory=1` to the end of the single line in `/boot/firmware/cmdline.txt`
> (older images: `/boot/cmdline.txt`) and reboot. Without it Docker silently ignores the memory limits.

These limits cover the server only. Editing happens in the visitor's browser, so heavy workbooks use the
memory of the PC doing the editing, not the Pi.

---

## Security

**Container**
- Runs as unprivileged user 101 and `cap_drop: ALL` with `no-new-privileges`
- Read-only root filesystem; only a small `noexec` tmpfs at `/tmp`
- Published on `127.0.0.1` only, so it is unreachable from the network except through the host nginx
- The stock start-up scripts that rewrite config are removed, as is `apk` from the runtime image. The only
  start-up step left is `docker/entrypoint.sh`, which writes the allow-list table to `/tmp`
- Multi-stage build: no Node, sources or build tools in the final image; `npm ci --ignore-scripts`
  (no dependency install hooks run); no source maps shipped
- Healthcheck, resource limits and rotated logs

**nginx (container)**
- Serves `GET`/`HEAD` only (anything else returns 405); 1 KB body limit; short timeouts; `server_tokens off`
- Allow-list routing: only `/`, `/index.html`, `/assets/*`, `/favicon.svg` and `/healthz`; everything else,
  including dotfiles, returns 404
- `absolute_redirect off`, so internal host and port never leak behind the proxy
- **IP allow-list** (`GREEK_ALLOWED_IPS` in `.env`): comma-separated IPv4/IPv6 addresses or CIDR ranges that may
  enter; everyone else gets `403`. `all` lets everyone in. It is checked on the real client address (`X-Real-IP`) and
  leaves `/healthz` open. Apply a change with `docker compose up -d`, which recreates the container (no rebuild, a
  brief interruption); `docker compose restart` keeps the old value. Compose refuses to start if the variable is
  missing or empty, and a bad entry stops the container (`docker logs greek-office` says why), so a mistake never
  opens the app. While it restricts, a request that carries no `X-Real-IP` (e.g. `curl` straight at the port) is
  refused too; use `curl -H 'X-Real-IP: <allowed ip>'` to test. Mirror the same addresses on the host nginx so both
  layers agree.
- Per-client rate limits (60 req/s, burst 300, 100 connections). The container restores the real visitor address
  from `X-Real-IP`, trusted only from private Docker bridge ranges. The host nginx must therefore keep
  `proxy_set_header X-Real-IP $remote_addr;` (the sample does); without it all visitors share one bucket
- Headers on every response:
  - a strict **Content-Security-Policy**: `script-src 'self'`, no `eval`, `connect-src 'self'`,
    `object-src 'none'`, `frame-ancestors 'none'`, `base-uri 'none'`, `form-action 'none'`
  - COOP + COEP (the page is cross-origin isolated) and CORP
  - `nosniff`, `X-Frame-Options: DENY`, `Referrer-Policy: no-referrer`, a restrictive `Permissions-Policy`
- `index.html` is `no-store`; hashed assets are `immutable` for a year and pre-gzipped at build time,
  so the Pi spends no CPU compressing

**Host nginx (sample)**: HSTS, `GET`-only, 1 KB body limit, optional per-client rate limit.

**The gate** (`frontend/src/lib/fileGate.ts`): every file is inspected before an editor sees it.
- Only `.docx`/`.xlsx`/`.pptx`, checked by extension **and** content; a Word file renamed `.xlsx` is refused
- Refused: macro-enabled files (`.docm`/`.xlsm`/`.pptm`/`.ppsm`, `vbaProject.bin`), ActiveX, templates
  (`.dotx`/`.xltx`/`.potx`), add-ins (`.ppam`), slide-show files (`.ppsx`), legacy `.doc`/`.xls`/`.ppt`,
  password-protected files, and anything that is not a valid OOXML package
- Pinax re-checks every XML part with the same prolog scan before the browser's XML parser sees it, and checks
  each file it builds (the gate again, every part re-parsed, every relationship target present) **before**
  writing it: if a save would produce a broken file, nothing is written, so autosave cannot damage the original
- Zip-bomb limits: 50 MB file, 1 GB expanded, 20,000 parts, compression ratio over 1000:1 refused
- Path-traversal names refused; any DTD/entity declaration (XXE, billion laughs) before the root element of a
  package XML part is refused however far into the prolog it sits, as are UTF-16/32-encoded parts
- External relationships that make Word or Excel fetch remote content when the saved file is later opened are
  refused: remote templates (`attachedTemplate`), linked OLE objects, sub-documents, frames. Plain external
  hyperlinks and images are allowed; workbook-to-workbook links are allowed but open without a write handle
- Workbooks whose charts, pivot tables, slicers or embedded objects cannot be kept are opened **without** a
  write handle: Save and autosave go to a new file (suggested `… (edited).xlsx`), so the original is never
  damaged

**Memory only**: the app never writes to localStorage, sessionStorage or IndexedDB. Univer's own preference
store is replaced with an in-memory map. It still opens an empty IndexedDB named `UniverLocalStorage` when it
loads, but nothing is written to it (the tests check this).

---

## How saving works

| Browser | Open | Save | Autosave (5 min) |
|---|---|---|---|
| Chrome / Edge | file picker or drag & drop, keeps a handle | writes back to the same file | silent write to the same file |
| Firefox / Safari | file picker or drag & drop | downloads a copy | not possible; the header says so |

- Silent writes need HTTPS, which your host nginx provides.
- Autosave runs only when there are unsaved changes; the header counts down to it only then.
- Renaming a document that was opened from a file makes the next Save a Save As, so the original file is never
  silently overwritten under its old name.
- Edits made while a save is in progress keep the document marked unsaved.
- Chrome asks once per file for write permission. The header then shows **Turn on autosave for this file**;
  your first **Save** also grants it.
- New documents show **Choose where to autosave** until saved once.
- `Ctrl+S` saves, `Ctrl+Shift+S` saves as, and `Ctrl+P` prints in Papyros.
- **Download PDF** (Papyros, Pinax) asks where to save first (Chrome/Edge) or downloads (Firefox/Safari), then
  builds the PDF in the tab.

## PDF export

PDFs are built in the browser with pdf-lib; nothing is uploaded and no server converter is involved.

- **Word:** Papyros lays the document out with docx-editor's own export layer, the same engine and fonts as the
  editor on screen, so page breaks match what you see. Text is real text (selectable, searchable) drawn glyph
  by glyph from the editor's shaping, with the fonts embedded as subsets. Covered: text formatting, lists,
  tables with shading and borders, pictures (crop, rotation), text boxes and simple shapes, headers, footers,
  page numbers, footnotes, page borders, column rules, hyperlinks, internal links. Characters the document's
  font cannot draw (polytonic Greek, for example) use the bundled GFS Didot. Equations, WordArt and EMF/WMF
  pictures with no raster copy are simplified, and Papyros says so in the message after saving; **Print** (the
  browser's print engine) stays available for an exact copy.
- **PowerPoint:** one page per visible slide at the slide's size. Shapes, pictures and tables are drawn as
  vectors; text is placed word by word where the screen put it, with the same font files. Gradients become
  their dominant colour; shadows, 3-D and animations are not drawn.

---

## What is supported, honestly

**Papyros (Word).** Built on docx-editor (Apache-2.0, free packages only). It edits real `.docx` files and,
on save, keeps parts of the file it does not understand. Not included: tracked changes and comments UI (those
are paid add-ons of the library). PDF: see above; Print uses the browser's print dialog with every page at its
real size.

**Pinax (PowerPoint).** Pinax's own editor (React, SVG, ProseMirror for text). It saves by patching the
original file: only what you changed is rewritten, so everything Pinax does not model survives untouched.

| Content | In Pinax | On save |
|---|---|---|
| Text in placeholders, text boxes, shapes, table cells | Edit with bold, italic, underline, strike, size, colour, font, alignment, bullets, numbering, list level | Only edited text is rewritten; run and paragraph settings Pinax does not model are kept, and theme colours stay theme colours unless you change them |
| Common shapes (about 60 presets), custom geometry, lines, connectors, arrowheads | Move, resize, rotate, fill and outline colour, order, duplicate, delete, copy and paste | Only what changed is rewritten |
| Pictures | Move, resize, rotate, delete; insert PNG, JPEG or GIF (10 MB at most) | Kept; new pictures are added to the file |
| Groups, tables | Shown; moved and resized as a whole; table cell text editable (rows and columns cannot be added yet) | Kept |
| Charts, SmartArt, embedded objects, video and audio | Shown as their saved picture or a labelled box; move, resize, delete | Kept untouched |
| Animations, transitions, comments, sections | Not shown; the slideshow runs without animations | Kept untouched |
| Slides | Add from any layout, duplicate, delete, drag to reorder, hide | |
| Speaker notes | Edit (a file with no notes layout keeps notes read-only) | |
| Masters, layouts, themes | Used for display (placeholders, theme colours and fonts, backgrounds) | Kept |

- New presentations start from a built-in 16:9 "Greek Office" theme with six layouts.
- Fonts: Calibri, Cambria, Arial, Times New Roman, Courier New and Century Gothic are drawn with the bundled
  metric-compatible faces; other families map to the closest of those, so a deck looks the same on every
  machine and in its PDF. Uncommon preset shapes draw as their bounding box (the shape itself is saved
  unchanged). Shadows, glow and 3-D effects are not drawn.
- Shortcuts: double-click or Enter edits text, Esc leaves it; Delete; arrows nudge (Shift for larger steps);
  Ctrl+D duplicates; Ctrl+C / X / V; Ctrl+Z / Y; F5 starts the slideshow, Shift+F5 from the current slide.
- Compatibility is tested with python-pptx and LibreOffice (the test suite opens every saved deck with both);
  there is no copy of PowerPoint in the test setup.

**Abakion (Excel).** Built on Univer (Apache-2.0) for the grid and formulas, with ExcelJS (MIT) reading and
writing the file.
- Round-trips: values, formulas (including newer functions such as `XLOOKUP` and `CONCAT`), number formats,
  fonts, fills, borders, alignment, wrap, rotation, merges, column widths, row heights, hidden rows, columns
  and sheets, frozen panes, gridlines, tab colours, sheet order, and hyperlinks.
- Data validation, conditional formatting, named ranges and notes are kept where ExcelJS supports them.
  They are not editable in the UI.
- Charts, pivot tables and slicers are not shown and are dropped on save. The gate warns you and protects
  the original file.

---

## Project layout

```
Dockerfile                    multi-stage, multi-arch image
docker-compose.yml            Pi limits, 127.0.0.1 binding, hardening
.env.example                  template for .env (GREEK_ALLOWED_IPS)
docker/entrypoint.sh          validates GREEK_ALLOWED_IPS and writes the nginx allow-list table
docker/nginx.conf             container nginx (read-only FS, /tmp only)
docker/security-headers.conf  CSP and the other headers
deploy/host-nginx-subpath.conf  snippet for your host nginx
frontend/                     React + Vite app
  src/hub/                    the Greek Office hub
  src/papyros/                Word editor, in-browser printing, Word -> PDF painter (docxPdf.ts)
  src/abakion/                Excel editor, .xlsx <-> Univer conversion, formula worker
  src/pinax/                  PowerPoint editor
    pptx/                     package, theme, reader, writer, built-in template, shape geometry
    render.tsx                one renderer for canvas, thumbnails, slideshow and PDF
    Canvas.tsx, TextEditor.tsx, Sorter.tsx, Slideshow.tsx, pdf.tsx
  src/lib/                    gate, file access, autosave session, config
  src/lib/pdf/                PDF builder (font subsetting, images, paths) shared by both exporters
tests/                        Playwright end-to-end checks and test files
```

## Development

```bash
cd frontend
npm ci
npm run dev        # http://127.0.0.1:5173
npm run build      # type-check + production build into dist/
```

## Tests

Real Chrome via Playwright (Python), against a running instance:

```bash
python3 tests/make_fixtures.py   # test files (needs python-docx, openpyxl, python-pptx, pillow)
python3 tests/e2e.py          https://your-domain/something/   # needs pypdf; uses LibreOffice if installed
python3 tests/pdf_check.py    https://your-domain/something/
python3 tests/print_check.py  https://your-domain/something/
python3 tests/autosave_check.py https://your-domain/something/
```

The suites check the hub, the gate refusing 15 kinds of hostile file, Word, Excel and PowerPoint
open/edit/save round-trips verified by opening the saved files (presentations also with LibreOffice),
chart-workbook protection, Word -> PDF (page count equal to the editor's, text on every page, embedded font
subsets, pictures, footnotes, page numbers), Pinax PDF export and slideshow, Print of a
multi-page document, the 5-minute autosave, the absence of CSP violations, and that nothing is persisted to
browser storage.

## Maintenance

- `docker compose build --pull && docker compose up -d` picks up patched base images.
- Pin base images by digest in the `Dockerfile` for reproducible builds.
- `npm audit --omit=dev` reports 2 *moderate* advisories (0 high or critical), one issue reached through
  ExcelJS: the `uuid` advisory covers v3/v5/v6 with a caller-supplied buffer, while ExcelJS only uses v4.

## Licences

App code Apache-2.0. Bundled: docx-editor (Apache-2.0; fonts OFL/GUST), Univer (Apache-2.0), ExcelJS (MIT),
pdf-lib and @pdf-lib/fontkit (MIT), ProseMirror (MIT), fflate (MIT), React (MIT), Cinzel / Alegreya Sans /
GFS Didot (OFL-1.1). nginx (BSD-2-Clause).
