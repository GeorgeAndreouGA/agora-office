/**
 * Print / Save as PDF for Papyros, done entirely in the browser.
 *
 * The editor only keeps pages near the viewport in the DOM, so a plain window.print() would
 * print blank pages. Instead each page is scrolled into view, allowed to render, cloned into a
 * print-only layer, and then the browser's print dialog (which also offers "Save as PDF") opens.
 */

const nextFrame = () => new Promise<void>((r) => requestAnimationFrame(() => r()));
const sleep = (ms: number) => new Promise<void>((r) => setTimeout(r, ms));

function scrollParent(el: HTMLElement): HTMLElement | null {
  let node: HTMLElement | null = el.parentElement;
  while (node) {
    const style = getComputedStyle(node);
    if (/(auto|scroll)/.test(style.overflowY) && node.scrollHeight > node.clientHeight) return node;
    node = node.parentElement;
  }
  return null;
}

function isRendered(page: Element): boolean {
  const content = page.querySelector('.docx-page-content');
  return !!content && content.childElementCount > 0;
}

async function renderPage(root: HTMLElement, index: number): Promise<HTMLElement | null> {
  let page = root.querySelectorAll<HTMLElement>('.docx-page')[index];
  if (!page) return null;
  page.scrollIntoView({ block: 'start', inline: 'nearest', behavior: 'instant' as ScrollBehavior });
  const deadline = performance.now() + 1500;
  while (performance.now() < deadline) {
    await nextFrame();
    page = root.querySelectorAll<HTMLElement>('.docx-page')[index] ?? page;
    if (isRendered(page)) break;
  }
  // One more frame so images and late runs settle.
  await nextFrame();
  await sleep(30);
  return root.querySelectorAll<HTMLElement>('.docx-page')[index] ?? page;
}

function clonePage(page: HTMLElement): HTMLElement {
  const clone = page.cloneNode(true) as HTMLElement;
  // Canvases do not carry their pixels when cloned; copy them as images.
  const src = page.querySelectorAll('canvas');
  const dst = clone.querySelectorAll('canvas');
  src.forEach((c, i) => {
    try {
      const img = document.createElement('img');
      img.src = c.toDataURL('image/png');
      img.style.cssText = c.style.cssText;
      img.width = c.width;
      img.height = c.height;
      dst[i]?.replaceWith(img);
    } catch {
      /* a tainted canvas stays blank rather than failing the print */
    }
  });
  // Editing furniture never prints.
  clone
    .querySelectorAll('.docx-hf-edit-hint, .docx-hf--placeholder, .docx-hf-hover, [contenteditable]')
    .forEach((n) => {
      if (n.hasAttribute('contenteditable')) n.removeAttribute('contenteditable');
      else n.remove();
    });
  clone.style.position = 'relative';
  clone.style.top = '0';
  clone.style.left = '0';
  clone.style.transform = 'none';
  clone.style.boxShadow = 'none';
  clone.style.margin = '0';
  return clone;
}

export interface PrintProgress {
  (done: number, total: number): void;
}

export async function printDocument(editorRoot: HTMLElement, onProgress?: PrintProgress): Promise<void> {
  const pages = editorRoot.querySelectorAll<HTMLElement>('.docx-page');
  if (!pages.length) throw new Error('There is nothing to print yet.');
  const scroller = scrollParent(pages[0]);
  const restoreTop = scroller?.scrollTop ?? 0;

  const host = document.createElement('div');
  host.className = 'greek-print-root docx-editor greek-docx';
  host.setAttribute('aria-hidden', 'true');
  const sheet = document.createElement('div');
  sheet.className = 'docx-pages';
  host.appendChild(sheet);

  const sizes = new Map<string, { w: number; h: number }>();
  const total = pages.length;
  try {
    for (let i = 0; i < total; i++) {
      const page = await renderPage(editorRoot, i);
      if (!page) continue;
      const w = Math.round(parseFloat(page.style.width) || page.offsetWidth);
      const h = Math.round(parseFloat(page.style.height) || page.offsetHeight);
      const name = `greek-p-${w}x${h}`;
      sizes.set(name, { w, h });
      const wrap = document.createElement('div');
      wrap.className = 'greek-print-page';
      wrap.style.width = `${w}px`;
      wrap.style.height = `${h}px`;
      wrap.style.setProperty('page', name);
      wrap.appendChild(clonePage(page));
      sheet.appendChild(wrap);
      onProgress?.(i + 1, total);
    }
  } finally {
    if (scroller) scroller.scrollTop = restoreTop;
  }

  const style = document.createElement('style');
  style.textContent = [...sizes.entries()]
    .map(([name, s]) => `@page ${name} { size: ${s.w}px ${s.h}px; margin: 0; }`)
    .concat(
      sizes.size ? [`@page { size: ${[...sizes.values()][0].w}px ${[...sizes.values()][0].h}px; margin: 0; }`] : [],
    )
    .join('\n');

  document.head.appendChild(style);
  document.body.appendChild(host);
  const cleanup = () => {
    host.remove();
    style.remove();
    window.removeEventListener('afterprint', cleanup);
  };
  window.addEventListener('afterprint', cleanup);
  try {
    await document.fonts.ready;
    await nextFrame();
    window.print();
  } finally {
    // afterprint normally cleans up; this is only a backstop for browsers that never fire it.
    setTimeout(cleanup, 120_000);
  }
}
