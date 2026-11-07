/**
 * "Delete page" for Papyros.
 *
 * Word documents have no page objects: a page exists because text, a table, a page break or a
 * section break sits on it. So deleting a page means deleting everything that makes it. The
 * editor publishes, on its painted pages, which paragraph and offset every line covers
 * (`data-paragraph-id`, `data-start`, `data-end`) and which table row a line sits in
 * (`data-table-id`, `data-row-id`). The page's content is read from there and removed with the
 * editor's own commands (`setSelection` + `deleteText`, `deleteRow`), so it is ordinary editing
 * that Ctrl+Z undoes.
 *
 * Text (and any table that lies wholly on the page) goes as one span:
 *  - any page but the last: up to the next page's first position, so the page break or empty
 *    paragraphs that end the page go too and the next page moves up;
 *  - the last page: from just before the break that ends the previous page, so no empty page is
 *    left behind.
 * A table that continues from or onto another page loses only the rows on this page (rows that
 * also show on a neighbouring page, such as repeated header rows, are left alone).
 *
 * Several deletions run back to front, so the positions read up front stay valid.
 */
import type { DocxEditorRef } from '@docx-editor.dev/react';

type Editor = NonNullable<ReturnType<DocxEditorRef['getEditor']>>;

interface Position {
  paragraphId: string;
  offset: number;
}

interface LineInfo {
  el: HTMLElement;
  paragraphId: string;
  tableId: string | null;
  rowId: string | null;
}

export interface DeletePageResult {
  /** The page is gone (the page count dropped). */
  deleted: boolean;
  /** Something was removed, even if the page still exists. */
  changed: boolean;
  message: string;
}

const BODY = '/word/document.xml#';
const PAGE_BREAK = '\f';

function pageElement(root: HTMLElement, index: number): HTMLElement | null {
  return root.querySelector<HTMLElement>(`.docx-page[data-page-index="${index}"]`);
}

/** Body lines of a page in reading order: no headers, footers, notes or text boxes. */
function lines(page: HTMLElement): LineInfo[] {
  return Array.from(page.querySelectorAll<HTMLElement>('.docx-line[data-paragraph-id]'))
    .filter(
      (el) =>
        (el.getAttribute('data-paragraph-id') ?? '').startsWith(BODY) &&
        !el.closest('[data-docx-hf], [data-docx-notes], [data-textbox-paragraph-id]'),
    )
    .map((el) => ({
      el,
      paragraphId: el.getAttribute('data-paragraph-id')!,
      tableId: el.closest('[data-table-id]')?.getAttribute('data-table-id') ?? null,
      rowId: el.closest('[data-row-id]')?.getAttribute('data-row-id') ?? null,
    }));
}

function runs(line: HTMLElement): { start: number; end: number; text: string }[] {
  const out: { start: number; end: number; text: string }[] = [];
  for (const el of Array.from(line.querySelectorAll<HTMLElement>('[data-start]'))) {
    const start = Number(el.getAttribute('data-start'));
    const end = Number(el.getAttribute('data-end') ?? start);
    if (Number.isFinite(start)) out.push({ start, end: Number.isFinite(end) ? end : start, text: el.textContent ?? '' });
  }
  for (const el of Array.from(line.querySelectorAll<HTMLElement>('[data-drawing-start]'))) {
    const start = Number(el.getAttribute('data-drawing-start'));
    if (Number.isFinite(start)) out.push({ start, end: start + 1, text: '' });
  }
  return out.sort((a, b) => a.start - b.start);
}

function startOf(line: LineInfo): Position {
  const r = runs(line.el);
  return { paragraphId: line.paragraphId, offset: r.length ? r[0].start : 0 };
}

/** The end of a line; with `beforeBreak`, in front of a page break that ends it. */
function endOf(line: LineInfo, beforeBreak: boolean): Position {
  const r = runs(line.el);
  const last = r[r.length - 1];
  let offset = last ? last.end : 0;
  if (beforeBreak && last && last.text === PAGE_BREAK) offset = last.start;
  return { paragraphId: line.paragraphId, offset };
}

const frame = () => new Promise<void>((r) => requestAnimationFrame(() => r()));

/** Scroll page `index` (0-based) into view and wait until the editor has painted it. */
async function painted(editor: Editor, root: HTMLElement, index: number): Promise<HTMLElement> {
  editor.scrollToPage(index + 1);
  const deadline = performance.now() + 4000;
  while (performance.now() < deadline) {
    await frame();
    const page = pageElement(root, index);
    if (page && page.getAttribute('data-materialized') !== 'false' && lines(page).length) return page;
  }
  throw new Error(`page ${index + 1} could not be read`);
}

const same = (a: Position, b: Position) => a.paragraphId === b.paragraphId && a.offset === b.offset;

export async function deletePage(editor: Editor, root: HTMLElement, pageNumber: number): Promise<DeletePageResult> {
  const total = editor.getTotalPages();
  if (pageNumber < 1 || pageNumber > total) return { deleted: false, changed: false, message: `There is no page ${pageNumber}.` };
  const index = pageNumber - 1;

  const prev = pageNumber > 1 ? lines(await painted(editor, root, index - 1)) : [];
  const next = pageNumber < total ? lines(await painted(editor, root, index + 1)) : [];
  const here = lines(await painted(editor, root, index));
  if (!here.length) return { deleted: false, changed: false, message: `Page ${pageNumber} has nothing that can be selected.` };

  // Tables that also show on a neighbouring page are cut row by row; rows that show on a
  // neighbouring page too (repeated headers, rows split across the break) are left alone.
  const neighbourTables = new Set([...prev, ...next].map((l) => l.tableId).filter(Boolean));
  const neighbourRows = new Set([...prev, ...next].map((l) => l.rowId).filter(Boolean));
  const partial = (l: LineInfo) => !!l.tableId && neighbourTables.has(l.tableId);
  const text = here.filter((l) => !partial(l));
  const leading = partial(here[0]);
  const trailing = partial(here[here.length - 1]);

  const rowsOf = (part: LineInfo[]) => {
    const seen = new Map<string, Position>();
    for (const l of part) if (l.rowId && !neighbourRows.has(l.rowId) && !seen.has(l.rowId)) seen.set(l.rowId, { paragraphId: l.paragraphId, offset: 0 });
    return [...seen.values()];
  };
  const firstText = text.length ? here.indexOf(text[0]) : here.length;
  const lastText = text.length ? here.indexOf(text[text.length - 1]) : -1;
  const leadingRows = leading ? rowsOf(here.slice(0, firstText)) : [];
  const trailingRows = trailing ? rowsOf(here.slice(lastText + 1)) : [];

  // The text span, widest first; narrower fallbacks for spans the editor will not take.
  const spans: [Position, Position][] = [];
  if (text.length) {
    const start = startOf(text[0]);
    const end = endOf(text[text.length - 1], false);
    const nextStart = next.length && !partial(next[0]) ? startOf(next[0]) : null;
    const prevEnd = prev.length && !prev[prev.length - 1].tableId ? endOf(prev[prev.length - 1], true) : null;
    if (!trailing && nextStart) spans.push([start, nextStart]);
    else if (!trailing && !leading && !next.length && prevEnd) spans.push([prevEnd, end]);
    spans.push([start, end]);
  }

  let changed = false;
  let refusal = '';
  const exec = (cmd: Parameters<Editor['exec']>[0]) => {
    const r = editor.exec(cmd);
    if (r.ok) changed ||= r.changed;
    else refusal = r.reason;
    return r.ok;
  };
  const deleteRows = (rows: Position[]) => {
    for (const at of [...rows].reverse()) {
      if (exec({ type: 'setSelection', range: { anchor: at, head: at } })) exec({ type: 'deleteRow' });
    }
  };

  // Back to front: rows after the text, the text, rows before it.
  deleteRows(trailingRows);
  for (const [anchor, head] of spans) {
    if (same(anchor, head)) continue;
    if (exec({ type: 'setSelection', range: { anchor, head } }) && exec({ type: 'deleteText' })) break;
  }
  deleteRows(leadingRows);

  // Let the editor lay the document out again, then see whether the page is really gone.
  for (let i = 0; i < 6; i++) await frame();
  const after = editor.getTotalPages();
  editor.scrollToPage(Math.min(pageNumber, after));
  if (after < total) return { deleted: true, changed: true, message: `Page ${pageNumber} deleted. Ctrl+Z brings it back.` };
  if (changed) {
    return {
      deleted: false,
      changed: true,
      message: `Page ${pageNumber} was emptied as far as possible, but a page is still there (usually a table that runs across the page break, or the paragraph Word keeps after a table). Ctrl+Z undoes this.`,
    };
  }
  return { deleted: false, changed: false, message: `Page ${pageNumber} was not deleted: ${refusal || 'the editor refused the edit.'}` };
}
