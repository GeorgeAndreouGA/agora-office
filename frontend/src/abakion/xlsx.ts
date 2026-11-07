/**
 * .xlsx <-> Univer conversion, done entirely in the browser with ExcelJS.
 *
 * Import maps values, formulas, styles, merges, sizes, frozen panes, hidden rows/columns/sheets
 * and tab colours into Univer's workbook model.
 *
 * Export writes back INTO the workbook that was opened, so parts Univer does not edit but
 * ExcelJS understands (data validation, conditional formatting, defined names, comments,
 * images) are kept where ExcelJS can keep them.
 */
import * as ExcelJS from 'exceljs';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { numfmt, type FUniver, type ICellData, type IStyleData, type IWorkbookData, type IWorksheetData } from '@univerjs/presets';
import { parseThemePalette, resolveColor, toArgb, type ExcelColor } from './colors';

// ---------- units ----------
const DEFAULT_COL_CHARS = 8.43;
const DEFAULT_ROW_PT = 15;
const charsToPx = (w: number) => Math.round(w * 7 + 5);
const pxToChars = (px: number) => Math.max(0, Math.round(((px - 5) / 7) * 100) / 100);
const ptToPx = (pt: number) => Math.round((pt * 4) / 3);
const pxToPt = (px: number) => Math.round(px * 0.75 * 100) / 100;

// Excel stores dates as days since 1899-12-30 (1900 date system). Univer uses the same system.
// A workbook saved in the 1904 system (date1904, mostly old Mac files) is normalised on import:
// ExcelJS already turns its serials into real JS Dates using the 1904 epoch, so converting those
// Dates with the 1900 epoch here yields the right calendar date in Univer's system. On export
// univerToXlsx therefore clears date1904, otherwise Excel would read our 1900 serials as 1904
// ones and every date would move forward by 1462 days (four years and a day).
const EXCEL_EPOCH_MS = Date.UTC(1899, 11, 30);
const dateToSerial = (d: Date) => (d.getTime() - EXCEL_EPOCH_MS) / 86_400_000;

// ---------- enums (numeric values from @univerjs/core) ----------
const BORDER_TO_UNIVER: Record<string, number> = {
  thin: 1, hair: 2, dotted: 3, dashed: 4, dashDot: 5, dashDotDot: 6, double: 7,
  medium: 8, mediumDashed: 9, mediumDashDot: 10, mediumDashDotDot: 11, slantDashDot: 12, thick: 13,
};
const BORDER_FROM_UNIVER = Object.fromEntries(
  Object.entries(BORDER_TO_UNIVER).map(([k, v]) => [v, k]),
) as Record<number, ExcelJS.BorderStyle>;

const H_TO_UNIVER: Record<string, number> = {
  left: 1, center: 2, right: 3, justify: 4, distributed: 6, fill: 1, centerContinuous: 2,
};
const H_FROM_UNIVER: Record<number, ExcelJS.Alignment['horizontal']> = {
  1: 'left', 2: 'center', 3: 'right', 4: 'justify', 5: 'justify', 6: 'distributed',
};
const V_TO_UNIVER: Record<string, number> = { top: 1, middle: 2, bottom: 3, justify: 2, distributed: 2 };
const V_FROM_UNIVER: Record<number, ExcelJS.Alignment['vertical']> = { 1: 'top', 2: 'middle', 3: 'bottom' };

const CELL_STRING = 1;
const CELL_NUMBER = 2;
const CELL_BOOLEAN = 3;

// Functions Excel stores with a _xlfn. prefix (added after Excel 2007).
const FUTURE_FUNCTIONS = [
  'CONCAT', 'TEXTJOIN', 'IFS', 'SWITCH', 'MAXIFS', 'MINIFS', 'XLOOKUP', 'XMATCH', 'UNIQUE',
  'SEQUENCE', 'RANDARRAY', 'SORTBY', 'LET', 'IFNA', 'XOR', 'DAYS', 'ISOWEEKNUM', 'AGGREGATE',
  'NUMBERVALUE', 'STDEV.S', 'STDEV.P', 'VAR.S', 'VAR.P', 'PERCENTILE.INC', 'PERCENTILE.EXC',
  'QUARTILE.INC', 'QUARTILE.EXC', 'RANK.EQ', 'RANK.AVG', 'MODE.SNGL', 'MODE.MULT', 'NORM.DIST',
  'NORM.S.DIST', 'NORM.INV', 'NORM.S.INV', 'CEILING.MATH', 'FLOOR.MATH', 'FORECAST.LINEAR',
  'CONFIDENCE.NORM', 'CONFIDENCE.T', 'T.DIST', 'T.INV', 'CHISQ.DIST', 'BINOM.DIST', 'EXPON.DIST',
  'GAMMA.DIST', 'LOGNORM.DIST', 'POISSON.DIST', 'WEIBULL.DIST', 'COVARIANCE.P', 'COVARIANCE.S',
  'BITAND', 'BITOR', 'BITXOR', 'IMAGE', 'TEXTBEFORE', 'TEXTAFTER', 'TEXTSPLIT', 'VSTACK', 'HSTACK',
  'TOCOL', 'TOROW', 'TAKE', 'DROP', 'CHOOSECOLS', 'CHOOSEROWS', 'EXPAND', 'WRAPROWS', 'WRAPCOLS',
];
const WORKSHEET_FUNCTIONS = ['FILTER', 'SORT'];
const escapeRe = (s: string) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
const FUTURE_RE = new RegExp(`(?<![A-Za-z0-9_.])(${FUTURE_FUNCTIONS.map(escapeRe).join('|')})\\(`, 'gi');
const WORKSHEET_RE = new RegExp(`(?<![A-Za-z0-9_.])(${WORKSHEET_FUNCTIONS.join('|')})\\(`, 'gi');

/** Excel file formula -> Univer formula ("=..." without storage prefixes). */
function formulaIn(formula: string): string {
  return `=${formula.replace(/_xlfn\.|_xlws\./gi, '')}`;
}

/** Univer formula -> Excel file formula (no "=", with the prefixes Excel needs). */
function formulaOut(formula: string): string {
  const body = formula.startsWith('=') ? formula.slice(1) : formula;
  return body
    .replace(FUTURE_RE, (_m, fn: string) => `_xlfn.${fn.toUpperCase()}(`)
    .replace(WORKSHEET_RE, (_m, fn: string) => `_xlfn._xlws.${fn.toUpperCase()}(`);
}

// Allow-list for hyperlink targets written to the file. "#..." is an in-workbook link, which ExcelJS
// writes as a location (documented form: '#\'Sheet2\'!A1'). Anything else (javascript:, data:,
// vbscript:, file:, ...) is dropped and the cell keeps just its text.
// Limitation: ExcelJS only reads external (r:id) hyperlinks, so an in-workbook link that came from
// a loaded file is not imported and cannot round-trip; only links already held by the editor can.
const SAFE_LINK_RE = /^(?:https?:|mailto:|tel:|ftp:|#)/i;

function safeSheetName(name: string, taken: Set<string>): string {
  // Excel rejects names that begin or end with an apostrophe, and "History" is reserved for the
  // change-tracking sheet. Strip the apostrophes after truncating, since slicing can expose a new one.
  const base = (name || 'Sheet').replace(/[[\]:*?/\\]/g, ' ').trim().slice(0, 31).replace(/^'+|'+$/g, '').trim() || 'Sheet';
  // Marking it taken makes the loop below pick "History (2)".
  if (base.toLowerCase() === 'history') taken.add('history');
  let candidate = base;
  let n = 2;
  while (taken.has(candidate.toLowerCase())) {
    const suffix = ` (${n++})`;
    candidate = `${base.slice(0, 31 - suffix.length)}${suffix}`;
  }
  taken.add(candidate.toLowerCase());
  return candidate;
}

// ---------- reading ----------

async function loadExcel(bytes: Uint8Array): Promise<ExcelJS.Workbook> {
  const wb = new ExcelJS.Workbook();
  const buffer = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
  await wb.xlsx.load(buffer);
  return wb;
}

function relativePath(fromDir: string, target: string): string {
  const from = fromDir.split('/').filter(Boolean);
  const to = target.replace(/^\/+/, '').split('/').filter(Boolean);
  let i = 0;
  while (i < from.length && i < to.length - 1 && from[i] === to[i]) i++;
  return [...Array(from.length - i).fill('..'), ...to.slice(i)].join('/');
}

/**
 * Relationship targets may be absolute ("/xl/drawings/drawing1.xml", written by some tools).
 * That is valid, but ExcelJS only resolves relative ones, so rewrite them before loading.
 */
function withRelativeRels(bytes: Uint8Array): Uint8Array | null {
  const parts = unzipSync(bytes);
  let changed = false;
  for (const name of Object.keys(parts)) {
    const m = /^(.*?)_rels\/[^/]+\.rels$/.exec(name);
    if (!m) continue;
    const baseDir = m[1].replace(/\/$/, '');
    const xml = strFromU8(parts[name]);
    const fixed = xml.replace(/Target="(\/[^"]+)"/g, (_all, target: string) => {
      changed = true;
      return `Target="${relativePath(baseDir, target)}"`;
    });
    parts[name] = strToU8(fixed);
  }
  return changed ? zipSync(parts, { level: 6 }) : null;
}

/** Last resort: leave out the drawing layer (charts, pictures, shapes) so the cells still open. */
function withoutDrawings(bytes: Uint8Array): Uint8Array {
  const parts = unzipSync(bytes);
  for (const name of Object.keys(parts)) {
    if (/^xl\/(drawings|charts)\//i.test(name)) delete parts[name];
    else if (/^xl\/worksheets\/sheet[^/]*\.xml$/i.test(name)) {
      parts[name] = strToU8(strFromU8(parts[name]).replace(/<drawing\b[^>]*\/>/g, ''));
    }
  }
  return zipSync(parts, { level: 6 });
}

export interface ReadResult {
  workbook: ExcelJS.Workbook;
  /** True when the drawing layer (charts, pictures, shapes) had to be left out to open the file. */
  droppedDrawings: boolean;
}

export async function readXlsx(bytes: Uint8Array): Promise<ReadResult> {
  let first: unknown;
  try {
    return { workbook: await loadExcel(bytes), droppedDrawings: false };
  } catch (err) {
    first = err;
  }
  const relative = withRelativeRels(bytes);
  if (relative) {
    try {
      return { workbook: await loadExcel(relative), droppedDrawings: false };
    } catch {
      /* fall through */
    }
  }
  try {
    return { workbook: await loadExcel(withoutDrawings(relative ?? bytes)), droppedDrawings: true };
  } catch {
    throw first;
  }
}

function styleToUniver(style: Partial<ExcelJS.Style> | undefined, theme: string[]): IStyleData | null {
  if (!style) return null;
  const out: IStyleData = {};
  const f = style.font;
  if (f) {
    if (f.name) out.ff = f.name;
    if (f.size) out.fs = f.size;
    if (f.bold) out.bl = 1;
    if (f.italic) out.it = 1;
    if (f.underline && f.underline !== 'none') out.ul = { s: 1 };
    if (f.strike) out.st = { s: 1 };
    const color = resolveColor(f.color as ExcelColor, theme);
    if (color && color !== '#000000') out.cl = { rgb: color };
    if (f.vertAlign === 'superscript') out.va = 3;
    if (f.vertAlign === 'subscript') out.va = 2;
  }
  const fill = style.fill;
  if (fill?.type === 'pattern' && fill.pattern && fill.pattern !== 'none') {
    const color = resolveColor(fill.fgColor as ExcelColor, theme) ?? resolveColor(fill.bgColor as ExcelColor, theme);
    if (color) out.bg = { rgb: color };
  } else if (fill?.type === 'gradient' && fill.stops?.length) {
    const color = resolveColor(fill.stops[0].color as ExcelColor, theme);
    if (color) out.bg = { rgb: color };
  }
  const b = style.border;
  if (b) {
    const side = (s: Partial<ExcelJS.Border> | undefined) =>
      s?.style && BORDER_TO_UNIVER[s.style]
        ? { s: BORDER_TO_UNIVER[s.style], cl: { rgb: resolveColor(s.color as ExcelColor, theme) ?? '#000000' } }
        : undefined;
    const bd: NonNullable<IStyleData['bd']> = {};
    const t = side(b.top), r = side(b.right), bo = side(b.bottom), l = side(b.left);
    if (t) bd.t = t;
    if (r) bd.r = r;
    if (bo) bd.b = bo;
    if (l) bd.l = l;
    const diag = b.diagonal as (Partial<ExcelJS.BorderDiagonal> | undefined);
    const d = side(diag);
    if (d && diag?.up) bd.bl_tr = d;
    if (d && diag?.down) bd.tl_br = d;
    if (Object.keys(bd).length) out.bd = bd;
  }
  const a = style.alignment;
  if (a) {
    if (a.horizontal && H_TO_UNIVER[a.horizontal]) out.ht = H_TO_UNIVER[a.horizontal];
    if (a.vertical && V_TO_UNIVER[a.vertical]) out.vt = V_TO_UNIVER[a.vertical];
    if (a.wrapText) out.tb = 3;
    const rot = a.textRotation;
    if (rot === 'vertical' || rot === 255) out.tr = { a: 0, v: 1 };
    else if (typeof rot === 'number' && rot !== 0) out.tr = { a: rot <= 90 ? rot : -(rot - 90) };
  }
  if (style.numFmt && style.numFmt !== 'General') out.n = { pattern: style.numFmt };
  return Object.keys(out).length ? out : null;
}

function resultToUniver(result: unknown): Pick<ICellData, 'v' | 't'> {
  if (result === null || result === undefined) return {};
  if (result instanceof Date) return { v: dateToSerial(result), t: CELL_NUMBER };
  if (typeof result === 'number') return { v: result, t: CELL_NUMBER };
  if (typeof result === 'boolean') return { v: result, t: CELL_BOOLEAN };
  if (typeof result === 'object' && 'error' in (result as object)) {
    return { v: String((result as { error: string }).error), t: CELL_STRING };
  }
  return { v: String(result), t: CELL_STRING };
}

function cellToUniver(cell: ExcelJS.Cell): ICellData | null {
  const V = ExcelJS.ValueType;
  const value = cell.value as unknown;
  switch (cell.type) {
    case V.Null:
    case V.Merge:
      return null;
    case V.Number:
      return { v: value as number, t: CELL_NUMBER };
    case V.String:
    case V.SharedString:
      return { v: String(value), t: CELL_STRING };
    case V.Boolean:
      return { v: !!value, t: CELL_BOOLEAN };
    case V.Date:
      return { v: dateToSerial(value as Date), t: CELL_NUMBER };
    case V.Hyperlink: {
      const h = value as ExcelJS.CellHyperlinkValue;
      const text = typeof h.text === 'string' ? h.text : (h.text as unknown as ExcelJS.CellRichTextValue)?.richText?.map((r) => r.text).join('') ?? '';
      return { v: text, t: CELL_STRING, custom: { hl: h.hyperlink } };
    }
    case V.RichText:
      return { v: (value as ExcelJS.CellRichTextValue).richText.map((r) => r.text).join(''), t: CELL_STRING };
    case V.Error:
      return { v: String((value as ExcelJS.CellErrorValue).error), t: CELL_STRING };
    case V.Formula: {
      const formula = cell.formula;
      const fv = value as { result?: unknown };
      if (!formula) return resultToUniver(fv?.result);
      return { f: formulaIn(formula), ...resultToUniver(fv?.result) };
    }
    default:
      return value === null || value === undefined ? null : { v: String(value), t: CELL_STRING };
  }
}

export interface ImportResult {
  data: Partial<IWorkbookData>;
  /** Univer sheet id -> ExcelJS worksheet id, so saving writes back into the same sheet. */
  idMap: Map<string, number>;
}

export function workbookToUniver(wb: ExcelJS.Workbook, name: string, locale: string): ImportResult {
  const theme = parseThemePalette((wb as unknown as { _themes?: Record<string, string> })._themes?.theme1);
  const styles: Record<string, IStyleData> = {};
  const styleIds = new Map<string, string>();
  const styleIdFor = (s: IStyleData | null): string | undefined => {
    if (!s) return undefined;
    const key = JSON.stringify(s);
    let id = styleIds.get(key);
    if (!id) {
      id = `s${styleIds.size + 1}`;
      styleIds.set(key, id);
      styles[id] = s;
    }
    return id;
  };
  // Many cells share one ExcelJS style object; convert each object once.
  const styleCache = new WeakMap<object, string | undefined>();
  const styleOf = (style: Partial<ExcelJS.Style> | undefined) => {
    if (!style || typeof style !== 'object') return undefined;
    if (styleCache.has(style)) return styleCache.get(style);
    const id = styleIdFor(styleToUniver(style, theme));
    styleCache.set(style, id);
    return id;
  };

  const sheets: Record<string, Partial<IWorksheetData>> = {};
  const sheetOrder: string[] = [];
  const idMap = new Map<string, number>();

  wb.worksheets.forEach((ws) => {
    const id = `sheet-${ws.id}`;
    idMap.set(id, ws.id);
    sheetOrder.push(id);

    const cellData: Record<number, Record<number, ICellData>> = {};
    const rowData: Record<number, { h?: number; hd?: number }> = {};
    const columnData: Record<number, { w?: number; hd?: number }> = {};
    const defaultRowPt = ws.properties.defaultRowHeight || DEFAULT_ROW_PT;
    const defaultColChars = ws.properties.defaultColWidth || DEFAULT_COL_CHARS;
    let maxCol = 0;

    for (let r = 1; r <= ws.rowCount; r++) {
      const row = ws.findRow(r);
      if (!row) continue;
      if (row.hidden) rowData[r - 1] = { ...(rowData[r - 1] ?? {}), hd: 1 };
      if (row.height && Math.abs(row.height - defaultRowPt) > 0.1) {
        rowData[r - 1] = { ...(rowData[r - 1] ?? {}), h: ptToPx(row.height) };
      }
      row.eachCell({ includeEmpty: true }, (cell, c) => {
        const data = cellToUniver(cell);
        const s = cell.type === ExcelJS.ValueType.Merge ? undefined : styleOf(cell.style);
        if (!data && !s) return;
        const out: ICellData = data ?? {};
        if (s) out.s = s;
        (cellData[r - 1] ??= {})[c - 1] = out;
        if (c > maxCol) maxCol = c;
      });
    }

    for (let c = 1; c <= ws.columnCount; c++) {
      const col = ws.getColumn(c);
      const entry: { w?: number; hd?: number } = {};
      if (col.width && Math.abs(col.width - defaultColChars) > 0.01) entry.w = charsToPx(col.width);
      if (col.hidden) entry.hd = 1;
      if (Object.keys(entry).length) columnData[c - 1] = entry;
    }

    const merges = Object.values((ws as unknown as { _merges: Record<string, { model: { top: number; left: number; bottom: number; right: number } }> })._merges ?? {});
    const mergeData = merges.map((m) => ({
      startRow: m.model.top - 1,
      startColumn: m.model.left - 1,
      endRow: m.model.bottom - 1,
      endColumn: m.model.right - 1,
    }));

    const view = (ws.views?.[0] ?? {}) as Partial<ExcelJS.WorksheetViewFrozen & ExcelJS.WorksheetViewCommon & { zoomScale?: number }>;
    const frozen = view.state === 'frozen';
    const xSplit = frozen ? Number(view.xSplit ?? 0) : 0;
    const ySplit = frozen ? Number(view.ySplit ?? 0) : 0;
    const tab = resolveColor(ws.properties.tabColor as ExcelColor | undefined, theme);

    sheets[id] = {
      id,
      name: ws.name,
      tabColor: tab ?? '',
      hidden: ws.state === 'veryHidden' ? 2 : ws.state === 'hidden' ? 1 : 0,
      rowCount: Math.max(ws.rowCount + 100, 1000),
      columnCount: Math.max(ws.columnCount, maxCol) + 10 > 26 ? Math.max(ws.columnCount, maxCol) + 10 : 26,
      zoomRatio: view.zoomScale ? view.zoomScale / 100 : 1,
      freeze: {
        xSplit,
        ySplit,
        startRow: ySplit > 0 ? ySplit : -1,
        startColumn: xSplit > 0 ? xSplit : -1,
      },
      defaultColumnWidth: charsToPx(defaultColChars),
      defaultRowHeight: ptToPx(defaultRowPt),
      showGridlines: view.showGridLines === false ? 0 : 1,
      rightToLeft: view.rightToLeft ? 1 : 0,
      mergeData,
      cellData,
      rowData,
      columnData,
    };
  });

  return {
    data: {
      id: 'greek-workbook',
      name,
      appVersion: '1.3.1',
      locale: locale as IWorkbookData['locale'],
      styles,
      sheetOrder,
      sheets,
    },
    idMap,
  };
}

export function blankWorkbook(name: string, locale: string): Partial<IWorkbookData> {
  return {
    id: 'greek-workbook',
    name,
    appVersion: '1.3.1',
    locale: locale as IWorkbookData['locale'],
    styles: {},
    sheetOrder: ['sheet-new-1'],
    sheets: {
      'sheet-new-1': {
        id: 'sheet-new-1',
        name: 'Sheet1',
        rowCount: 1000,
        columnCount: 26,
        defaultColumnWidth: 64,
        defaultRowHeight: 20,
        cellData: {},
      },
    },
  };
}

// ---------- writing ----------

function styleFromUniver(s: IStyleData): Partial<ExcelJS.Style> {
  const out: Partial<ExcelJS.Style> = {};
  const font: Partial<ExcelJS.Font> = {};
  if (s.ff) font.name = s.ff;
  if (s.fs) font.size = s.fs;
  if (s.bl === 1) font.bold = true;
  if (s.it === 1) font.italic = true;
  if (s.ul?.s === 1) font.underline = true;
  if (s.st?.s === 1) font.strike = true;
  const color = toArgb(s.cl?.rgb);
  if (color) font.color = { argb: color };
  if (s.va === 3) font.vertAlign = 'superscript';
  if (s.va === 2) font.vertAlign = 'subscript';
  if (Object.keys(font).length) out.font = font;

  const bg = toArgb(s.bg?.rgb);
  if (bg) out.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };

  if (s.bd) {
    const side = (d: { s: number; cl?: { rgb?: string | null | void } } | null | undefined | void): Partial<ExcelJS.Border> | undefined =>
      d && d.s && BORDER_FROM_UNIVER[d.s]
        ? { style: BORDER_FROM_UNIVER[d.s], color: { argb: toArgb(d.cl?.rgb) ?? 'FF000000' } }
        : undefined;
    const border: Partial<ExcelJS.Borders> = {};
    const t = side(s.bd.t), r = side(s.bd.r), b = side(s.bd.b), l = side(s.bd.l);
    if (t) border.top = t;
    if (r) border.right = r;
    if (b) border.bottom = b;
    if (l) border.left = l;
    const up = side(s.bd.bl_tr), down = side(s.bd.tl_br);
    if (up || down) border.diagonal = { ...(up ?? down)!, up: !!up, down: !!down } as ExcelJS.BorderDiagonal;
    if (Object.keys(border).length) out.border = border;
  }

  const align: Partial<ExcelJS.Alignment> = {};
  if (s.ht && H_FROM_UNIVER[s.ht]) align.horizontal = H_FROM_UNIVER[s.ht];
  if (s.vt && V_FROM_UNIVER[s.vt]) align.vertical = V_FROM_UNIVER[s.vt];
  if (s.tb === 3) align.wrapText = true;
  if (s.tr) {
    if (s.tr.v === 1) align.textRotation = 'vertical';
    else if (s.tr.a > 0) align.textRotation = Math.min(90, s.tr.a);
    else if (s.tr.a < 0) align.textRotation = Math.min(180, 90 - s.tr.a);
  }
  if (Object.keys(align).length) out.alignment = align;

  if (s.n?.pattern) out.numFmt = s.n.pattern;
  return out;
}

/** Plain text of a rich-text cell (Univer document model). */
function richText(p: ICellData['p']): string | undefined {
  const stream = p?.body?.dataStream;
  if (typeof stream !== 'string') return undefined;
  return stream.replace(/\r\n$/, '').replace(/[\r\n]+$/, '').replace(/\r/g, '\n');
}

function resetWorksheet(ws: ExcelJS.Worksheet) {
  const merges = Object.keys((ws as unknown as { _merges: Record<string, unknown> })._merges ?? {});
  for (const addr of merges) ws.unMergeCells(addr);
  ws.eachRow({ includeEmpty: true }, (row) => {
    row.eachCell({ includeEmpty: true }, (cell) => {
      cell.value = null;
      cell.style = {};
    });
    (row as { height?: number }).height = undefined;
    row.hidden = false;
  });
  for (let c = 1; c <= ws.columnCount; c++) {
    const col = ws.getColumn(c);
    (col as { width?: number }).width = undefined;
    col.hidden = false;
  }
}

export async function univerToXlsx(
  univerAPI: FUniver,
  base: ExcelJS.Workbook | null,
  idMap: Map<string, number>,
): Promise<{ bytes: Uint8Array; workbook: ExcelJS.Workbook }> {
  const fwb = univerAPI.getActiveWorkbook();
  if (!fwb) throw new Error('No workbook is open.');
  const snapshot = fwb.save();

  const wb = base ?? new ExcelJS.Workbook();
  if (!base) {
    wb.creator = 'Greek Office';
    wb.created = new Date();
  }
  wb.modified = new Date();
  // Dates were normalised to the 1900 system on import (see EXCEL_EPOCH_MS), so the file must say so.
  wb.properties = { ...(wb.properties ?? {}), date1904: false };
  wb.calcProperties = { ...(wb.calcProperties ?? {}), fullCalcOnLoad: true };

  const styleCache = new Map<string, Partial<ExcelJS.Style>>();
  const resolveStyle = (s: ICellData['s']): Partial<ExcelJS.Style> | undefined => {
    if (!s) return undefined;
    if (typeof s === 'string') {
      if (!styleCache.has(s)) {
        const def = snapshot.styles[s];
        styleCache.set(s, def ? styleFromUniver(def) : {});
      }
      return styleCache.get(s);
    }
    return styleFromUniver(s);
  };

  const keep = new Set<number>();
  const takenNames = new Set<string>();

  snapshot.sheetOrder.forEach((sheetId, order) => {
    const sd = snapshot.sheets[sheetId] as Partial<IWorksheetData>;
    if (!sd) return;
    const name = safeSheetName(sd.name ?? `Sheet${order + 1}`, takenNames);

    const mapped = idMap.get(sheetId);
    let ws = mapped !== undefined ? wb.getWorksheet(mapped) : undefined;
    if (ws) {
      resetWorksheet(ws);
      if (ws.name !== name) ws.name = name;
    } else {
      ws = wb.addWorksheet(name);
      idMap.set(sheetId, ws.id);
    }
    keep.add(ws.id);
    (ws as unknown as { orderNo: number }).orderNo = order;

    // Formulas: ask Univer, which resolves shared formulas into one per cell.
    const cellData = (sd.cellData ?? {}) as Record<string, Record<string, ICellData>>;
    let maxRow = -1;
    let maxCol = -1;
    for (const r of Object.keys(cellData)) {
      maxRow = Math.max(maxRow, Number(r));
      for (const c of Object.keys(cellData[r] ?? {})) maxCol = Math.max(maxCol, Number(c));
    }
    let formulas: string[][] = [];
    if (maxRow >= 0 && maxCol >= 0) {
      const fws = fwb.getSheetBySheetId(sheetId);
      formulas = fws ? fws.getRange(0, 0, maxRow + 1, maxCol + 1).getFormulas() : [];
    }

    for (const [rKey, cols] of Object.entries(cellData)) {
      const r = Number(rKey);
      for (const [cKey, cd] of Object.entries(cols ?? {})) {
        const c = Number(cKey);
        if (!cd) continue;
        const cell = ws.getCell(r + 1, c + 1);
        const style = resolveStyle(cd.s);
        if (style && Object.keys(style).length) cell.style = style;

        const formula = formulas[r]?.[c] || cd.f || '';
        const text = richText(cd.p);
        const v = text ?? cd.v;
        if (formula) {
          cell.value = { formula: formulaOut(formula), result: v === null || v === undefined ? undefined : (v as ExcelJS.CellValue as number) } as ExcelJS.CellFormulaValue;
          continue;
        }
        if (v === null || v === undefined || v === '') continue;
        const hl = (cd.custom as { hl?: string } | undefined)?.hl;
        if (cd.t === CELL_NUMBER || (cd.t !== CELL_STRING && typeof v === 'number')) {
          const n = typeof v === 'number' ? v : Number(v);
          cell.value = Number.isFinite(n) ? n : String(v);
        } else if (cd.t === CELL_BOOLEAN || typeof v === 'boolean') {
          cell.value = v === true || v === 1 || v === 'TRUE' || v === 'true';
        } else if (hl && SAFE_LINK_RE.test(hl)) {
          cell.value = { text: String(v), hyperlink: hl };
        } else {
          cell.value = String(v);
        }
      }
    }

    for (const m of sd.mergeData ?? []) {
      if (m.endRow > m.startRow || m.endColumn > m.startColumn) {
        try {
          ws.mergeCells(m.startRow + 1, m.startColumn + 1, m.endRow + 1, m.endColumn + 1);
        } catch {
          // Overlapping merges cannot be written; skip rather than fail the whole save.
        }
      }
    }

    const defaultRowPx = sd.defaultRowHeight ?? 20;
    for (const [rKey, rd] of Object.entries((sd.rowData ?? {}) as Record<string, { h?: number; hd?: number; ia?: number }>)) {
      const row = ws.getRow(Number(rKey) + 1);
      const hidden = rd.hd === 1;
      const custom = !!rd.h && rd.ia !== 1 && Math.abs(rd.h - defaultRowPx) > 0.5;
      if (!hidden && !custom) continue;
      if (hidden) row.hidden = true;
      if (custom) row.height = pxToPt(rd.h!);
      // ExcelJS does not write rows without cells; an empty styled cell keeps the row.
      if (!row.hasValues) {
        const anchor = row.getCell(1);
        if (!anchor.style || Object.keys(anchor.style).length === 0) anchor.style = { alignment: {} };
      }
    }
    const defaultColPx = sd.defaultColumnWidth ?? 64;
    for (const [cKey, cd] of Object.entries((sd.columnData ?? {}) as Record<string, { w?: number; hd?: number }>)) {
      const col = ws.getColumn(Number(cKey) + 1);
      if (cd.hd === 1) col.hidden = true;
      if (cd.w && Math.abs(cd.w - defaultColPx) > 0.5) col.width = pxToChars(cd.w);
    }

    ws.properties.defaultColWidth = pxToChars(defaultColPx);
    ws.properties.defaultRowHeight = pxToPt(defaultRowPx);
    const tab = toArgb(sd.tabColor);
    if (tab) ws.properties.tabColor = { argb: tab };
    else delete (ws.properties as { tabColor?: unknown }).tabColor;
    ws.state = sd.hidden === 2 ? 'veryHidden' : sd.hidden === 1 ? 'hidden' : 'visible';

    const fz = sd.freeze;
    const frozen = !!fz && (fz.xSplit > 0 || fz.ySplit > 0);
    const common = {
      showGridLines: sd.showGridlines !== 0,
      rightToLeft: sd.rightToLeft === 1,
      zoomScale: Math.round((sd.zoomRatio ?? 1) * 100),
    };
    ws.views = frozen
      ? [{ state: 'frozen', xSplit: fz!.xSplit, ySplit: fz!.ySplit, ...common }]
      : [{ state: 'normal', ...common }];
  });

  for (const ws of [...wb.worksheets]) {
    if (!keep.has(ws.id)) wb.removeWorksheet(ws.id);
  }

  const buffer = await wb.xlsx.writeBuffer();
  return { bytes: new Uint8Array(buffer as ArrayBuffer), workbook: wb };
}

type CsvCell = ICellData | null | undefined;
type StyleMap = Record<string, IStyleData | null | undefined | void>;

/** Number format pattern of a cell, from its style (a shared style id or an inline style). */
function patternOf(cd: ICellData, styles: StyleMap): string | undefined {
  const s = (typeof cd.s === 'string' ? styles[cd.s] : cd.s) as IStyleData | null | undefined;
  const pattern = s?.n?.pattern;
  return typeof pattern === 'string' && pattern ? pattern : undefined;
}

/**
 * Text of one cell as it should appear in the CSV, plus whether it is text.
 * Only text can carry CSV injection, so numbers and booleans are never flagged.
 */
function csvValue(cd: CsvCell, styles: StyleMap): { text: string; isText: boolean } {
  if (!cd) return { text: '', isText: false };
  const rich = richText(cd.p);
  if (rich !== undefined) return { text: rich, isText: true };
  const v = cd.v;
  if (v === null || v === undefined) return { text: '', isText: false };

  if (typeof v === 'boolean' || (cd.t === CELL_BOOLEAN && (v === 1 || v === 0 || /^(true|false)$/i.test(String(v))))) {
    return { text: v === true || v === 1 || /^true$/i.test(String(v)) ? 'TRUE' : 'FALSE', isText: false };
  }

  const numeric = typeof v === 'number' || (cd.t === CELL_NUMBER && String(v).trim() !== '' && Number.isFinite(Number(v)));
  if (!numeric) return { text: String(v), isText: true };

  // Univer's FRange.getDisplayValues() is not usable here: in 1.3.1 it returns String(cell.v) without
  // applying the number format. The format engine itself (numfmt, shared with the grid) is pure, so
  // run it directly on the stored pattern: dates become dates, percentages and currencies keep their form.
  const n = Number(v);
  const pattern = patternOf(cd, styles);
  if (pattern && pattern !== 'General') {
    try {
      const shown = numfmt.format(pattern, n, { locale: 'en', throws: false });
      if (typeof shown === 'string' && shown !== '') return { text: shown, isText: false };
    } catch {
      /* unreadable pattern: fall back to the raw number */
    }
  }
  return { text: String(v), isText: false };
}

/** Spreadsheet apps run a cell that starts with one of these characters as a formula (CSV injection). */
const CSV_FORMULA_START = /^[=+\-@\t\r]/;

/**
 * Pure CSV builder (RFC 4180 quoting) over Univer cell data, so it can be tested without a Univer instance.
 * Text that could be run as a formula is prefixed with an apostrophe, per OWASP guidance, before quoting.
 */
export function buildCsv(
  cellData: Record<string, Record<string, ICellData | null | undefined> | undefined>,
  styles: StyleMap = {},
): string {
  let maxRow = -1;
  let maxCol = -1;
  for (const r of Object.keys(cellData)) {
    for (const [c, cd] of Object.entries(cellData[r] ?? {})) {
      const v = richText(cd?.p) ?? cd?.v;
      if (v === null || v === undefined || v === '') continue;
      maxRow = Math.max(maxRow, Number(r));
      maxCol = Math.max(maxCol, Number(c));
    }
  }
  const quote = (s: string) => (/[",\r\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s);
  const lines: string[] = [];
  for (let r = 0; r <= maxRow; r++) {
    const cells: string[] = [];
    for (let c = 0; c <= maxCol; c++) {
      const { text, isText } = csvValue(cellData[r]?.[c], styles);
      cells.push(quote(isText && CSV_FORMULA_START.test(text) ? `'${text}` : text));
    }
    lines.push(cells.join(','));
  }
  return lines.join('\r\n');
}

/** CSV of one sheet's displayed values. */
export function sheetToCsv(univerAPI: FUniver): { csv: string; sheetName: string } {
  const fwb = univerAPI.getActiveWorkbook();
  const fws = fwb?.getActiveSheet();
  if (!fwb || !fws) throw new Error('No sheet is open.');
  const snapshot = fwb.save();
  const sd = snapshot.sheets[fws.getSheetId()] as Partial<IWorksheetData>;
  const cellData = (sd.cellData ?? {}) as Record<string, Record<string, ICellData>>;
  return { csv: buildCsv(cellData, snapshot.styles ?? {}), sheetName: sd.name ?? 'Sheet' };
}
