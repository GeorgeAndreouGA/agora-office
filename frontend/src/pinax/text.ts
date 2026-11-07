/** Resolved text styling shared by the renderer, the text editor and the PDF exporter. */
import { mergeLevel } from './pptx/read';
import type { Bullet, LevelStyle, Paragraph, RunProps, RunStyle, TextBody } from './model';
import { PX_PER_PT } from './model';

export function paraStyle(body: TextBody, p: Paragraph): LevelStyle {
  const base = body.levels[p.level] ?? body.levels[0];
  return mergeLevel(base, p.props);
}

export function runStyle(lvl: LevelStyle, props: RunProps): RunStyle {
  const out: RunStyle = { ...lvl.run };
  for (const [k, v] of Object.entries(props)) if (v !== undefined) (out as unknown as Record<string, unknown>)[k] = v;
  return out;
}

export const ALIGN_CSS: Record<string, 'left' | 'center' | 'right' | 'justify'> = {
  l: 'left',
  ctr: 'center',
  r: 'right',
  just: 'justify',
  dist: 'justify',
};

/** True when no paragraph has any visible text. */
export function isEmptyBody(body: TextBody | null): boolean {
  if (!body) return true;
  return body.paragraphs.every((p) => p.items.every((i) => i.kind === 'br' || !i.text.trim()));
}

/** Plain text of a body (for search, accessibility and tests). */
export function plainText(body: TextBody | null): string {
  if (!body) return '';
  return body.paragraphs.map((p) => p.items.map((i) => (i.kind === 'br' ? '\n' : i.text)).join('')).join('\n');
}

// Wingdings / Symbol bullet characters and their Unicode look-alikes.
const WINGDINGS: Record<string, string> = {
  'Ø': '➢',
  '§': '▪',
  q: '❑',
  v: '❖',
  ü: '✓',
  l: '●',
  n: '■',
  'Ÿ': '•',
  'ª': '•',
  è: '➔',
  à: '➔',
  o: '□',
  p: '◆',
};

export function bulletChar(b: Extract<Bullet, { kind: 'char' }>): string {
  const font = (b.font ?? '').toLowerCase();
  if (font.startsWith('wingdings')) return WINGDINGS[b.char] ?? '•';
  if (font === 'symbol') return b.char === '·' || b.char === '' ? '•' : b.char;
  if (/^[-]$/.test(b.char)) return WINGDINGS[String.fromCharCode(b.char.charCodeAt(0) - 0xf000)] ?? '•';
  return b.char || '•';
}

function roman(n: number): string {
  const map: [number, string][] = [
    [1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i'],
  ];
  let out = '';
  for (const [v, s] of map) while (n >= v) {
    out += s;
    n -= v;
  }
  return out;
}

function alpha(n: number): string {
  let out = '';
  while (n > 0) {
    n -= 1;
    out = String.fromCharCode(97 + (n % 26)) + out;
    n = Math.floor(n / 26);
  }
  return out;
}

export function numberText(scheme: string, n: number): string {
  let core: string;
  if (/^romanUc/.test(scheme)) core = roman(n).toUpperCase();
  else if (/^romanLc/.test(scheme)) core = roman(n);
  else if (/^alphaUc/.test(scheme)) core = alpha(n).toUpperCase();
  else if (/^alphaLc/.test(scheme)) core = alpha(n);
  else core = String(n);
  if (/ParenBoth$/.test(scheme)) return `(${core})`;
  if (/ParenR$/.test(scheme)) return `${core})`;
  if (/Plain$/.test(scheme)) return core;
  return `${core}.`;
}

/** The marker text for every paragraph (null for none), numbering counted per level. */
export function markers(body: TextBody): (string | null)[] {
  const counters: number[] = Array(9).fill(0);
  return body.paragraphs.map((p) => {
    const st = paraStyle(body, p);
    for (let l = p.level + 1; l < 9; l++) counters[l] = 0;
    const b = st.bullet;
    const empty = p.items.every((i) => i.kind === 'br' || !i.text);
    if (b.kind === 'none' || empty) {
      if (b.kind !== 'num') counters[p.level] = 0;
      return null;
    }
    if (b.kind === 'char') {
      counters[p.level] = 0;
      return bulletChar(b);
    }
    counters[p.level] = counters[p.level] ? counters[p.level] + 1 : b.startAt;
    return numberText(b.scheme, counters[p.level]);
  });
}

/** CSS line-height for a paragraph. */
export function lineHeight(st: LevelStyle, fontScale: number, reduction: number, sizePt: number): string {
  if (st.lineSpacing.pts) return `${st.lineSpacing.pts * PX_PER_PT * fontScale}px`;
  const pct = (st.lineSpacing.pct ?? 1) * (1 - reduction);
  return `${Math.max(0.5, pct * 1.2) * sizePt * PX_PER_PT * fontScale}px`;
}

/** Space before/after in px. */
export function spacingPx(s: { pct?: number; pts?: number }, sizePt: number, fontScale: number): number {
  if (s.pts) return s.pts * PX_PER_PT * fontScale;
  if (s.pct) return s.pct * 1.2 * sizePt * PX_PER_PT * fontScale;
  return 0;
}

/** Largest run size in a paragraph (for line height and empty paragraphs). */
export function paragraphSize(st: LevelStyle, p: Paragraph): number {
  let size = 0;
  for (const it of p.items) if (it.kind !== 'br') size = Math.max(size, runStyle(st, it.props).size);
  return size || p.endSize || st.run.size;
}
