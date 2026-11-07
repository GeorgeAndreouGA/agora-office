/**
 * DrawingML colours, fills and lines, resolved against a theme.
 *
 * Colour references (schemeClr "accent1", "tx1"...) go through the slide's colour map to the
 * theme's colour scheme; modifiers (lumMod, tint, alpha...) are then applied in order.
 */
import { attr, is, kid, kids, num } from './xml';

export interface Theme {
  colors: Record<string, string>;
  majorFont: string;
  minorFont: string;
  fillStyles: Element[];
  lineStyles: Element[];
  bgFillStyles: Element[];
}

export const DEFAULT_THEME: Theme = {
  colors: {
    dk1: '000000',
    lt1: 'FFFFFF',
    dk2: '44546A',
    lt2: 'E7E6E6',
    accent1: '4472C4',
    accent2: 'ED7D31',
    accent3: 'A5A5A5',
    accent4: 'FFC000',
    accent5: '5B9BD5',
    accent6: '70AD47',
    hlink: '0563C1',
    folHlink: '954F72',
  },
  majorFont: 'Calibri Light',
  minorFont: 'Calibri',
  fillStyles: [],
  lineStyles: [],
  bgFillStyles: [],
};

export function readTheme(doc: Document | null): Theme {
  const root = doc?.documentElement;
  const elements = kid(root, 'a:themeElements');
  if (!elements) return DEFAULT_THEME;
  const colors: Record<string, string> = { ...DEFAULT_THEME.colors };
  for (const c of kids(kid(elements, 'a:clrScheme'))) {
    const v = kid(c, 'a:srgbClr') ? attr(kid(c, 'a:srgbClr'), 'val') : attr(kid(c, 'a:sysClr'), 'lastClr');
    if (v && /^[0-9a-f]{6}$/i.test(v)) colors[c.localName] = v.toUpperCase();
  }
  const fonts = kid(elements, 'a:fontScheme');
  const fmt = kid(elements, 'a:fmtScheme');
  return {
    colors,
    majorFont: attr(kid(kid(fonts, 'a:majorFont'), 'a:latin'), 'typeface') || DEFAULT_THEME.majorFont,
    minorFont: attr(kid(kid(fonts, 'a:minorFont'), 'a:latin'), 'typeface') || DEFAULT_THEME.minorFont,
    fillStyles: kids(kid(fmt, 'a:fillStyleLst')),
    lineStyles: kids(kid(fmt, 'a:lnStyleLst')),
    bgFillStyles: kids(kid(fmt, 'a:bgFillStyleLst')),
  };
}

/** bg1 -> lt1, tx1 -> dk1, ... */
export type ColorMap = Record<string, string>;

export const DEFAULT_CLR_MAP: ColorMap = {
  bg1: 'lt1',
  tx1: 'dk1',
  bg2: 'lt2',
  tx2: 'dk2',
  accent1: 'accent1',
  accent2: 'accent2',
  accent3: 'accent3',
  accent4: 'accent4',
  accent5: 'accent5',
  accent6: 'accent6',
  hlink: 'hlink',
  folHlink: 'folHlink',
};

export function readClrMap(el: Element | null, base: ColorMap = DEFAULT_CLR_MAP): ColorMap {
  if (!el) return base;
  const out: ColorMap = { ...base };
  for (const k of Object.keys(DEFAULT_CLR_MAP)) {
    const v = attr(el, k);
    if (v) out[k] = v;
  }
  return out;
}

export interface ColorContext {
  theme: Theme;
  clrMap: ColorMap;
}

export interface Color {
  hex: string;
  alpha: number;
}

const PRESET: Record<string, string> = {
  black: '000000',
  white: 'FFFFFF',
  red: 'FF0000',
  green: '008000',
  blue: '0000FF',
  yellow: 'FFFF00',
  gray: '808080',
  grey: '808080',
  darkGray: 'A9A9A9',
  lightGray: 'D3D3D3',
  orange: 'FFA500',
  navy: '000080',
};

const COLOR_TAGS = ['a:srgbClr', 'a:schemeClr', 'a:sysClr', 'a:prstClr', 'a:scrgbClr', 'a:hslClr'];

/** The colour element inside a fill or colour container, if any. */
export function colorChild(el: Element | null): Element | null {
  for (const c of kids(el)) if (COLOR_TAGS.some((t) => is(c, t))) return c;
  return null;
}

/** Resolve a colour element. `phClr` stands in for the placeholder colour of style matrices. */
export function resolveColor(el: Element | null, ctx: ColorContext, phClr?: Color | null): Color | null {
  if (!el) return null;
  let rgb: [number, number, number] | null = null;
  let alpha = 1;
  if (is(el, 'a:srgbClr')) rgb = hexToRgb(attr(el, 'val'));
  else if (is(el, 'a:sysClr')) rgb = hexToRgb(attr(el, 'lastClr') ?? (attr(el, 'val') === 'window' ? 'FFFFFF' : '000000'));
  else if (is(el, 'a:prstClr')) rgb = hexToRgb(PRESET[attr(el, 'val') ?? ''] ?? '000000');
  else if (is(el, 'a:scrgbClr')) {
    rgb = [num(el, 'r', 0), num(el, 'g', 0), num(el, 'b', 0)].map((v) => linToSrgb(v / 100000) * 255) as [number, number, number];
  } else if (is(el, 'a:hslClr')) {
    rgb = hslToRgb(num(el, 'hue', 0) / 60000 / 360, num(el, 'sat', 0) / 100000, num(el, 'lum', 0) / 100000);
  } else if (is(el, 'a:schemeClr')) {
    const val = attr(el, 'val') ?? '';
    if (val === 'phClr') {
      if (!phClr) return null;
      rgb = hexToRgb(phClr.hex);
      alpha = phClr.alpha;
    } else {
      const mapped = ctx.clrMap[val] ?? val;
      rgb = hexToRgb(ctx.theme.colors[mapped] ?? ctx.theme.colors[val] ?? '000000');
    }
  }
  if (!rgb) return null;
  // Modifiers, in document order.
  for (const m of kids(el)) {
    const v = num(m, 'val', 0) / 100000;
    switch (m.localName) {
      case 'alpha':
        alpha = v;
        break;
      case 'alphaMod':
        alpha *= v;
        break;
      case 'alphaOff':
        alpha = Math.min(1, Math.max(0, alpha + v));
        break;
      case 'lumMod':
      case 'lumOff':
      case 'satMod':
      case 'satOff':
      case 'hueMod':
      case 'hueOff': {
        const [h, s, l] = rgbToHsl(rgb);
        let hh = h;
        let ss = s;
        let ll = l;
        if (m.localName === 'lumMod') ll = l * v;
        if (m.localName === 'lumOff') ll = l + v;
        if (m.localName === 'satMod') ss = s * v;
        if (m.localName === 'satOff') ss = s + v;
        if (m.localName === 'hueMod') hh = (h * v) % 1;
        if (m.localName === 'hueOff') hh = (h + num(m, 'val', 0) / 60000 / 360) % 1;
        rgb = hslToRgb(hh, clamp01(ss), clamp01(ll));
        break;
      }
      case 'tint':
        rgb = rgb.map((c) => c * v + 255 * (1 - v)) as [number, number, number];
        break;
      case 'shade':
        rgb = rgb.map((c) => c * v) as [number, number, number];
        break;
      case 'inv':
        rgb = rgb.map((c) => 255 - c) as [number, number, number];
        break;
      case 'gray': {
        const g: number = rgb[0] * 0.3 + rgb[1] * 0.59 + rgb[2] * 0.11;
        rgb = [g, g, g];
        break;
      }
      case 'comp': {
        const [h, s, l] = rgbToHsl(rgb);
        rgb = hslToRgb((h + 0.5) % 1, s, l);
        break;
      }
      default:
        break;
    }
  }
  return { hex: rgbToHex(rgb), alpha: clamp01(alpha) };
}

export type Fill =
  | { kind: 'none' }
  | { kind: 'solid'; color: Color }
  | { kind: 'gradient'; stops: { pos: number; color: Color }[]; angle: number; path: 'linear' | 'radial' }
  | { kind: 'image'; rId: string; owner: string }
  | { kind: 'pattern'; fg: Color; bg: Color };

const FILL_TAGS = ['a:noFill', 'a:solidFill', 'a:gradFill', 'a:blipFill', 'a:pattFill', 'a:grpFill'];

export function fillChild(el: Element | null): Element | null {
  for (const c of kids(el)) if (FILL_TAGS.some((t) => is(c, t))) return c;
  return null;
}

/**
 * Read a fill element (a:solidFill etc.). `owner` is the part the element came from (needed
 * to resolve an image fill's relationship). Null for grpFill (caller supplies the group's).
 */
export function readFill(el: Element | null, ctx: ColorContext, owner: string, phClr?: Color | null): Fill | null {
  if (!el) return null;
  if (is(el, 'a:noFill')) return { kind: 'none' };
  if (is(el, 'a:solidFill')) {
    const c = resolveColor(colorChild(el), ctx, phClr);
    return c ? { kind: 'solid', color: c } : { kind: 'none' };
  }
  if (is(el, 'a:gradFill')) {
    const stops = kids(kid(el, 'a:gsLst'))
      .map((gs) => ({ pos: num(gs, 'pos', 0) / 100000, color: resolveColor(colorChild(gs), ctx, phClr) }))
      .filter((s): s is { pos: number; color: Color } => !!s.color)
      .sort((a, b) => a.pos - b.pos);
    if (!stops.length) return { kind: 'none' };
    const lin = kid(el, 'a:lin');
    return {
      kind: 'gradient',
      stops,
      angle: lin ? num(lin, 'ang', 0) / 60000 : 90,
      path: kid(el, 'a:path') ? 'radial' : 'linear',
    };
  }
  if (is(el, 'a:blipFill')) {
    const blip = kid(el, 'a:blip');
    const rId = attr(blip, 'r:embed');
    return rId ? { kind: 'image', rId, owner } : { kind: 'none' };
  }
  if (is(el, 'a:pattFill')) {
    const fg = resolveColor(colorChild(kid(el, 'a:fgClr')), ctx, phClr) ?? { hex: '000000', alpha: 1 };
    const bg = resolveColor(colorChild(kid(el, 'a:bgClr')), ctx, phClr) ?? { hex: 'FFFFFF', alpha: 1 };
    return { kind: 'pattern', fg, bg };
  }
  return null;
}

export interface Line {
  fill: Fill;
  /** Width in EMU. */
  width: number;
  dash: string | null;
  head: string | null;
  tail: string | null;
  cap: 'flat' | 'rnd' | 'sq';
}

export function readLine(ln: Element | null, ctx: ColorContext, owner: string, phClr?: Color | null, base?: Line | null): Line | null {
  if (!ln) return base ?? null;
  const fill = readFill(fillChild(ln), ctx, owner, phClr) ?? base?.fill ?? null;
  return {
    fill: fill ?? { kind: 'none' },
    width: num(ln, 'w') ?? base?.width ?? 9525,
    dash: attr(kid(ln, 'a:prstDash'), 'val') ?? base?.dash ?? null,
    head: attr(kid(ln, 'a:headEnd'), 'type') ?? base?.head ?? null,
    tail: attr(kid(ln, 'a:tailEnd'), 'type') ?? base?.tail ?? null,
    cap: (attr(ln, 'cap') as Line['cap']) ?? base?.cap ?? 'flat',
  };
}

/** The colour a fill shows on average (used where gradients cannot be drawn exactly). */
export function fillColor(fill: Fill | null): Color | null {
  if (!fill) return null;
  if (fill.kind === 'solid') return fill.color;
  if (fill.kind === 'gradient') return fill.stops[Math.floor(fill.stops.length / 2)]?.color ?? null;
  if (fill.kind === 'pattern') return fill.fg;
  return null;
}

/** CSS for a fill (images are handled by the caller). */
export function fillCss(fill: Fill | null): string {
  if (!fill || fill.kind === 'none' || fill.kind === 'image') return 'transparent';
  if (fill.kind === 'solid') return cssColor(fill.color);
  if (fill.kind === 'pattern') return cssColor(fill.fg);
  const stops = fill.stops.map((s) => `${cssColor(s.color)} ${(s.pos * 100).toFixed(1)}%`).join(', ');
  if (fill.path === 'radial') return `radial-gradient(circle, ${stops})`;
  return `linear-gradient(${(fill.angle + 90) % 360}deg, ${stops})`;
}

export function cssColor(c: Color | null | undefined): string {
  if (!c) return 'transparent';
  if (c.alpha >= 1) return `#${c.hex}`;
  const [r, g, b] = hexToRgb(c.hex) ?? [0, 0, 0];
  return `rgba(${Math.round(r)}, ${Math.round(g)}, ${Math.round(b)}, ${c.alpha.toFixed(3)})`;
}

// ---- colour maths ----
function clamp01(v: number) {
  return Math.min(1, Math.max(0, v));
}

export function hexToRgb(hex: string | null): [number, number, number] | null {
  if (!hex || !/^[0-9a-f]{6}$/i.test(hex)) return null;
  const v = parseInt(hex, 16);
  return [(v >> 16) & 255, (v >> 8) & 255, v & 255];
}

export function rgbToHex(rgb: readonly number[]): string {
  return rgb
    .map((c) => Math.round(Math.min(255, Math.max(0, c))).toString(16).padStart(2, '0'))
    .join('')
    .toUpperCase();
}

function linToSrgb(v: number): number {
  return v <= 0.0031308 ? 12.92 * v : 1.055 * Math.pow(v, 1 / 2.4) - 0.055;
}

function rgbToHsl([r, g, b]: [number, number, number]): [number, number, number] {
  r /= 255;
  g /= 255;
  b /= 255;
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l = (max + min) / 2;
  if (max === min) return [0, 0, l];
  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  let h = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
  h /= 6;
  return [h, s, l];
}

function hslToRgb(h: number, s: number, l: number): [number, number, number] {
  if (s === 0) return [l * 255, l * 255, l * 255];
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  const f = (t: number) => {
    if (t < 0) t += 1;
    if (t > 1) t -= 1;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
  };
  return [f(h + 1 / 3) * 255, f(h) * 255, f(h - 1 / 3) * 255];
}
