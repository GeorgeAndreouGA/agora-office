/** Excel colour resolution: ARGB, theme + tint, and the legacy indexed palette. */

export interface ExcelColor {
  argb?: string;
  theme?: number;
  tint?: number;
  indexed?: number;
}

// The 64-entry legacy palette (indices 0-63), as defined by ECMA-376.
const INDEXED = [
  '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
  '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
  '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080',
  '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF',
  '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF',
  '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99',
  '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696',
  '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333',
];

// Office 2013+ default theme, in Excel's index order: lt1, dk1, lt2, dk2, accent1-6, hlink, folHlink.
const DEFAULT_THEME = [
  'FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31',
  'A5A5A5', 'FFC000', '5B9BD5', '70AD47', '0563C1', '954F72',
];

/** Read the colour scheme out of the workbook's theme XML (theme1.xml). */
export function parseThemePalette(themeXml: string | undefined): string[] {
  if (!themeXml) return DEFAULT_THEME;
  const scheme = /<a:clrScheme[\s\S]*?<\/a:clrScheme>/.exec(themeXml)?.[0];
  if (!scheme) return DEFAULT_THEME;
  const pick = (tag: string): string | undefined => {
    const block = new RegExp(`<a:${tag}>([\\s\\S]*?)</a:${tag}>`).exec(scheme)?.[1];
    if (!block) return undefined;
    const srgb = /<a:srgbClr\s+val="([0-9A-Fa-f]{6})"/.exec(block)?.[1];
    const sys = /lastClr="([0-9A-Fa-f]{6})"/.exec(block)?.[1];
    return (srgb ?? sys)?.toUpperCase();
  };
  const order = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];
  return order.map((tag, i) => pick(tag) ?? DEFAULT_THEME[i]);
}

function hexToHsl(hex: string): [number, number, number] {
  const r = parseInt(hex.slice(0, 2), 16) / 255;
  const g = parseInt(hex.slice(2, 4), 16) / 255;
  const b = parseInt(hex.slice(4, 6), 16) / 255;
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l = (max + min) / 2;
  if (max === min) return [0, 0, l];
  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  let h = 0;
  if (max === r) h = (g - b) / d + (g < b ? 6 : 0);
  else if (max === g) h = (b - r) / d + 2;
  else h = (r - g) / d + 4;
  return [h / 6, s, l];
}

function hslToHex(h: number, s: number, l: number): string {
  const f = (p: number, q: number, t: number) => {
    if (t < 0) t += 1;
    if (t > 1) t -= 1;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
  };
  let r: number, g: number, b: number;
  if (s === 0) {
    r = g = b = l;
  } else {
    const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
    const p = 2 * l - q;
    r = f(p, q, h + 1 / 3);
    g = f(p, q, h);
    b = f(p, q, h - 1 / 3);
  }
  const to = (x: number) => Math.round(Math.min(1, Math.max(0, x)) * 255).toString(16).padStart(2, '0');
  return `${to(r)}${to(g)}${to(b)}`.toUpperCase();
}

function applyTint(hex: string, tint: number): string {
  if (!tint) return hex;
  const [h, s, l] = hexToHsl(hex);
  const nl = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint;
  return hslToHex(h, s, nl);
}

/** Resolve an Excel colour to '#RRGGBB', or undefined when it means "automatic". */
export function resolveColor(color: ExcelColor | undefined, theme: string[]): string | undefined {
  if (!color) return undefined;
  if (color.argb && /^[0-9A-Fa-f]{8}$/.test(color.argb)) return `#${color.argb.slice(2).toUpperCase()}`;
  if (color.argb && /^[0-9A-Fa-f]{6}$/.test(color.argb)) return `#${color.argb.toUpperCase()}`;
  if (typeof color.theme === 'number' && theme[color.theme]) {
    return `#${applyTint(theme[color.theme], color.tint ?? 0)}`;
  }
  if (typeof color.indexed === 'number' && INDEXED[color.indexed]) return `#${INDEXED[color.indexed]}`;
  return undefined;
}

export function toArgb(rgb: string | null | undefined | void): string | undefined {
  if (!rgb) return undefined;
  const m = /^#?([0-9a-f]{6})$/i.exec(rgb.trim());
  if (m) return `FF${m[1].toUpperCase()}`;
  const rgbFn = /^rgba?\((\d+),\s*(\d+),\s*(\d+)/i.exec(rgb.trim());
  if (rgbFn) {
    const h = (n: string) => Math.min(255, parseInt(n, 10)).toString(16).padStart(2, '0');
    return `FF${h(rgbFn[1])}${h(rgbFn[2])}${h(rgbFn[3])}`.toUpperCase();
  }
  return undefined;
}
