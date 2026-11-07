/**
 * A small PDF builder for Greek Office's exporters: pdf-lib holds the document (objects, images,
 * compression) and fontkit parses and subsets fonts. Everything runs in the browser tab.
 *
 * Why text is not drawn with pdf-lib's drawText: the exporters already know every glyph's exact
 * ID and position (Papyros gets them from the editor's own shaping engine). drawText would shape
 * the text a second time and drift from the screen. So fonts are embedded here as Type0 subsets
 * whose glyphs are drawn by ID, with a ToUnicode map so the PDF's text can be selected, searched
 * and copied.
 *
 * Coordinates: each page's content stream first flips to a top-left, y-down space in points, the
 * space both editors lay out in. Callers never think about PDF's bottom-left origin, except
 * here for annotation rectangles.
 */
import { PDFDocument, PDFName, PDFString, type PDFPage, type PDFRef } from 'pdf-lib';
import fontkit from '@pdf-lib/fontkit';
import { hasGlyfOutlines, subsetTrueType, woffToSfnt } from './sfnt';

type FkFont = ReturnType<typeof fontkit.create>;

/** RGB in 0..1. */
export type Rgb = readonly [number, number, number];

export const BLACK: Rgb = [0, 0, 0];

/** `RRGGBB` (with or without '#') to Rgb; anything malformed gives the fallback. */
export function rgb(hex: string | null | undefined, fallback: Rgb = BLACK): Rgb {
  const m = /^#?([0-9a-f]{6})$/i.exec(hex ?? '');
  if (!m) return fallback;
  const v = parseInt(m[1], 16);
  return [((v >> 16) & 255) / 255, ((v >> 8) & 255) / 255, (v & 255) / 255];
}

/** Compact, locale-free number formatting for content streams. */
export function num(v: number): string {
  if (!Number.isFinite(v)) return '0';
  const r = Math.round(v * 1000) / 1000;
  return Object.is(r, -0) ? '0' : String(r);
}

export interface Point {
  x: number;
  y: number;
}

/** Affine matrix [a b c d e f]: x' = a*x + c*y + e, y' = b*x + d*y + f. */
export type Matrix = readonly [number, number, number, number, number, number];

function utf16Hex(text: string): string {
  let out = '';
  for (let i = 0; i < text.length; i++) out += text.charCodeAt(i).toString(16).padStart(4, '0');
  return out;
}

function sanitizeName(name: string): string {
  return name.replace(/[^A-Za-z0-9-]/g, '').slice(0, 60) || 'Font';
}

/**
 * One embedded font face. CIDs are the font's own glyph IDs (CIDToGIDMap Identity), so a glyph
 * shaped anywhere else can be drawn as-is; the subset only empties the glyphs never drawn.
 */
export class PdfFont {
  readonly ref: PDFRef;
  readonly unitsPerEm: number;
  private readonly widths = new Map<number, number>();
  private readonly unicode = new Map<number, string>();

  constructor(
    private readonly doc: PDFDocument,
    readonly font: FkFont,
    private readonly sfnt: Uint8Array,
    private readonly tag: string,
  ) {
    this.ref = doc.context.nextRef();
    this.unitsPerEm = font.unitsPerEm || 1000;
  }

  /** The CID to draw glyph `gid` with; `text` (the characters it stands for) feeds copy/search. */
  cid(gid: number, text?: string): number {
    if (!this.widths.has(gid)) this.widths.set(gid, (this.font.getGlyph(gid).advanceWidth * 1000) / this.unitsPerEm);
    // .notdef never maps to text: it would make every missing character copy as the first one.
    if (gid !== 0 && text && !this.unicode.has(gid)) this.unicode.set(gid, text);
    return gid;
  }

  /** Advance of a drawn CID in thousandths of an em (what the PDF viewer will move by). */
  widthOfCid(cid: number): number {
    return this.widths.get(cid) ?? 0;
  }

  /** Underline position (negative = below baseline) and thickness, in ems. */
  get underline(): { position: number; thickness: number } {
    const pos = this.font.underlinePosition || -0.1 * this.unitsPerEm;
    const th = this.font.underlineThickness || 0.05 * this.unitsPerEm;
    return { position: pos / this.unitsPerEm, thickness: th / this.unitsPerEm };
  }

  /** Writes the font program, widths and ToUnicode map. Called once, by PdfBuilder.save(). */
  async finish(): Promise<void> {
    const ctx = this.doc.context;
    const trueType = hasGlyfOutlines(this.sfnt);
    // TrueType outlines are subset; CFF (OpenType) faces are embedded whole, addressed by GID.
    const program = (trueType && subsetTrueType(this.sfnt, this.widths.keys())) || this.sfnt;
    const isCFF = !trueType;
    const scale = 1000 / this.unitsPerEm;
    const name = `${this.tag}+${sanitizeName(this.font.postscriptName ?? 'Font')}`;

    const fontFile = ctx.register(ctx.flateStream(program, isCFF ? { Subtype: 'OpenType' } : {}));
    const bbox = this.font.bbox;
    const italic = (this.font.italicAngle ?? 0) !== 0;
    const fixed = !!(this.font as unknown as { post?: { isFixedPitch?: number } }).post?.isFixedPitch;
    const descriptor = ctx.register(
      ctx.obj({
        Type: 'FontDescriptor',
        FontName: name,
        Flags: 4 | (italic ? 64 : 0) | (fixed ? 1 : 0),
        FontBBox: [bbox.minX * scale, bbox.minY * scale, bbox.maxX * scale, bbox.maxY * scale],
        ItalicAngle: this.font.italicAngle ?? 0,
        Ascent: this.font.ascent * scale,
        Descent: this.font.descent * scale,
        CapHeight: (this.font.capHeight || this.font.ascent) * scale,
        StemV: 0,
        [isCFF ? 'FontFile3' : 'FontFile2']: fontFile,
      }),
    );

    // W array: runs of consecutive CIDs share one bracket.
    const cids = [...this.widths.keys()].sort((a, b) => a - b);
    const w: (number | number[])[] = [];
    let run: number[] = [];
    let prev = -2;
    for (const c of cids) {
      if (c !== prev + 1) {
        if (run.length) w.push(run);
        w.push(c);
        run = [];
      }
      run.push(Math.round(this.widths.get(c) ?? 0));
      prev = c;
    }
    if (run.length) w.push(run);

    const cidFont = ctx.register(
      ctx.obj({
        Type: 'Font',
        Subtype: isCFF ? 'CIDFontType0' : 'CIDFontType2',
        BaseFont: name,
        CIDSystemInfo: { Registry: PDFString.of('Adobe'), Ordering: PDFString.of('Identity'), Supplement: 0 },
        FontDescriptor: descriptor,
        W: w,
        ...(isCFF ? {} : { CIDToGIDMap: 'Identity' }),
      }),
    );

    const entries = [...this.unicode.entries()].sort((a, b) => a[0] - b[0]);
    const blocks: string[] = [];
    for (let i = 0; i < entries.length; i += 100) {
      const chunk = entries.slice(i, i + 100);
      blocks.push(
        `${chunk.length} beginbfchar\n` +
          chunk.map(([c, t]) => `<${c.toString(16).padStart(4, '0')}> <${utf16Hex(t)}>`).join('\n') +
          '\nendbfchar',
      );
    }
    const cmap =
      '/CIDInit /ProcSet findresource begin\n12 dict begin\nbegincmap\n' +
      '/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def\n' +
      '/CMapName /Adobe-Identity-UCS def\n/CMapType 2 def\n' +
      '1 begincodespacerange\n<0000> <ffff>\nendcodespacerange\n' +
      blocks.join('\n') +
      '\nendcmap\nCMapName currentdict /CMap defineresource pop\nend\nend';
    const toUnicode = ctx.register(ctx.flateStream(cmap));

    ctx.assign(
      this.ref,
      ctx.obj({
        Type: 'Font',
        Subtype: 'Type0',
        BaseFont: name,
        Encoding: 'Identity-H',
        DescendantFonts: [cidFont],
        ToUnicode: toUnicode,
      }),
    );
  }
}

export interface PdfImage {
  ref: PDFRef;
  width: number;
  height: number;
}

/** A glyph to draw: its CID and the pen position (baseline origin) in page points. */
export interface PlacedGlyph {
  cid: number;
  x: number;
  y: number;
}

export interface GlyphRunStyle {
  /** Font size in points. */
  size: number;
  /** Horizontal scale, 1 = normal. */
  hScale?: number;
  fill?: Rgb | null;
  fillAlpha?: number;
  /** Outline the glyphs (text outline effect). */
  stroke?: { color: Rgb; width: number } | null;
}

/** One page being drawn. All coordinates are top-left origin, y down, in points. */
export class PdfPage {
  private readonly ops: string[] = [];
  private readonly fontNames = new Map<PdfFont, string>();
  private readonly imageNames = new Map<PDFRef, string>();
  private readonly gsNames = new Map<string, string>();
  private readonly annots: PDFRef[] = [];

  constructor(
    private readonly builder: PdfBuilder,
    readonly page: PDFPage,
    readonly width: number,
    readonly height: number,
  ) {
    this.ops.push(`1 0 0 -1 0 ${num(height)} cm`);
  }

  // ---- state ----
  save(): this {
    this.ops.push('q');
    return this;
  }
  restore(): this {
    this.ops.push('Q');
    return this;
  }
  transform(m: Matrix): this {
    this.ops.push(`${m.map(num).join(' ')} cm`);
    return this;
  }
  /** Rotate by `degrees` clockwise (screen sense) about (cx, cy). */
  rotateAbout(degrees: number, cx: number, cy: number): this {
    if (!degrees) return this;
    const r = (degrees * Math.PI) / 180;
    const c = Math.cos(r);
    const s = Math.sin(r);
    return this.transform([c, s, -s, c, cx - c * cx + s * cy, cy - s * cx - c * cy]);
  }
  fillColor(c: Rgb): this {
    this.ops.push(`${num(c[0])} ${num(c[1])} ${num(c[2])} rg`);
    return this;
  }
  strokeColor(c: Rgb): this {
    this.ops.push(`${num(c[0])} ${num(c[1])} ${num(c[2])} RG`);
    return this;
  }
  lineWidth(w: number): this {
    this.ops.push(`${num(w)} w`);
    return this;
  }
  dash(pattern: readonly number[] | null): this {
    this.ops.push(pattern && pattern.length ? `[${pattern.map(num).join(' ')}] 0 d` : '[] 0 d');
    return this;
  }
  lineCap(cap: 'butt' | 'round' | 'square'): this {
    this.ops.push(`${cap === 'round' ? 1 : cap === 'square' ? 2 : 0} J`);
    return this;
  }
  lineJoin(join: 'miter' | 'round' | 'bevel'): this {
    this.ops.push(`${join === 'round' ? 1 : join === 'bevel' ? 2 : 0} j`);
    return this;
  }
  /** Fill and stroke opacity, 0..1 (only emitted when below 1). */
  alpha(fill: number, stroke: number = fill): this {
    if (fill >= 1 && stroke >= 1) return this;
    const key = `${num(fill)}/${num(stroke)}`;
    let name = this.gsNames.get(key);
    if (!name) {
      const ctx = this.builder.doc.context;
      const ref = ctx.register(ctx.obj({ Type: 'ExtGState', ca: Math.max(0, fill), CA: Math.max(0, stroke) }));
      name = this.page.node.newExtGState('GS', ref).toString();
      this.gsNames.set(key, name);
    }
    this.ops.push(`${name} gs`);
    return this;
  }

  // ---- paths ----
  rect(x: number, y: number, w: number, h: number): this {
    this.ops.push(`${num(x)} ${num(y)} ${num(w)} ${num(h)} re`);
    return this;
  }
  moveTo(x: number, y: number): this {
    this.ops.push(`${num(x)} ${num(y)} m`);
    return this;
  }
  lineTo(x: number, y: number): this {
    this.ops.push(`${num(x)} ${num(y)} l`);
    return this;
  }
  curveTo(x1: number, y1: number, x2: number, y2: number, x: number, y: number): this {
    this.ops.push(`${num(x1)} ${num(y1)} ${num(x2)} ${num(y2)} ${num(x)} ${num(y)} c`);
    return this;
  }
  close(): this {
    this.ops.push('h');
    return this;
  }
  polygon(points: readonly Point[], closed = true): this {
    if (!points.length) return this;
    this.moveTo(points[0].x, points[0].y);
    for (let i = 1; i < points.length; i++) this.lineTo(points[i].x, points[i].y);
    if (closed) this.close();
    return this;
  }
  /** An SVG path (M L H V C S Q T Z, absolute or relative; arcs become straight lines). */
  svgPath(d: string, map: (x: number, y: number) => Point = (x, y) => ({ x, y })): this {
    const tokens = d.match(/[MmLlHhVvCcSsQqTtAaZz]|-?(?:\d+\.?\d*|\.\d+)(?:e[-+]?\d+)?/g) ?? [];
    let i = 0;
    let cmd = '';
    let cx = 0;
    let cy = 0;
    let sx = 0;
    let sy = 0;
    let lcx = 0; // last control point, for S/T
    let lcy = 0;
    let lastCubic = false;
    let lastQuad = false;
    const nextNum = () => parseFloat(tokens[i++] ?? '0');
    const isCmd = (t: string | undefined) => !!t && /^[A-Za-z]$/.test(t);
    const emitMove = (x: number, y: number) => {
      const p = map(x, y);
      this.moveTo(p.x, p.y);
    };
    const emitLine = (x: number, y: number) => {
      const p = map(x, y);
      this.lineTo(p.x, p.y);
    };
    const emitCubic = (x1: number, y1: number, x2: number, y2: number, x: number, y: number) => {
      const a = map(x1, y1);
      const b = map(x2, y2);
      const c = map(x, y);
      this.curveTo(a.x, a.y, b.x, b.y, c.x, c.y);
    };
    while (i < tokens.length) {
      if (isCmd(tokens[i])) cmd = tokens[i++];
      else if (!cmd) {
        i++;
        continue;
      }
      const rel = cmd === cmd.toLowerCase();
      const C = cmd.toUpperCase();
      if (C === 'Z') {
        this.close();
        cx = sx;
        cy = sy;
        lastCubic = lastQuad = false;
        cmd = '';
        continue;
      }
      if (C === 'M') {
        let x = nextNum();
        let y = nextNum();
        if (rel) {
          x += cx;
          y += cy;
        }
        emitMove(x, y);
        cx = sx = x;
        cy = sy = y;
        cmd = rel ? 'l' : 'L'; // implicit lineto for following pairs
        lastCubic = lastQuad = false;
      } else if (C === 'L') {
        let x = nextNum();
        let y = nextNum();
        if (rel) {
          x += cx;
          y += cy;
        }
        emitLine(x, y);
        cx = x;
        cy = y;
        lastCubic = lastQuad = false;
      } else if (C === 'H') {
        let x = nextNum();
        if (rel) x += cx;
        emitLine(x, cy);
        cx = x;
        lastCubic = lastQuad = false;
      } else if (C === 'V') {
        let y = nextNum();
        if (rel) y += cy;
        emitLine(cx, y);
        cy = y;
        lastCubic = lastQuad = false;
      } else if (C === 'C' || C === 'S') {
        let x1: number;
        let y1: number;
        if (C === 'C') {
          x1 = nextNum();
          y1 = nextNum();
          if (rel) {
            x1 += cx;
            y1 += cy;
          }
        } else {
          x1 = lastCubic ? 2 * cx - lcx : cx;
          y1 = lastCubic ? 2 * cy - lcy : cy;
        }
        let x2 = nextNum();
        let y2 = nextNum();
        let x = nextNum();
        let y = nextNum();
        if (rel) {
          x2 += cx;
          y2 += cy;
          x += cx;
          y += cy;
        }
        emitCubic(x1, y1, x2, y2, x, y);
        lcx = x2;
        lcy = y2;
        cx = x;
        cy = y;
        lastCubic = true;
        lastQuad = false;
      } else if (C === 'Q' || C === 'T') {
        let qx: number;
        let qy: number;
        if (C === 'Q') {
          qx = nextNum();
          qy = nextNum();
          if (rel) {
            qx += cx;
            qy += cy;
          }
        } else {
          qx = lastQuad ? 2 * cx - lcx : cx;
          qy = lastQuad ? 2 * cy - lcy : cy;
        }
        let x = nextNum();
        let y = nextNum();
        if (rel) {
          x += cx;
          y += cy;
        }
        // Quadratic -> cubic.
        emitCubic(cx + (2 / 3) * (qx - cx), cy + (2 / 3) * (qy - cy), x + (2 / 3) * (qx - x), y + (2 / 3) * (qy - y), x, y);
        lcx = qx;
        lcy = qy;
        cx = x;
        cy = y;
        lastQuad = true;
        lastCubic = false;
      } else if (C === 'A') {
        // rx ry rot large sweep x y: approximated by a straight line (rare in our sources).
        i += 5;
        let x = nextNum();
        let y = nextNum();
        if (rel) {
          x += cx;
          y += cy;
        }
        emitLine(x, y);
        cx = x;
        cy = y;
        lastCubic = lastQuad = false;
      } else {
        i++;
      }
    }
    return this;
  }
  fill(evenOdd = false): this {
    this.ops.push(evenOdd ? 'f*' : 'f');
    return this;
  }
  stroke(): this {
    this.ops.push('S');
    return this;
  }
  fillStroke(evenOdd = false): this {
    this.ops.push(evenOdd ? 'B*' : 'B');
    return this;
  }
  clip(evenOdd = false): this {
    this.ops.push(evenOdd ? 'W* n' : 'W n');
    return this;
  }
  endPath(): this {
    this.ops.push('n');
    return this;
  }

  fillRect(x: number, y: number, w: number, h: number, color: Rgb, alpha = 1): this {
    if (w <= 0 || h <= 0) return this;
    this.save();
    if (alpha < 1) this.alpha(alpha);
    this.fillColor(color).rect(x, y, w, h).fill();
    return this.restore();
  }

  line(x1: number, y1: number, x2: number, y2: number, color: Rgb, width: number, dash?: readonly number[] | null): this {
    this.save().strokeColor(color).lineWidth(width);
    if (dash) this.dash(dash);
    this.moveTo(x1, y1).lineTo(x2, y2).stroke();
    return this.restore();
  }

  // ---- images ----
  /**
   * Draw an image so its top-left, top-right and bottom-left corners land on the given points
   * (this handles scale, rotation, flips and skew in one matrix).
   */
  image(img: PdfImage, tl: Point, tr: Point, bl: Point): this {
    let name = this.imageNames.get(img.ref);
    if (!name) {
      name = this.page.node.newXObject('Im', img.ref).toString();
      this.imageNames.set(img.ref, name);
    }
    // Image space: (0,0) bottom-left of the picture, (0,1) its top-left.
    this.save().transform([tr.x - tl.x, tr.y - tl.y, tl.x - bl.x, tl.y - bl.y, bl.x, bl.y]);
    this.ops.push(`${name} Do`);
    return this.restore();
  }

  imageInBox(img: PdfImage, x: number, y: number, w: number, h: number): this {
    return this.image(img, { x, y }, { x: x + w, y }, { x, y: y + h });
  }

  // ---- text ----
  /**
   * Draw positioned glyphs of one font at one size. Glyphs on the same baseline become one TJ
   * array whose kerning numbers move the pen exactly to each glyph's x.
   */
  glyphs(font: PdfFont, style: GlyphRunStyle, glyphs: readonly PlacedGlyph[]): this {
    if (!glyphs.length || style.size <= 0) return this;
    let name = this.fontNames.get(font);
    if (!name) {
      name = this.page.node.newFontDictionary('F', font.ref).toString();
      this.fontNames.set(font, name);
    }
    const hs = style.hScale && style.hScale > 0 ? style.hScale : 1;
    const sx = style.size * hs;
    this.save();
    if ((style.fillAlpha ?? 1) < 1) this.alpha(style.fillAlpha ?? 1, 1);
    if (style.fill !== null) this.fillColor(style.fill ?? BLACK);
    if (style.stroke) this.strokeColor(style.stroke.color).lineWidth(style.stroke.width);
    const mode = style.stroke ? (style.fill === null ? 1 : 2) : 0;
    this.ops.push('BT', `${name} 1 Tf`);
    if (mode) this.ops.push(`${mode} Tr`);
    let i = 0;
    while (i < glyphs.length) {
      const y = glyphs[i].y;
      let j = i + 1;
      while (j < glyphs.length && Math.abs(glyphs[j].y - y) < 1e-4) j++;
      this.ops.push(`${num(sx)} 0 0 ${num(-style.size)} ${num(glyphs[i].x)} ${num(y)} Tm`);
      let tj = '';
      for (let k = i; k < j; k++) {
        const g = glyphs[k];
        tj += `<${g.cid.toString(16).padStart(4, '0')}>`;
        if (k + 1 < j) {
          const pen = g.x + (font.widthOfCid(g.cid) / 1000) * sx;
          const adj = ((pen - glyphs[k + 1].x) * 1000) / sx;
          if (Math.abs(adj) > 0.05) tj += num(adj);
        }
      }
      this.ops.push(`[${tj}] TJ`);
      i = j;
    }
    this.ops.push('ET');
    return this.restore();
  }

  // ---- links ----
  /** A clickable area (top-left coordinates) opening a URL. */
  linkUri(x: number, y: number, w: number, h: number, uri: string): this {
    const ctx = this.builder.doc.context;
    this.annots.push(
      ctx.register(
        ctx.obj({
          Type: 'Annot',
          Subtype: 'Link',
          Rect: [x, this.height - (y + h), x + w, this.height - y],
          Border: [0, 0, 0],
          A: { Type: 'Action', S: 'URI', URI: PDFString.of(uri) },
        }),
      ),
    );
    return this;
  }

  /** A clickable area jumping to a point on another page of this document. */
  linkToPage(x: number, y: number, w: number, h: number, target: PdfPage, targetY: number): this {
    const ctx = this.builder.doc.context;
    this.annots.push(
      ctx.register(
        ctx.obj({
          Type: 'Annot',
          Subtype: 'Link',
          Rect: [x, this.height - (y + h), x + w, this.height - y],
          Border: [0, 0, 0],
          Dest: [target.page.ref, 'XYZ', 0, target.height - targetY, 0],
        }),
      ),
    );
    return this;
  }

  /** Writes the content stream and annotations. Called once per page by PdfBuilder.save(). */
  finish(): void {
    const ctx = this.builder.doc.context;
    const stream = ctx.flateStream(this.ops.join('\n'));
    this.page.node.addContentStream(ctx.register(stream));
    if (this.annots.length) this.page.node.set(PDFName.of('Annots'), ctx.obj(this.annots));
  }
}

export interface PdfMetadata {
  title: string;
  author?: string;
  subject?: string;
  creator: string;
}

/** The document being built. */
export class PdfBuilder {
  private readonly fonts = new Map<string, Promise<PdfFont | null>>();
  private readonly images = new Map<string, Promise<PdfImage | null>>();
  private readonly pages: PdfPage[] = [];
  private fontCounter = 0;

  private constructor(readonly doc: PDFDocument) {}

  static async create(meta: PdfMetadata): Promise<PdfBuilder> {
    const doc = await PDFDocument.create({ updateMetadata: false });
    doc.setTitle(meta.title, { showInWindowTitleBar: true });
    if (meta.author) doc.setAuthor(meta.author);
    if (meta.subject) doc.setSubject(meta.subject);
    doc.setCreator(meta.creator);
    doc.setProducer('Greek Office (in-browser, pdf-lib)');
    const now = new Date();
    doc.setCreationDate(now);
    doc.setModificationDate(now);
    return new PdfBuilder(doc);
  }

  addPage(width: number, height: number): PdfPage {
    const page = this.doc.addPage([width, height]);
    const p = new PdfPage(this, page, width, height);
    this.pages.push(p);
    return p;
  }

  /**
   * Embed a font face once per key (e.g. its content hash). Resolves null when fontkit cannot
   * read the bytes, so callers can fall back instead of failing the export.
   */
  font(key: string, bytes: () => Uint8Array | Promise<Uint8Array | null> | null): Promise<PdfFont | null> {
    let p = this.fonts.get(key);
    if (!p) {
      p = (async () => {
        try {
          const data = await bytes();
          if (!data) return null;
          const sfnt = woffToSfnt(data);
          const fk = fontkit.create(sfnt);
          if (!fk || !fk.unitsPerEm) return null;
          return new PdfFont(this.doc, fk, sfnt, this.nextTag());
        } catch {
          return null;
        }
      })();
      this.fonts.set(key, p);
    }
    return p;
  }

  /**
   * Embed an image once per key. PNG and JPEG go in as they are; any other format the browser
   * can decode (GIF, BMP, WebP, SVG) is converted to PNG first. Null when it cannot be decoded.
   */
  image(key: string, bytes: Uint8Array, mime: string): Promise<PdfImage | null> {
    let p = this.images.get(key);
    if (!p) {
      p = (async () => {
        try {
          if (mime === 'image/png' || isPng(bytes)) {
            const im = await this.doc.embedPng(bytes);
            return { ref: im.ref, width: im.width, height: im.height };
          }
          if (mime === 'image/jpeg' || isJpeg(bytes)) {
            const im = await this.doc.embedJpg(bytes);
            return { ref: im.ref, width: im.width, height: im.height };
          }
          const png = await rasterToPng(bytes, mime);
          if (!png) return null;
          const im = await this.doc.embedPng(png);
          return { ref: im.ref, width: im.width, height: im.height };
        } catch {
          return null;
        }
      })();
      this.images.set(key, p);
    }
    return p;
  }

  async save(): Promise<Uint8Array> {
    for (const p of this.pages) p.finish();
    for (const f of this.fonts.values()) {
      const font = await f;
      if (font) await font.finish();
    }
    return this.doc.save({ useObjectStreams: true });
  }

  private nextTag(): string {
    // Subset tags are six capital letters (PDF 32000 §9.6.4).
    let n = this.fontCounter++;
    let tag = '';
    for (let i = 0; i < 6; i++) {
      tag = String.fromCharCode(65 + (n % 26)) + tag;
      n = Math.floor(n / 26);
    }
    return tag;
  }
}

function isPng(b: Uint8Array): boolean {
  return b.length > 8 && b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4e && b[3] === 0x47;
}

function isJpeg(b: Uint8Array): boolean {
  return b.length > 3 && b[0] === 0xff && b[1] === 0xd8 && b[2] === 0xff;
}

/** Decode with the browser and re-encode as PNG. Null where the browser cannot decode it. */
async function rasterToPng(bytes: Uint8Array, mime: string): Promise<Uint8Array | null> {
  if (typeof createImageBitmap !== 'function' || typeof OffscreenCanvas === 'undefined') return null;
  const blob = new Blob([bytes as BlobPart], { type: mime || 'application/octet-stream' });
  let bitmap: ImageBitmap;
  if (mime === 'image/svg+xml') {
    // SVG needs an <img> to decode; render it at a fixed, generous size.
    const url = URL.createObjectURL(blob);
    try {
      const img = new Image();
      img.decoding = 'async';
      img.src = url;
      await img.decode();
      const w = Math.min(4096, Math.max(1, Math.round((img.naturalWidth || 512) * 2)));
      const h = Math.min(4096, Math.max(1, Math.round((img.naturalHeight || 512) * 2)));
      bitmap = await createImageBitmap(img, { resizeWidth: w, resizeHeight: h });
    } finally {
      URL.revokeObjectURL(url);
    }
  } else {
    bitmap = await createImageBitmap(blob);
  }
  const canvas = new OffscreenCanvas(bitmap.width, bitmap.height);
  const g = canvas.getContext('2d');
  if (!g) return null;
  g.drawImage(bitmap, 0, 0);
  bitmap.close();
  const out = await canvas.convertToBlob({ type: 'image/png' });
  return new Uint8Array(await out.arrayBuffer());
}
