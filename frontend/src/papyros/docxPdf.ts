/**
 * Word -> PDF, entirely in the browser.
 *
 * The document is laid out by docx-editor's own export layer (the same engine, fonts and
 * pagination as the editor on screen), which publishes every page as plain records: lines,
 * spans with their shaped glyphs, table border strokes, pictures, headers, footers and notes.
 * This module paints those records into a PDF with lib/pdf/writer. Page breaks therefore match
 * the screen, text stays real text (selectable, searchable), and fonts are embedded as subsets.
 *
 * What it cannot draw faithfully is counted and reported back (see DocxPdfResult.simplified),
 * so the caller can say so instead of handing over a quietly incomplete file.
 */
import {
  exportDestinationNamed,
  openFontBackedDocumentForExport,
  type ExportSemanticLayout,
  type FontBackedExportCapabilities,
} from '@docx-editor.dev/core/export';
import {
  baselineShiftPtOf,
  displayText,
  glyphSizeFactorOf,
  groupPictureLayerDrawing,
  TAB_LEADER_GLYPH,
} from '@docx-editor.dev/core/layout';
import { packagedFonts } from '@docx-editor.dev/fonts';
import { fallbackFonts } from '../lib/pdf/fallback';
import { BLACK, PdfBuilder, rgb, type PdfFont, type PdfPage, type PlacedGlyph, type Point, type Rgb } from '../lib/pdf/writer';

// Record types, derived from the layout type itself so they always match the installed engine.
type Layout = ExportSemanticLayout;
type PageRec = Layout['pages'][number];
type Block = PageRec['fragments'][number];
type Para = Extract<Block, { kind: 'paragraph' }>;
type Table = Extract<Block, { kind: 'table' }>;
type Line = Para['lines'][number];
type Span = Line['spans'][number];
type RunStyle = Span['style'];
type InlineDrawing = NonNullable<Line['drawings']>[number];
type AnchoredDrawing = NonNullable<PageRec['anchoredDrawings']>[number];
type Drawing = InlineDrawing | AnchoredDrawing;
type Marker = NonNullable<Para['marker']>;
type Story = NonNullable<InlineDrawing['textboxStory']>;
type NoteArea = NonNullable<PageRec['footnotes']>;
type Edge = NonNullable<Para['borders']>[number]['edge'];

export interface DocxPdfResult {
  pdf: Uint8Array;
  pages: number;
  /** Human-readable notes on what was simplified, e.g. "2 equations". Empty when nothing was. */
  simplified: string[];
}

/** One sentence for the person, naming anything the PDF simplified. */
export function pdfSavedMessage(simplified: string[]): string {
  if (!simplified.length) return 'PDF saved.';
  return `PDF saved. Simplified: ${simplified.join(', ')}. Print gives an exact copy.`;
}

export interface DocxPdfOptions {
  /** PDF title (usually the file name without extension). */
  title: string;
  onProgress?(done: number, total: number): void;
  signal?: AbortSignal;
}

const HIGHLIGHT: Record<string, string> = {
  yellow: 'FFFF00',
  green: '00FF00',
  cyan: '00FFFF',
  magenta: 'FF00FF',
  blue: '0000FF',
  red: 'FF0000',
  darkBlue: '000080',
  darkCyan: '008080',
  darkGreen: '008000',
  darkMagenta: '800080',
  darkRed: '800000',
  darkYellow: '808000',
  darkGray: '808080',
  lightGray: 'C0C0C0',
  black: '000000',
  white: 'FFFFFF',
};

const REVISION_COLOR = rgb('2E74B5');


const nextTick = () => new Promise<void>((r) => setTimeout(r, 0));

export async function docxToPdf(bytes: Uint8Array, options: DocxPdfOptions): Promise<DocxPdfResult> {
  const opened = await openFontBackedDocumentForExport(bytes, {
    fonts: packagedFonts({ onFailure: () => undefined }),
    signal: options.signal,
  });
  if (!opened.ok) {
    throw new Error(
      opened.reason === 'aborted' ? 'The conversion was cancelled.' : 'This document could not be laid out for PDF.',
    );
  }
  const session = opened.session;
  try {
    const layout = await session.layout();
    const painter = new DocxPainter(session, layout, await PdfBuilder.create({
      title: options.title,
      author: layout.documentMetadata?.creator || undefined,
      subject: layout.documentMetadata?.subject || undefined,
      creator: 'Greek Office · Papyros',
    }));
    const total = layout.pages.length;
    for (let i = 0; i < total; i++) {
      if (options.signal?.aborted) throw new Error('The conversion was cancelled.');
      await painter.paintPage(layout.pages[i]);
      options.onProgress?.(i + 1, total);
      await nextTick(); // keep the tab responsive on long documents
    }
    painter.resolveInternalLinks();
    for (const w of layout.contentWarnings ?? []) {
      painter.note(
        w.code === 'legacy-textbox'
          ? 'legacy text boxes'
          : w.code === 'legacy-drawing'
            ? 'legacy drawings'
            : w.code === 'scan-limit'
              ? 'content beyond the scan limit'
              : 'unusual style separators',
      );
    }
    const pdf = await painter.builder.save();
    return { pdf, pages: total, simplified: painter.simplifiedNotes() };
  } finally {
    session.dispose();
  }
}

interface Frame {
  /** Offset from record coordinates to page coordinates. */
  ox: number;
  oy: number;
}

interface PendingLink {
  page: PdfPage;
  x: number;
  y: number;
  w: number;
  h: number;
  anchor: string;
}

class DocxPainter {
  private readonly pdfPages: PdfPage[] = [];
  private readonly counts = new Map<string, number>();
  private readonly pendingLinks: PendingLink[] = [];

  constructor(
    private readonly session: FontBackedExportCapabilities,
    private readonly layout: Layout,
    readonly builder: PdfBuilder,
  ) {}

  note(what: string, n = 1) {
    this.counts.set(what, (this.counts.get(what) ?? 0) + n);
  }

  simplifiedNotes(): string[] {
    return [...this.counts.entries()].map(([what, n]) => (n > 1 ? `${what} (${n})` : what));
  }

  async paintPage(page: PageRec): Promise<void> {
    const pg = this.builder.addPage(page.box.width, page.box.height);
    this.pdfPages.push(pg);
    const px = page.box.x;
    const py = page.box.y;
    const body: Frame = { ox: page.contentBox.x - px, oy: page.contentBox.y - py };

    if (page.pageBorders?.zOrder === 'back') this.pageBorders(pg, page);

    // Floating objects behind the text first.
    const behind = (d: AnchoredDrawing) => d.behindDocument;
    for (const story of [page.header, page.footer]) {
      if (!story) continue;
      const f = { ox: story.box.x - px, oy: story.box.y - py };
      for (const d of story.anchoredDrawings ?? []) if (behind(d)) await this.drawing(pg, d, f);
    }
    for (const d of page.anchoredDrawings ?? []) if (behind(d)) await this.drawing(pg, d, body);

    for (const story of [page.header, page.footer]) {
      if (!story) continue;
      const f = { ox: story.box.x - px, oy: story.box.y - py };
      await this.blocks(pg, story.fragments, f);
      for (const d of story.anchoredDrawings ?? []) if (!behind(d)) await this.drawing(pg, d, f);
    }

    await this.blocks(pg, page.fragments, body);

    for (const sep of page.columnSeparators ?? []) {
      pg.fillRect(body.ox + sep.x, body.oy + sep.y, Math.max(sep.width, 0.5), sep.height, BLACK);
    }

    for (const area of [page.footnotes, page.endnotes]) if (area) await this.noteArea(pg, area, px, py);

    for (const d of page.anchoredDrawings ?? []) if (!behind(d)) await this.drawing(pg, d, body);

    if (page.pageBorders && page.pageBorders.zOrder !== 'back') this.pageBorders(pg, page);
  }

  resolveInternalLinks() {
    for (const l of this.pendingLinks) {
      const dest = exportDestinationNamed(this.layout, l.anchor);
      if (!dest) continue;
      const target = this.pdfPages[dest.pageIndex];
      const page = this.layout.pages[dest.pageIndex];
      if (!target || !page) continue;
      const y = page.contentBox.y - page.box.y + dest.pageContent.y;
      l.page.linkToPage(l.x, l.y, l.w, l.h, target, Math.max(0, y - 6));
    }
  }

  // ---------------------------------------------------------------- blocks

  private async blocks(pg: PdfPage, blocks: readonly Block[], f: Frame): Promise<void> {
    for (const b of blocks) {
      if (b.kind === 'paragraph') await this.paragraph(pg, b, f);
      else await this.table(pg, b, f);
    }
  }

  private async paragraph(pg: PdfPage, p: Para, f: Frame): Promise<void> {
    if (p.outOfFlow && !p.positionedFrame) return;
    const clip = !!p.clipToBox;
    if (clip) pg.save().rect(f.ox + p.box.x, f.oy + p.box.y, p.box.width, p.box.height).clip();

    if (p.shading && p.shadingBox) {
      const b = p.shadingBox;
      pg.fillRect(f.ox + b.x, f.oy + b.y, b.width, b.height, rgb(p.shading));
    }
    const strokes = p.borders ?? (p.bottomBorder ? [{ side: 'bottom' as const, ...p.bottomBorder }] : []);
    for (const s of strokes) this.borderBox(pg, s.edge, f.ox + s.box.x, f.oy + s.box.y, s.box.width, s.box.height);

    if (p.marker) await this.marker(pg, p, p.marker, f);

    for (const line of p.lines) await this.line(pg, line, f);

    if (clip) pg.restore();
  }

  private async table(pg: PdfPage, t: Table, f: Frame): Promise<void> {
    for (const row of t.rows) {
      for (const cell of row.cells) {
        if (cell.paintInert) continue;
        const b = cell.box;
        if (cell.shading) pg.fillRect(f.ox + b.x, f.oy + b.y, b.width, b.height, rgb(cell.shading));
        if (cell.textDirection === 'btLr') {
          // Bottom-to-top cells: rotate the content about the cell centre.
          const cx = f.ox + b.x + b.width / 2;
          const cy = f.oy + b.y + b.height / 2;
          pg.save().rotateAbout(-90, cx, cy);
          await this.blocks(pg, cell.blocks, f);
          pg.restore();
        } else {
          await this.blocks(pg, cell.blocks, f);
        }
      }
      // Borders after every cell of the row, so a neighbour's fill never covers a shared edge.
      for (const cell of row.cells) {
        if (cell.paintInert || !cell.borders) continue;
        const b = cell.box;
        for (const s of cell.borders.strokes ?? []) {
          const color = rgb(s.color);
          const x = f.ox + b.x + s.x;
          const y = f.oy + b.y + s.y;
          if (s.cssStyle === 'solid') pg.fillRect(x, y, s.width, s.height, color);
          else {
            const horizontal = s.width >= s.height;
            const w = horizontal ? s.height : s.width;
            const dash = s.cssStyle === 'dotted' ? [w, w * 1.5] : [w * 3, w * 2];
            if (horizontal) pg.line(x, y + s.height / 2, x + s.width, y + s.height / 2, color, w, dash);
            else pg.line(x + s.width / 2, y, x + s.width / 2, y + s.height, color, w, dash);
          }
        }
        for (const d of cell.borders.diagonals ?? []) {
          pg.save().rect(f.ox + b.x, f.oy + b.y, d.clipWidth || b.width, d.clipHeight || b.height).clip();
          pg.line(
            f.ox + b.x + d.x1,
            f.oy + b.y + d.y1,
            f.ox + b.x + d.x2,
            f.oy + b.y + d.y2,
            rgb(d.color),
            d.widthPt,
            d.cssStyle === 'solid' ? null : d.cssStyle === 'dotted' ? [d.widthPt, d.widthPt * 1.5] : [d.widthPt * 3, d.widthPt * 2],
          );
          pg.restore();
        }
      }
    }
  }

  private async line(pg: PdfPage, line: Line, f: Frame): Promise<void> {
    // Backgrounds under every span first, then pictures, then glyphs, then lines on top.
    for (const span of line.spans) {
      if (span.style.hidden) continue;
      const hl = span.style.highlight ? HIGHLIGHT[span.style.highlight] : null;
      const fill = hl ?? span.style.shading;
      if (fill) pg.fillRect(f.ox + span.box.x, f.oy + span.box.y, span.box.width, span.box.height, rgb(fill));
    }
    for (const d of line.drawings ?? []) await this.drawing(pg, d, f);

    let linkRun: { id: string; x: number; y: number; r: number; b: number; href: string | null; anchor?: string } | null = null;
    const flushLink = () => {
      if (!linkRun) return;
      const { x, y, r, b, href, anchor } = linkRun;
      if (href && /^(https?:|mailto:)/i.test(href)) pg.linkUri(x, y, r - x, b - y, href);
      else if (anchor) this.pendingLinks.push({ page: pg, x, y, w: r - x, h: b - y, anchor });
      linkRun = null;
    };

    for (const span of line.spans) {
      if (span.style.hidden) continue;
      if (span.equation) {
        this.note('equations');
        continue;
      }
      if (span.unsupported) {
        this.note('unsupported inline content');
        continue;
      }
      const font = await this.spanText(pg, span, line, f);
      this.decorations(pg, span, line, f, font);

      const link = span.link;
      if (link && (link.href || link.anchor)) {
        const x = f.ox + span.box.x;
        const y = f.oy + span.box.y;
        if (linkRun && linkRun.id === link.id) {
          linkRun.r = x + span.box.width;
          linkRun.b = Math.max(linkRun.b, y + span.box.height);
        } else {
          flushLink();
          linkRun = { id: link.id, x, y, r: x + span.box.width, b: y + span.box.height, href: link.href, anchor: link.anchor };
        }
      } else flushLink();
    }
    flushLink();
  }

  // ---------------------------------------------------------------- text

  private fillOf(style: RunStyle, span?: Span): Rgb {
    if (span?.revisions?.some((r) => r.kind !== 'format')) return REVISION_COLOR;
    return rgb(style.color);
  }

  /**
   * Draw one span's glyphs. Returns the font used (for underline metrics), or null when the span
   * drew nothing.
   */
  private async spanText(pg: PdfPage, span: Span, line: Line, f: Frame): Promise<PdfFont | null> {
    const style = span.style;
    const baseY = f.oy + line.box.y + line.baseline - baselineShiftPtOf(style);
    const x0 = f.ox + span.box.x + (span.glyphOffsetPt ?? 0);
    const factor = glyphSizeFactorOf(style);
    const size = style.fontSizePt * factor;
    const hs = (style.horizontalScalePercent || 100) / 100;
    const fill = this.fillOf(style, span);
    const outline = style.textOutline ? { color: rgb(style.textOutline.color), width: style.textOutline.widthPt } : null;

    if (span.tabLeader && span.tabLeaderAdvancePt) {
      const ch = TAB_LEADER_GLYPH.get(span.tabLeader) ?? '.';
      const face = await this.faceFor(span);
      if (face) {
        const g = face.font.glyphForCodePoint(ch.codePointAt(0) ?? 46);
        const cid = face.cid(g.id, ch);
        const step = span.tabLeaderAdvancePt;
        const glyphs: PlacedGlyph[] = [];
        for (let x = x0; x + step <= x0 + span.box.width + 0.01; x += step) glyphs.push({ cid, x, y: baseY });
        pg.glyphs(face, { size, hScale: hs, fill }, glyphs);
      }
      return face;
    }

    let text = span.text;
    if (span.optionalHyphenBreak) text = '-';
    if (!text || /^[\s­\f​]*$/.test(text) && !span.optionalHyphenBreak) return this.faceFor(span);

    const shaped = span.optionalHyphenBreak ? null : this.session.shapeLaidOutText(span);
    if (!shaped || shaped.run.glyphs.some((g) => g.id === 0)) {
      return this.fallbackText(pg, displayText(text, style), x0, baseY, size, span.box.width, fill);
    }

    const run = shaped.run;
    const sc = shaped.fixedPointScale || 1;
    const cs = style.characterSpacingPt || 0;
    // glyph -> (cluster text if this glyph starts its cluster, is-last-glyph-of-cluster)
    const startText = new Map<number, string>();
    const clusterEnd = new Set<number>();
    for (const c of run.clusters) {
      startText.set(c.glyphStart, run.text.slice(c.textStart, c.textEnd));
      clusterEnd.add(c.glyphEnd - 1);
    }
    const faceOf = (gi: number) => {
      if (!shaped.fonts) return shaped.font;
      const k = run.fontSpans.findIndex((s) => gi >= s.glyphStart && gi < s.glyphEnd);
      return shaped.fonts[k] ?? shaped.font;
    };

    // Group consecutive glyphs that share a face and a draw scale; each group is one TJ run.
    let pen = 0;
    let group: { face: PdfFont; scale: number; glyphs: PlacedGlyph[] } | null = null;
    const flush = () => {
      if (group) pg.glyphs(group.face, { size: size * group.scale, hScale: hs, fill, stroke: outline }, group.glyphs);
      group = null;
    };
    let used: PdfFont | null = null;
    for (let gi = 0; gi < run.glyphs.length; gi++) {
      const g = run.glyphs[gi];
      const ident = faceOf(gi);
      const face = await this.builder.font(ident.identity, () => this.session.admittedFontFace(ident.request)?.bytes ?? null);
      const adv = ((run.exactAdvancesX?.[gi] ?? g.advanceX) / sc) * factor * hs;
      if (!face) {
        pen += adv + (clusterEnd.has(gi) ? cs : 0);
        continue;
      }
      used ??= face;
      const scale = g.drawScale ?? 1;
      if (!group || group.face !== face || group.scale !== scale) {
        flush();
        group = { face, scale, glyphs: [] };
      }
      group.glyphs.push({
        cid: face.cid(g.id, startText.get(gi)),
        x: x0 + pen + (g.offsetX / sc) * factor * hs,
        y: baseY - (g.offsetY / sc) * factor,
      });
      pen += adv + (clusterEnd.has(gi) ? cs : 0);
    }
    flush();
    return used;
  }

  /** The admitted face a span's text resolves to (for leaders, hyphens and decoration metrics). */
  private async faceFor(span: Span): Promise<PdfFont | null> {
    const shaped = this.session.shapeLaidOutText(span);
    if (!shaped) return null;
    const ident = shaped.font;
    return this.builder.font(ident.identity, () => this.session.admittedFontFace(ident.request)?.bytes ?? null);
  }

  /**
   * Text the document's font cannot draw (missing glyphs) or that layout measured without
   * shaping. Drawn with a bundled fallback face, squeezed or stretched into the span's box so
   * nothing after it moves.
   */
  private async fallbackText(
    pg: PdfPage,
    text: string,
    x0: number,
    baseY: number,
    size: number,
    boxWidth: number,
    fill: Rgb,
  ): Promise<PdfFont | null> {
    const fonts = await fallbackFonts(this.builder);
    if (!fonts.length || !/\S/.test(text)) return null;
    // Split into runs, each in the first fallback face that has the character.
    const pick = (ch: string) => {
      if (/\s/.test(ch)) return null;
      const cp = ch.codePointAt(0) ?? 0;
      return fonts.find((pf) => pf.font.hasGlyphForCodePoint(cp)) ?? null;
    };
    const runs: { font: PdfFont; text: string }[] = [];
    let missing = 0;
    for (const ch of text) {
      let font = pick(ch);
      if (!font) {
        if (!/\s/.test(ch)) missing++;
        font = runs.length ? runs[runs.length - 1].font : fonts[0];
      }
      const last = runs[runs.length - 1];
      if (last && last.font === font) last.text += ch;
      else runs.push({ font, text: ch });
    }
    if (missing) this.note('characters with no available font', missing);

    const laid = runs.map((r) => ({ font: r.font, run: r.font.font.layout(r.text) }));
    const k = (pf: PdfFont) => size / pf.unitsPerEm;
    const natural = laid.reduce((a, l) => a + l.run.positions.reduce((b, p) => b + p.xAdvance, 0) * k(l.font), 0);
    const fit = natural > 0 && boxWidth > 0 ? Math.min(1.4, Math.max(0.6, boxWidth / natural)) : 1;
    let pen = 0;
    for (const { font, run } of laid) {
      const glyphs: PlacedGlyph[] = [];
      run.glyphs.forEach((g, i) => {
        const pos = run.positions[i];
        glyphs.push({
          cid: font.cid(g.id, g.codePoints.length ? String.fromCodePoint(...g.codePoints) : undefined),
          x: x0 + (pen + pos.xOffset * k(font)) * fit,
          y: baseY - pos.yOffset * k(font),
        });
        pen += pos.xAdvance * k(font);
      });
      pg.glyphs(font, { size, hScale: fit, fill }, glyphs);
    }
    return laid[0]?.font ?? null;
  }

  private decorations(pg: PdfPage, span: Span, line: Line, f: Frame, font: PdfFont | null) {
    const style = span.style;
    const revisionDeleted = span.revisions?.some((r) => r.kind === 'delete' || r.kind === 'moveFrom');
    const revisionInserted = span.revisions?.some((r) => r.kind === 'insert' || r.kind === 'moveTo');
    const underline = style.underline ?? (revisionInserted ? { variant: 'single', color: null } : null);
    const strike = style.strike || style.doubleStrike || revisionDeleted;
    if (!underline && !strike) return;
    if (span.lineEndWhitespace && /^\s*$/.test(span.text)) return;

    const size = style.fontSizePt * glyphSizeFactorOf(style);
    const baseY = f.oy + line.box.y + line.baseline - baselineShiftPtOf(style);
    const x = f.ox + span.box.x;
    const w = span.box.width;
    const color = this.fillOf(style, span);
    const metrics = font?.underline ?? { position: -0.1, thickness: 0.05 };
    const t = Math.max(0.4, metrics.thickness * size);

    if (underline && underline.variant !== 'none' && !(underline.variant === 'words' && /^\s+$/.test(span.text))) {
      const v = underline.variant;
      const c = underline.color ? rgb(underline.color) : color;
      const y = baseY - metrics.position * size;
      const th = /thick|Heavy/i.test(v) ? t * 2 : t;
      const dash = /^dot(?!Dash)|dotted/i.test(v)
        ? [th, th * 1.5]
        : /dotDotDash/i.test(v)
          ? [th * 3, th * 2, th, th * 2, th, th * 2]
          : /dotDash/i.test(v)
            ? [th * 3, th * 2, th, th * 2]
            : /dash/i.test(v)
              ? [th * 3, th * 2]
              : null;
      if (v === 'double') {
        pg.line(x, y - th, x + w, y - th, c, th * 0.8);
        pg.line(x, y + th * 1.2, x + w, y + th * 1.2, c, th * 0.8);
      } else pg.line(x, y + th / 2, x + w, y + th / 2, c, th, dash);
    }
    if (strike) {
      const y = baseY - size * 0.28;
      if (style.doubleStrike) {
        pg.line(x, y - t, x + w, y - t, color, t * 0.8);
        pg.line(x, y + t, x + w, y + t, color, t * 0.8);
      } else pg.line(x, y, x + w, y, color, t);
    }
  }

  private async marker(pg: PdfPage, p: Para, m: Marker, f: Frame): Promise<void> {
    const first = p.lines[0];
    const baseY = first ? f.oy + first.box.y + first.baseline : f.oy + m.box.y + m.box.height * 0.8;
    if (m.picture?.resource.kind === 'ready') {
      const bytes = this.session.validatedImageBytes(m.picture);
      const res = m.picture.resource;
      const img = bytes ? await this.builder.image(res.contentId, bytes, res.mime) : null;
      if (img) {
        const b = m.picture.box;
        pg.imageInBox(img, f.ox + b.x, f.oy + b.y, b.width, b.height);
        return;
      }
    }
    for (const piece of m.pieces ?? [m]) {
      const pseudo = {
        text: piece.text,
        style: piece.style,
        box: piece.box,
        range: p.range,
        props: [],
      } as unknown as Span;
      const fakeLine = { box: { ...piece.box, y: baseY - f.oy }, baseline: 0 } as unknown as Line;
      await this.spanText(pg, pseudo, fakeLine, f);
    }
  }

  // ---------------------------------------------------------------- borders

  private borderBox(pg: PdfPage, edge: Edge, x: number, y: number, w: number, h: number) {
    const color = rgb(edge.color);
    const v = edge.val;
    if (!v || v === 'nil' || v === 'none') return;
    if (v === 'double' && Math.min(w, h) >= 1.5) {
      const horizontal = w >= h;
      if (horizontal) {
        const t = h / 3;
        pg.fillRect(x, y, w, t, color).fillRect(x, y + h - t, w, t, color);
      } else {
        const t = w / 3;
        pg.fillRect(x, y, t, h, color).fillRect(x + w - t, y, t, h, color);
      }
      return;
    }
    if (/dot|dash/i.test(v)) {
      const horizontal = w >= h;
      const t = horizontal ? h : w;
      const dash = /dot(?!Dash)/i.test(v) ? [t, t * 1.5] : [t * 3, t * 2];
      if (horizontal) pg.line(x, y + h / 2, x + w, y + h / 2, color, t, dash);
      else pg.line(x + w / 2, y, x + w / 2, y + h, color, t, dash);
      return;
    }
    pg.fillRect(x, y, w, h, color);
  }

  private pageBorders(pg: PdfPage, page: PageRec) {
    for (const s of page.pageBorders?.strokes ?? []) this.borderBox(pg, s.edge, s.box.x, s.box.y, s.box.width, s.box.height);
  }

  private async noteArea(pg: PdfPage, area: NoteArea, px: number, py: number): Promise<void> {
    for (const sep of [area.separator, area.continuationNotice]) {
      if (!sep) continue;
      const f = { ox: sep.box.x - px, oy: sep.box.y - py };
      if (sep.fragments.length && !sep.synthetic) await this.blocks(pg, sep.fragments, f);
      else if (sep.ruleStyle) {
        const color = rgb(sep.ruleColor ?? null);
        const y = f.oy + sep.box.height / 2;
        pg.line(f.ox, y, f.ox + sep.box.width, y, color, 0.5);
        if (sep.ruleStyle === 'double') pg.line(f.ox, y + 1.5, f.ox + sep.box.width, y + 1.5, color, 0.5);
      }
    }
    for (const n of area.notes) await this.blocks(pg, n.fragments, { ox: n.box.x - px, oy: n.box.y - py });
  }

  // ---------------------------------------------------------------- drawings

  private async drawing(pg: PdfPage, d: Drawing, f: Frame): Promise<void> {
    const g = d.geometry;
    const at = (p: Point): Point => ({ x: f.ox + p.x, y: f.oy + p.y });
    const corners = g.transformedCorners.length === 4 ? g.transformedCorners.map(at) : null;
    let drew = false;

    // Clip to the drawing's crop/clip polygon when it has one.
    const clip = g.clipPolygon && g.clipPolygon.length >= 3 ? g.clipPolygon.map(at) : null;
    if (clip) pg.save().polygon(clip).clip();

    if (d.vectorShape && corners) {
      this.vectorShape(pg, d.vectorShape, corners);
      drew = true;
    }

    let resource = d.resource;
    let source: Drawing = d;
    if (d.chart || d.placeholderGraphicKind === 'chart') {
      const fb = d.chartFallbackResource ?? (await this.session.chartFallbackResource?.(d));
      if (fb?.kind === 'ready') {
        resource = fb;
        source = { ...d, resource: fb } as Drawing;
      } else if (resource.kind !== 'ready') this.note('charts');
    }
    if (resource.kind === 'ready') {
      drew = (await this.picture(pg, source, resource.contentId, resource.mime, d.effects?.opacity, at)) || drew;
    } else if (resource.kind === 'unrenderable' && !d.vectorShape && !d.textboxStory && !d.textArt) {
      if (resource.mime === 'image/x-emf' || resource.mime === 'image/x-wmf') this.note('EMF/WMF pictures');
      else if (resource.reason !== 'non-picture-graphic') this.note('pictures that could not be decoded');
    }
    for (const layer of d.groupPictureLayers ?? []) {
      const ld = groupPictureLayerDrawing(d, layer);
      if (ld.resource.kind === 'ready') {
        drew = (await this.picture(pg, ld, ld.resource.contentId, ld.resource.mime, ld.effects?.opacity, at)) || drew;
      }
    }

    if (d.textArt) drew = this.textArt(pg, d.textArt, d, f) || drew;

    if (clip) pg.restore();

    if (d.textboxStory) {
      await this.textbox(pg, d.textboxStory, d.x, d.y, d.width, d.height, f, !!d.vectorShape);
      drew = true;
    }
    for (const m of d.groupTextboxStories ?? []) {
      pg.save().rotateAbout(m.rotationDegrees, f.ox + d.x + m.box.x + m.box.width / 2, f.oy + d.y + m.box.y + m.box.height / 2);
      await this.textbox(pg, m.story, d.x + m.box.x, d.y + m.box.y, m.box.width, m.box.height, f, true);
      pg.restore();
      drew = true;
    }

    if (!drew && d.placeholderGraphicKind) this.note(String(d.placeholderGraphicKind).replace(/([a-z])([A-Z])/g, '$1 $2').toLowerCase() + ' objects');
  }

  private async picture(
    pg: PdfPage,
    d: Drawing,
    key: string,
    mime: string,
    opacity: number | undefined,
    at: (p: Point) => Point,
  ): Promise<boolean> {
    const bytes = this.session.validatedImageBytes(d);
    if (!bytes) return false;
    const img = await this.builder.image(key, bytes, mime);
    if (!img) {
      this.note('pictures that could not be decoded');
      return false;
    }
    const c = (d.geometry.imageTransformCorners?.length === 4 ? d.geometry.imageTransformCorners : d.geometry.transformedCorners).map(at);
    if (c.length !== 4) return false;
    if (opacity !== undefined && opacity < 1) pg.save().alpha(opacity);
    pg.image(img, c[0], c[1], c[3]);
    if (opacity !== undefined && opacity < 1) pg.restore();
    return true;
  }

  private vectorShape(pg: PdfPage, v: NonNullable<Drawing['vectorShape']>, corners: Point[]) {
    const [tl, tr, , bl] = corners;
    const cx = v.extentEmu.cx || 1;
    const cy = v.extentEmu.cy || 1;
    const map = (p: { x: number; y: number }): Point => {
      const u = p.x / cx;
      const w = p.y / cy;
      return { x: tl.x + u * (tr.x - tl.x) + w * (bl.x - tl.x), y: tl.y + u * (tr.y - tl.y) + w * (bl.y - tl.y) };
    };
    // Points per EMU along the drawing's own axes, for stroke widths.
    const scale = Math.hypot(tr.x - tl.x, tr.y - tl.y) / cx;
    for (const comp of v.components) {
      if (comp.shadowOnly) continue;
      const paths = comp.subpathsEmu;
      const closed = comp.subpathsClosed;
      const outline = (fn: () => void) => {
        pg.save();
        if (comp.clipEmu?.length) {
          for (const c of comp.clipEmu) pg.polygon(c.map(map));
          pg.clip();
        }
        fn();
        pg.restore();
      };
      if (comp.fillHex && paths.some((_, i) => closed?.[i] !== false)) {
        outline(() => {
          pg.alpha(comp.fillAlpha ?? 1);
          pg.fillColor(rgb(comp.fillHex));
          paths.forEach((sp, i) => closed?.[i] !== false && pg.polygon(sp.map(map), true));
          pg.fill(true);
        });
      }
      if (comp.strokeHex && comp.strokeWidthEmu > 0) {
        const width = comp.strokeWidthEmu * scale;
        outline(() => {
          pg.alpha(1, comp.strokeAlpha ?? 1);
          pg.strokeColor(rgb(comp.strokeHex)).lineWidth(width);
          if (comp.strokeDash?.length) pg.dash(comp.strokeDash.map((x) => x * width));
          if (comp.strokeCap) pg.lineCap(comp.strokeCap);
          if (comp.strokeJoin) pg.lineJoin(comp.strokeJoin);
          paths.forEach((sp, i) => pg.polygon(sp.map(map), closed?.[i] !== false));
          pg.stroke();
        });
      }
      if (comp.arrowheadsEmu?.length && comp.strokeHex) {
        pg.save().alpha(comp.strokeAlpha ?? 1).fillColor(rgb(comp.strokeHex));
        for (const a of comp.arrowheadsEmu) pg.polygon(a.map(map), true);
        pg.fill().restore();
      }
    }
  }

  private textArt(pg: PdfPage, art: NonNullable<Drawing['textArt']>, d: Drawing, f: Frame): boolean {
    if (!art.path || !art.fit) {
      this.note('WordArt');
      return false;
    }
    const x = f.ox + d.x;
    const y = f.oy + d.y;
    const fit = art.fit;
    pg.save().rotateAbout(art.rotationDegrees, x + d.width / 2, y + d.height / 2);
    const map = (px: number, py: number) => ({ x: x + fit.scaleX * px + fit.offsetX, y: y + fit.scaleY * py + fit.offsetY });
    pg.svgPath(art.path, map);
    if (art.fillHex && art.strokeHex) {
      pg.alpha(art.fillAlpha, art.strokeAlpha).fillColor(rgb(art.fillHex)).strokeColor(rgb(art.strokeHex)).lineWidth(art.strokeWidthPt);
      pg.fillStroke();
    } else if (art.fillHex) pg.alpha(art.fillAlpha).fillColor(rgb(art.fillHex)).fill();
    else if (art.strokeHex) pg.alpha(1, art.strokeAlpha).strokeColor(rgb(art.strokeHex)).lineWidth(art.strokeWidthPt).stroke();
    else pg.endPath();
    pg.restore();
    return true;
  }

  private async textbox(pg: PdfPage, s: Story, x: number, y: number, w: number, h: number, f: Frame, chromeDrawn: boolean) {
    const bx = f.ox + x;
    const by = f.oy + y;
    const width = s.extentWidth ?? w;
    const height = s.extentHeight ?? h;
    if (!chromeDrawn && !s.chromeByVector && (s.fillHex || (s.strokeHex && s.strokeWidthPt > 0))) {
      pg.save().rotateAbout(s.chromeRotationDegrees ?? 0, bx + width / 2, by + height / 2);
      pg.rect(bx, by, width, height);
      if (s.fillHex) pg.fillColor(rgb(s.fillHex));
      if (s.strokeHex && s.strokeWidthPt > 0) {
        pg.strokeColor(rgb(s.strokeHex)).lineWidth(s.strokeWidthPt);
        if (s.strokeDashPt?.length) pg.dash(s.strokeDashPt);
      }
      if (s.fillHex && s.strokeHex && s.strokeWidthPt > 0) pg.fillStroke();
      else if (s.fillHex) pg.fill();
      else pg.stroke();
      pg.restore();
    }
    const inner: Frame = { ox: bx + s.contentOffset.x, oy: by + s.contentOffset.y };
    pg.save();
    if (s.contentRotation) pg.rotateAbout(s.contentRotation.degrees, bx + s.contentRotation.pivot.x, by + s.contentRotation.pivot.y);
    if (!s.overflowVisible) pg.rect(bx, by, width, height).clip();
    await this.blocks(pg, s.fragments, inner);
    pg.restore();
  }
}

