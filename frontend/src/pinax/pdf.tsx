/**
 * Presentation -> PDF, in the browser. One page per slide at the slide's own size.
 *
 * Shapes, fills, outlines and pictures are drawn from the model (the same geometry as the
 * screen). Text is placed from the rendered slide: each slide is rendered off-screen at 1:1
 * without rotation, every word's box is measured, and the word is drawn there with the very font
 * file the screen used. Rotation is then applied in the PDF, so rotated text stays real text.
 */
import { flushSync } from 'react-dom';
import { createRoot } from 'react-dom/client';
import { fallbackFonts, fallbackFor } from '../lib/pdf/fallback';
import { BLACK, PdfBuilder, type PdfFont, type PdfPage, type PlacedGlyph, type Rgb, rgb } from '../lib/pdf/writer';
import type { FontRegistry } from './fonts';
import type { MediaStore } from './media';
import { type Deck, type GroupEl, type Slide, type SlideElement, type Xfrm, emuToPx } from './model';
import { arrowHead, dashPattern, endPoints, geometryPaths, isLinePreset, type Cmd } from './pptx/geometry';
import { type Fill, type Line, fillColor } from './pptx/theme';
import { SlideView, mapIntoGroup, type RenderEnv } from './render';

const K = 0.75; // points per CSS px

interface Job {
  builder: PdfBuilder;
  deck: Deck;
  media: MediaStore;
  fonts: FontRegistry;
  fallbacks: PdfFont[];
  stage: HTMLElement;
}

const frame = () => new Promise<void>((r) => requestAnimationFrame(() => r()));

export async function presentationToPdf(opts: {
  deck: Deck;
  slides: Slide[];
  media: MediaStore;
  fonts: FontRegistry;
  title: string;
  onProgress?(done: number, total: number): void;
}): Promise<Uint8Array> {
  const { deck, slides, media, fonts } = opts;
  if (!slides.length) throw new Error('There are no visible slides to export.');
  const builder = await PdfBuilder.create({ title: opts.title, creator: 'Greek Office · Pinax' });
  const host = document.createElement('div');
  host.className = 'px-pdf-stage';
  host.setAttribute('aria-hidden', 'true');
  document.body.appendChild(host);
  const root = createRoot(host);
  const env: RenderEnv = { deck, media, fonts, mode: 'pdf', flat: true };
  const job: Job = { builder, deck, media, fonts, fallbacks: await fallbackFonts(builder), stage: host };
  try {
    for (let i = 0; i < slides.length; i++) {
      const slide = slides[i];
      flushSync(() => root.render(<SlideView env={env} slide={slide} scale={1} />));
      await document.fonts.ready;
      await Promise.all(Array.from(host.querySelectorAll('img')).map((img) => img.decode().catch(() => undefined)));
      await frame();
      const slideEl = host.querySelector('.px-slide') as HTMLElement;
      await paintSlide(job, slide, slideEl);
      opts.onProgress?.(i + 1, slides.length);
    }
  } finally {
    root.unmount();
    host.remove();
  }
  return builder.save();
}

async function paintSlide(job: Job, slide: Slide, slideEl: HTMLElement) {
  const { deck } = job;
  const W = emuToPx(deck.size.cx) * K;
  const H = emuToPx(deck.size.cy) * K;
  const pg = job.builder.addPage(W, H);
  const layout = deck.layouts.get(slide.layout);
  const master = layout ? deck.masters.get(layout.master) : undefined;
  const bg = slide.background ?? layout?.background ?? master?.background ?? null;
  await paintFill(job, pg, bg, () => pg.rect(0, 0, W, H), { x: 0, y: 0, w: W, h: H }, { r: 1, g: 1, b: 1 });

  const els: SlideElement[] = [];
  if (slide.showMasterShapes && layout?.showMasterShapes !== false && master) els.push(...master.decorations);
  if (slide.showMasterShapes && layout) els.push(...layout.decorations);
  els.push(...slide.elements);
  for (const el of els) {
    const dom = slideEl.querySelector(`[data-el="${el.id}"]`) as HTMLElement | null;
    await paintElement(job, pg, el, el.xfrm, dom);
  }
}

interface Box {
  x: number;
  y: number;
  w: number;
  h: number;
}

function boxOf(x: Xfrm): Box {
  return { x: emuToPx(x.x) * K, y: emuToPx(x.y) * K, w: emuToPx(x.w) * K, h: emuToPx(x.h) * K };
}

async function paintElement(job: Job, pg: PdfPage, el: SlideElement, xfrm: Xfrm, dom: HTMLElement | null) {
  const b = boxOf(xfrm);
  pg.save().rotateAbout(xfrm.rot, b.x + b.w / 2, b.y + b.h / 2);
  try {
    if (el.type === 'group') {
      await paintGroup(job, pg, el, b, dom);
    } else if (el.type === 'shape') {
      const lineOnly = el.connector || isLinePreset(el.geom);
      paintGeometry(pg, el.geom, b, xfrm, lineOnly ? null : el.fill, el.line, job);
      if (!lineOnly && el.fill?.kind === 'image') await paintImageFill(job, pg, el.fill, el.geom, b, xfrm);
      if (dom) await paintText(job, pg, dom.querySelector(':scope > .px-text'), dom, b.x, b.y);
    } else if (el.type === 'picture') {
      await paintPicture(job, pg, el, b, xfrm);
    } else if (el.type === 'table') {
      if (dom) await paintTable(job, pg, dom, b);
    } else if (el.type === 'frame') {
      const bytes = el.preview ? job.media.bytes(el.preview) : null;
      const mime = el.preview ? job.media.mimeOf(el.preview) : null;
      const img = bytes && mime ? await job.builder.image(el.preview!, bytes, mime) : null;
      if (img) pg.imageInBox(img, b.x, b.y, b.w, b.h);
      else if (el.kind !== 'other') {
        pg.save().fillColor([0.95, 0.94, 0.92]).strokeColor([0.7, 0.67, 0.62]).lineWidth(0.75).rect(b.x, b.y, b.w, b.h).fillStroke().restore();
        if (dom) await paintText(job, pg, dom, dom, b.x, b.y);
      }
    }
  } finally {
    pg.restore();
  }
}

async function paintGroup(job: Job, pg: PdfPage, g: GroupEl, b: Box, dom: HTMLElement | null) {
  pg.save().transform([1, 0, 0, 1, b.x, b.y]);
  for (const c of g.children) {
    const childDom = dom?.querySelector(`:scope > [data-el="${c.id}"]`) as HTMLElement | null;
    await paintElement(job, pg, c, mapIntoGroup(g, c.xfrm), childDom);
  }
  pg.restore();
}

// ---------------------------------------------------------------- geometry

function emitPath(pg: PdfPage, cmds: Cmd[], b: Box, x: Xfrm) {
  const fx = (v: number) => b.x + (x.flipH ? b.w - v : v);
  const fy = (v: number) => b.y + (x.flipV ? b.h - v : v);
  for (const c of cmds) {
    if (c[0] === 'M') pg.moveTo(fx(c[1]), fy(c[2]));
    else if (c[0] === 'L') pg.lineTo(fx(c[1]), fy(c[2]));
    else if (c[0] === 'C') pg.curveTo(fx(c[1]), fy(c[2]), fx(c[3]), fy(c[4]), fx(c[5]), fy(c[6]));
    else pg.close();
  }
}

function solidOf(fill: Fill | null): { rgb: Rgb; alpha: number } | null {
  const c = fill ? fillColor(fill) : null;
  if (!c) return null;
  return { rgb: rgb(c.hex), alpha: c.alpha };
}

function paintGeometry(pg: PdfPage, geom: Parameters<typeof geometryPaths>[0], b: Box, x: Xfrm, fill: Fill | null, line: Line | null, job: Job) {
  void job;
  const paths = geometryPaths(geom, b.w, b.h);
  const f = fill && fill.kind !== 'image' ? solidOf(fill) : null;
  if (f) {
    pg.save().alpha(f.alpha).fillColor(f.rgb);
    for (const p of paths) if (p.fill) emitPath(pg, p.cmds, b, x);
    pg.fill(paths.some((p) => p.evenOdd)).restore();
  }
  const s = line && line.fill.kind !== 'none' ? solidOf(line.fill) : null;
  if (s && line) {
    const width = Math.max(emuToPx(line.width) * K, 0.5);
    pg.save().alpha(1, s.alpha).strokeColor(s.rgb).lineWidth(width).lineJoin('round');
    const dash = dashPattern(line.dash);
    if (dash) pg.dash(dash.map((d) => d * width));
    if (line.cap === 'rnd') pg.lineCap('round');
    for (const p of paths) if (p.stroke) emitPath(pg, p.cmds, b, x);
    pg.stroke().restore();
    if (line.head || line.tail) {
      const ends = endPoints(paths);
      if (ends) {
        pg.save().alpha(s.alpha).fillColor(s.rgb);
        const head = arrowHead(line.head, ends.start, ends.second, width);
        const tail = arrowHead(line.tail, ends.end, ends.beforeEnd, width);
        if (head) emitPath(pg, head, b, x);
        if (tail) emitPath(pg, tail, b, x);
        pg.fill().restore();
      }
    }
  }
}

async function paintFill(job: Job, pg: PdfPage, fill: Fill | null, path: () => void, b: Box, fallback: { r: number; g: number; b: number }) {
  if (fill?.kind === 'image') {
    const rel = job.deck.pkg.rel(fill.owner, fill.rId);
    const bytes = rel && !rel.external ? job.media.bytes(rel.target) : null;
    const mime = rel ? job.media.mimeOf(rel.target) : null;
    const img = bytes && mime ? await job.builder.image(rel!.target, bytes, mime) : null;
    if (img) {
      pg.save();
      path();
      pg.clip();
      pg.imageInBox(img, b.x, b.y, b.w, b.h);
      pg.restore();
      return;
    }
  }
  const s = solidOf(fill) ?? { rgb: [fallback.r, fallback.g, fallback.b] as Rgb, alpha: 1 };
  pg.save().alpha(s.alpha).fillColor(s.rgb);
  path();
  pg.fill().restore();
}

async function paintImageFill(job: Job, pg: PdfPage, fill: Fill, geom: Parameters<typeof geometryPaths>[0], b: Box, x: Xfrm) {
  const paths = geometryPaths(geom, b.w, b.h);
  await paintFill(job, pg, fill, () => paths.filter((p) => p.fill).forEach((p) => emitPath(pg, p.cmds, b, x)), b, { r: 1, g: 1, b: 1 });
}

async function paintPicture(job: Job, pg: PdfPage, el: Extract<SlideElement, { type: 'picture' }>, b: Box, x: Xfrm) {
  const bytes = el.media ? job.media.bytes(el.media) : null;
  const mime = el.media ? job.media.mimeOf(el.media) : null;
  const img = bytes && mime ? await job.builder.image(el.media!, bytes, mime) : null;
  const paths = geometryPaths(el.geom, b.w, b.h);
  if (img) {
    const { l, t, r, b: bt } = el.crop;
    const fw = b.w / Math.max(0.01, 1 - l - r);
    const fh = b.h / Math.max(0.01, 1 - t - bt);
    const ix = b.x - l * fw;
    const iy = b.y - t * fh;
    pg.save();
    emitPath(pg, paths[0]?.cmds ?? [], b, x);
    pg.clip();
    const L = x.flipH ? ix + fw : ix;
    const R = x.flipH ? ix : ix + fw;
    const T = x.flipV ? iy + fh : iy;
    const B = x.flipV ? iy : iy + fh;
    pg.image(img, { x: L, y: T }, { x: R, y: T }, { x: L, y: B });
    pg.restore();
  }
  if (el.line) paintGeometry(pg, el.geom, b, x, null, el.line, job);
}

// ---------------------------------------------------------------- tables

function cssRgb(v: string): { rgb: Rgb; alpha: number } | null {
  const m = /rgba?\(([^)]+)\)/.exec(v);
  if (!m) return null;
  const parts = m[1].split(/[\s,/]+/).filter(Boolean).map(Number);
  const alpha = parts.length > 3 ? parts[3] : 1;
  if (alpha <= 0) return null;
  return { rgb: [parts[0] / 255, parts[1] / 255, parts[2] / 255], alpha };
}

async function paintTable(job: Job, pg: PdfPage, dom: HTMLElement, b: Box) {
  const base = dom.getBoundingClientRect();
  const cells = Array.from(dom.querySelectorAll('td')) as HTMLElement[];
  for (const td of cells) {
    const r = td.getBoundingClientRect();
    const x = b.x + (r.left - base.left) * K;
    const y = b.y + (r.top - base.top) * K;
    const w = r.width * K;
    const h = r.height * K;
    const cs = getComputedStyle(td);
    const bgc = cssRgb(cs.backgroundColor);
    if (bgc) pg.fillRect(x, y, w, h, bgc.rgb, bgc.alpha);
  }
  for (const td of cells) {
    const r = td.getBoundingClientRect();
    const x = b.x + (r.left - base.left) * K;
    const y = b.y + (r.top - base.top) * K;
    const w = r.width * K;
    const h = r.height * K;
    const cs = getComputedStyle(td);
    const side = (width: string, style: string, color: string) => {
      const px = parseFloat(width) || 0;
      const c = cssRgb(color);
      return px > 0 && style !== 'none' && c ? { t: px * K, c } : null;
    };
    const top = side(cs.borderTopWidth, cs.borderTopStyle, cs.borderTopColor);
    const bottom = side(cs.borderBottomWidth, cs.borderBottomStyle, cs.borderBottomColor);
    const left = side(cs.borderLeftWidth, cs.borderLeftStyle, cs.borderLeftColor);
    const right = side(cs.borderRightWidth, cs.borderRightStyle, cs.borderRightColor);
    if (top) pg.fillRect(x, y - top.t / 2, w, top.t, top.c.rgb, top.c.alpha);
    if (bottom) pg.fillRect(x, y + h - bottom.t / 2, w, bottom.t, bottom.c.rgb, bottom.c.alpha);
    if (left) pg.fillRect(x - left.t / 2, y, left.t, h, left.c.rgb, left.c.alpha);
    if (right) pg.fillRect(x + w - right.t / 2, y, right.t, h, right.c.rgb, right.c.alpha);
    await paintText(job, pg, td, dom, b.x, b.y);
  }
}

// ---------------------------------------------------------------- text

async function faceFor(job: Job, family: string, bold: boolean, italic: boolean): Promise<PdfFont | null> {
  const face = job.fonts.face(family, bold, italic);
  if (!face) return null;
  return job.builder.font(face.key, () => face.bytes);
}

/**
 * Draw the text under `container`. Positions are measured relative to `origin` (the element's
 * own box in the flat render) and offset by (ox, oy), the element's box in the current PDF space.
 */
async function paintText(job: Job, pg: PdfPage, container: Element | null, origin: HTMLElement, ox: number, oy: number) {
  if (!container) return;
  const base = origin.getBoundingClientRect();
  const walker = document.createTreeWalker(container, NodeFilter.SHOW_TEXT);
  for (let n = walker.nextNode() as Text | null; n; n = walker.nextNode() as Text | null) {
    const text = n.nodeValue ?? '';
    if (!/[^\s​]/.test(text)) continue;
    const el = n.parentElement;
    if (!el) continue;
    const cs = getComputedStyle(el);
    if (cs.visibility === 'hidden' || cs.display === 'none') continue;
    const size = parseFloat(cs.fontSize) * K;
    const bold = (parseInt(cs.fontWeight, 10) || 400) >= 600;
    const italic = cs.fontStyle === 'italic' || cs.fontStyle.startsWith('oblique');
    const font = await faceFor(job, cs.fontFamily, bold, italic);
    const color = cssRgb(cs.color) ?? { rgb: BLACK, alpha: 1 };
    const deco = cs.textDecorationLine || '';
    const decoColor = cssRgb(cs.textDecorationColor) ?? color;
    const opacity = parseFloat(cs.opacity) || 1;
    for (const m of text.matchAll(/[^\s​]+/g)) {
      const range = document.createRange();
      range.setStart(n, m.index!);
      range.setEnd(n, m.index! + m[0].length);
      const rects = Array.from(range.getClientRects()).filter((r) => r.width > 0);
      if (!rects.length) continue;
      if (rects.length > 1) {
        // A word broken across lines: place it character by character.
        for (let i = 0; i < m[0].length; i++) {
          const cr = document.createRange();
          cr.setStart(n, m.index! + i);
          cr.setEnd(n, m.index! + i + 1);
          const r = cr.getClientRects()[0];
          if (r) drawWord(job, pg, font, m[0][i], r, base, ox, oy, size, color, opacity);
        }
        continue;
      }
      const r = rects[0];
      drawWord(job, pg, font, m[0], r, base, ox, oy, size, color, opacity);
      if (/underline|line-through/.test(deco)) {
        const x = ox + (r.left - base.left) * K;
        const top = oy + (r.top - base.top) * K;
        const baseline = baselineOf(font, top, r.height * K, size);
        const t = Math.max(0.4, size * 0.05);
        if (deco.includes('underline')) pg.fillRect(x, baseline + size * 0.1, r.width * K, t, decoColor.rgb, decoColor.alpha * opacity);
        if (deco.includes('line-through')) pg.fillRect(x, baseline - size * 0.28, r.width * K, t, decoColor.rgb, decoColor.alpha * opacity);
      }
    }
  }
}

function baselineOf(font: PdfFont | null, top: number, h: number, size: number): number {
  if (!font) return top + h * 0.8;
  const upem = font.unitsPerEm;
  const asc = (font.font.ascent / upem) * size;
  const desc = (-font.font.descent / upem) * size;
  return top + (h - (asc + desc)) / 2 + asc;
}

function drawWord(
  job: Job,
  pg: PdfPage,
  font: PdfFont | null,
  word: string,
  r: DOMRect,
  base: DOMRect,
  ox: number,
  oy: number,
  size: number,
  color: { rgb: Rgb; alpha: number },
  opacity: number,
) {
  const x0 = ox + (r.left - base.left) * K;
  const top = oy + (r.top - base.top) * K;
  const h = r.height * K;
  // Split the word into runs: its own face where it has the glyph, a fallback face otherwise.
  const runs: { font: PdfFont; text: string }[] = [];
  for (const ch of word) {
    const cp = ch.codePointAt(0) ?? 0;
    const f = font && font.font.hasGlyphForCodePoint(cp) ? font : fallbackFor(job.fallbacks, ch) ?? font;
    if (!f) continue;
    const last = runs[runs.length - 1];
    if (last && last.font === f) last.text += ch;
    else runs.push({ font: f, text: ch });
  }
  if (!runs.length) return;
  const laid = runs.map((run) => ({ font: run.font, layout: run.font.font.layout(run.text) }));
  const natural = laid.reduce((a, l) => a + l.layout.positions.reduce((s, p) => s + p.xAdvance, 0) * (size / l.font.unitsPerEm), 0);
  const target = r.width * K;
  const fit = natural > 0 && target > 0 && Math.abs(target / natural - 1) < 0.2 ? target / natural : 1;
  const baseline = baselineOf(font ?? runs[0].font, top, h, size);
  let pen = 0;
  for (const l of laid) {
    const k = size / l.font.unitsPerEm;
    const glyphs: PlacedGlyph[] = l.layout.glyphs.map((g, i) => {
      const p = l.layout.positions[i];
      const placed = { cid: l.font.cid(g.id, g.codePoints.length ? String.fromCodePoint(...g.codePoints) : undefined), x: x0 + (pen + p.xOffset * k) * fit, y: baseline - p.yOffset * k };
      pen += p.xAdvance * k;
      return placed;
    });
    pg.glyphs(l.font, { size, hScale: fit, fill: color.rgb, fillAlpha: color.alpha * opacity }, glyphs);
  }
}
