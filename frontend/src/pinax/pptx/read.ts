/**
 * .pptx -> Pinax model.
 *
 * PowerPoint content is mostly inherited: a slide's placeholder takes its position, text styles
 * and body settings from the matching placeholder on its layout, which takes them from the
 * master, whose theme supplies colours and fonts. This module resolves that cascade once, when
 * the file is opened, so rendering is a plain walk over resolved values.
 */
import { strFromU8 } from 'fflate';
import {
  type Align,
  type BodyProps,
  type Context,
  type CustomPath,
  type Deck,
  type DeckState,
  type Geometry,
  type InlineItem,
  type Layout,
  type LevelProps,
  type LevelStyle,
  type Master,
  type Paragraph,
  type Placeholder,
  type PlaceholderDef,
  type RunProps,
  type Slide,
  type SlideElement,
  type Spacing,
  type TableCell,
  type TextBody,
  type Xfrm,
  newId,
} from '../model';
import { Pkg, REL } from './package';
import {
  type Color,
  type ColorContext,
  type Fill,
  type Line,
  type Theme,
  DEFAULT_CLR_MAP,
  colorChild,
  fillChild,
  readClrMap,
  readFill,
  readLine,
  readTheme,
  resolveColor,
} from './theme';
import { attr, bool, find, findAll, is, kid, kids, num, parseXml, path, serialize } from './xml';

export class DeckError extends Error {}

// ---------------------------------------------------------------- entry

export function openDeck(bytes: Uint8Array): { deck: Deck; state: DeckState } {
  const pkg = Pkg.open(bytes);
  const root = readRootRels(pkg);
  const presPart = root.find((r) => r.type === REL.officeDocument)?.target ?? 'ppt/presentation.xml';
  const pres = pkg.xml(presPart);
  if (!pres) throw new DeckError('This presentation is damaged and cannot be opened.');
  const pr = pres.documentElement;

  const sz = kid(pr, 'p:sldSz');
  const size = { cx: num(sz, 'cx', 12192000), cy: num(sz, 'cy', 6858000) };

  const deck: Deck = {
    pkg,
    presPart,
    size,
    layouts: new Map(),
    masters: new Map(),
    layoutOrder: [],
    defaultTextStyle: kid(pr, 'p:defaultTextStyle'),
    notesMaster: pkg.relOfType(presPart, REL.notesMaster)?.target ?? null,
    tableStyles: null,
    fonts: new Set(),
  };
  const ts = pkg.relOfType(presPart, REL.tableStyles);
  if (ts) deck.tableStyles = pkg.xml(ts.target);

  for (const id of kids(kid(pr, 'p:sldMasterIdLst'), 'p:sldMasterId')) {
    const rel = pkg.rel(presPart, attr(id, 'r:id'));
    if (!rel || deck.masters.has(rel.target)) continue;
    const master = readMaster(deck, rel.target);
    if (!master) continue;
    deck.masters.set(master.part, master);
    const mdoc = pkg.xml(master.part);
    for (const lid of kids(kid(mdoc?.documentElement, 'p:sldLayoutIdLst'), 'p:sldLayoutId')) {
      const lrel = pkg.rel(master.part, attr(lid, 'r:id'));
      if (!lrel || deck.layouts.has(lrel.target)) continue;
      const layout = readLayout(deck, lrel.target, master);
      if (layout) {
        deck.layouts.set(layout.part, layout);
        deck.layoutOrder.push(layout.part);
      }
    }
  }
  if (!deck.masters.size) throw new DeckError('This presentation has no slide master and cannot be opened.');

  const slides: Slide[] = [];
  for (const id of kids(kid(pr, 'p:sldIdLst'), 'p:sldId')) {
    const rel = pkg.rel(presPart, attr(id, 'r:id'));
    if (!rel) continue;
    const slide = readSlide(deck, rel.target);
    if (slide) slides.push(slide);
  }
  return { deck, state: { slides } };
}

function readRootRels(pkg: Pkg) {
  const doc = pkg.xml('_rels/.rels');
  return kids(doc?.documentElement).map((el) => ({
    type: attr(el, 'Type') ?? '',
    target: (attr(el, 'Target') ?? '').replace(/^\//, ''),
  }));
}

// ---------------------------------------------------------------- masters, layouts, slides

function readMaster(deck: Deck, part: string): Master | null {
  const doc = deck.pkg.xml(part);
  if (!doc) return null;
  const root = doc.documentElement;
  const themeRel = deck.pkg.relOfType(part, REL.theme);
  const theme = readTheme(themeRel ? deck.pkg.xml(themeRel.target) : null);
  deck.fonts.add(theme.majorFont);
  deck.fonts.add(theme.minorFont);
  const ctx: Context = { theme, color: { theme, clrMap: readClrMap(kid(root, 'p:clrMap'), DEFAULT_CLR_MAP) } };
  const tx = kid(root, 'p:txStyles');
  const master: Master = {
    part,
    decorations: [],
    placeholders: [],
    background: readBackground(kid(kid(root, 'p:cSld'), 'p:bg'), ctx, part),
    ctx,
    titleStyle: kid(tx, 'p:titleStyle'),
    bodyStyle: kid(tx, 'p:bodyStyle'),
    otherStyle: kid(tx, 'p:otherStyle'),
  };
  const tree = path(root, 'p:cSld', 'p:spTree');
  master.placeholders = placeholderDefs(tree, part);
  master.decorations = readElements(tree, { deck, owner: part, ctx, master, layout: null, decorations: true });
  return master;
}

function readLayout(deck: Deck, part: string, master: Master): Layout | null {
  const doc = deck.pkg.xml(part);
  if (!doc) return null;
  const root = doc.documentElement;
  const ctx: Context = {
    theme: master.ctx.theme,
    color: { theme: master.ctx.theme, clrMap: overrideMap(root, master.ctx.color.clrMap) },
  };
  const layout: Layout = {
    part,
    name: attr(kid(root, 'p:cSld'), 'name') || 'Layout',
    master: master.part,
    decorations: [],
    placeholders: [],
    background: readBackground(kid(kid(root, 'p:cSld'), 'p:bg'), ctx, part),
    ctx,
    showMasterShapes: bool(root, 'showMasterSp', true),
  };
  const tree = path(root, 'p:cSld', 'p:spTree');
  layout.placeholders = placeholderDefs(tree, part);
  layout.decorations = readElements(tree, { deck, owner: part, ctx, master, layout: null, decorations: true });
  return layout;
}

function readSlide(deck: Deck, part: string): Slide | null {
  const doc = deck.pkg.xml(part);
  if (!doc) return null;
  const layoutPart = deck.pkg.relOfType(part, REL.slideLayout)?.target ?? '';
  return buildSlide(deck, doc, part, layoutPart, part);
}

/**
 * A slide model from slide XML. `part` is the source part (null for a slide made in Pinax);
 * `owner` is the part whose relationships the XML's r:ids refer to.
 */
export function buildSlide(deck: Deck, doc: Document, part: string | null, layoutPart: string, owner: string): Slide | null {
  const root = doc.documentElement;
  const layout = deck.layouts.get(layoutPart) ?? [...deck.layouts.values()][0];
  if (!layout) return null;
  const master = deck.masters.get(layout.master)!;
  const ctx: Context = {
    theme: master.ctx.theme,
    color: { theme: master.ctx.theme, clrMap: overrideMap(root, layout.ctx.color.clrMap) },
  };
  const tree = path(root, 'p:cSld', 'p:spTree');
  return {
    id: newId('s'),
    part,
    doc,
    layout: layout.part,
    elements: readElements(tree, { deck, owner, ctx, master, layout, decorations: false }),
    background: readBackground(kid(kid(root, 'p:cSld'), 'p:bg'), ctx, owner),
    hidden: attr(root, 'show') === '0' || attr(root, 'show') === 'false',
    notes: part ? readNotes(deck, part) : deck.notesMaster ? [] : null,
    ctx,
    showMasterShapes: bool(root, 'showMasterSp', true),
  };
}

/** Elements from generated XML, read against a slide's context (for elements made in Pinax). */
export function readFragment(deck: Deck, slide: Slide, tree: Element, owner: string): SlideElement[] {
  const layout = deck.layouts.get(slide.layout) ?? null;
  const master = deck.masters.get(layout?.master ?? '') ?? [...deck.masters.values()][0];
  return readElements(tree, { deck, owner, ctx: slide.ctx, master, layout, decorations: false });
}

function overrideMap(root: Element, base: ColorContext['clrMap']) {
  const ovr = path(root, 'p:clrMapOvr', 'a:overrideClrMapping');
  return ovr ? readClrMap(ovr, base) : base;
}

function readBackground(bg: Element | null, ctx: Context, owner: string): Fill | null {
  if (!bg) return null;
  const pr = kid(bg, 'p:bgPr');
  if (pr) return readFill(fillChild(pr), ctx.color, owner);
  const ref = kid(bg, 'p:bgRef');
  if (ref) {
    const idx = num(ref, 'idx', 0);
    const ph = resolveColor(colorChild(ref), ctx.color);
    const list = idx >= 1001 ? ctx.theme.bgFillStyles : ctx.theme.fillStyles;
    const el = list[(idx >= 1001 ? idx - 1001 : idx - 1)] ?? null;
    return readFill(el, ctx.color, owner, ph);
  }
  return null;
}

function readNotes(deck: Deck, part: string): string[] | null {
  const rel = deck.pkg.relOfType(part, REL.notesSlide);
  if (!rel) return null;
  const doc = deck.pkg.xml(rel.target);
  if (!doc) return null;
  for (const sp of findAll(doc, 'p:sp')) {
    const ph = path(sp, 'p:nvSpPr', 'p:nvPr', 'p:ph');
    if (attr(ph, 'type') !== 'body') continue;
    return kids(kid(sp, 'p:txBody'), 'a:p').map((p) =>
      kids(p)
        .map((c) => (is(c, 'a:r') || is(c, 'a:fld') ? kid(c, 'a:t')?.textContent ?? '' : is(c, 'a:br') ? '\n' : ''))
        .join(''),
    );
  }
  return [];
}

// ---------------------------------------------------------------- placeholders

function placeholderDefs(tree: Element | null, owner: string): PlaceholderDef[] {
  const out: PlaceholderDef[] = [];
  for (const sp of kids(tree)) {
    if (!is(sp, 'p:sp') && !is(sp, 'p:pic')) continue;
    const nv = kid(sp, is(sp, 'p:sp') ? 'p:nvSpPr' : 'p:nvPicPr');
    const ph = path(nv, 'p:nvPr', 'p:ph');
    if (!ph) continue;
    out.push({
      type: attr(ph, 'type') ?? 'obj',
      idx: num(ph, 'idx', 0),
      xfrm: readXfrm(path(sp, 'p:spPr', 'a:xfrm')),
      sp,
      owner,
    });
  }
  return out;
}

/** Master placeholders are matched by kind: titles to title, text to body. */
function masterType(type: string): string {
  if (type === 'ctrTitle' || type === 'title') return 'title';
  if (type === 'subTitle' || type === 'obj' || type === 'body' || type === 'tbl' || type === 'chart' || type === 'pic' || type === 'clipArt' || type === 'dgm' || type === 'media')
    return 'body';
  return type;
}

function findLayoutPh(layout: Layout | null, type: string, idx: number | null): PlaceholderDef | null {
  if (!layout) return null;
  if (idx !== null && idx !== 0) {
    const byIdx = layout.placeholders.find((d) => d.idx === idx);
    if (byIdx) return byIdx;
  }
  return (
    layout.placeholders.find((d) => d.type === type) ??
    layout.placeholders.find((d) => masterType(d.type) === masterType(type) && (type === 'title' || type === 'ctrTitle')) ??
    null
  );
}

function findMasterPh(master: Master, type: string): PlaceholderDef | null {
  const t = masterType(type);
  return master.placeholders.find((d) => masterType(d.type) === t) ?? null;
}

const PROMPTS: Record<string, string> = {
  title: 'Click to add title',
  ctrTitle: 'Click to add title',
  subTitle: 'Click to add subtitle',
  body: 'Click to add text',
  obj: 'Click to add text',
  pic: 'Click to add picture',
  tbl: 'Table placeholder',
  chart: 'Chart placeholder',
  dt: '',
  ftr: '',
  sldNum: '',
};

// ---------------------------------------------------------------- elements

interface Env {
  deck: Deck;
  owner: string;
  ctx: Context;
  master: Master;
  /** The slide's layout (null when reading a layout or master itself). */
  layout: Layout | null;
  /** Reading a layout/master: skip placeholders, they are prompts, not content. */
  decorations: boolean;
  groupFill?: Fill | null;
}

export function readElements(tree: Element | null, env: Env): SlideElement[] {
  const out: SlideElement[] = [];
  for (const node of kids(tree)) {
    const el = readElement(node, env);
    if (el) out.push(el);
  }
  return out;
}

function readElement(node: Element, env: Env): SlideElement | null {
  if (is(node, 'p:nvGrpSpPr') || is(node, 'p:grpSpPr') || is(node, 'p:extLst')) return null;
  if (is(node, 'mc:AlternateContent')) {
    const branch = kid(node, 'mc:Fallback') ?? kid(node, 'mc:Choice');
    const inner = kids(branch)[0];
    if (!inner) return null;
    const el = readElement(inner, env);
    if (el) el.xml = node; // patch the whole AlternateContent on save
    return el;
  }
  if (is(node, 'p:sp')) return readShape(node, env, false);
  if (is(node, 'p:cxnSp')) return readShape(node, env, true);
  if (is(node, 'p:pic')) return readPicture(node, env);
  if (is(node, 'p:grpSp')) return readGroup(node, env);
  if (is(node, 'p:graphicFrame')) return readGraphicFrame(node, env);
  if (env.decorations) return null;
  // Anything else (ink, content parts...) is kept as an opaque object.
  const x = readXfrm(find(node, 'a:xfrm') ?? find(node, 'p:xfrm'));
  return {
    type: 'frame',
    id: newId(),
    name: 'Object',
    xfrm: x ?? { x: 0, y: 0, w: 0, h: 0, rot: 0, flipH: false, flipV: false },
    xml: node,
    owner: env.owner,
    changed: {},
    kind: 'other',
    label: 'Object',
    preview: null,
  };
}

export function readXfrm(x: Element | null): Xfrm | null {
  if (!x) return null;
  const off = kid(x, 'a:off');
  const ext = kid(x, 'a:ext');
  if (!off || !ext) return null;
  return {
    x: num(off, 'x', 0),
    y: num(off, 'y', 0),
    w: num(ext, 'cx', 0),
    h: num(ext, 'cy', 0),
    rot: num(x, 'rot', 0) / 60000,
    flipH: bool(x, 'flipH', false),
    flipV: bool(x, 'flipV', false),
  };
}

function readPlaceholder(nvPr: Element | null): { type: string; idx: number | null } | null {
  const ph = kid(nvPr, 'p:ph');
  if (!ph) return null;
  const idx = num(ph, 'idx');
  return { type: attr(ph, 'type') ?? 'obj', idx };
}

function readGeometry(spPr: Element | null): Geometry | null {
  const prst = kid(spPr, 'a:prstGeom');
  if (prst) {
    const adj: Record<string, number> = {};
    for (const gd of kids(kid(prst, 'a:avLst'), 'a:gd')) {
      const m = /^val\s+(-?\d+)/.exec(attr(gd, 'fmla') ?? '');
      if (m) adj[attr(gd, 'name') ?? ''] = Number(m[1]);
    }
    return { kind: 'preset', name: attr(prst, 'prst') ?? 'rect', adj };
  }
  const cust = kid(spPr, 'a:custGeom');
  if (cust) {
    const paths: CustomPath[] = [];
    for (const p of kids(kid(cust, 'a:pathLst'), 'a:path')) {
      const cmds: (string | number)[][] = [];
      // Coordinates may be guide names (formulas); those stay NaN and the shape draws as its box.
      const n = (el: Element | null, k: string) => Number(attr(el, k) ?? NaN);
      const pt = (el: Element | null) => [n(el, 'x'), n(el, 'y')];
      for (const c of kids(p)) {
        if (is(c, 'a:moveTo')) cmds.push(['M', ...pt(kid(c, 'a:pt'))]);
        else if (is(c, 'a:lnTo')) cmds.push(['L', ...pt(kid(c, 'a:pt'))]);
        else if (is(c, 'a:cubicBezTo')) cmds.push(['C', ...kids(c, 'a:pt').flatMap(pt)]);
        else if (is(c, 'a:quadBezTo')) cmds.push(['Q', ...kids(c, 'a:pt').flatMap(pt)]);
        else if (is(c, 'a:arcTo'))
          cmds.push(['A', n(c, 'wR'), n(c, 'hR'), n(c, 'stAng'), n(c, 'swAng')]);
        else if (is(c, 'a:close')) cmds.push(['Z']);
      }
      paths.push({
        w: num(p, 'w', 0),
        h: num(p, 'h', 0),
        fill: attr(p, 'fill') !== 'none',
        stroke: bool(p, 'stroke', true),
        cmds,
      });
    }
    return { kind: 'custom', paths };
  }
  return null;
}

interface StyleRefs {
  fill: Fill | null;
  line: Line | null;
  fontColor: Color | null;
  fontFamily: string | null;
}

function readStyleRefs(style: Element | null, env: Env): StyleRefs {
  const out: StyleRefs = { fill: null, line: null, fontColor: null, fontFamily: null };
  if (!style) return out;
  const theme = env.ctx.theme;
  const fr = kid(style, 'a:fillRef');
  if (fr) {
    const idx = num(fr, 'idx', 0);
    const ph = resolveColor(colorChild(fr), env.ctx.color);
    if (idx === 0) out.fill = { kind: 'none' };
    else {
      const list = idx >= 1001 ? theme.bgFillStyles : theme.fillStyles;
      const el = list[idx >= 1001 ? idx - 1001 : idx - 1];
      out.fill = el ? readFill(el, env.ctx.color, env.owner, ph) : ph ? { kind: 'solid', color: ph } : null;
    }
  }
  const lr = kid(style, 'a:lnRef');
  if (lr) {
    const idx = num(lr, 'idx', 0);
    const ph = resolveColor(colorChild(lr), env.ctx.color);
    const el = idx > 0 ? theme.lineStyles[idx - 1] : null;
    out.line = el ? readLine(el, env.ctx.color, env.owner, ph) : null;
  }
  const fo = kid(style, 'a:fontRef');
  if (fo) {
    out.fontColor = resolveColor(colorChild(fo), env.ctx.color);
    const idx = attr(fo, 'idx');
    out.fontFamily = idx === 'major' ? theme.majorFont : idx === 'minor' ? theme.minorFont : null;
  }
  return out;
}

function readShape(sp: Element, env: Env, connector: boolean): SlideElement | null {
  const nv = kid(sp, connector ? 'p:nvCxnSpPr' : 'p:nvSpPr');
  const cNvPr = kid(nv, 'p:cNvPr');
  const nvPr = kid(nv, 'p:nvPr');
  const phRef = connector ? null : readPlaceholder(nvPr);
  if (phRef && env.decorations) return null;
  const spPr = kid(sp, 'p:spPr');

  const layoutDef = phRef ? findLayoutPh(env.layout, phRef.type, phRef.idx) : null;
  const masterDef = phRef ? findMasterPh(env.master, layoutDef?.type ?? phRef.type) : null;
  const chain = [masterDef, layoutDef].filter((d): d is PlaceholderDef => !!d);

  const xfrm =
    readXfrm(kid(spPr, 'a:xfrm')) ?? layoutDef?.xfrm ?? masterDef?.xfrm ?? { x: 0, y: 0, w: 0, h: 0, rot: 0, flipH: false, flipV: false };

  let geom = readGeometry(spPr);
  for (const d of [...chain].reverse()) geom ??= readGeometry(kid(d.sp, 'p:spPr'));

  const refs = readStyleRefs(kid(sp, 'p:style'), env);

  let fill: Fill | null = null;
  const ownFill = fillChild(spPr);
  if (ownFill && is(ownFill, 'a:grpFill')) fill = env.groupFill ?? null;
  else fill = readFill(ownFill, env.ctx.color, env.owner);
  for (const d of [...chain].reverse()) fill ??= readFill(fillChild(kid(d.sp, 'p:spPr')), env.ctx.color, d.owner);
  fill ??= refs.fill;

  let line: Line | null = readLine(kid(spPr, 'a:ln'), env.ctx.color, env.owner, null, refs.line);
  if (!kid(spPr, 'a:ln')) {
    for (const d of [...chain].reverse()) {
      const ln = path(d.sp, 'p:spPr', 'a:ln');
      if (ln) {
        line = readLine(ln, env.ctx.color, d.owner, null, refs.line);
        break;
      }
    }
  }

  const ph: Placeholder | null = phRef
    ? {
        type: phRef.type,
        idx: phRef.idx ?? 0,
        prompt: promptFor(layoutDef, phRef.type),
      }
    : null;

  const textBox = bool(kid(nv, 'p:cNvSpPr'), 'txBox', false);
  const text = connector ? null : readTextBody(kid(sp, 'p:txBody'), env, phRef, chain, refs);

  return {
    type: 'shape',
    id: newId(),
    name: attr(cNvPr, 'name') ?? 'Shape',
    xfrm,
    xml: sp,
    owner: env.owner,
    changed: {},
    geom: geom ?? { kind: 'preset', name: 'rect', adj: {} },
    fill,
    line,
    text,
    ph,
    textBox,
    connector,
  };
}

function promptFor(def: PlaceholderDef | null, type: string): string {
  if (def) {
    const ph = path(def.sp, 'p:nvSpPr', 'p:nvPr', 'p:ph');
    if (bool(ph, 'hasCustomPrompt', false)) {
      const t = findAll(def.sp, 'a:t')
        .map((x) => x.textContent ?? '')
        .join('')
        .trim();
      if (t) return t;
    }
  }
  return PROMPTS[type] ?? 'Click to add text';
}

function readPicture(pic: Element, env: Env): SlideElement | null {
  const nv = kid(pic, 'p:nvPicPr');
  const nvPr = kid(nv, 'p:nvPr');
  const phRef = readPlaceholder(nvPr);
  if (phRef && env.decorations) return null;
  const spPr = kid(pic, 'p:spPr');
  const layoutDef = phRef ? findLayoutPh(env.layout, phRef.type, phRef.idx) : null;
  const xfrm = readXfrm(kid(spPr, 'a:xfrm')) ?? layoutDef?.xfrm ?? { x: 0, y: 0, w: 0, h: 0, rot: 0, flipH: false, flipV: false };
  const blipFill = kid(pic, 'p:blipFill');
  const blip = kid(blipFill, 'a:blip');
  const rId = attr(blip, 'r:embed') ?? '';
  const rel = env.deck.pkg.rel(env.owner, rId);
  // Pictures inserted in Pinax reference their media part directly until saved.
  const media = rId.startsWith('__media:') ? rId.slice(8) : rel && !rel.external && env.deck.pkg.has(rel.target) ? rel.target : null;
  const src = kid(blipFill, 'a:srcRect');
  const mediaKind = kid(nvPr, 'a:videoFile') ? 'video' : kid(nvPr, 'a:audioFile') ? 'audio' : find(nvPr, 'p14:media') ? 'video' : null;
  return {
    type: 'picture',
    id: newId(),
    name: attr(kid(nv, 'p:cNvPr'), 'name') ?? 'Picture',
    xfrm,
    xml: pic,
    owner: env.owner,
    changed: {},
    media,
    crop: {
      l: num(src, 'l', 0) / 100000,
      t: num(src, 't', 0) / 100000,
      r: num(src, 'r', 0) / 100000,
      b: num(src, 'b', 0) / 100000,
    },
    geom: readGeometry(spPr) ?? { kind: 'preset', name: 'rect', adj: {} },
    line: readLine(kid(spPr, 'a:ln'), env.ctx.color, env.owner, null, readStyleRefs(kid(pic, 'p:style'), env).line),
    ph: phRef ? { type: phRef.type, idx: phRef.idx ?? 0, prompt: '' } : null,
    mediaKind,
  };
}

function readGroup(grp: Element, env: Env): SlideElement | null {
  const pr = kid(grp, 'p:grpSpPr');
  const x = kid(pr, 'a:xfrm');
  const xfrm = readXfrm(x) ?? { x: 0, y: 0, w: 0, h: 0, rot: 0, flipH: false, flipV: false };
  const chOff = kid(x, 'a:chOff');
  const chExt = kid(x, 'a:chExt');
  const groupFill = readFill(fillChild(pr), env.ctx.color, env.owner) ?? env.groupFill ?? null;
  const children = readElements(grp, { ...env, decorations: false, groupFill });
  return {
    type: 'group',
    id: newId(),
    name: attr(path(grp, 'p:nvGrpSpPr', 'p:cNvPr'), 'name') ?? 'Group',
    xfrm,
    xml: grp,
    owner: env.owner,
    changed: {},
    children,
    child: {
      x: num(chOff, 'x', xfrm.x),
      y: num(chOff, 'y', xfrm.y),
      w: num(chExt, 'cx', xfrm.w) || xfrm.w || 1,
      h: num(chExt, 'cy', xfrm.h) || xfrm.h || 1,
    },
  };
}

const URI = {
  table: 'http://schemas.openxmlformats.org/drawingml/2006/table',
  chart: 'http://schemas.openxmlformats.org/drawingml/2006/chart',
  chartex: 'http://schemas.microsoft.com/office/drawing/2014/chartex',
  diagram: 'http://schemas.openxmlformats.org/drawingml/2006/diagram',
  ole: 'http://schemas.openxmlformats.org/presentationml/2006/ole',
};

function readGraphicFrame(gf: Element, env: Env): SlideElement | null {
  const nv = kid(gf, 'p:nvGraphicFramePr');
  const phRef = readPlaceholder(kid(nv, 'p:nvPr'));
  if (phRef && env.decorations) return null;
  const layoutDef = phRef ? findLayoutPh(env.layout, phRef.type, phRef.idx) : null;
  const xfrm = readXfrm(kid(gf, 'p:xfrm')) ?? layoutDef?.xfrm ?? { x: 0, y: 0, w: 0, h: 0, rot: 0, flipH: false, flipV: false };
  const data = path(gf, 'a:graphic', 'a:graphicData');
  const uri = attr(data, 'uri') ?? '';
  const base = {
    id: newId(),
    name: attr(kid(nv, 'p:cNvPr'), 'name') ?? 'Object',
    xfrm,
    xml: gf,
    owner: env.owner,
    changed: {},
  };
  if (uri === URI.table) {
    const tbl = kid(data, 'a:tbl');
    if (tbl) return readTable(tbl, env, base);
  }
  if (uri === URI.chart || uri === URI.chartex) {
    return { ...base, type: 'frame', kind: 'chart', label: chartTitle(env, data) ?? 'Chart', preview: null };
  }
  if (uri === URI.diagram) return { ...base, type: 'frame', kind: 'diagram', label: 'SmartArt graphic', preview: null };
  if (uri === URI.ole) {
    const pic = find(data, 'p:pic');
    const blip = find(pic, 'a:blip');
    const rel = env.deck.pkg.rel(env.owner, attr(blip, 'r:embed'));
    const prog = attr(find(data, 'p:oleObj'), 'progId') ?? '';
    return {
      ...base,
      type: 'frame',
      kind: 'ole',
      label: prog ? `Embedded object (${prog.split('.')[0]})` : 'Embedded object',
      preview: rel && !rel.external && env.deck.pkg.has(rel.target) ? rel.target : null,
    };
  }
  return { ...base, type: 'frame', kind: 'other', label: 'Object', preview: null };
}

function chartTitle(env: Env, data: Element | null): string | null {
  const chart = kid(data, 'c:chart');
  const rel = env.deck.pkg.rel(env.owner, attr(chart, 'r:id'));
  if (!rel) return null;
  const doc = env.deck.pkg.xml(rel.target);
  const title = find(doc?.documentElement, 'c:title');
  const t = findAll(title, 'a:t')
    .map((x) => x.textContent ?? '')
    .join('')
    .trim();
  return t ? `Chart: ${t}` : null;
}

// ---------------------------------------------------------------- tables

const MEDIUM_STYLE_2_ACCENT_1 = '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';

/** PowerPoint's default table style, used when the file references it without defining it. */
const BUILTIN_TABLE_STYLE = `<a:tblStyle xmlns:a="${A}" styleId="${MEDIUM_STYLE_2_ACCENT_1}" styleName="Medium Style 2 - Accent 1">
<a:wholeTbl><a:tcTxStyle><a:fontRef idx="minor"/><a:schemeClr val="dk1"/></a:tcTxStyle><a:tcStyle><a:tcBdr>
<a:left><a:ln w="12700"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln></a:left>
<a:right><a:ln w="12700"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln></a:right>
<a:top><a:ln w="12700"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln></a:top>
<a:bottom><a:ln w="12700"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln></a:bottom>
<a:insideH><a:ln w="12700"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln></a:insideH>
<a:insideV><a:ln w="12700"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln></a:insideV>
</a:tcBdr><a:fill><a:solidFill><a:schemeClr val="accent1"><a:tint val="20000"/></a:schemeClr></a:solidFill></a:fill></a:tcStyle></a:wholeTbl>
<a:band1H><a:tcStyle><a:fill><a:solidFill><a:schemeClr val="accent1"><a:tint val="40000"/></a:schemeClr></a:solidFill></a:fill></a:tcStyle></a:band1H>
<a:band1V><a:tcStyle><a:fill><a:solidFill><a:schemeClr val="accent1"><a:tint val="40000"/></a:schemeClr></a:solidFill></a:fill></a:tcStyle></a:band1V>
<a:lastCol><a:tcTxStyle b="on"><a:schemeClr val="lt1"/></a:tcTxStyle><a:tcStyle><a:fill><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:fill></a:tcStyle></a:lastCol>
<a:firstCol><a:tcTxStyle b="on"><a:schemeClr val="lt1"/></a:tcTxStyle><a:tcStyle><a:fill><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:fill></a:tcStyle></a:firstCol>
<a:lastRow><a:tcTxStyle b="on"><a:schemeClr val="lt1"/></a:tcTxStyle><a:tcStyle><a:fill><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:fill></a:tcStyle></a:lastRow>
<a:firstRow><a:tcTxStyle b="on"><a:schemeClr val="lt1"/></a:tcTxStyle><a:tcStyle><a:fill><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:fill></a:tcStyle></a:firstRow>
</a:tblStyle>`;

let builtinStyle: Element | null = null;

function tableStyle(deck: Deck, id: string | null): Element | null {
  const wanted = id ?? attr(deck.tableStyles?.documentElement, 'def');
  if (!wanted) return null;
  for (const s of findAll(deck.tableStyles, 'a:tblStyle')) if (attr(s, 'styleId') === wanted) return s;
  if (wanted === MEDIUM_STYLE_2_ACCENT_1) {
    builtinStyle ??= parseXml(BUILTIN_TABLE_STYLE).documentElement;
    return builtinStyle;
  }
  return null;
}

interface CellStyle {
  fill: Fill | null;
  bold: boolean | null;
  italic: boolean | null;
  color: Color | null;
  borders: Partial<Record<'left' | 'right' | 'top' | 'bottom' | 'insideH' | 'insideV', Line | null>>;
}

function readTable(tbl: Element, env: Env, base: { id: string; name: string; xfrm: Xfrm; xml: Element; owner: string; changed: object }): SlideElement {
  const tblPr = kid(tbl, 'a:tblPr');
  const flags = {
    firstRow: bool(tblPr, 'firstRow', false),
    lastRow: bool(tblPr, 'lastRow', false),
    firstCol: bool(tblPr, 'firstCol', false),
    lastCol: bool(tblPr, 'lastCol', false),
    bandRow: bool(tblPr, 'bandRow', false),
    bandCol: bool(tblPr, 'bandCol', false),
  };
  const style = tableStyle(env.deck, kid(tblPr, 'a:tableStyleId')?.textContent?.trim() ?? null);
  const part = (name: string): CellStyle | null => {
    const el = kid(style, `a:${name}`);
    if (!el) return null;
    const tc = kid(el, 'a:tcStyle');
    const tx = kid(el, 'a:tcTxStyle');
    const fillEl = kid(tc, 'a:fill');
    let fill: Fill | null = fillEl ? readFill(fillChild(fillEl), env.ctx.color, env.owner) : null;
    const fillRef = kid(tc, 'a:fillRef');
    if (!fill && fillRef) {
      const ph = resolveColor(colorChild(fillRef), env.ctx.color);
      fill = ph ? { kind: 'solid', color: ph } : null;
    }
    const borders: CellStyle['borders'] = {};
    for (const side of ['left', 'right', 'top', 'bottom', 'insideH', 'insideV'] as const) {
      const s = kid(kid(tc, 'a:tcBdr'), `a:${side}`);
      if (!s) continue;
      const ln = kid(s, 'a:ln');
      if (ln) borders[side] = readLine(ln, env.ctx.color, env.owner);
      const lnRef = kid(s, 'a:lnRef');
      if (lnRef && !ln) {
        const c = resolveColor(colorChild(lnRef), env.ctx.color);
        borders[side] = c ? { fill: { kind: 'solid', color: c }, width: 12700, dash: null, head: null, tail: null, cap: 'flat' } : null;
      }
    }
    const fontRefColor = resolveColor(colorChild(kid(tx, 'a:fontRef')), env.ctx.color);
    return {
      fill,
      bold: tx ? (attr(tx, 'b') === 'on' ? true : attr(tx, 'b') === 'off' ? false : null) : null,
      italic: tx ? (attr(tx, 'i') === 'on' ? true : attr(tx, 'i') === 'off' ? false : null) : null,
      color: resolveColor(colorChild(tx), env.ctx.color) ?? fontRefColor,
      borders,
    };
  };
  const parts = {
    wholeTbl: part('wholeTbl'),
    band1H: part('band1H'),
    band2H: part('band2H'),
    band1V: part('band1V'),
    band2V: part('band2V'),
    firstRow: part('firstRow'),
    lastRow: part('lastRow'),
    firstCol: part('firstCol'),
    lastCol: part('lastCol'),
  };

  const cols = kids(kid(tbl, 'a:tblGrid'), 'a:gridCol').map((g) => num(g, 'w', 0));
  const trs = kids(tbl, 'a:tr');
  const rows = trs.map((tr, r) => {
    const tcs = kids(tr, 'a:tc');
    const cells: TableCell[] = tcs.map((tc, c) => {
      const tcPr = kid(tc, 'a:tcPr');
      // Which style parts apply to this cell, weakest first.
      const applied: (CellStyle | null)[] = [parts.wholeTbl];
      const dataRow = flags.firstRow ? r - 1 : r;
      if (flags.bandRow && !(flags.firstRow && r === 0) && !(flags.lastRow && r === trs.length - 1)) {
        applied.push(dataRow % 2 === 0 ? parts.band1H : parts.band2H);
      }
      if (flags.bandCol) applied.push((flags.firstCol ? c - 1 : c) % 2 === 0 ? parts.band1V : parts.band2V);
      if (flags.lastCol && c === tcs.length - 1) applied.push(parts.lastCol);
      if (flags.firstCol && c === 0) applied.push(parts.firstCol);
      if (flags.lastRow && r === trs.length - 1) applied.push(parts.lastRow);
      if (flags.firstRow && r === 0) applied.push(parts.firstRow);
      const eff: CellStyle = { fill: null, bold: null, italic: null, color: null, borders: {} };
      for (const s of applied) {
        if (!s) continue;
        if (s.fill) eff.fill = s.fill;
        if (s.bold !== null) eff.bold = s.bold;
        if (s.italic !== null) eff.italic = s.italic;
        if (s.color) eff.color = s.color;
        Object.assign(eff.borders, s.borders);
      }
      const ownFill = readFill(fillChild(tcPr), env.ctx.color, env.owner);
      const side = (own: string, outer: 'left' | 'right' | 'top' | 'bottom', inner: 'insideH' | 'insideV', isOuter: boolean) => {
        const ln = kid(tcPr, own);
        if (ln) return readLine(ln, env.ctx.color, env.owner);
        return (isOuter ? eff.borders[outer] : eff.borders[inner]) ?? null;
      };
      const textOverrides: RunProps = {};
      if (eff.bold !== null) textOverrides.b = eff.bold;
      if (eff.italic !== null) textOverrides.i = eff.italic;
      if (eff.color) textOverrides.color = eff.color;
      return {
        text: readTextBody(kid(tc, 'a:txBody'), env, null, [], { fill: null, line: null, fontColor: null, fontFamily: null }, textOverrides, true)!,
        fill: ownFill ?? eff.fill,
        borders: {
          l: side('a:lnL', 'left', 'insideV', c === 0),
          r: side('a:lnR', 'right', 'insideV', c === tcs.length - 1),
          t: side('a:lnT', 'top', 'insideH', r === 0),
          b: side('a:lnB', 'bottom', 'insideH', r === trs.length - 1),
        },
        colSpan: num(tc, 'gridSpan', 1),
        rowSpan: num(tc, 'rowSpan', 1),
        merged: bool(tc, 'hMerge', false) || bool(tc, 'vMerge', false),
        margins: {
          l: num(tcPr, 'marL', 91440),
          r: num(tcPr, 'marR', 91440),
          t: num(tcPr, 'marT', 45720),
          b: num(tcPr, 'marB', 45720),
        },
        anchor: (attr(tcPr, 'anchor') as TableCell['anchor']) ?? 't',
      };
    });
    return { h: num(tr, 'h', 0), cells };
  });
  return { ...base, type: 'table', cols, rows } as SlideElement;
}

// ---------------------------------------------------------------- text

const DEFAULT_BODY: BodyProps = {
  lIns: 91440,
  tIns: 45720,
  rIns: 91440,
  bIns: 45720,
  anchor: 't',
  wrap: true,
  autofit: 'none',
  fontScale: 1,
  lnSpcReduction: 0,
  vert: null,
};

function applyBodyPr(base: BodyProps, el: Element | null): BodyProps {
  if (!el) return base;
  const out = { ...base };
  const n = (k: 'lIns' | 'tIns' | 'rIns' | 'bIns') => {
    const v = num(el, k);
    if (v !== null) out[k] = v;
  };
  n('lIns');
  n('tIns');
  n('rIns');
  n('bIns');
  const anchor = attr(el, 'anchor');
  if (anchor === 't' || anchor === 'ctr' || anchor === 'b') out.anchor = anchor;
  const wrap = attr(el, 'wrap');
  if (wrap) out.wrap = wrap !== 'none';
  const vert = attr(el, 'vert');
  if (vert) out.vert = vert === 'horz' ? null : vert;
  if (kid(el, 'a:noAutofit')) out.autofit = 'none';
  if (kid(el, 'a:spAutoFit')) out.autofit = 'shape';
  const norm = kid(el, 'a:normAutofit');
  if (norm) {
    out.autofit = 'norm';
    out.fontScale = num(norm, 'fontScale', 100000) / 100000;
    out.lnSpcReduction = num(norm, 'lnSpcReduction', 0) / 100000;
  }
  return out;
}

function seedLevel(ctx: Context): LevelStyle {
  return {
    marL: 0,
    indent: 0,
    align: 'l',
    lineSpacing: { pct: 1 },
    spaceBefore: { pts: 0 },
    spaceAfter: { pts: 0 },
    bullet: { kind: 'none' },
    run: {
      size: 18,
      b: false,
      i: false,
      u: false,
      strike: false,
      color: resolveColor(schemeEl('tx1'), ctx.color) ?? { hex: '000000', alpha: 1 },
      font: ctx.theme.minorFont,
      baseline: 0,
    },
  };
}

let scratchDoc: Document | null = null;
function schemeEl(val: string): Element {
  scratchDoc ??= parseXml(`<a:root xmlns:a="${A}"/>`);
  const el = scratchDoc.createElementNS(A, 'a:schemeClr');
  el.setAttribute('val', val);
  return el;
}

function resolveFont(name: string | null, theme: Theme): string | null {
  if (!name) return null;
  if (name.startsWith('+mj')) return theme.majorFont;
  if (name.startsWith('+mn')) return theme.minorFont;
  return name;
}

function readSpacing(el: Element | null): Spacing | undefined {
  if (!el) return undefined;
  const pct = kid(el, 'a:spcPct');
  if (pct) return { pct: num(pct, 'val', 100000) / 100000 };
  const pts = kid(el, 'a:spcPts');
  if (pts) return { pts: num(pts, 'val', 0) / 100 };
  return undefined;
}

/** Parse a:rPr / a:defRPr / a:endParaRPr. */
export function readRunProps(el: Element | null, ctx: Context, deck?: Deck): RunProps {
  const out: RunProps = {};
  if (!el) return out;
  const sz = num(el, 'sz');
  if (sz !== null) out.size = sz / 100;
  const b = attr(el, 'b');
  if (b !== null) out.b = b === '1' || b === 'true';
  const i = attr(el, 'i');
  if (i !== null) out.i = i === '1' || i === 'true';
  const u = attr(el, 'u');
  if (u !== null) out.u = u !== 'none';
  const strike = attr(el, 'strike');
  if (strike !== null) out.strike = strike !== 'noStrike';
  const baseline = num(el, 'baseline');
  if (baseline !== null) out.baseline = baseline / 1000;
  const fill = kid(el, 'a:solidFill');
  if (fill) {
    const c = resolveColor(colorChild(fill), ctx.color);
    if (c) out.color = c;
  } else if (kid(el, 'a:gradFill')) {
    const first = kid(kid(kid(el, 'a:gradFill'), 'a:gsLst'), 'a:gs');
    const c = resolveColor(colorChild(first), ctx.color);
    if (c) out.color = c;
  }
  const font = resolveFont(attr(kid(el, 'a:latin'), 'typeface'), ctx.theme);
  if (font) {
    out.font = font;
    deck?.fonts.add(font);
  }
  return out;
}

function readBullet(pPr: Element, ctx: Context): Partial<LevelStyle> {
  const out: Partial<LevelStyle> = {};
  const color = kid(pPr, 'a:buClrTx') ? null : resolveColor(colorChild(kid(pPr, 'a:buClr')), ctx.color);
  const sizePct = kid(pPr, 'a:buSzPct') ? num(kid(pPr, 'a:buSzPct'), 'val', 100000) / 100000 : 1;
  const font = kid(pPr, 'a:buFontTx') ? null : resolveFont(attr(kid(pPr, 'a:buFont'), 'typeface'), ctx.theme);
  if (kid(pPr, 'a:buNone')) out.bullet = { kind: 'none' };
  else if (kid(pPr, 'a:buChar')) out.bullet = { kind: 'char', char: attr(kid(pPr, 'a:buChar'), 'char') ?? '•', font, color, size: sizePct };
  else if (kid(pPr, 'a:buAutoNum')) {
    const an = kid(pPr, 'a:buAutoNum');
    out.bullet = { kind: 'num', scheme: attr(an, 'type') ?? 'arabicPeriod', startAt: num(an, 'startAt', 1), font, color, size: sizePct };
  } else if (kid(pPr, 'a:buBlip')) out.bullet = { kind: 'char', char: '•', font: null, color, size: sizePct };
  return out;
}

/** Parse a:pPr / a:lvlNpPr into partial level properties. */
export function readLevelProps(pPr: Element | null, ctx: Context, deck?: Deck): LevelProps {
  const out: LevelProps = {};
  if (!pPr) return out;
  const marL = num(pPr, 'marL');
  if (marL !== null) out.marL = marL;
  const indent = num(pPr, 'indent');
  if (indent !== null) out.indent = indent;
  const algn = attr(pPr, 'algn');
  if (algn) out.align = (['l', 'ctr', 'r', 'just', 'dist'].includes(algn) ? algn : 'l') as Align;
  const ls = readSpacing(kid(pPr, 'a:lnSpc'));
  if (ls) out.lineSpacing = ls;
  const sb = readSpacing(kid(pPr, 'a:spcBef'));
  if (sb) out.spaceBefore = sb;
  const sa = readSpacing(kid(pPr, 'a:spcAft'));
  if (sa) out.spaceAfter = sa;
  Object.assign(out, readBullet(pPr, ctx));
  // Bullet colour/size/font may be set without the bullet itself.
  const buClr = kid(pPr, 'a:buClr');
  const buSz = kid(pPr, 'a:buSzPct');
  const buFont = kid(pPr, 'a:buFont');
  if (!out.bullet && (buClr || buSz || buFont)) {
    (out as LevelProps & { bulletLook?: object }).bulletLook = {
      color: resolveColor(colorChild(buClr), ctx.color),
      size: buSz ? num(buSz, 'val', 100000) / 100000 : null,
      font: resolveFont(attr(buFont, 'typeface'), ctx.theme),
    };
  }
  const def = kid(pPr, 'a:defRPr');
  if (def) out.run = readRunProps(def, ctx, deck);
  return out;
}

export function mergeLevel(base: LevelStyle, p: LevelProps = {}): LevelStyle {
  const look = (p as LevelProps & { bulletLook?: { color: Color | null; size: number | null; font: string | null } }).bulletLook;
  let bullet = p.bullet ?? base.bullet;
  if (look && bullet.kind !== 'none') {
    bullet = {
      ...bullet,
      color: look.color ?? bullet.color,
      size: look.size ?? bullet.size,
      font: look.font ?? bullet.font,
    };
  }
  return {
    marL: p.marL ?? base.marL,
    indent: p.indent ?? base.indent,
    align: p.align ?? base.align,
    lineSpacing: p.lineSpacing ?? base.lineSpacing,
    spaceBefore: p.spaceBefore ?? base.spaceBefore,
    spaceAfter: p.spaceAfter ?? base.spaceAfter,
    bullet,
    run: { ...base.run, ...stripUndefined(p.run ?? {}) },
  };
}

function stripUndefined<T extends object>(o: T): T {
  const out = {} as T;
  for (const [k, v] of Object.entries(o)) if (v !== undefined) (out as Record<string, unknown>)[k] = v;
  return out;
}

function levelsOf(list: Element | null, ctx: Context, deck: Deck): LevelProps[] {
  const out: LevelProps[] = [];
  for (let i = 1; i <= 9; i++) out.push(readLevelProps(kid(list, `a:lvl${i}pPr`), ctx, deck));
  return out;
}

function applyLevels(base: LevelStyle[], list: Element | null, ctx: Context, deck: Deck): LevelStyle[] {
  if (!list) return base;
  const props = levelsOf(list, ctx, deck);
  return base.map((b, i) => mergeLevel(b, props[i]));
}

function readTextBody(
  txBody: Element | null,
  env: Env,
  ph: { type: string; idx: number | null } | null,
  chain: PlaceholderDef[],
  refs: StyleRefs,
  overrides: RunProps = {},
  tableCell = false,
): TextBody | null {
  const ctx = env.ctx;
  const deck = env.deck;
  // 1. The base style list: the master's title/body/other style, or the presentation default.
  const seed = seedLevel(ctx);
  let levels: LevelStyle[] = Array.from({ length: 9 }, () => seed);
  const type = ph?.type ?? null;
  let baseList: Element | null;
  if (type && masterType(type) === 'title') baseList = env.master.titleStyle;
  else if (type && masterType(type) === 'body') baseList = env.master.bodyStyle;
  else if (type) baseList = env.master.otherStyle;
  else baseList = deck.defaultTextStyle ?? env.master.otherStyle;
  if (!type && !tableCell && deck.defaultTextStyle) levels = applyLevels(levels, env.master.otherStyle, ctx, deck);
  levels = applyLevels(levels, baseList, ctx, deck);
  if (tableCell) levels = levels.map((l) => ({ ...l, run: { ...l.run, size: Math.min(l.run.size, 18) } }));

  // 2. The shape style's font reference.
  if (refs.fontColor || refs.fontFamily) {
    levels = levels.map((l) => ({
      ...l,
      run: { ...l.run, ...(refs.fontColor ? { color: refs.fontColor } : {}), ...(refs.fontFamily ? { font: refs.fontFamily } : {}) },
    }));
  }
  if (Object.keys(overrides).length) levels = levels.map((l) => ({ ...l, run: { ...l.run, ...stripUndefined(overrides) } }));

  // 3. Inherited placeholder list styles and body settings (master, then layout).
  let body = { ...DEFAULT_BODY };
  for (const d of chain) {
    const tb = kid(d.sp, 'p:txBody');
    levels = applyLevels(levels, kid(tb, 'a:lstStyle'), ctx, deck);
    body = applyBodyPr(body, kid(tb, 'a:bodyPr'));
  }
  // 4. The shape's own.
  levels = applyLevels(levels, kid(txBody, 'a:lstStyle'), ctx, deck);
  body = applyBodyPr(body, kid(txBody, 'a:bodyPr'));

  if (!txBody) {
    // A placeholder with no text yet still needs its styles (for typing into it).
    if (!ph && !tableCell) return null;
    return { paragraphs: [emptyParagraph()], body, levels };
  }

  const paragraphs: Paragraph[] = kids(txBody, 'a:p').map((p) => {
    const pPr = kid(p, 'a:pPr');
    const level = Math.min(8, Math.max(0, num(pPr, 'lvl', 0)));
    const items: InlineItem[] = [];
    for (const c of kids(p)) {
      if (is(c, 'a:r')) {
        const rPr = kid(c, 'a:rPr');
        const props = readRunProps(rPr, ctx, deck);
        const link = kid(rPr, 'a:hlinkClick');
        if (link) {
          props.color ??= resolveColor(schemeEl('hlink'), ctx.color) ?? undefined;
          props.u ??= true;
        }
        items.push({ kind: 'text', text: kid(c, 'a:t')?.textContent ?? '', props, rPr: rPr ? serialize(rPr) : null, link: attr(link, 'r:id') });
      } else if (is(c, 'a:br')) {
        const rPr = kid(c, 'a:rPr');
        items.push({ kind: 'br', rPr: rPr ? serialize(rPr) : null });
      } else if (is(c, 'a:fld')) {
        items.push({ kind: 'field', text: kid(c, 'a:t')?.textContent ?? '', xml: serialize(c), props: readRunProps(kid(c, 'a:rPr'), ctx, deck) });
      }
    }
    const end = kid(p, 'a:endParaRPr');
    return {
      items,
      level,
      props: readLevelProps(pPr, ctx, deck),
      pPr: pPr ? serialize(pPr) : null,
      endRPr: end ? serialize(end) : null,
      endSize: num(end, 'sz') !== null ? num(end, 'sz', 1800) / 100 : null,
    };
  });
  return { paragraphs: paragraphs.length ? paragraphs : [emptyParagraph()], body, levels };
}

export function emptyParagraph(): Paragraph {
  return { items: [], level: 0, props: {}, pPr: null, endRPr: null, endSize: null };
}

/** Levels and body settings for a new text box (no placeholder). */
export function textBoxStyle(deck: Deck, slide: Slide): { levels: LevelStyle[]; body: BodyProps } {
  const master = deck.masters.get(deck.layouts.get(slide.layout)?.master ?? '') ?? [...deck.masters.values()][0];
  const env: Env = { deck, owner: slide.part ?? '', ctx: slide.ctx, master, layout: deck.layouts.get(slide.layout) ?? null, decorations: false };
  const tb = readTextBody(null, env, { type: '__textbox', idx: null }, [], { fill: null, line: null, fontColor: null, fontFamily: null });
  const base = tb ?? { levels: [], body: DEFAULT_BODY };
  // Text boxes use the presentation default style rather than "other".
  const levels = applyLevels(
    Array.from({ length: 9 }, () => seedLevel(slide.ctx)),
    deck.defaultTextStyle ?? master.otherStyle,
    slide.ctx,
    deck,
  );
  return { levels, body: { ...base.body, wrap: true, autofit: 'shape' } };
}

/** Text styles for a shape placed on `slide` that takes the given placeholder. */
export function placeholderText(deck: Deck, slide: Slide, type: string, idx: number | null): TextBody | null {
  const layout = deck.layouts.get(slide.layout) ?? null;
  const master = deck.masters.get(layout?.master ?? '') ?? [...deck.masters.values()][0];
  const env: Env = { deck, owner: slide.part ?? '', ctx: slide.ctx, master, layout, decorations: false };
  const layoutDef = findLayoutPh(layout, type, idx);
  const masterDef = findMasterPh(master, layoutDef?.type ?? type);
  const chain = [masterDef, layoutDef].filter((d): d is PlaceholderDef => !!d);
  return readTextBody(null, env, { type, idx }, chain, { fill: null, line: null, fontColor: null, fontFamily: null });
}

export { findLayoutPh, findMasterPh, promptFor };

/** Raw text of an XML part, for writers that need it. */
export function partText(pkg: Pkg, name: string): string | null {
  const b = pkg.bytes(name);
  return b ? strFromU8(b) : null;
}

export type { Env };
