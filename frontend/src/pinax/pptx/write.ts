/**
 * Pinax model -> .pptx bytes.
 *
 * Every slide starts from its original XML (cloned, never mutated) and only what changed is
 * patched in: positions, fills and outlines that were edited, text bodies that were edited, and
 * the element list itself (added, deleted, reordered). Anything Pinax does not model rides along
 * untouched. The result is checked before it is handed back: if it is not a valid package,
 * nothing is saved.
 */
import { strFromU8, unzipSync } from 'fflate';
import { GateError, inspectPackage, isSafeProlog } from '../../lib/fileGate';
import type { MediaStore } from '../media';
import type { Deck, DeckState, Paragraph, RunProps, Slide, SlideElement, TableEl, TextBody, Xfrm } from '../model';
import { CT, PkgWriter, REL, resolveTarget } from './package';
import { readRunProps } from './read';
import type { Color, Fill, Line } from './theme';
import { NS, attr, insertOrdered, is, kid, kids, make, parseXml, path, removeKids, setAttr } from './xml';

export class SaveError extends Error {}

export const NEW_OWNER = '__new__';
export const MEDIA_REF = '__media:';

type RelRec = { id: string; type: string; target: string; external: boolean };

export function writeDeck(deck: Deck, state: DeckState, media: MediaStore): Uint8Array {
  const pkg = deck.pkg;
  const w = new PkgWriter(pkg);
  for (const [part, bytes] of media.addedParts()) w.parts.set(part, bytes);

  const pres = pkg.xml(deck.presPart)!.cloneNode(true) as Document;
  const presRoot = pres.documentElement;
  const presRels = w.relsOf(pkg, deck.presPart);

  // Existing slide ids by part, so unchanged slides keep their ids.
  const idByPart = new Map<string, string>();
  for (const sid of kids(kid(presRoot, 'p:sldIdLst'), 'p:sldId')) {
    const rel = presRels.find((r) => r.id === attr(sid, 'r:id'));
    if (rel) idByPart.set(rel.target, attr(sid, 'id') ?? '');
  }
  let nextSlideId = Math.max(255, ...[...idByPart.values()].map(Number).filter(Number.isFinite)) + 1;

  const used = new Set<string>();
  const entries: { part: string; id: string }[] = [];
  for (const slide of state.slides) {
    const reuse = slide.part && !used.has(slide.part) ? slide.part : null;
    const target = reuse ?? w.freshName('ppt/slides', 'slide', 'xml');
    used.add(target);
    if (!reuse) w.parts.set(target, new Uint8Array()); // reserve the name
    writeSlide(deck, w, slide, target);
    if (!reuse) w.overrides.set(target, CT.slide);
    entries.push({ part: target, id: (reuse && idByPart.get(reuse)) || String(nextSlideId++) });
  }

  // presentation.xml: slide relationships and the slide list.
  const kept = presRels.filter((r) => r.type !== REL.slide || used.has(r.target));
  w.setRels(deck.presPart, kept);
  const list = kid(presRoot, 'p:sldIdLst') ?? insertSlideList(pres);
  removeKids(list, ['p:sldId']);
  for (const e of entries) {
    const rId = w.addRel(pkg, deck.presPart, REL.slide, e.part);
    list.appendChild(make(pres, 'p:sldId', { id: e.id, 'r:id': rId }));
  }
  if (!entries.length) presRoot.removeChild(list);
  w.setXml(deck.presPart, pres);

  const bytes = w.finish();
  verify(bytes);
  return bytes;
}

function insertSlideList(pres: Document): Element {
  const root = pres.documentElement;
  const el = make(pres, 'p:sldIdLst');
  insertOrdered(root, el, ['p:sldMasterIdLst', 'p:notesMasterIdLst', 'p:handoutMasterIdLst', 'p:sldIdLst', 'p:sldSz', 'p:notesSz']);
  return el;
}

/** The saved file must be a package the gate itself would accept, with every part parseable. */
function verify(bytes: Uint8Array) {
  try {
    inspectPackage(bytes, 'pptx');
  } catch (err) {
    throw new SaveError(`Pinax built an invalid file (${err instanceof GateError ? err.message : 'unknown problem'}), so nothing was saved.`);
  }
  const files = unzipSync(bytes);
  for (const [name, data] of Object.entries(files)) {
    if (!/\.(xml|rels)$/i.test(name)) continue;
    if (!isSafeProlog(data)) throw new SaveError(`Pinax built an invalid part (${name}), so nothing was saved.`);
    let doc: Document;
    try {
      doc = parseXml(strFromU8(data));
    } catch {
      throw new SaveError(`Pinax built an invalid part (${name}), so nothing was saved.`);
    }
    if (name.endsWith('.rels')) {
      const owner = name.replace(/_rels\/([^/]+)\.rels$/, '$1');
      for (const r of kids(doc.documentElement)) {
        if ((attr(r, 'TargetMode') ?? '').toLowerCase() === 'external') continue;
        const t = attr(r, 'Target') ?? '';
        const resolved = name === '_rels/.rels' ? t.replace(/^\//, '') : resolveTarget(owner, t);
        if (!(resolved in files)) throw new SaveError(`Pinax built a broken link (${name} -> ${t}), so nothing was saved.`);
      }
    }
  }
}

// ---------------------------------------------------------------- slides

function writeSlide(deck: Deck, w: PkgWriter, slide: Slide, target: string) {
  const pkg = deck.pkg;
  const src = slide.part;
  const doc = (slide.doc ?? parseXml(EMPTY_SLIDE)).cloneNode(true) as Document;
  const root = doc.documentElement;

  // Relationships: a copy of the source slide's, without its notes (handled below).
  const srcRels: RelRec[] = src ? pkg.rels(src).map((r) => ({ ...r })) : [];
  const srcNotes = srcRels.find((r) => r.type === REL.notesSlide) ?? null;
  const rels = srcRels.filter((r) => r.type !== REL.notesSlide && r.type !== REL.slideLayout);
  w.setRels(target, rels);
  w.addRel(pkg, target, REL.slideLayout, slide.layout);

  const ctx: PatchCtx = { deck, w, target, src, doc, slide, nextId: 2 };

  const tree = path(root, 'p:cSld', 'p:spTree');
  if (tree) {
    ctx.nextId = maxShapeId(tree) + 1;
    for (const c of kids(tree)) {
      if (is(c, 'p:nvGrpSpPr') || is(c, 'p:grpSpPr') || is(c, 'p:extLst')) continue;
      tree.removeChild(c);
    }
    const ext = kid(tree, 'p:extLst');
    for (const el of slide.elements) {
      const node = elementNode(ctx, el);
      if (ext) tree.insertBefore(node, ext);
      else tree.appendChild(node);
    }
  }

  if (slide.hidden) root.setAttribute('show', '0');
  else root.removeAttribute('show');

  writeNotes(ctx, srcNotes);
  w.setXml(target, doc);
}

const EMPTY_SLIDE = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`;

interface PatchCtx {
  deck: Deck;
  w: PkgWriter;
  target: string;
  src: string | null;
  doc: Document;
  slide: Slide;
  nextId: number;
}

function maxShapeId(tree: Element): number {
  let max = 1;
  for (const el of Array.from(tree.getElementsByTagNameNS(NS.p, 'cNvPr'))) max = Math.max(max, Number(attr(el, 'id')) || 0);
  return max;
}

/** The XML for one element in the target slide: its original node, cloned and patched. */
function elementNode(ctx: PatchCtx, el: SlideElement): Element {
  const node = ctx.doc.importNode(el.xml!, true) as Element;
  stripNsDecl(node);
  remapRelIds(ctx, node, el.owner);
  if (el.owner === NEW_OWNER || el.owner !== (ctx.src ?? '')) freshIds(ctx, node);
  const targets = is(node, 'mc:AlternateContent')
    ? kids(node).map((b) => kids(b)[0]).filter((x): x is Element => !!x)
    : [node];
  for (const t of targets) patchElement(ctx, t, el);
  return node;
}

/** Imported fragments carry their own xmlns declarations; the slide root already has them. */
function stripNsDecl(el: Element) {
  for (const a of Array.from(el.attributes)) {
    if (a.name === 'xmlns:a' || a.name === 'xmlns:p' || a.name === 'xmlns:r') el.removeAttribute(a.name);
  }
}

/** Give pasted / generated shapes ids that do not collide with the slide's own. */
function freshIds(ctx: PatchCtx, node: Element) {
  for (const c of Array.from(node.getElementsByTagNameNS(NS.p, 'cNvPr'))) c.setAttribute('id', String(ctx.nextId++));
  if (is(node, 'p:cNvPr')) node.setAttribute('id', String(ctx.nextId++));
}

/**
 * r:ids in a node refer to the relationships of the part it came from. When it lands on a
 * different part (pasted, duplicated, or made in Pinax), add matching relationships to the
 * target and rewrite the ids.
 */
function remapRelIds(ctx: PatchCtx, node: Element, owner: string) {
  const sameOwner = owner === ctx.src && owner !== NEW_OWNER;
  const all = [node, ...Array.from(node.getElementsByTagName('*'))];
  for (const el of all) {
    for (const a of Array.from(el.attributes)) {
      if (a.namespaceURI !== NS.r) continue;
      const v = a.value;
      if (v.startsWith(MEDIA_REF)) {
        a.value = ctx.w.addRel(ctx.deck.pkg, ctx.target, REL.image, v.slice(MEDIA_REF.length));
        continue;
      }
      if (sameOwner) continue;
      const rel = owner && owner !== NEW_OWNER ? ctx.deck.pkg.rel(owner, v) : null;
      if (rel) a.value = ctx.w.addRel(ctx.deck.pkg, ctx.target, rel.type, rel.target, rel.external);
    }
  }
}

// ---------------------------------------------------------------- patching

const SPPR_ORDER = ['a:xfrm', 'a:custGeom', 'a:prstGeom', 'a:noFill', 'a:solidFill', 'a:gradFill', 'a:blipFill', 'a:pattFill', 'a:grpFill', 'a:ln', 'a:effectLst', 'a:effectDag', 'a:scene3d', 'a:sp3d', 'a:extLst'];
const FILLS = ['a:noFill', 'a:solidFill', 'a:gradFill', 'a:blipFill', 'a:pattFill', 'a:grpFill'];

function patchElement(ctx: PatchCtx, node: Element, el: SlideElement) {
  const doc = ctx.doc;
  if (is(node, 'p:sp') || is(node, 'p:cxnSp') || is(node, 'p:pic')) {
    const spPr = kid(node, 'p:spPr') ?? insertSpPr(doc, node);
    if (el.changed.xfrm) setXfrm(doc, spPr, el.xfrm, 'a:xfrm');
    if (el.changed.fill && el.type === 'shape') setFill(doc, spPr, el.fill);
    if (el.changed.line && (el.type === 'shape' || el.type === 'picture')) setLine(doc, spPr, el.line);
    if (el.changed.text && el.type === 'shape' && el.text) writeTextBody(ctx, node, el.text, 'p:txBody');
  } else if (is(node, 'p:grpSp')) {
    if (el.changed.xfrm) {
      const pr = kid(node, 'p:grpSpPr') ?? node.insertBefore(make(doc, 'p:grpSpPr'), kids(node)[1] ?? null);
      setXfrm(doc, pr, el.xfrm, 'a:xfrm');
    }
  } else if (is(node, 'p:graphicFrame')) {
    if (el.changed.xfrm) {
      let x = kid(node, 'p:xfrm');
      if (!x) {
        x = make(doc, 'p:xfrm');
        node.insertBefore(x, kid(node, 'a:graphic'));
      }
      setXfrm(doc, node, el.xfrm, 'p:xfrm');
    }
    if (el.changed.cells && el.type === 'table') writeCells(ctx, node, el);
  }
}

function insertSpPr(doc: Document, node: Element): Element {
  const spPr = make(doc, 'p:spPr');
  const after = kids(node)[0];
  node.insertBefore(spPr, after?.nextElementSibling ?? null);
  return spPr;
}

function setXfrm(doc: Document, parent: Element, x: Xfrm, tag: 'a:xfrm' | 'p:xfrm') {
  let xf = kid(parent, tag);
  if (!xf) {
    xf = make(doc, tag);
    parent.insertBefore(xf, parent.firstElementChild);
  }
  const rot = Math.round(x.rot * 60000) % 21600000;
  setAttr(xf, 'rot', rot ? String(rot < 0 ? rot + 21600000 : rot) : null);
  setAttr(xf, 'flipH', x.flipH ? '1' : null);
  setAttr(xf, 'flipV', x.flipV ? '1' : null);
  let off = kid(xf, 'a:off');
  if (!off) {
    off = make(doc, 'a:off');
    xf.insertBefore(off, xf.firstElementChild);
  }
  off.setAttribute('x', String(Math.round(x.x)));
  off.setAttribute('y', String(Math.round(x.y)));
  let ext = kid(xf, 'a:ext');
  if (!ext) {
    ext = make(doc, 'a:ext');
    xf.insertBefore(ext, off.nextElementSibling);
  }
  ext.setAttribute('cx', String(Math.max(0, Math.round(x.w))));
  ext.setAttribute('cy', String(Math.max(0, Math.round(x.h))));
}

function colorEl(doc: Document, c: Color): Element {
  const el = make(doc, 'a:srgbClr', { val: c.hex });
  if (c.alpha < 1) el.appendChild(make(doc, 'a:alpha', { val: Math.round(c.alpha * 100000) }));
  return el;
}

function fillEl(doc: Document, fill: Fill | null): Element | null {
  if (!fill || fill.kind === 'none') return make(doc, 'a:noFill');
  if (fill.kind === 'solid') {
    const s = make(doc, 'a:solidFill');
    s.appendChild(colorEl(doc, fill.color));
    return s;
  }
  return null; // other fills are never produced by Pinax edits
}

function setFill(doc: Document, spPr: Element, fill: Fill | null) {
  const el = fillEl(doc, fill);
  if (!el) return;
  removeKids(spPr, FILLS);
  insertOrdered(spPr, el, SPPR_ORDER);
}

function setLine(doc: Document, spPr: Element, line: Line | null) {
  let ln = kid(spPr, 'a:ln');
  if (!ln) {
    ln = make(doc, 'a:ln');
    insertOrdered(spPr, ln, SPPR_ORDER);
  }
  if (!line || line.fill.kind === 'none') {
    removeKids(ln, ['a:noFill', 'a:solidFill', 'a:gradFill', 'a:pattFill']);
    ln.insertBefore(make(doc, 'a:noFill'), ln.firstElementChild);
    return;
  }
  ln.setAttribute('w', String(Math.round(line.width)));
  const f = fillEl(doc, line.fill);
  if (f) {
    removeKids(ln, ['a:noFill', 'a:solidFill', 'a:gradFill', 'a:pattFill']);
    ln.insertBefore(f, ln.firstElementChild);
  }
}

// ---------------------------------------------------------------- text

function importFragment(doc: Document, xml: string): Element {
  const el = doc.importNode(parseXml(xml).documentElement, true) as Element;
  stripNsDecl(el);
  return el;
}

/** Replace the paragraphs of a txBody (creating it when the shape had none). */
function writeTextBody(ctx: PatchCtx, node: Element, body: TextBody, tag: 'p:txBody' | 'a:txBody') {
  const doc = ctx.doc;
  let tb = kid(node, tag);
  if (!tb) {
    tb = make(doc, tag);
    tb.appendChild(make(doc, 'a:bodyPr'));
    tb.appendChild(make(doc, 'a:lstStyle'));
    const before = kid(node, 'p:extLst') ?? (tag === 'a:txBody' ? kid(node, 'a:tcPr') : null);
    node.insertBefore(tb, before);
  }
  removeKids(tb, ['a:p']);
  const ext = kid(tb, 'a:extLst');
  for (const p of body.paragraphs) {
    const pel = paragraphEl(ctx, p);
    if (ext) tb.insertBefore(pel, ext);
    else tb.appendChild(pel);
  }
}

const PPR_ORDER = ['a:lnSpc', 'a:spcBef', 'a:spcAft', 'a:buClrTx', 'a:buClr', 'a:buSzTx', 'a:buSzPct', 'a:buSzPts', 'a:buFontTx', 'a:buFont', 'a:buNone', 'a:buAutoNum', 'a:buChar', 'a:buBlip', 'a:tabLst', 'a:defRPr', 'a:extLst'];
const RPR_ORDER = ['a:ln', 'a:noFill', 'a:solidFill', 'a:gradFill', 'a:blipFill', 'a:pattFill', 'a:grpFill', 'a:effectLst', 'a:effectDag', 'a:highlight', 'a:uLnTx', 'a:uLn', 'a:uFillTx', 'a:uFill', 'a:latin', 'a:ea', 'a:cs', 'a:sym', 'a:hlinkClick', 'a:hlinkMouseOver', 'a:rtl', 'a:extLst'];

function spacingEl(doc: Document, tag: string, s: { pct?: number; pts?: number }): Element {
  const el = make(doc, tag);
  if (s.pts !== undefined) el.appendChild(make(doc, 'a:spcPts', { val: Math.round(s.pts * 100) }));
  else el.appendChild(make(doc, 'a:spcPct', { val: Math.round((s.pct ?? 1) * 100000) }));
  return el;
}

function paragraphEl(ctx: PatchCtx, p: Paragraph): Element {
  const doc = ctx.doc;
  const pel = make(doc, 'a:p');
  const pPr = p.pPr ? importFragment(doc, p.pPr) : make(doc, 'a:pPr');
  setAttr(pPr, 'lvl', p.level ? String(p.level) : null);
  const pr = p.props;
  setAttr(pPr, 'algn', pr.align ?? null);
  setAttr(pPr, 'marL', pr.marL !== undefined ? String(Math.round(pr.marL)) : null);
  setAttr(pPr, 'indent', pr.indent !== undefined ? String(Math.round(pr.indent)) : null);
  for (const [key, tag] of [['lineSpacing', 'a:lnSpc'], ['spaceBefore', 'a:spcBef'], ['spaceAfter', 'a:spcAft']] as const) {
    removeKids(pPr, [tag]);
    const v = pr[key];
    if (v) insertOrdered(pPr, spacingEl(doc, tag, v), PPR_ORDER);
  }
  if (pr.bullet) {
    removeKids(pPr, ['a:buNone', 'a:buAutoNum', 'a:buChar', 'a:buBlip']);
    const b = pr.bullet;
    const el =
      b.kind === 'none'
        ? make(doc, 'a:buNone')
        : b.kind === 'char'
          ? make(doc, 'a:buChar', { char: b.char })
          : make(doc, 'a:buAutoNum', { type: b.scheme, startAt: b.startAt !== 1 ? b.startAt : null });
    insertOrdered(pPr, el, PPR_ORDER);
    if (b.kind !== 'none' && b.font && !kid(pPr, 'a:buFont')) insertOrdered(pPr, make(doc, 'a:buFont', { typeface: b.font }), PPR_ORDER);
  }
  if (pPr.attributes.length || pPr.firstElementChild) pel.appendChild(pPr);

  for (const it of p.items) {
    if (it.kind === 'br') {
      const br = make(doc, 'a:br');
      if (it.rPr) br.appendChild(importFragment(doc, it.rPr));
      pel.appendChild(br);
    } else if (it.kind === 'field') {
      pel.appendChild(importFragment(doc, it.xml));
    } else {
      const r = make(doc, 'a:r');
      r.appendChild(runPr(ctx, it.rPr, it.props));
      const t = make(doc, 'a:t');
      t.textContent = it.text;
      r.appendChild(t);
      pel.appendChild(r);
    }
  }
  if (p.endRPr) pel.appendChild(importFragment(doc, p.endRPr));
  else if (!p.items.length) pel.appendChild(make(doc, 'a:endParaRPr', { lang: 'en-US', dirty: '0' }));
  return pel;
}

/** The run's a:rPr: the original with Pinax's edits applied, theme references kept when unchanged. */
function runPr(ctx: PatchCtx, original: string | null, props: RunProps): Element {
  const doc = ctx.doc;
  const el = original ? importFragment(doc, original) : make(doc, 'a:rPr', { lang: 'en-US', dirty: '0' });
  const before = original ? readRunProps(parseXml(original).documentElement, ctx.slide.ctx) : {};
  const flag = (name: string, v: boolean | undefined, on = '1', off = '0') => setAttr(el, name, v === undefined ? null : v ? on : off);
  flag('b', props.b);
  flag('i', props.i);
  if (props.u !== before.u) setAttr(el, 'u', props.u === undefined ? null : props.u ? 'sng' : 'none');
  if (props.strike !== before.strike) setAttr(el, 'strike', props.strike === undefined ? null : props.strike ? 'sngStrike' : 'noStrike');
  setAttr(el, 'sz', props.size !== undefined ? String(Math.round(props.size * 100)) : null);
  setAttr(el, 'baseline', props.baseline ? String(Math.round(props.baseline * 1000)) : null);
  const sameColor = props.color && before.color && props.color.hex === before.color.hex && props.color.alpha === before.color.alpha;
  if (!sameColor) {
    if (props.color) {
      removeKids(el, ['a:noFill', 'a:solidFill', 'a:gradFill', 'a:blipFill', 'a:pattFill', 'a:grpFill']);
      const s = make(doc, 'a:solidFill');
      s.appendChild(colorEl(doc, props.color));
      insertOrdered(el, s, RPR_ORDER);
    } else if (before.color) {
      removeKids(el, ['a:solidFill', 'a:gradFill']);
    }
  }
  if (props.font !== before.font) {
    removeKids(el, ['a:latin']);
    if (props.font) insertOrdered(el, make(doc, 'a:latin', { typeface: props.font }), RPR_ORDER);
  }
  return el;
}

// ---------------------------------------------------------------- tables

function writeCells(ctx: PatchCtx, node: Element, el: TableEl) {
  const tbl = path(node, 'a:graphic', 'a:graphicData', 'a:tbl');
  if (!tbl) return;
  kids(tbl, 'a:tr').forEach((tr, r) => {
    kids(tr, 'a:tc').forEach((tc, c) => {
      const cell = el.rows[r]?.cells[c];
      if (cell?.changed) writeTextBody(ctx, tc, cell.text, 'a:txBody');
    });
  });
}

// ---------------------------------------------------------------- notes

function writeNotes(ctx: PatchCtx, srcNotes: RelRec | null) {
  const { deck, w, slide, target } = ctx;
  const sameSlide = target === slide.part;
  if (slide.notes === null) return;
  if (sameSlide && srcNotes && !slide.notesChanged) {
    w.addRel(deck.pkg, target, REL.notesSlide, srcNotes.target);
    return;
  }
  const paragraphs = slide.notes.length ? slide.notes : [''];
  if (sameSlide && srcNotes) {
    // Update the existing notes slide in place.
    const doc = deck.pkg.xml(srcNotes.target)?.cloneNode(true) as Document | undefined;
    const body = doc ? notesBody(doc) : null;
    if (doc && body) {
      fillNotes(doc, body, paragraphs);
      w.setXml(srcNotes.target, doc);
      w.addRel(deck.pkg, target, REL.notesSlide, srcNotes.target);
      return;
    }
  }
  if (!deck.notesMaster) return; // no notes master to build a notes page from
  if (!slide.notes.some((t) => t.trim())) return;
  const part = w.freshName('ppt/notesSlides', 'notesSlide', 'xml');
  const doc = parseXml(NOTES_SLIDE);
  fillNotes(doc, notesBody(doc)!, paragraphs);
  w.setXml(part, doc);
  w.overrides.set(part, CT.notesSlide);
  w.setRels(part, [
    { id: 'rId1', type: REL.notesMaster, target: deck.notesMaster, external: false },
    { id: 'rId2', type: REL.slide, target, external: false },
  ]);
  w.addRel(deck.pkg, target, REL.notesSlide, part);
}

function notesBody(doc: Document): Element | null {
  for (const sp of Array.from(doc.getElementsByTagNameNS(NS.p, 'sp'))) {
    const ph = path(sp, 'p:nvSpPr', 'p:nvPr', 'p:ph');
    if (attr(ph, 'type') !== 'body') continue;
    let tb = kid(sp, 'p:txBody');
    if (!tb) {
      tb = make(doc, 'p:txBody');
      tb.appendChild(make(doc, 'a:bodyPr'));
      tb.appendChild(make(doc, 'a:lstStyle'));
      sp.appendChild(tb);
    }
    return tb;
  }
  return null;
}

function fillNotes(doc: Document, tb: Element, paragraphs: string[]) {
  removeKids(tb, ['a:p']);
  for (const text of paragraphs) {
    const p = make(doc, 'a:p');
    if (text) {
      const r = make(doc, 'a:r');
      r.appendChild(make(doc, 'a:rPr', { lang: 'en-US', dirty: '0' }));
      const t = make(doc, 'a:t');
      t.textContent = text;
      r.appendChild(t);
      p.appendChild(r);
    } else p.appendChild(make(doc, 'a:endParaRPr', { lang: 'en-US', dirty: '0' }));
    tb.appendChild(p);
  }
}

const NOTES_SLIDE = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:notes xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldImg"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp><p:sp><p:nvSpPr><p:cNvPr id="3" name="Notes Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/></p:txBody></p:sp></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>`;

