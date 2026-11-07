/**
 * Editing operations on the Pinax model. All of them return new objects (the model is
 * immutable, which is what makes undo a stack of snapshots).
 *
 * New content is generated as PresentationML and read back through the same reader as a file,
 * so a shape made in Pinax is modelled exactly as it will be saved.
 */
import { type Deck, type Paragraph, type ShapeEl, type Slide, type SlideElement, type TextBody, type Xfrm, newId } from './model';
import type { Fill, Line } from './pptx/theme';
import { buildSlide, readFragment } from './pptx/read';
import { MEDIA_REF, NEW_OWNER } from './pptx/write';
import { NS, esc, parseXml } from './pptx/xml';

const HEAD = `xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}"`;

function fragment(deck: Deck, slide: Slide, inner: string): SlideElement[] {
  const doc = parseXml(`<p:spTree ${HEAD}>${inner}</p:spTree>`);
  return readFragment(deck, slide, doc.documentElement, NEW_OWNER);
}

const xfrmXml = (x: Xfrm) =>
  `<a:xfrm${x.rot ? ` rot="${Math.round(x.rot * 60000)}"` : ''}><a:off x="${Math.round(x.x)}" y="${Math.round(x.y)}"/><a:ext cx="${Math.round(x.w)}" cy="${Math.round(x.h)}"/></a:xfrm>`;

// ---------------------------------------------------------------- slides

/** A new slide from a layout, with that layout's (empty) placeholders. */
export function newSlide(deck: Deck, layoutPart: string): Slide | null {
  const layout = deck.layouts.get(layoutPart);
  if (!layout) return null;
  let id = 2;
  const shapes = layout.placeholders
    .filter((d) => !['dt', 'ftr', 'sldNum', 'sldImg', 'hdr'].includes(d.type))
    .map((d) => {
      const attrs = [d.type !== 'obj' ? `type="${d.type}"` : '', d.idx ? `idx="${d.idx}"` : ''].filter(Boolean).join(' ');
      const name = d.type === 'title' || d.type === 'ctrTitle' ? 'Title' : d.type === 'subTitle' ? 'Subtitle' : 'Content Placeholder';
      return `<p:sp><p:nvSpPr><p:cNvPr id="${id++}" name="${name} ${id - 1}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph ${attrs}/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US" dirty="0"/></a:p></p:txBody></p:sp>`;
    })
    .join('');
  const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld ${HEAD}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>${shapes}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`;
  const slide = buildSlide(deck, parseXml(xml), null, layoutPart, NEW_OWNER);
  return slide ? { ...slide, changed: true } : null;
}

export function duplicateSlide(slide: Slide): Slide {
  return {
    ...slide,
    id: newId('s'),
    elements: slide.elements.map((e) => ({ ...e, id: newId() })),
    notesChanged: slide.notes ? true : undefined,
    changed: true,
  };
}

// ---------------------------------------------------------------- elements

export function withXfrm<T extends SlideElement>(el: T, xfrm: Xfrm): T {
  return { ...el, xfrm, changed: { ...el.changed, xfrm: true } };
}

export function withText(el: ShapeEl, text: TextBody): ShapeEl {
  return { ...el, text, changed: { ...el.changed, text: true } };
}

export function withFill(el: ShapeEl, fill: Fill | null): ShapeEl {
  return { ...el, fill, changed: { ...el.changed, fill: true } };
}

export function withLine<T extends SlideElement & { line: Line | null }>(el: T, line: Line | null): T {
  return { ...el, line, changed: { ...el.changed, line: true } };
}

export function mapElements(slide: Slide, ids: ReadonlySet<string>, fn: (el: SlideElement) => SlideElement): Slide {
  return { ...slide, elements: slide.elements.map((e) => (ids.has(e.id) ? fn(e) : e)), changed: true };
}

export function deleteElements(slide: Slide, ids: ReadonlySet<string>): Slide {
  return { ...slide, elements: slide.elements.filter((e) => !ids.has(e.id)), changed: true };
}

/** Copies of elements (new ids, same original XML), offset so they are visible. */
export function copiesOf(elements: SlideElement[], offset: number): SlideElement[] {
  return elements.map((e) => withXfrm({ ...e, id: newId() }, { ...e.xfrm, x: e.xfrm.x + offset, y: e.xfrm.y + offset }));
}

export type ZMove = 'front' | 'back' | 'forward' | 'backward';

export function reorder(slide: Slide, ids: ReadonlySet<string>, move: ZMove): Slide {
  const els = [...slide.elements];
  const picked = els.filter((e) => ids.has(e.id));
  if (!picked.length) return slide;
  let out: SlideElement[];
  if (move === 'front') out = [...els.filter((e) => !ids.has(e.id)), ...picked];
  else if (move === 'back') out = [...picked, ...els.filter((e) => !ids.has(e.id))];
  else {
    out = els;
    const step = move === 'forward' ? 1 : -1;
    const order = move === 'forward' ? [...out.keys()].reverse() : [...out.keys()];
    for (const i of order) {
      if (!ids.has(out[i].id)) continue;
      const j = i + step;
      if (j < 0 || j >= out.length || ids.has(out[j].id)) continue;
      [out[i], out[j]] = [out[j], out[i]];
    }
  }
  return { ...slide, elements: out, changed: true };
}

// ---------------------------------------------------------------- new content

export function newTextBox(deck: Deck, slide: Slide, x: Xfrm): SlideElement | null {
  return (
    fragment(
      deck,
      slide,
      `<p:sp><p:nvSpPr><p:cNvPr id="2" name="TextBox"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr>${xfrmXml(x)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr><p:txBody><a:bodyPr wrap="square" rtlCol="0"><a:spAutoFit/></a:bodyPr><a:lstStyle/><a:p><a:endParaRPr lang="en-US" dirty="0"/></a:p></p:txBody></p:sp>`,
    )[0] ?? null
  );
}

export const SHAPE_CHOICES: { prst: string; label: string }[] = [
  { prst: 'rect', label: 'Rectangle' },
  { prst: 'roundRect', label: 'Rounded rectangle' },
  { prst: 'ellipse', label: 'Ellipse' },
  { prst: 'triangle', label: 'Triangle' },
  { prst: 'diamond', label: 'Diamond' },
  { prst: 'hexagon', label: 'Hexagon' },
  { prst: 'star5', label: 'Star' },
  { prst: 'rightArrow', label: 'Arrow' },
  { prst: 'chevron', label: 'Chevron' },
  { prst: 'wedgeRoundRectCallout', label: 'Callout' },
  { prst: 'line', label: 'Line' },
  { prst: 'arrowLine', label: 'Line with arrow' },
];

export function newShape(deck: Deck, slide: Slide, prst: string, x: Xfrm): SlideElement | null {
  if (prst === 'line' || prst === 'arrowLine') {
    const arrow = prst === 'arrowLine' ? '<a:tailEnd type="triangle"/>' : '';
    return (
      fragment(
        deck,
        slide,
        `<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="2" name="Line"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr><p:spPr>${xfrmXml(x)}<a:prstGeom prst="line"><a:avLst/></a:prstGeom><a:ln w="28575"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill>${arrow}</a:ln></p:spPr></p:cxnSp>`,
      )[0] ?? null
    );
  }
  return (
    fragment(
      deck,
      slide,
      `<p:sp><p:nvSpPr><p:cNvPr id="2" name="Shape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>${xfrmXml(x)}<a:prstGeom prst="${esc(prst)}"><a:avLst/></a:prstGeom></p:spPr><p:style><a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="50000"/></a:schemeClr></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style><p:txBody><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:endParaRPr lang="en-US" dirty="0"/></a:p></p:txBody></p:sp>`,
    )[0] ?? null
  );
}

export function newPicture(deck: Deck, slide: Slide, mediaPart: string, x: Xfrm): SlideElement | null {
  return (
    fragment(
      deck,
      slide,
      `<p:pic><p:nvPicPr><p:cNvPr id="2" name="Picture"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="${MEDIA_REF}${esc(mediaPart)}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr>${xfrmXml(x)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`,
    )[0] ?? null
  );
}

/** Replace the plain text of a shape (used for notes-like quick edits and tests). */
export function plainParagraphs(texts: string[]): Paragraph[] {
  return texts.map((t) => ({
    items: t ? [{ kind: 'text' as const, text: t, props: {}, rPr: null, link: null }] : [],
    level: 0,
    props: {},
    pPr: null,
    endRPr: null,
    endSize: null,
  }));
}
