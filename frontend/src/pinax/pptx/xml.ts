/** Small, namespace-aware helpers over the browser's XML DOM for PresentationML / DrawingML. */

export const NS = {
  p: 'http://schemas.openxmlformats.org/presentationml/2006/main',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  mc: 'http://schemas.openxmlformats.org/markup-compatibility/2006',
  rel: 'http://schemas.openxmlformats.org/package/2006/relationships',
  ct: 'http://schemas.openxmlformats.org/package/2006/content-types',
  c: 'http://schemas.openxmlformats.org/drawingml/2006/chart',
  dgm: 'http://schemas.openxmlformats.org/drawingml/2006/diagram',
  p14: 'http://schemas.microsoft.com/office/powerpoint/2010/main',
} as const;

export type Prefix = keyof typeof NS;

/** 'a:solidFill' -> [namespace, local name] */
function split(qname: string): [string, string] {
  const i = qname.indexOf(':');
  const prefix = qname.slice(0, i) as Prefix;
  return [NS[prefix], qname.slice(i + 1)];
}

export function is(el: Element | null | undefined, qname: string): boolean {
  if (!el) return false;
  const [ns, local] = split(qname);
  return el.namespaceURI === ns && el.localName === local;
}

/** Element children, optionally only those with the given qualified name. */
export function kids(el: Element | null | undefined, qname?: string): Element[] {
  if (!el) return [];
  const out: Element[] = [];
  for (let c = el.firstElementChild; c; c = c.nextElementSibling) {
    if (!qname || is(c, qname)) out.push(c);
  }
  return out;
}

export function kid(el: Element | null | undefined, qname: string): Element | null {
  if (!el) return null;
  for (let c = el.firstElementChild; c; c = c.nextElementSibling) if (is(c, qname)) return c;
  return null;
}

/** Follow a chain of child names: path(sp, 'p:spPr', 'a:xfrm'). */
export function path(el: Element | null | undefined, ...names: string[]): Element | null {
  let cur: Element | null | undefined = el;
  for (const n of names) {
    cur = kid(cur, n);
    if (!cur) return null;
  }
  return cur ?? null;
}

export function attr(el: Element | null | undefined, name: string): string | null {
  if (!el) return null;
  if (name.includes(':')) {
    const [ns, local] = split(name);
    return el.getAttributeNS(ns, local);
  }
  return el.hasAttribute(name) ? el.getAttribute(name) : null;
}

export function num(el: Element | null | undefined, name: string, fallback: number): number;
export function num(el: Element | null | undefined, name: string): number | null;
export function num(el: Element | null | undefined, name: string, fallback: number | null = null): number | null {
  const v = attr(el, name);
  if (v === null || v === '') return fallback;
  const n = Number(v);
  return Number.isFinite(n) ? n : fallback;
}

/** OOXML booleans: "1", "true", "on". Absent gives the fallback. */
export function bool(el: Element | null | undefined, name: string, fallback: boolean): boolean {
  const v = attr(el, name);
  if (v === null) return fallback;
  return v === '1' || v === 'true' || v === 'on';
}

/** First descendant (depth-first) with the qualified name. */
export function find(el: Element | null | undefined, qname: string): Element | null {
  if (!el) return null;
  const [ns, local] = split(qname);
  const list = el.getElementsByTagNameNS(ns, local);
  return list.length ? list[0] : null;
}

export function findAll(el: Element | Document | null | undefined, qname: string): Element[] {
  if (!el) return [];
  const [ns, local] = split(qname);
  return Array.from(el.getElementsByTagNameNS(ns, local));
}

/** Create a namespaced element in `doc`. */
export function make(doc: Document, qname: string, attrs?: Record<string, string | number | boolean | null | undefined>): Element {
  const [ns] = split(qname);
  const el = doc.createElementNS(ns, qname);
  if (attrs) for (const [k, v] of Object.entries(attrs)) if (v !== null && v !== undefined) setAttr(el, k, String(v === true ? 1 : v === false ? 0 : v));
  return el;
}

export function setAttr(el: Element, name: string, value: string | null) {
  if (name.includes(':') && !name.startsWith('xmlns')) {
    const [ns] = split(name);
    if (value === null) el.removeAttributeNS(ns, name.split(':')[1]);
    else el.setAttributeNS(ns, name, value);
  } else if (value === null) el.removeAttribute(name);
  else el.setAttribute(name, value);
}

/**
 * Insert `child` into `parent` respecting a schema order: it goes before the first existing
 * child whose name appears AFTER it in `order`.
 */
export function insertOrdered(parent: Element, child: Element, order: readonly string[]) {
  const myIndex = order.findIndex((q) => is(child, q));
  for (let c = parent.firstElementChild; c; c = c.nextElementSibling) {
    const idx = order.findIndex((q) => is(c, q));
    if (idx > myIndex && myIndex >= 0) {
      parent.insertBefore(child, c);
      return;
    }
  }
  parent.appendChild(child);
}

export function removeKids(parent: Element, qnames: readonly string[]) {
  for (const c of kids(parent)) if (qnames.some((q) => is(c, q))) parent.removeChild(c);
}

const parser = () => new DOMParser();

/** Parse XML; throws on malformed input (the DOMParser reports errors as a document). */
export function parseXml(text: string): Document {
  const doc = parser().parseFromString(text, 'application/xml');
  if (doc.getElementsByTagName('parsererror').length) throw new Error('malformed XML');
  return doc;
}

export function serialize(doc: Document | Element): string {
  const s = new XMLSerializer().serializeToString(doc);
  return s.startsWith('<?xml') ? s : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${s}`;
}

/** Escape text for inclusion in generated XML strings. */
export function esc(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

export const NS_DECL = `xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}"`;
