/**
 * The .pptx package (an OPC zip): parts, content types and relationships.
 *
 * The file already passed the gate (lib/fileGate.ts) before it gets here. Each XML part is
 * checked again with the gate's prolog scan before the browser's DOMParser sees it, because the
 * gate only decompressed the control parts and relationships.
 */
import { strFromU8, strToU8, unzipSync, zipSync, type Zippable } from 'fflate';
import { isSafeProlog } from '../../lib/fileGate';
import { NS, attr, kids, parseXml, serialize } from './xml';

export interface Rel {
  id: string;
  type: string;
  /** Resolved part name (no leading slash) for internal targets; the raw target for external ones. */
  target: string;
  external: boolean;
}

export const REL = {
  officeDocument: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
  slide: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide',
  slideLayout: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout',
  slideMaster: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster',
  theme: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
  notesSlide: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide',
  notesMaster: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster',
  image: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
  hyperlink: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
  tableStyles: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles',
} as const;

export const CT = {
  slide: 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml',
  notesSlide: 'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml',
  rels: 'application/vnd.openxmlformats-package.relationships+xml',
} as const;

/** Directory of a part name ('ppt/slides/slide1.xml' -> 'ppt/slides'). */
export function dirOf(name: string): string {
  const i = name.lastIndexOf('/');
  return i < 0 ? '' : name.slice(0, i);
}

export function relsPathOf(name: string): string {
  const d = dirOf(name);
  const base = name.slice(d.length ? d.length + 1 : 0);
  return `${d ? d + '/' : ''}_rels/${base}.rels`;
}

/** Resolve a relationship target against the part it belongs to. */
export function resolveTarget(fromPart: string, target: string): string {
  if (target.startsWith('/')) return normalize(target.slice(1));
  return normalize(`${dirOf(fromPart)}/${target}`);
}

function normalize(p: string): string {
  const out: string[] = [];
  for (const seg of p.split('/')) {
    if (!seg || seg === '.') continue;
    if (seg === '..') out.pop();
    else out.push(seg);
  }
  return out.join('/');
}

/** Relative target from one part to another, as PowerPoint writes it. */
export function relativeTarget(fromPart: string, toPart: string): string {
  const from = dirOf(fromPart).split('/').filter(Boolean);
  const to = toPart.split('/');
  let i = 0;
  while (i < from.length && i < to.length - 1 && from[i] === to[i]) i++;
  return [...from.slice(i).map(() => '..'), ...to.slice(i)].join('/');
}

export class Pkg {
  private readonly xmlCache = new Map<string, Document>();
  private readonly relsCache = new Map<string, Rel[]>();
  readonly defaults = new Map<string, string>();
  readonly overrides = new Map<string, string>();

  private constructor(readonly parts: Map<string, Uint8Array>) {
    const ct = this.xml('[Content_Types].xml');
    for (const el of kids(ct?.documentElement)) {
      if (el.localName === 'Default') this.defaults.set((attr(el, 'Extension') ?? '').toLowerCase(), attr(el, 'ContentType') ?? '');
      if (el.localName === 'Override') this.overrides.set((attr(el, 'PartName') ?? '').replace(/^\//, ''), attr(el, 'ContentType') ?? '');
    }
  }

  static open(bytes: Uint8Array): Pkg {
    const files = unzipSync(bytes);
    const parts = new Map<string, Uint8Array>();
    for (const [name, data] of Object.entries(files)) if (!name.endsWith('/')) parts.set(name, data);
    return new Pkg(parts);
  }

  has(name: string): boolean {
    return this.parts.has(name);
  }

  bytes(name: string): Uint8Array | null {
    return this.parts.get(name) ?? null;
  }

  /** Parsed XML of a part (cached; treat as read-only). Null when absent, unsafe or malformed. */
  xml(name: string): Document | null {
    const cached = this.xmlCache.get(name);
    if (cached) return cached;
    const data = this.parts.get(name);
    if (!data || !isSafeProlog(data)) return null;
    try {
      const doc = parseXml(strFromU8(data));
      this.xmlCache.set(name, doc);
      return doc;
    } catch {
      return null;
    }
  }

  rels(name: string): Rel[] {
    const cached = this.relsCache.get(name);
    if (cached) return cached;
    const doc = this.xml(relsPathOf(name));
    const out: Rel[] = [];
    for (const el of kids(doc?.documentElement)) {
      if (el.namespaceURI !== NS.rel || el.localName !== 'Relationship') continue;
      const external = (attr(el, 'TargetMode') ?? '').toLowerCase() === 'external';
      const raw = attr(el, 'Target') ?? '';
      out.push({
        id: attr(el, 'Id') ?? '',
        type: attr(el, 'Type') ?? '',
        external,
        target: external ? raw : resolveTarget(name, raw),
      });
    }
    this.relsCache.set(name, out);
    return out;
  }

  rel(name: string, id: string | null): Rel | null {
    if (!id) return null;
    return this.rels(name).find((r) => r.id === id) ?? null;
  }

  relOfType(name: string, type: string): Rel | null {
    return this.rels(name).find((r) => r.type === type) ?? null;
  }

  contentType(name: string): string {
    return this.overrides.get(name) ?? this.defaults.get((name.split('.').pop() ?? '').toLowerCase()) ?? '';
  }
}

/** A writable copy of a package's parts, for building the saved file. */
export class PkgWriter {
  readonly parts: Map<string, Uint8Array>;
  readonly defaults: Map<string, string>;
  readonly overrides: Map<string, string>;
  /** Relationships per source part, editable. */
  private readonly rels = new Map<string, { id: string; type: string; target: string; external: boolean }[]>();

  constructor(src: Pkg) {
    this.parts = new Map(src.parts);
    this.defaults = new Map(src.defaults);
    this.overrides = new Map(src.overrides);
  }

  /** Relationships of a part (loaded from the source on first use). */
  relsOf(pkg: Pkg, name: string) {
    let list = this.rels.get(name);
    if (!list) {
      list = pkg.rels(name).map((r) => ({ ...r }));
      this.rels.set(name, list);
    }
    return list;
  }

  setRels(name: string, list: { id: string; type: string; target: string; external: boolean }[]) {
    this.rels.set(name, list);
  }

  /** Add (or reuse) a relationship from `from` to `target`; returns its id. */
  addRel(pkg: Pkg, from: string, type: string, target: string, external = false): string {
    const list = this.relsOf(pkg, from);
    const existing = list.find((r) => r.type === type && r.target === target && r.external === external);
    if (existing) return existing.id;
    let n = list.length + 1;
    const ids = new Set(list.map((r) => r.id));
    while (ids.has(`rId${n}`)) n++;
    const id = `rId${n}`;
    list.push({ id, type, target, external });
    return id;
  }

  setXml(name: string, doc: Document | string) {
    this.parts.set(name, strToU8(typeof doc === 'string' ? doc : serialize(doc)));
  }

  /** A fresh part name like 'ppt/slides/slide7.xml' that is not taken. */
  freshName(dir: string, stem: string, ext: string): string {
    for (let i = 1; ; i++) {
      const name = `${dir}/${stem}${i}.${ext}`;
      if (!this.parts.has(name)) return name;
    }
  }

  /**
   * Write relationship parts, drop every part nothing reaches any more (deleted slides, their
   * notes, orphaned media), keep [Content_Types].xml in step, and zip.
   */
  finish(): Uint8Array {
    for (const [name, list] of this.rels) {
      const body = list
        .map(
          (r) =>
            `<Relationship Id="${xmlAttr(r.id)}" Type="${xmlAttr(r.type)}" Target="${xmlAttr(
              r.external ? r.target : relativeTarget(name, r.target),
            )}"${r.external ? ' TargetMode="External"' : ''}/>`,
        )
        .join('');
      this.parts.set(
        relsPathOf(name),
        strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${NS.rel}">${body}</Relationships>`),
      );
    }

    // Reachability from the package root.
    const reachable = new Set<string>(['[Content_Types].xml', '_rels/.rels']);
    const queue = [''];
    const seen = new Set<string>();
    while (queue.length) {
      const part = queue.pop()!;
      if (seen.has(part)) continue;
      seen.add(part);
      const relsName = part ? relsPathOf(part) : '_rels/.rels';
      const data = this.parts.get(relsName);
      if (!data) continue;
      reachable.add(relsName);
      let doc: Document;
      try {
        doc = parseXml(strFromU8(data));
      } catch {
        continue;
      }
      for (const el of kids(doc.documentElement)) {
        if ((attr(el, 'TargetMode') ?? '').toLowerCase() === 'external') continue;
        const target = resolveTarget(part || '_', attr(el, 'Target') ?? '');
        const resolved = part ? target : normalize(attr(el, 'Target') ?? '');
        if (this.parts.has(resolved)) {
          reachable.add(resolved);
          queue.push(resolved);
        }
      }
    }
    for (const name of [...this.parts.keys()]) if (!reachable.has(name)) this.parts.delete(name);
    for (const name of [...this.overrides.keys()]) if (!this.parts.has(name)) this.overrides.delete(name);

    // Every extension in use needs a Default (or every such part an Override).
    for (const name of this.parts.keys()) {
      const ext = (name.split('.').pop() ?? '').toLowerCase();
      if (this.overrides.has(name) || this.defaults.has(ext)) continue;
      const mime = MEDIA_TYPES[ext];
      if (mime) this.defaults.set(ext, mime);
    }
    const ct =
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="${NS.ct}">` +
      [...this.defaults].map(([e, t]) => `<Default Extension="${xmlAttr(e)}" ContentType="${xmlAttr(t)}"/>`).join('') +
      [...this.overrides].map(([p, t]) => `<Override PartName="/${xmlAttr(p)}" ContentType="${xmlAttr(t)}"/>`).join('') +
      '</Types>';
    this.parts.set('[Content_Types].xml', strToU8(ct));

    const files: Zippable = {};
    // [Content_Types].xml first, as Office writes it.
    files['[Content_Types].xml'] = [this.parts.get('[Content_Types].xml')!, { level: 6 }];
    for (const [name, data] of this.parts) {
      if (name === '[Content_Types].xml') continue;
      const stored = /\.(png|jpe?g|gif|mp4|m4a|mp3|wav|wmv|zip)$/i.test(name);
      files[name] = [data, { level: stored ? 0 : 6 }];
    }
    return zipSync(files);
  }
}

export const MEDIA_TYPES: Record<string, string> = {
  rels: CT.rels,
  xml: 'application/xml',
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  bmp: 'image/bmp',
  svg: 'image/svg+xml',
  tif: 'image/tiff',
  tiff: 'image/tiff',
  emf: 'image/x-emf',
  wmf: 'image/x-wmf',
};

function xmlAttr(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;');
}
