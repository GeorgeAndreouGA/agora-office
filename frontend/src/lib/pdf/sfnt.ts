/**
 * Minimal TrueType handling for PDF embedding: WOFF -> sfnt, and glyph subsetting.
 *
 * The subsetter keeps every glyph SLOT and empties the ones not drawn, so glyph IDs never change
 * and the PDF can address glyphs by their original ID (CIDToGIDMap Identity). Glyph outlines are
 * copied byte for byte, plus the components of any composite glyph. Nothing is re-encoded, which
 * is what makes it safe: fontkit's own subsetter corrupted glyphs of some fonts (Carlito).
 */
import { unzlibSync } from 'fflate';

const tagOf = (b: Uint8Array, o: number) => String.fromCharCode(b[o], b[o + 1], b[o + 2], b[o + 3]);

interface TableRecord {
  tag: string;
  offset: number;
  length: number;
}

function view(b: Uint8Array): DataView {
  return new DataView(b.buffer, b.byteOffset, b.byteLength);
}

/** WOFF 1.0 (zlib per table) to a plain sfnt. Returns the input unchanged when it is not WOFF. */
export function woffToSfnt(bytes: Uint8Array): Uint8Array {
  if (bytes.length < 44 || tagOf(bytes, 0) !== 'wOFF') return bytes;
  const dv = view(bytes);
  const flavor = dv.getUint32(4);
  const numTables = dv.getUint16(12);
  const tables: { tag: string; data: Uint8Array; checksum: number }[] = [];
  for (let i = 0; i < numTables; i++) {
    const p = 44 + i * 20;
    const tag = tagOf(bytes, p);
    const offset = dv.getUint32(p + 4);
    const compLength = dv.getUint32(p + 8);
    const origLength = dv.getUint32(p + 12);
    const checksum = dv.getUint32(p + 16);
    const raw = bytes.subarray(offset, offset + compLength);
    const data = compLength < origLength ? unzlibSync(raw) : raw.slice();
    tables.push({ tag, data, checksum });
  }
  return writeSfnt(flavor, tables);
}

function readTables(bytes: Uint8Array): Map<string, TableRecord> {
  const dv = view(bytes);
  const n = dv.getUint16(4);
  const out = new Map<string, TableRecord>();
  for (let i = 0; i < n; i++) {
    const p = 12 + i * 16;
    out.set(tagOf(bytes, p), { tag: tagOf(bytes, p), offset: dv.getUint32(p + 8), length: dv.getUint32(p + 12) });
  }
  return out;
}

function checksum(data: Uint8Array): number {
  let sum = 0;
  const padded = data.length % 4 ? new Uint8Array(data.length + (4 - (data.length % 4))) : data;
  if (padded !== data) padded.set(data);
  const dv = view(padded);
  for (let i = 0; i < padded.length; i += 4) sum = (sum + dv.getUint32(i)) >>> 0;
  return sum;
}

function writeSfnt(flavor: number, tables: { tag: string; data: Uint8Array; checksum?: number }[]): Uint8Array {
  const sorted = [...tables].sort((a, b) => (a.tag < b.tag ? -1 : a.tag > b.tag ? 1 : 0));
  const n = sorted.length;
  let pow = 1;
  let log = 0;
  while (pow * 2 <= n) {
    pow *= 2;
    log++;
  }
  const headerLen = 12 + n * 16;
  let total = headerLen;
  for (const t of sorted) total += (t.data.length + 3) & ~3;
  const out = new Uint8Array(total);
  const dv = view(out);
  dv.setUint32(0, flavor);
  dv.setUint16(4, n);
  dv.setUint16(6, pow * 16);
  dv.setUint16(8, log);
  dv.setUint16(10, n * 16 - pow * 16);
  let offset = headerLen;
  let headOffset = -1;
  sorted.forEach((t, i) => {
    const p = 12 + i * 16;
    for (let k = 0; k < 4; k++) out[p + k] = t.tag.charCodeAt(k);
    dv.setUint32(p + 4, checksum(t.data));
    dv.setUint32(p + 8, offset);
    dv.setUint32(p + 12, t.data.length);
    out.set(t.data, offset);
    if (t.tag === 'head') headOffset = offset;
    offset += (t.data.length + 3) & ~3;
  });
  if (headOffset >= 0) {
    dv.setUint32(headOffset + 8, 0);
    dv.setUint32(headOffset + 8, (0xb1b0afba - checksum(out)) >>> 0);
  }
  return out;
}

/** True when the font has TrueType outlines (a glyf table) and can be subset here. */
export function hasGlyfOutlines(bytes: Uint8Array): boolean {
  try {
    return readTables(bytes).has('glyf');
  } catch {
    return false;
  }
}

/**
 * Subset a TrueType font to the given glyph IDs (plus .notdef and composite components).
 * Returns null when the font is not TrueType-outlined or its tables are malformed; the caller
 * then embeds the whole font instead.
 */
export function subsetTrueType(bytes: Uint8Array, used: Iterable<number>): Uint8Array | null {
  try {
    const tables = readTables(bytes);
    const head = tables.get('head');
    const maxp = tables.get('maxp');
    const loca = tables.get('loca');
    const glyf = tables.get('glyf');
    if (!head || !maxp || !loca || !glyf) return null;
    const dv = view(bytes);
    const numGlyphs = dv.getUint16(maxp.offset + 4);
    const longLoca = dv.getInt16(head.offset + 50) === 1;
    const offsets = new Array<number>(numGlyphs + 1);
    for (let i = 0; i <= numGlyphs; i++) {
      offsets[i] = longLoca ? dv.getUint32(loca.offset + i * 4) : dv.getUint16(loca.offset + i * 2) * 2;
    }
    const glyphBytes = (gid: number) =>
      gid >= 0 && gid < numGlyphs && offsets[gid + 1] > offsets[gid]
        ? bytes.subarray(glyf.offset + offsets[gid], glyf.offset + offsets[gid + 1])
        : null;

    // Closure over composite components.
    const keep = new Set<number>([0]);
    const queue = [...used].filter((g) => g >= 0 && g < numGlyphs);
    while (queue.length) {
      const gid = queue.pop()!;
      if (keep.has(gid) && gid !== 0) continue;
      keep.add(gid);
      const g = glyphBytes(gid);
      if (!g || g.length < 10) continue;
      const gv = view(g);
      if (gv.getInt16(0) >= 0) continue; // simple glyph
      let p = 10;
      for (let guard = 0; guard < 256 && p + 4 <= g.length; guard++) {
        const flags = gv.getUint16(p);
        const comp = gv.getUint16(p + 2);
        if (!keep.has(comp)) queue.push(comp);
        p += 4 + (flags & 0x0001 ? 4 : 2);
        if (flags & 0x0008) p += 2;
        else if (flags & 0x0040) p += 4;
        else if (flags & 0x0080) p += 8;
        if (!(flags & 0x0020)) break;
      }
    }

    let glyfLen = 0;
    for (const gid of keep) {
      const g = glyphBytes(gid);
      if (g) glyfLen += (g.length + 3) & ~3;
    }
    const newGlyf = new Uint8Array(glyfLen);
    const newLoca = new Uint8Array((numGlyphs + 1) * 4);
    const lv = view(newLoca);
    let o = 0;
    for (let gid = 0; gid < numGlyphs; gid++) {
      lv.setUint32(gid * 4, o);
      if (!keep.has(gid)) continue;
      const g = glyphBytes(gid);
      if (!g) continue;
      newGlyf.set(g, o);
      o += (g.length + 3) & ~3;
    }
    lv.setUint32(numGlyphs * 4, o);

    const newHead = bytes.slice(head.offset, head.offset + head.length);
    view(newHead).setInt16(50, 1); // long loca
    const copy = (tag: string) => {
      const t = tables.get(tag);
      return t ? bytes.slice(t.offset, t.offset + t.length) : null;
    };
    const out: { tag: string; data: Uint8Array }[] = [
      { tag: 'head', data: newHead },
      { tag: 'loca', data: newLoca },
      { tag: 'glyf', data: newGlyf },
    ];
    for (const tag of ['hhea', 'maxp', 'hmtx', 'cvt ', 'fpgm', 'prep', 'OS/2']) {
      const d = copy(tag);
      if (d) out.push({ tag, data: d });
    }
    // post format 3 (no glyph names): the original header with the version changed.
    const post = copy('post');
    if (post && post.length >= 32) {
      const p3 = post.slice(0, 32);
      view(p3).setUint32(0, 0x00030000);
      out.push({ tag: 'post', data: p3 });
    }
    return writeSfnt(0x00010000, out);
  } catch {
    return null;
  }
}
