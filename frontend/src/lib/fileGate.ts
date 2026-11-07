/**
 * The gate: every file is inspected here before an editor ever sees it.
 *
 * What it enforces:
 *  - size limit, and only .docx / .xlsx / .pptx by name AND by content
 *  - a real ZIP (OOXML) package; legacy binary and password-protected files are refused
 *  - no macros (vbaProject), no ActiveX controls, no macro-enabled or template content types
 *  - zip-bomb limits: entry count, total uncompressed size, per-part compression ratio
 *  - no path tricks in entry names
 *  - no DTD/entity declarations in the package's control XML (XXE / billion-laughs), and no
 *    UTF-16/32 encoded parts that would hide one
 *  - no external relationships to remote templates, linked objects or other workbooks
 *
 * Nothing here executes file content. The file never leaves the browser.
 */
import { unzipSync, strFromU8 } from 'fflate';
import {
  COMPRESSION_RATIO_FLOOR_BYTES,
  KIND_INFO,
  MAX_COMPRESSION_RATIO,
  MAX_FILE_BYTES,
  MAX_UNCOMPRESSED_BYTES,
  MAX_ZIP_ENTRIES,
  type DocKind,
} from './config';

export class GateError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'GateError';
  }
}

export interface GatedFile {
  kind: DocKind;
  name: string;
  bytes: Uint8Array;
  /** Parts the editor cannot keep when it saves (e.g. charts in a workbook). Empty when none. */
  lossyParts: string[];
}

/** Workbook parts ExcelJS/Univer drop on save, by folder. */
const XLSX_LOSSY: [RegExp, string][] = [
  [/^xl\/charts\//i, 'charts'],
  [/^xl\/pivot(Tables|Cache)\//i, 'pivot tables'],
  [/^xl\/slicers?\/|^xl\/slicerCaches\//i, 'slicers'],
  [/^xl\/timelines?\//i, 'timelines'],
  [/^xl\/externalLinks\//i, 'links to other workbooks'],
  [/^xl\/embeddings\//i, 'embedded objects'],
];

const REFUSED_EXTENSIONS: Record<string, string> = {
  '.docm': 'Macro-enabled Word files (.docm) are refused. Save it as .docx in Word first.',
  '.xlsm': 'Macro-enabled Excel files (.xlsm) are refused. Save it as .xlsx in Excel first.',
  '.xlsb': 'Binary Excel files (.xlsb) are not supported. Save it as .xlsx in Excel first.',
  '.doc': 'Old Word files (.doc) are not supported. Save it as .docx in Word first.',
  '.xls': 'Old Excel files (.xls) are not supported. Save it as .xlsx in Excel first.',
  '.dotx': 'Word templates (.dotx) are not supported. Only .docx documents can be opened.',
  '.dotm': 'Macro-enabled templates are refused.',
  '.xltx': 'Excel templates (.xltx) are not supported. Only .xlsx workbooks can be opened.',
  '.xltm': 'Macro-enabled templates are refused.',
  '.pptm': 'Macro-enabled PowerPoint files (.pptm) are refused. Save it as .pptx in PowerPoint first.',
  '.ppsm': 'Macro-enabled slide shows (.ppsm) are refused.',
  '.potm': 'Macro-enabled templates are refused.',
  '.ppam': 'PowerPoint add-ins (.ppam) are refused.',
  '.ppt': 'Old PowerPoint files (.ppt) are not supported. Save it as .pptx in PowerPoint first.',
  '.pps': 'Old PowerPoint slide shows (.pps) are not supported. Save it as .pptx in PowerPoint first.',
  '.pot': 'Old PowerPoint templates (.pot) are not supported.',
  '.potx': 'PowerPoint templates (.potx) are not supported. Only .pptx presentations can be opened.',
  '.ppsx': 'PowerPoint slide shows (.ppsx) are not supported here. Save it as .pptx in PowerPoint first.',
};

const MAIN_PART: Record<DocKind, string> = {
  docx: 'word/document.xml',
  xlsx: 'xl/workbook.xml',
  pptx: 'ppt/presentation.xml',
};

const MAIN_CONTENT_TYPE: Record<DocKind, string> = {
  docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
  xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
  pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml',
};

/** Strip control characters and path parts; keep the name short and printable. */
export function cleanFileName(name: string): string {
  const base = name.split(/[\\/]/).pop() ?? '';
  // eslint-disable-next-line no-control-regex
  const cleaned = base.replace(/[\u0000-\u001f\u007f<>:"|?*]/g, '').trim();
  return (cleaned || 'Untitled').slice(0, 140);
}

export function extensionOf(name: string): string {
  const m = /\.[a-z0-9]+$/i.exec(name);
  return m ? m[0].toLowerCase() : '';
}

function kindFromExtension(ext: string): DocKind | null {
  if (ext === '.docx') return 'docx';
  if (ext === '.xlsx') return 'xlsx';
  if (ext === '.pptx') return 'pptx';
  return null;
}

/** How far into a part we look for the root element. Real Office parts open with a short prolog. */
const PROLOG_SCAN_BYTES = 1024 * 1024;

/**
 * True when the part's XML prolog is acceptable: single-byte-compatible encoding, and nothing
 * but an XML declaration, comments, processing instructions and whitespace before the root
 * element. A DOCTYPE/ENTITY declaration can only live before the root, so anything else in the
 * prolog (including any `<!...>` declaration) is refused. Comments and PIs are skipped properly
 * so a fake `<root>` hidden inside a comment cannot end the scan early, and a DOCTYPE hidden
 * behind a long comment is still found. Real Office files never carry a DTD.
 */
export function isSafeProlog(data: Uint8Array): boolean {
  if (data.length === 0) return true; // nothing to parse, nothing to exploit
  // UTF-16/UTF-32 BOMs: the ASCII-based scan below would see nothing. (EF BB BF = UTF-8, fine.)
  const [b0, b1, b2, b3] = data;
  if (
    (b0 === 0xff && b1 === 0xfe) ||
    (b0 === 0xfe && b1 === 0xff) ||
    (b0 === 0x00 && b1 === 0x00 && b2 === 0xfe && b3 === 0xff)
  ) {
    return false;
  }
  // One byte per char (Latin-1) keeps indexes aligned with bytes and never throws.
  const head = strFromU8(data.subarray(0, PROLOG_SCAN_BYTES), true);
  // BOM-less UTF-16/32 is full of NULs; no legitimate XML part has any.
  if (head.includes('\0')) return false;

  const start = head.charCodeAt(0) === 0xef && head.charCodeAt(1) === 0xbb && head.charCodeAt(2) === 0xbf ? 3 : 0;
  let i = start;
  while (i < head.length) {
    const c = head.charCodeAt(i);
    if (c === 0x20 || c === 0x09 || c === 0x0a || c === 0x0d) {
      i++;
    } else if (head.startsWith('<?', i)) {
      if (i === start) {
        // A declared UTF-16/32/EBCDIC encoding is not something we can scan reliably.
        const decl = /^<\?xml\b[^>]*?\bencoding\s*=\s*["']\s*([^"'\s]*)/i.exec(head.slice(i, i + 512));
        if (decl && /^(utf-?(16|32)|ucs-?[24]|ebcdic|cp0?37|ibm0?37)/i.test(decl[1])) return false;
      }
      const end = head.indexOf('?>', i + 2);
      if (end < 0) return false;
      i = end + 2;
    } else if (head.startsWith('<!--', i)) {
      const end = head.indexOf('-->', i + 4);
      if (end < 0) return false;
      i = end + 3;
    } else if (head.startsWith('<!', i)) {
      return false; // DOCTYPE, ENTITY or any other declaration before the root
    } else if (c === 0x3c && /[A-Za-z_:\u0080-\uffff]/.test(head[i + 1] ?? '')) {
      return true; // root start tag reached: no DTD is possible any more
    } else {
      return false; // not well-formed XML
    }
  }
  // Ran out of bytes (or hit the scan cap) without finding a root element.
  return false;
}

/** Relationship types that, with TargetMode="External", pull in or point at content outside the file. */
const EXTERNAL_REL_REFUSED: [RegExp, string][] = [
  [/\/attachedTemplate$/i, 'a remote template'],
  [/\/oleObject$/i, 'a linked object'],
  // Deliberately NOT listed: externalLink / externalLinkPath. A workbook that links to other
  // workbooks is ordinary (finance models); Abakion cannot keep those links, so XLSX_LOSSY above
  // warns and opens it without a write handle instead of refusing it.
  [/\/subDocument$/i, 'a linked sub-document'],
  [/\/frame$/i, 'a linked frame'],
  [/\/package$/i, 'an embedded package'],
];

/** Decode XML character/entity references in an attribute value, as a real parser would. */
function decodeXmlAttr(v: string): string {
  return v
    .replace(/&#x([0-9a-f]+);/gi, (_, h: string) => String.fromCodePoint(Math.min(parseInt(h, 16), 0x10ffff)))
    .replace(/&#(\d+);/g, (_, d: string) => String.fromCodePoint(Math.min(parseInt(d, 10), 0x10ffff)))
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&amp;/g, '&')
    .trim();
}

/** What an external relationship points at, if it is one we refuse; otherwise null. */
function refusedExternalRel(relsXml: string): string | null {
  // Structured on purpose: every attribute is name="value" or name='value', so a '>' inside a
  // value cannot cut the tag short, and attribute order, quote style and whitespace do not matter.
  const tagRe =
    /<(?:[\w.-]+:)?Relationship\b((?:\s+[^\s=>/]+\s*=\s*(?:"[^"]*"|'[^']*'))*)\s*\/?>/gi;
  const attrRe = /([^\s=>/]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;
  for (const tag of relsXml.matchAll(tagRe)) {
    const attrs = new Map<string, string>();
    for (const a of tag[1].matchAll(attrRe)) attrs.set(a[1], decodeXmlAttr(a[2] ?? a[3] ?? ''));
    if (attrs.get('TargetMode')?.toLowerCase() !== 'external') continue;
    const type = attrs.get('Type') ?? '';
    for (const [re, what] of EXTERNAL_REL_REFUSED) if (re.test(type)) return what;
  }
  return null;
}

function isSuspiciousPath(name: string): boolean {
  return (
    name.startsWith('/') ||
    name.includes('\\') ||
    name.includes('\0') ||
    /(^|\/)\.\.(\/|$)/.test(name) ||
    /^[a-z]:/i.test(name)
  );
}

function isActiveContent(name: string): boolean {
  const lower = name.toLowerCase();
  return (
    lower.endsWith('vbaproject.bin') ||
    lower.endsWith('vbadata.xml') ||
    lower.includes('/activex/') ||
    lower.startsWith('activex/')
  );
}

/**
 * Inspect a file. Resolves with its bytes when it is a clean .docx or .xlsx,
 * rejects with a GateError carrying a message fit to show the person.
 */
export async function gateFile(file: File, expected?: DocKind): Promise<GatedFile> {
  const name = cleanFileName(file.name);
  const ext = extensionOf(name);

  if (REFUSED_EXTENSIONS[ext]) throw new GateError(REFUSED_EXTENSIONS[ext]);

  const extKind = kindFromExtension(ext);
  if (!extKind) {
    throw new GateError('Only Word (.docx), Excel (.xlsx) and PowerPoint (.pptx) files can be opened.');
  }
  if (expected && extKind !== expected) {
    throw new GateError(
      `${KIND_INFO[expected].app} opens ${KIND_INFO[expected].ext} files. ` +
        `Open this ${ext} file from the hub instead.`,
    );
  }
  if (file.size === 0) throw new GateError('This file is empty.');
  if (file.size > MAX_FILE_BYTES) {
    throw new GateError(
      `This file is ${(file.size / 1048576).toFixed(1)} MB. The limit is ${MAX_FILE_BYTES / 1048576} MB.`,
    );
  }

  const bytes = new Uint8Array(await file.arrayBuffer());
  const names = inspectPackage(bytes, extKind);
  const lossyParts =
    extKind === 'xlsx'
      ? [...new Set(XLSX_LOSSY.filter(([re]) => names.some((n) => re.test(n))).map(([, label]) => label))]
      : [];
  return { kind: extKind, name, bytes, lossyParts };
}

/** Structural checks on the package bytes. Throws GateError on any problem; returns the part names. */
export function inspectPackage(bytes: Uint8Array, kind: DocKind): string[] {
  // Magic bytes
  if (bytes[0] === 0xd0 && bytes[1] === 0xcf && bytes[2] === 0x11 && bytes[3] === 0xe0) {
    throw new GateError(
      'This file is password-protected or in an old binary format. Remove the password or save it as ' +
        `${KIND_INFO[kind].ext} in Office first.`,
    );
  }
  if (!(bytes[0] === 0x50 && bytes[1] === 0x4b && bytes[2] === 0x03 && bytes[3] === 0x04)) {
    throw new GateError(`This is not a valid ${KIND_INFO[kind].description}.`);
  }

  const wanted = new Set<string>(['[Content_Types].xml', '_rels/.rels', MAIN_PART[kind]]);
  const names: string[] = [];
  let totalUncompressed = 0;
  let problem: string | null = null;

  let parts: Record<string, Uint8Array>;
  try {
    parts = unzipSync(bytes, {
      filter(info) {
        if (problem) return false;
        names.push(info.name);
        if (names.length > MAX_ZIP_ENTRIES) {
          problem = 'This file contains too many parts to be a normal Office file.';
          return false;
        }
        if (isSuspiciousPath(info.name)) {
          problem = 'This file has a malformed internal structure.';
          return false;
        }
        if (isActiveContent(info.name)) {
          problem = 'This file contains macros or ActiveX controls, which are refused.';
          return false;
        }
        totalUncompressed += info.originalSize;
        if (totalUncompressed > MAX_UNCOMPRESSED_BYTES) {
          problem = 'This file expands to an unsafe size and was refused.';
          return false;
        }
        if (
          info.originalSize > COMPRESSION_RATIO_FLOOR_BYTES &&
          info.size > 0 &&
          info.originalSize / info.size > MAX_COMPRESSION_RATIO
        ) {
          problem = 'This file is compressed suspiciously (a possible zip bomb) and was refused.';
          return false;
        }
        // Decompress only the small control parts and every relationships file.
        return wanted.has(info.name) || info.name.endsWith('.rels');
      },
    });
  } catch {
    throw new GateError(`This ${KIND_INFO[kind].description} is damaged and cannot be opened.`);
  }
  if (problem) throw new GateError(problem);

  const ctBytes = parts['[Content_Types].xml'];
  if (!ctBytes || !parts[MAIN_PART[kind]]) {
    const other = (Object.keys(MAIN_PART) as DocKind[]).find((k) => k !== kind && names.includes(MAIN_PART[k]));
    if (other) {
      throw new GateError(
        `The file name says ${KIND_INFO[kind].ext}, but it contains a ${KIND_INFO[other].description}.`,
      );
    }
    throw new GateError(`This is not a valid ${KIND_INFO[kind].description}.`);
  }

  for (const [partName, data] of Object.entries(parts)) {
    if (!isSafeProlog(data)) {
      throw new GateError(`This file has unsafe XML in ${partName} and was refused.`);
    }
  }

  // Relationships that reach outside the file (remote templates, linked objects, other workbooks).
  // Plain external hyperlinks and images stay allowed: they are inert until clicked or fetched.
  for (const [partName, data] of Object.entries(parts)) {
    if (!partName.endsWith('.rels')) continue;
    const what = refusedExternalRel(strFromU8(data));
    if (what) throw new GateError(`This file links to external content (${what}), which is refused.`);
  }

  const contentTypes = strFromU8(ctBytes);
  if (/macroEnabled/i.test(contentTypes) || /vbaProject/i.test(contentTypes)) {
    throw new GateError('This file is macro-enabled, which is refused.');
  }
  if (/\.template\.main\+xml/i.test(contentTypes)) {
    throw new GateError('Templates are not supported. Save it as a normal document first.');
  }
  if (!contentTypes.includes(MAIN_CONTENT_TYPE[kind])) {
    throw new GateError(`This is not a standard ${KIND_INFO[kind].description}.`);
  }
  return names;
}
