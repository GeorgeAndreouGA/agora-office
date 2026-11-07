/**
 * Fonts for Pinax: the same bundled, metric-compatible faces Papyros uses (Carlito for Calibri,
 * Liberation for Arial/Times/Courier...), registered under private "Pinax ..." names.
 *
 * Every family a presentation names maps to one of those faces, and the screen never falls back
 * to whatever the visitor has installed. That keeps a deck looking the same on every machine and
 * lets the PDF exporter embed exactly the bytes the screen drew with.
 */
import { packagedFonts } from '@docx-editor.dev/fonts';

/** The Word/Office families the bundled faces stand in for. */
type Target = 'Calibri' | 'Cambria' | 'Arial' | 'Times New Roman' | 'Courier New' | 'Century Gothic';

const ALIASES: Record<string, Target> = {
  calibri: 'Calibri',
  'calibri light': 'Calibri',
  carlito: 'Calibri',
  aptos: 'Calibri',
  'aptos display': 'Calibri',
  'aptos narrow': 'Calibri',
  'segoe ui': 'Calibri',
  'segoe ui light': 'Calibri',
  'segoe ui semibold': 'Calibri',
  candara: 'Calibri',
  corbel: 'Calibri',
  cambria: 'Cambria',
  caladea: 'Cambria',
  constantia: 'Cambria',
  arial: 'Arial',
  'arial black': 'Arial',
  'arial narrow': 'Arial',
  helvetica: 'Arial',
  'helvetica neue': 'Arial',
  'liberation sans': 'Arial',
  tahoma: 'Arial',
  verdana: 'Arial',
  'trebuchet ms': 'Arial',
  'gill sans mt': 'Arial',
  'franklin gothic book': 'Arial',
  'franklin gothic medium': 'Arial',
  'lucida sans': 'Arial',
  'times new roman': 'Times New Roman',
  times: 'Times New Roman',
  'liberation serif': 'Times New Roman',
  georgia: 'Times New Roman',
  garamond: 'Times New Roman',
  'book antiqua': 'Times New Roman',
  'palatino linotype': 'Times New Roman',
  'bookman old style': 'Times New Roman',
  'courier new': 'Courier New',
  courier: 'Courier New',
  'liberation mono': 'Courier New',
  consolas: 'Courier New',
  'lucida console': 'Courier New',
  'century gothic': 'Century Gothic',
  'tex gyre adventor': 'Century Gothic',
  futura: 'Century Gothic',
};

function targetOf(family: string): Target {
  const f = family.trim().toLowerCase();
  if (ALIASES[f]) return ALIASES[f];
  if (/mono|code|courier/.test(f)) return 'Courier New';
  if (/serif/.test(f) && !/sans/.test(f)) return 'Times New Roman';
  if (/roman|garamond|baskerville|didot|bodoni|minion|caslon/.test(f)) return 'Times New Roman';
  return 'Calibri';
}

export interface FaceBytes {
  bytes: Uint8Array;
  key: string;
}

export interface FontRegistry {
  /** CSS font-family list for a family named in the file. */
  css(family: string): string;
  /** The font file the screen uses for (family, bold, italic), for PDF export. */
  face(cssFamily: string, bold: boolean, italic: boolean): FaceBytes | null;
  dispose(): void;
}

const FALLBACK = '"GFS Didot", sans-serif';

/** Load (from the app's own bundle) and register the faces a deck needs. */
export async function loadFonts(families: Iterable<string>): Promise<FontRegistry> {
  const targets = new Set<Target>(['Calibri']);
  for (const f of families) if (f) targets.add(targetOf(f));
  const fragment = await packagedFonts({ onFailure: () => undefined })({
    families: [...targets],
    defaultFamily: 'Calibri',
  });

  // substitute family (e.g. "Carlito") for each target
  const substituteOf = new Map<string, string>();
  for (const s of fragment.substitutions) substituteOf.set(s.from.family.toLowerCase(), s.to.family);

  const faces = new Map<string, FaceBytes>(); // "Pinax Carlito|700|italic" -> bytes
  const registered: FontFace[] = [];
  for (const src of fragment.sources) {
    const name = `Pinax ${src.request.family}`;
    const weight = src.request.weight >= 600 ? 700 : 400;
    const style = src.request.style;
    faces.set(`${name}|${weight}|${style}`, { bytes: src.bytes, key: src.hash || `${name}|${weight}|${style}` });
    try {
      const face = new FontFace(name, src.bytes.slice().buffer, { weight: String(weight), style });
      document.fonts.add(face);
      registered.push(face);
      void face.load().catch(() => undefined);
    } catch {
      /* a face the browser rejects just falls back */
    }
  }
  await Promise.all(registered.map((f) => f.loaded.catch(() => undefined)));

  const cssCache = new Map<string, string>();
  return {
    css(family: string) {
      let v = cssCache.get(family);
      if (!v) {
        const target = targetOf(family || 'Calibri');
        const sub = substituteOf.get(target.toLowerCase()) ?? 'Carlito';
        v = `"Pinax ${sub}", ${FALLBACK}`;
        cssCache.set(family, v);
      }
      return v;
    },
    face(cssFamily: string, bold: boolean, italic: boolean) {
      const first = cssFamily.split(',')[0].trim().replace(/^["']|["']$/g, '');
      const weight = bold ? 700 : 400;
      const style = italic ? 'italic' : 'normal';
      return (
        faces.get(`${first}|${weight}|${style}`) ??
        faces.get(`${first}|${weight}|normal`) ??
        faces.get(`${first}|400|${style}`) ??
        faces.get(`${first}|400|normal`) ??
        null
      );
    },
    dispose() {
      for (const f of registered) document.fonts.delete(f);
    },
  };
}
