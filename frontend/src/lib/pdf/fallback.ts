/**
 * Fallback faces for characters a document's own fonts cannot draw (polytonic Greek, mostly):
 * GFS Didot, already bundled for the app's own lettering. Loaded only when needed.
 */
import greekFallbackUrl from '@fontsource/gfs-didot/files/gfs-didot-greek-400-normal.woff?url';
import greekExtFallbackUrl from '@fontsource/gfs-didot/files/gfs-didot-greek-ext-400-normal.woff?url';
import latinFallbackUrl from '@fontsource/gfs-didot/files/gfs-didot-latin-400-normal.woff?url';
import type { PdfBuilder, PdfFont } from './writer';

export const FALLBACK_FONT_URLS = [latinFallbackUrl, greekFallbackUrl, greekExtFallbackUrl];

const cache = new WeakMap<PdfBuilder, Promise<PdfFont[]>>();

export function fallbackFonts(builder: PdfBuilder): Promise<PdfFont[]> {
  let p = cache.get(builder);
  if (!p) {
    p = Promise.all(
      FALLBACK_FONT_URLS.map((url) =>
        builder.font(`fallback:${url}`, async () => {
          const res = await fetch(url);
          return res.ok ? new Uint8Array(await res.arrayBuffer()) : null;
        }),
      ),
    ).then((list) => list.filter((f): f is PdfFont => !!f));
    cache.set(builder, p);
  }
  return p;
}

/** The first fallback face that has the character, or null. */
export function fallbackFor(fonts: PdfFont[], ch: string): PdfFont | null {
  const cp = ch.codePointAt(0) ?? 0;
  return fonts.find((f) => f.font.hasGlyphForCodePoint(cp)) ?? null;
}
