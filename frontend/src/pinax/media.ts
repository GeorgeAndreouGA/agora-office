/**
 * Pictures inside the presentation, served to <img> as object URLs made from the package's own
 * bytes (nothing is fetched). Formats a browser cannot show (EMF, WMF, TIFF) report null so the
 * slide can draw a labelled placeholder instead.
 */
import type { Pkg } from './pptx/package';

const SHOWABLE: Record<string, string> = {
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  jpe: 'image/jpeg',
  gif: 'image/gif',
  bmp: 'image/bmp',
  webp: 'image/webp',
  svg: 'image/svg+xml',
};

export class MediaStore {
  private readonly urls = new Map<string, string | null>();
  /** Bytes of pictures added in Pinax (not yet in the package). */
  private readonly added = new Map<string, Uint8Array>();

  constructor(private readonly pkg: Pkg) {}

  mimeOf(part: string): string | null {
    const ext = (part.split('.').pop() ?? '').toLowerCase();
    return SHOWABLE[ext] ?? null;
  }

  bytes(part: string): Uint8Array | null {
    return this.added.get(part) ?? this.pkg.bytes(part);
  }

  /** Register a picture inserted in Pinax under a provisional part name. */
  add(part: string, bytes: Uint8Array) {
    this.added.set(part, bytes);
  }

  addedParts(): Map<string, Uint8Array> {
    return this.added;
  }

  url(part: string | null): string | null {
    if (!part) return null;
    if (this.urls.has(part)) return this.urls.get(part)!;
    const mime = this.mimeOf(part);
    const data = mime ? this.bytes(part) : null;
    const url = data && mime ? URL.createObjectURL(new Blob([data as BlobPart], { type: mime })) : null;
    this.urls.set(part, url);
    return url;
  }

  dispose() {
    for (const u of this.urls.values()) if (u) URL.revokeObjectURL(u);
    this.urls.clear();
  }
}
