/** Central limits and timings. Change them here, not across the code. */

/** Largest file the gate will open. */
export const MAX_FILE_BYTES = 50 * 1024 * 1024;

/** Largest total uncompressed size the gate will accept inside one package. */
export const MAX_UNCOMPRESSED_BYTES = 1024 * 1024 * 1024;

/** Most entries a package may contain. Real Office files have dozens to a few thousand. */
export const MAX_ZIP_ENTRIES = 20_000;

/** A part this many times bigger than its compressed form (and over the floor below) is a zip bomb. */
export const MAX_COMPRESSION_RATIO = 1000;
export const COMPRESSION_RATIO_FLOOR_BYTES = 10 * 1024 * 1024;

/** Autosave interval. */
export const AUTOSAVE_MS = 5 * 60 * 1000;

export type DocKind = 'docx' | 'xlsx' | 'pptx';

export const KIND_INFO: Record<
  DocKind,
  { app: string; greek: string; ext: string; mime: string; description: string }
> = {
  docx: {
    app: 'Papyros',
    greek: 'Πάπυρος',
    ext: '.docx',
    mime: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    description: 'Word document',
  },
  xlsx: {
    app: 'Abakion',
    greek: 'Ἀβάκιον',
    ext: '.xlsx',
    mime: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    description: 'Excel workbook',
  },
  pptx: {
    app: 'Pinax',
    greek: 'Πίναξ',
    ext: '.pptx',
    mime: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    description: 'PowerPoint presentation',
  },
};
