/**
 * Talking to the person's own disk. Two paths:
 *  - Chrome / Edge: File System Access API. Open keeps a handle, so Save and the
 *    5-minute autosave write straight back to the original file.
 *  - Other browsers: a normal file input to open, a download to save.
 *
 * Nothing is ever sent over the network and nothing is put in browser storage.
 */
import { KIND_INFO, type DocKind } from './config';
import { extensionOf } from './fileGate';

export const canWriteInPlace =
  typeof window !== 'undefined' && typeof window.showOpenFilePicker === 'function';

export interface PickedFile {
  file: File;
  handle: FileSystemFileHandle | null;
}

function pickerTypes(kinds: DocKind[]): FilePickerAcceptType[] {
  return kinds.map((k) => ({
    description: `${KIND_INFO[k].description} (${KIND_INFO[k].ext})`,
    accept: { [KIND_INFO[k].mime]: [KIND_INFO[k].ext] },
  }));
}

function isAbort(err: unknown): boolean {
  // Only a real cancel. NotAllowedError is a permission or security failure the person should see.
  return err instanceof DOMException && err.name === 'AbortError';
}

/** Ask the person for a file. Resolves null when they cancel. */
export async function pickFile(kinds: DocKind[]): Promise<PickedFile | null> {
  if (canWriteInPlace && window.showOpenFilePicker) {
    try {
      const [handle] = await window.showOpenFilePicker({
        id: 'greek-open',
        multiple: false,
        excludeAcceptAllOption: true,
        types: pickerTypes(kinds),
      });
      return { file: await handle.getFile(), handle };
    } catch (err) {
      if (isAbort(err)) return null;
      throw err;
    }
  }
  return pickWithInput(kinds);
}

function pickWithInput(kinds: DocKind[]): Promise<PickedFile | null> {
  return new Promise((resolve) => {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = kinds.map((k) => `${KIND_INFO[k].ext},${KIND_INFO[k].mime}`).join(',');
    input.style.display = 'none';
    let settled = false;
    let timer = 0;
    const done = (value: PickedFile | null) => {
      if (settled) return;
      settled = true;
      window.clearTimeout(timer);
      window.removeEventListener('focus', onFocus);
      input.removeEventListener('change', onChange);
      input.removeEventListener('cancel', onCancel);
      input.remove();
      resolve(value);
    };
    const onChange = () => {
      const file = input.files?.[0];
      done(file ? { file, handle: null } : null);
    };
    const onCancel = () => done(null);
    // Older browsers never fire 'cancel'. The window regains focus when the picker closes, but
    // that can come just before 'change', so wait a moment and only give up if nothing arrived.
    const onFocus = () => {
      window.clearTimeout(timer);
      timer = window.setTimeout(() => {
        if (!settled && !input.files?.length) done(null);
      }, 1000);
    };
    input.addEventListener('change', onChange);
    input.addEventListener('cancel', onCancel);
    document.body.appendChild(input);
    // Listen only once the picker is about to open, so the click itself cannot trigger it.
    window.addEventListener('focus', onFocus, { once: true });
    input.click();
  });
}

/** Turn a drop into a file (and a writable handle where the browser offers one). */
export async function fileFromDrop(dt: DataTransfer): Promise<PickedFile | null> {
  const item = Array.from(dt.items).find((i) => i.kind === 'file');
  if (!item) return null;
  // getAsFileSystemHandle must be called synchronously within the drop event.
  const handlePromise = item.getAsFileSystemHandle ? item.getAsFileSystemHandle() : null;
  const file = item.getAsFile();
  if (!file) return null;
  let handle: FileSystemFileHandle | null = null;
  if (handlePromise) {
    try {
      const h = await handlePromise;
      if (h && h.kind === 'file') handle = h as FileSystemFileHandle;
    } catch {
      handle = null;
    }
  }
  return { file, handle };
}

export type PermissionResult = 'granted' | 'prompt' | 'denied';

export async function writePermission(handle: FileSystemFileHandle): Promise<PermissionResult> {
  if (!handle.queryPermission) return 'granted';
  return (await handle.queryPermission({ mode: 'readwrite' })) as PermissionResult;
}

/** Must run inside a click or key press: browsers only show the prompt after a gesture. */
export async function requestWritePermission(handle: FileSystemFileHandle): Promise<boolean> {
  if (!handle.requestPermission) return true;
  return (await handle.requestPermission({ mode: 'readwrite' })) === 'granted';
}

/** Write bytes to the handle. The browser writes to a temp file and swaps it in on close. */
export async function writeToHandle(handle: FileSystemFileHandle, bytes: Uint8Array, kind: DocKind) {
  await writeBlob(handle, bytes, KIND_INFO[kind].mime);
}

async function writeBlob(handle: FileSystemFileHandle, bytes: Uint8Array, mime: string) {
  const writable = await handle.createWritable();
  try {
    await writable.write(new Blob([bytes as BlobPart], { type: mime }));
    await writable.close();
  } catch (err) {
    await writable.abort().catch(() => undefined);
    throw err;
  }
}

/** Where a generated PDF will go, chosen BEFORE it is built. */
export interface PdfTarget {
  name: string;
  write(bytes: Uint8Array): Promise<void>;
  /** The PDF could not be built: remove the empty file the save dialog created. */
  discard(): Promise<void>;
}

/**
 * Ask where to save a PDF. Call this straight from a click: the save dialog needs the click's
 * user activation, which building a long PDF would outlast. Chrome / Edge get the save dialog;
 * other browsers download the file once it is written. Resolves null when cancelled.
 */
export async function choosePdfTarget(baseName: string): Promise<PdfTarget | null> {
  const name = `${baseName.replace(/\.(docx|xlsx|pptx|pdf)$/i, '').trim() || 'Document'}.pdf`;
  if (canWriteInPlace && window.showSaveFilePicker) {
    try {
      const handle = await window.showSaveFilePicker({
        id: 'greek-pdf',
        suggestedName: name,
        excludeAcceptAllOption: true,
        types: [{ description: 'PDF document (.pdf)', accept: { 'application/pdf': ['.pdf'] } }],
      });
      return {
        name: handle.name,
        write: (bytes) => writeBlob(handle, bytes, 'application/pdf'),
        discard: async () => {
          const removable = handle as FileSystemFileHandle & { remove?: () => Promise<void> };
          await removable.remove?.().catch(() => undefined);
        },
      };
    } catch (err) {
      if (isAbort(err)) return null;
      throw err;
    }
  }
  return {
    name,
    write: async (bytes) => downloadBytes(bytes, name, 'application/pdf'),
    discard: async () => undefined,
  };
}

export function ensureExtension(name: string, kind: DocKind): string {
  const ext = KIND_INFO[kind].ext;
  return extensionOf(name) === ext ? name : `${name.replace(/\.[a-z0-9]+$/i, '')}${ext}`;
}

/**
 * Save As. Chrome / Edge: a save dialog that returns a new handle.
 * Elsewhere: a download (returns null handle). Resolves undefined when cancelled.
 */
export async function saveAs(
  bytes: Uint8Array,
  suggestedName: string,
  kind: DocKind,
): Promise<{ handle: FileSystemFileHandle | null; name: string } | undefined> {
  const name = ensureExtension(suggestedName, kind);
  if (canWriteInPlace && window.showSaveFilePicker) {
    try {
      const handle = await window.showSaveFilePicker({
        id: 'greek-save',
        suggestedName: name,
        excludeAcceptAllOption: true,
        types: pickerTypes([kind]),
      });
      await writeToHandle(handle, bytes, kind);
      return { handle, name: handle.name };
    } catch (err) {
      if (isAbort(err)) return undefined;
      throw err;
    }
  }
  downloadBytes(bytes, name, KIND_INFO[kind].mime);
  return { handle: null, name };
}

export function downloadBytes(bytes: Uint8Array | string, name: string, mime: string) {
  const blob = new Blob([bytes as BlobPart], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = name;
  a.rel = 'noopener';
  a.style.display = 'none';
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 30_000);
}
