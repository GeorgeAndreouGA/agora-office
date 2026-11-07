import type { DocKind } from './config';

/** A document handed from the hub to an editor. bytes === null means a new, blank file. */
export interface OpenedDoc {
  kind: DocKind;
  name: string;
  bytes: Uint8Array | null;
  handle: FileSystemFileHandle | null;
  /** Shown above the editor when opening, e.g. features that will not survive saving. */
  notice?: string;
}

export interface EditorProps {
  doc: OpenedDoc;
  onExit(): void;
}
