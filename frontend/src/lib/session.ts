/**
 * One open document's lifecycle: name, file handle, unsaved changes, Save / Save As,
 * the 5-minute autosave and the leave-page warning.
 *
 * The document itself lives only in the editor's memory. This hook never touches
 * localStorage, IndexedDB or any network API.
 */
import { useCallback, useEffect, useRef, useState } from 'react';
import { AUTOSAVE_MS, type DocKind } from './config';
import {
  canWriteInPlace,
  ensureExtension,
  requestWritePermission,
  saveAs as saveAsFile,
  writePermission,
  writeToHandle,
} from './fileAccess';

export type AutosaveIssue = null | 'needs-permission' | 'needs-location' | 'unsupported';

export interface SessionState {
  name: string;
  hasHandle: boolean;
  dirty: boolean;
  saving: boolean;
  lastSavedAt: number | null;
  nextAutosaveAt: number;
  autosaveIssue: AutosaveIssue;
  error: string | null;
}

export interface DocSession extends SessionState {
  markDirty(): void;
  save(): Promise<boolean>;
  saveAs(): Promise<boolean>;
  enableAutosave(): Promise<void>;
  rename(name: string): void;
}

interface Options {
  kind: DocKind;
  initialName: string;
  initialHandle: FileSystemFileHandle | null;
  /** Serialise the current document. Called only when saving. */
  serialize: () => Promise<Uint8Array>;
  onSaved?: (how: 'save' | 'autosave' | 'saveAs') => void;
}

export function useDocSession({ kind, initialName, initialHandle, serialize, onSaved }: Options): DocSession {
  const handleRef = useRef<FileSystemFileHandle | null>(initialHandle);
  const serializeRef = useRef(serialize);
  serializeRef.current = serialize;
  const onSavedRef = useRef(onSaved);
  onSavedRef.current = onSaved;
  const dirtyRef = useRef(false);
  const savingRef = useRef(false);
  // Bumped on every edit. A save remembers the value from just before it serialised, and only
  // clears "unsaved" if nothing changed meanwhile; otherwise an edit made while the write was
  // in flight would be reported as saved and skip the leave-page warning.
  const changeRef = useRef(0);

  const [state, setState] = useState<SessionState>(() => ({
    name: ensureExtension(initialName, kind),
    hasHandle: !!initialHandle,
    dirty: false,
    saving: false,
    lastSavedAt: null,
    nextAutosaveAt: Date.now() + AUTOSAVE_MS,
    autosaveIssue: canWriteInPlace ? null : 'unsupported',
    error: null,
  }));
  const nameRef = useRef(state.name);
  nameRef.current = state.name;

  const patch = useCallback((p: Partial<SessionState>) => setState((s) => ({ ...s, ...p })), []);

  const markDirty = useCallback(() => {
    changeRef.current += 1;
    dirtyRef.current = true;
    // Returning the same object when already dirty keeps typing from re-rendering the editor.
    setState((s) => (s.dirty ? s : { ...s, dirty: true }));
  }, []);

  const finishSave = useCallback(
    (how: 'save' | 'autosave' | 'saveAs', seenChange: number, extra: Partial<SessionState> = {}) => {
      const unchanged = changeRef.current === seenChange;
      if (unchanged) dirtyRef.current = false;
      patch({
        dirty: !unchanged,
        lastSavedAt: Date.now(),
        error: null,
        autosaveIssue: null,
        ...extra,
      });
      onSavedRef.current?.(how);
    },
    [patch],
  );

  const doSaveAs = useCallback(async (): Promise<boolean> => {
    if (savingRef.current) return false;
    savingRef.current = true;
    patch({ saving: true });
    try {
      const seenChange = changeRef.current;
      const bytes = await serializeRef.current();
      const result = await saveAsFile(bytes, nameRef.current, kind);
      if (!result) return false;
      handleRef.current = result.handle;
      finishSave('saveAs', seenChange, {
        name: result.name,
        hasHandle: !!result.handle,
        autosaveIssue: result.handle ? null : canWriteInPlace ? 'needs-location' : 'unsupported',
      });
      return true;
    } catch (err) {
      patch({ error: `Could not save: ${(err as Error).message}` });
      return false;
    } finally {
      savingRef.current = false;
      patch({ saving: false });
    }
  }, [finishSave, kind, patch]);

  /** Manual save: called from a click or Ctrl+S, so permission prompts are allowed. */
  const save = useCallback(async (): Promise<boolean> => {
    const handle = handleRef.current;
    if (!handle) return doSaveAs();
    if (savingRef.current) return false;
    savingRef.current = true;
    patch({ saving: true });
    try {
      if ((await writePermission(handle)) !== 'granted' && !(await requestWritePermission(handle))) {
        patch({ autosaveIssue: 'needs-permission', error: 'Permission to write the file was not given.' });
        return false;
      }
      const seenChange = changeRef.current;
      const bytes = await serializeRef.current();
      // Renamed while we waited: the handle was dropped, so do not overwrite the old file.
      if (handleRef.current !== handle) return false;
      await writeToHandle(handle, bytes, kind);
      finishSave('save', seenChange);
      return true;
    } catch (err) {
      patch({ error: `Could not save: ${(err as Error).message}` });
      return false;
    } finally {
      savingRef.current = false;
      patch({ saving: false });
    }
  }, [doSaveAs, finishSave, kind, patch]);

  /** Click handler for the "turn on autosave" chip. */
  const enableAutosave = useCallback(async () => {
    const handle = handleRef.current;
    if (!handle) {
      await doSaveAs();
      return;
    }
    if (await requestWritePermission(handle)) {
      patch({ autosaveIssue: null });
      if (dirtyRef.current) await save();
    }
  }, [doSaveAs, patch, save]);

  // Check write permission up front so the header can offer to turn autosave on.
  useEffect(() => {
    const handle = handleRef.current;
    if (!canWriteInPlace) return;
    if (!handle) {
      patch({ autosaveIssue: 'needs-location' });
      return;
    }
    writePermission(handle)
      .then((p) => patch({ autosaveIssue: p === 'granted' ? null : 'needs-permission' }))
      .catch(() => patch({ autosaveIssue: 'needs-permission' }));
  }, [patch]);

  // The autosave clock. Runs every 5 minutes; writes only when there are unsaved changes
  // and the browser already allows silent writes to the file.
  useEffect(() => {
    let timer = 0;
    const tick = async () => {
      patch({ nextAutosaveAt: Date.now() + AUTOSAVE_MS });
      timer = window.setTimeout(tick, AUTOSAVE_MS);
      if (!dirtyRef.current || savingRef.current) return;
      const handle = handleRef.current;
      if (!canWriteInPlace) {
        patch({ autosaveIssue: 'unsupported' });
        return;
      }
      if (!handle) {
        patch({ autosaveIssue: 'needs-location' });
        return;
      }
      if ((await writePermission(handle)) !== 'granted') {
        patch({ autosaveIssue: 'needs-permission' });
        return;
      }
      // The permission check yielded: a manual save may have started or finished meanwhile.
      if (savingRef.current || !dirtyRef.current || handleRef.current !== handle) return;
      savingRef.current = true;
      patch({ saving: true });
      try {
        const seenChange = changeRef.current;
        const bytes = await serializeRef.current();
        if (handleRef.current !== handle) return;
        await writeToHandle(handle, bytes, kind);
        finishSave('autosave', seenChange);
      } catch (err) {
        patch({ error: `Autosave failed: ${(err as Error).message}` });
      } finally {
        savingRef.current = false;
        patch({ saving: false });
      }
    };
    timer = window.setTimeout(tick, AUTOSAVE_MS);
    return () => window.clearTimeout(timer);
  }, [finishSave, kind, patch]);

  // Warn before the tab is closed or reloaded with unsaved work: it lives only in memory.
  useEffect(() => {
    const onBeforeUnload = (e: BeforeUnloadEvent) => {
      if (!dirtyRef.current) return;
      e.preventDefault();
      e.returnValue = '';
    };
    window.addEventListener('beforeunload', onBeforeUnload);
    return () => window.removeEventListener('beforeunload', onBeforeUnload);
  }, []);

  // Ctrl/Cmd+S and Ctrl/Cmd+Shift+S, captured before the editors see them.
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (!(e.ctrlKey || e.metaKey) || e.altKey || e.key.toLowerCase() !== 's') return;
      e.preventDefault();
      e.stopPropagation();
      void (e.shiftKey ? doSaveAs() : save());
    };
    window.addEventListener('keydown', onKey, { capture: true });
    return () => window.removeEventListener('keydown', onKey, { capture: true });
  }, [doSaveAs, save]);

  const rename = useCallback(
    (name: string) => {
      const clean = name.trim();
      if (!clean) return;
      const next = ensureExtension(clean, kind);
      if (next === nameRef.current) return;
      nameRef.current = next;
      const handle = handleRef.current;
      if (handle && next !== handle.name) {
        // The handle points at the file under its old name. Keeping it would make Save overwrite
        // that file while the header shows the new name, so the next Save goes through Save As.
        handleRef.current = null;
        patch({
          name: next,
          hasHandle: false,
          autosaveIssue: canWriteInPlace ? 'needs-location' : 'unsupported',
        });
        return;
      }
      patch({ name: next });
    },
    [kind, patch],
  );

  return { ...state, markDirty, save, saveAs: doSaveAs, enableAutosave, rename };
}
