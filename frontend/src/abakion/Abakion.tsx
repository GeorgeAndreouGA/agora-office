/**
 * Abakion: the Excel editor. Grid and formula engine from Univer (Apache-2.0),
 * .xlsx reading/writing from ExcelJS (MIT). Everything runs in this tab.
 */
import { useCallback, useEffect, useRef, useState } from 'react';
import type * as ExcelJS from 'exceljs';
import {
  createUniver,
  defaultTheme,
  ILocalStorageService,
  LocaleType,
  mergeLocales,
  type FUniver,
  type Univer,
} from '@univerjs/presets';
import { UniverSheetsCorePreset } from '@univerjs/preset-sheets-core';
import UniverPresetSheetsCoreEnUS from '@univerjs/preset-sheets-core/locales/en-US';
import '@univerjs/preset-sheets-core/lib/index.css';
import './abakion.css';
import { ConfirmDialog } from '../components/ConfirmDialog';
import { EditorHeader } from '../components/EditorHeader';
import { downloadBytes } from '../lib/fileAccess';
import { useDocSession } from '../lib/session';
import { useToast } from '../lib/toast';
import type { EditorProps } from '../lib/types';
import { blankWorkbook, readXlsx, sheetToCsv, univerToXlsx, workbookToUniver } from './xlsx';

// Olive primary scale for the Univer chrome, matching the Greek Office palette.
const oliveTheme = {
  ...defaultTheme,
  primary: {
    50: '#F4F6EA',
    100: '#E6EBD0',
    200: '#CDD6A3',
    300: '#AEBB73',
    400: '#8E9E4B',
    500: '#6F7F33',
    600: '#4E5B23',
    700: '#3F4A1C',
    800: '#313915',
    900: '#22280E',
  },
};

/**
 * Univer normally keeps small UI preferences (recent number formats, emoji) in IndexedDB.
 * Greek Office keeps everything in memory, so that storage is replaced with a Map that dies with the tab.
 */
class MemoryStorageService implements ILocalStorageService {
  private readonly store = new Map<string, unknown>();
  async getItem<T>(key: string): Promise<T | null> {
    return (this.store.get(key) as T | undefined) ?? null;
  }
  async setItem<T>(key: string, value: T): Promise<T> {
    this.store.set(key, value);
    return value;
  }
  async removeItem(key: string): Promise<void> {
    this.store.delete(key);
  }
  async clear(): Promise<void> {
    this.store.clear();
  }
  async key(index: number): Promise<string | null> {
    return [...this.store.keys()][index] ?? null;
  }
  async keys(): Promise<string[]> {
    return [...this.store.keys()];
  }
  async iterate<T, U>(iteratee: (value: T, key: string, iterationNumber: number) => U): Promise<U> {
    let i = 0;
    let result: U | undefined;
    for (const [k, v] of this.store) {
      result = iteratee(v as T, k, i++);
      if (result !== undefined) return result;
    }
    return result as U;
  }
}

// Mutations that do not represent a user edit (layout, formula bookkeeping, selection).
const NON_EDIT = /formula\.mutation|calculation|auto-height|set-worksheet-row-auto|selection|scroll|zoom|activate|set-worksheet-active/i;
const MUTATION = 2;

export default function Abakion({ doc, onExit }: EditorProps) {
  const toast = useToast();
  const containerRef = useRef<HTMLDivElement>(null);
  const apiRef = useRef<FUniver | null>(null);
  const baseRef = useRef<ExcelJS.Workbook | null>(null);
  const idMapRef = useRef(new Map<string, number>());
  const [phase, setPhase] = useState<'loading' | 'ready' | 'error'>('loading');
  const [failure, setFailure] = useState('');
  const [confirmLeave, setConfirmLeave] = useState(false);
  const [notice, setNotice] = useState(doc.notice ?? '');

  const serialize = useCallback(async () => {
    const api = apiRef.current;
    if (!api) throw new Error('The workbook is not ready yet.');
    const fwb = api.getActiveWorkbook() as (ReturnType<FUniver['getActiveWorkbook']> & { endEditingAsync?: (save?: boolean) => Promise<boolean> });
    await fwb?.endEditingAsync?.(true);
    const { bytes, workbook } = await univerToXlsx(api, baseRef.current, idMapRef.current);
    baseRef.current = workbook;
    return bytes;
  }, []);

  const session = useDocSession({
    kind: 'xlsx',
    initialName: doc.name,
    initialHandle: doc.handle,
    serialize,
    onSaved: (how) => toast.show(how === 'autosave' ? 'Autosaved.' : 'Saved.', 'success'),
  });
  const markDirtyRef = useRef(session.markDirty);
  markDirtyRef.current = session.markDirty;

  useEffect(() => {
    let disposed = false;
    let univer: Univer | null = null;
    let worker: Worker | null = null;

    (async () => {
      let data;
      const workbookName = doc.name.replace(/\.xlsx$/i, '');
      if (doc.bytes) {
        const { workbook: wb, droppedDrawings } = await readXlsx(doc.bytes);
        if (disposed) return;
        baseRef.current = wb;
        if (droppedDrawings) {
          setNotice((n) =>
            n
              ? `${n} Charts, pictures and shapes are not shown.`
              : 'Charts, pictures and shapes in this workbook could not be read and are not shown.',
          );
        }
        const imported = workbookToUniver(wb, workbookName, LocaleType.EN_US);
        idMapRef.current = imported.idMap;
        data = imported.data;
      } else {
        data = blankWorkbook(workbookName, LocaleType.EN_US);
      }
      if (disposed || !containerRef.current) return;

      worker = new Worker(new URL('./univer.worker.ts', import.meta.url), { type: 'module', name: 'abakion-formula' });
      const created = createUniver({
        locale: LocaleType.EN_US,
        locales: { [LocaleType.EN_US]: mergeLocales(UniverPresetSheetsCoreEnUS) },
        theme: oliveTheme,
        override: [[ILocalStorageService, { useClass: MemoryStorageService }]],
        presets: [
          UniverSheetsCorePreset({
            container: containerRef.current,
            workerURL: worker,
            header: true,
            toolbar: true,
            footer: { sheetBar: true, statisticBar: true, menus: true, zoomSlider: true },
          }),
        ],
      });
      univer = created.univer;
      apiRef.current = created.univerAPI;
      const fwb = created.univerAPI.createWorkbook(data);
      fwb.onCommandExecuted((cmd) => {
        if (cmd.type !== MUTATION || NON_EDIT.test(cmd.id)) return;
        // Formula results written by the engine (on open, or after an edit) are not user edits.
        const params = (cmd.params ?? {}) as Record<string, unknown>;
        if ('sourceFormulaData' in params) return;
        markDirtyRef.current();
      });
      setPhase('ready');
    })().catch((err: unknown) => {
      if (disposed) return;
      setFailure(err instanceof Error ? err.message : String(err));
      setPhase('error');
    });

    return () => {
      disposed = true;
      apiRef.current = null;
      univer?.dispose();
      worker?.terminate();
    };
  }, [doc.bytes, doc.name]);

  const exportCsv = () => {
    const api = apiRef.current;
    if (!api) return;
    try {
      const { csv, sheetName } = sheetToCsv(api);
      const base = session.name.replace(/\.xlsx$/i, '');
      downloadBytes(`﻿${csv}`, `${base} - ${sheetName}.csv`, 'text/csv;charset=utf-8');
    } catch (err) {
      toast.show(`Could not export CSV: ${(err as Error).message}`, 'error');
    }
  };

  const back = () => (session.dirty ? setConfirmLeave(true) : onExit());

  return (
    <div className="editor-shell">
      <EditorHeader
        kind="xlsx"
        session={session}
        onBack={back}
        extraActions={
          <button type="button" className="btn btn--ghost-dark" onClick={exportCsv} disabled={phase !== 'ready'}>
            Export CSV
          </button>
        }
      />
      {session.error && (
        <div className="editor-alert" role="alert">
          {session.error}
        </div>
      )}
      {notice && (
        <div className="editor-notice" role="status">
          <span>{notice}</span>
          <button type="button" className="editor-notice__close" onClick={() => setNotice('')} aria-label="Dismiss notice">
            ×
          </button>
        </div>
      )}
      <div className="editor-body abakion">
        {phase === 'loading' && (
          <div className="loading loading--overlay" role="status">
            <span className="loading__mark" aria-hidden="true" />
            <span>Reading the workbook…</span>
          </div>
        )}
        {phase === 'error' && (
          <div className="editor-failure" role="alert">
            <h2>This workbook could not be opened</h2>
            <p>{failure}</p>
            <button type="button" className="btn btn--xlsx" onClick={onExit}>
              Back to the Greek Office
            </button>
          </div>
        )}
        <div ref={containerRef} className="abakion__grid" />
      </div>
      <ConfirmDialog
        open={confirmLeave}
        title="Leave without saving?"
        body="This workbook lives only in this tab. If you go back to the Greek Office now, your unsaved changes are lost."
        confirmLabel="Discard and leave"
        onCancel={() => setConfirmLeave(false)}
        onConfirm={onExit}
      />
    </div>
  );
}
