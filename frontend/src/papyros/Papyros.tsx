/**
 * Papyros: the Word editor. Built on @docx-editor.dev (Apache-2.0, free packages only).
 * Parsing, editing and saving all happen in this tab; the document is never uploaded.
 */
import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { DocxEditor, useFonts, type DocxEditorRef } from '@docx-editor.dev/react';
import { packagedFonts } from '@docx-editor.dev/fonts';
import '@docx-editor.dev/core/styles/editor.css';
import './papyros.css';
import { ConfirmDialog } from '../components/ConfirmDialog';
import { EditorHeader } from '../components/EditorHeader';
import { choosePdfTarget } from '../lib/fileAccess';
import { useDocSession } from '../lib/session';
import { useToast } from '../lib/toast';
import type { EditorProps } from '../lib/types';
import { DeletePageDialog } from './DeletePageDialog';
import { printDocument } from './print';

// The packaged Export and Print rows need a server-side PDF converter (a paid add-on).
// Papyros prints in the browser instead (header button / Ctrl+P), so those rows are hidden.
// Open is hidden too: files come in through the Greek Office gate.
const FILE_MENU_OVERRIDES = (
  <DocxEditor.Menu.File preset={false}>
    <DocxEditor.Menu.Save />
    <DocxEditor.Menu.Separator />
    <DocxEditor.Menu.PageSetup />
  </DocxEditor.Menu.File>
);

export default function Papyros({ doc, onExit }: EditorProps) {
  const toast = useToast();
  const editorRef = useRef<DocxEditorRef>(null);
  const fonts = useFonts(packagedFonts());
  const [confirmLeave, setConfirmLeave] = useState(false);
  const [printing, setPrinting] = useState<string | null>(null);
  const [pageDialog, setPageDialog] = useState<{ current: number; total: number } | null>(null);
  const [marks, setMarks] = useState(false);
  const bodyRef = useRef<HTMLDivElement>(null);

  // A stable source: a new array identity would make the editor reload the document.
  const source = useMemo(() => doc.bytes ?? ('blank' as const), [doc.bytes]);

  const serialize = useCallback(async () => {
    const buf = await editorRef.current?.save();
    if (!buf) throw new Error('The editor is not ready yet.');
    return new Uint8Array(buf);
  }, []);

  const session = useDocSession({
    kind: 'docx',
    initialName: doc.name,
    initialHandle: doc.handle,
    serialize,
    onSaved: (how) => toast.show(how === 'autosave' ? 'Autosaved.' : 'Saved.', 'success'),
  });

  const title = session.name.replace(/\.docx$/i, '');
  const menu = useMemo(
    () => ({ reportIssue: false, onSave: () => void session.save(), children: FILE_MENU_OVERRIDES }),
    [session.save],
  );

  const print = useCallback(async () => {
    const root = bodyRef.current;
    if (!root || printing) return;
    setPrinting('Preparing pages…');
    try {
      await printDocument(root, (done, total) => setPrinting(`Preparing page ${done} of ${total}…`));
    } catch (err) {
      toast.show(`Could not print: ${(err as Error).message}`, 'error');
    } finally {
      setPrinting(null);
    }
  }, [printing, toast]);

  // One-click PDF, built in this tab from the editor's own layout (papyros/docxPdf.ts).
  const downloadPdf = useCallback(async () => {
    if (printing) return;
    let target;
    try {
      target = await choosePdfTarget(title);
    } catch (err) {
      toast.show(`Could not save the PDF: ${(err as Error).message}`, 'error');
      return;
    }
    if (!target) return;
    setPrinting('Building PDF…');
    try {
      const bytes = await serialize();
      const { docxToPdf, pdfSavedMessage } = await import('./docxPdf');
      const result = await docxToPdf(bytes, {
        title,
        onProgress: (done, total) => setPrinting(`Building PDF: page ${done} of ${total}…`),
      });
      await target.write(result.pdf);
      toast.show(pdfSavedMessage(result.simplified), result.simplified.length ? 'info' : 'success');
    } catch (err) {
      await target.discard();
      toast.show(`Could not build the PDF: ${(err as Error).message}`, 'error');
    } finally {
      setPrinting(null);
    }
  }, [printing, serialize, title, toast]);

  // Delete page: Word has no page objects, so this removes what makes the page (papyros/deletePage.ts).
  const openDeletePage = useCallback(() => {
    const editor = editorRef.current?.getEditor();
    if (!editor) return;
    const total = editor.getTotalPages();
    if (!total) return;
    setPageDialog({ current: Math.min(Math.max(editor.getCurrentPage('caret') || 1, 1), total), total });
  }, []);

  const confirmDeletePage = useCallback(
    async (page: number) => {
      setPageDialog(null);
      const editor = editorRef.current?.getEditor();
      const root = bodyRef.current;
      if (!editor || !root || printing) return;
      setPrinting(`Deleting page ${page}…`);
      try {
        const { deletePage } = await import('./deletePage');
        const result = await deletePage(editor, root, page);
        toast.show(result.message, result.deleted ? 'success' : result.changed ? 'info' : 'error');
      } catch (err) {
        toast.show(`Page ${page} was not deleted: ${(err as Error).message}`, 'error');
      } finally {
        setPrinting(null);
        editorRef.current?.focus();
      }
    },
    [printing, toast],
  );

  // ¶ shows paragraph marks and breaks, which is what makes an empty page visible.
  const toggleMarks = useCallback(() => {
    const r = editorRef.current?.exec({ type: 'toggleParagraphMarks' });
    if (r?.ok) setMarks((m) => !m);
  }, []);

  // Ctrl/Cmd+P prints through Papyros (the packaged shortcut needs the paid converter).
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (!(e.ctrlKey || e.metaKey) || e.altKey || e.shiftKey || e.key.toLowerCase() !== 'p') return;
      e.preventDefault();
      e.stopPropagation();
      void print();
    };
    window.addEventListener('keydown', onKey, { capture: true });
    return () => window.removeEventListener('keydown', onKey, { capture: true });
  }, [print]);

  const back = () => (session.dirty ? setConfirmLeave(true) : onExit());

  return (
    <div className="editor-shell">
      <EditorHeader
        kind="docx"
        session={session}
        onBack={back}
        extraActions={
          <>
            <button
              type="button"
              className={`btn btn--ghost-dark ${marks ? 'btn--pressed' : ''}`}
              onClick={toggleMarks}
              aria-pressed={marks}
              title="Show paragraph marks and breaks"
              aria-label="Show paragraph marks and breaks"
            >
              ¶
            </button>
            <button type="button" className="btn btn--ghost-dark" onClick={openDeletePage} disabled={!!printing}>
              Delete page
            </button>
            <button type="button" className="btn btn--ghost-dark" onClick={() => void print()} disabled={!!printing}>
              Print
            </button>
            <button type="button" className="btn btn--ghost-dark" onClick={() => void downloadPdf()} disabled={!!printing}>
              Download PDF
            </button>
          </>
        }
      />
      {session.error && (
        <div className="editor-alert" role="alert">
          {session.error}
        </div>
      )}
      <div className="editor-body papyros" ref={bodyRef}>
        {printing && (
          <div className="print-progress" role="status">
            {printing}
          </div>
        )}
        <DocxEditor
          ref={editorRef}
          className="greek-docx"
          document={source}
          fonts={fonts}
          colorMode="light"
          title={title}
          onTitleChange={session.rename}
          onSave={() => void session.save()}
          menu={menu}
          onChange={(change) => {
            if (!change.source) session.markDirty();
          }}
        />
      </div>
      <DeletePageDialog
        open={!!pageDialog}
        initialPage={pageDialog?.current ?? 1}
        totalPages={pageDialog?.total ?? 1}
        onCancel={() => setPageDialog(null)}
        onConfirm={(page) => void confirmDeletePage(page)}
      />
      <ConfirmDialog
        open={confirmLeave}
        title="Leave without saving?"
        body="This document lives only in this tab. If you go back to the Greek Office now, your unsaved changes are lost."
        confirmLabel="Discard and leave"
        onCancel={() => setConfirmLeave(false)}
        onConfirm={onExit}
      />
    </div>
  );
}
