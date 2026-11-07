import { useCallback, useEffect, useRef, useState } from 'react';
import {
  AbacusIcon,
  ClockIcon,
  DropIcon,
  Meander,
  MemoryIcon,
  Pediment,
  PinaxIcon,
  ScrollIcon,
  ShieldIcon,
} from '../components/Ornaments';
import { type DocKind } from '../lib/config';
import { preloadEditor } from '../lib/editors';
import { canWriteInPlace, fileFromDrop, pickFile, type PickedFile } from '../lib/fileAccess';
import { GateError, gateFile } from '../lib/fileGate';
import { useToast } from '../lib/toast';
import type { OpenedDoc } from '../lib/types';

interface Props {
  onOpen(doc: OpenedDoc): void;
}


export function Hub({ onOpen }: Props) {
  const toast = useToast();
  const [busy, setBusy] = useState(false);
  const [dragging, setDragging] = useState(false);
  const dragDepth = useRef(0);
  // State is stale inside the long-lived drop listener, so overlap is guarded with a ref.
  const busyRef = useRef(false);

  const admit = useCallback(
    async (picked: PickedFile | null, expected?: DocKind) => {
      if (!picked || busyRef.current) return;
      busyRef.current = true;
      setBusy(true);
      try {
        const gated = await gateFile(picked.file, expected);
        if (gated.lossyParts.length) {
          // Never write over the original silently: it would lose these parts. Saving goes to a new file.
          const list = gated.lossyParts.join(', ');
          onOpen({
            kind: gated.kind,
            name: gated.name.replace(/\.xlsx$/i, ' (edited).xlsx'),
            bytes: gated.bytes,
            handle: null,
            notice:
              `This workbook contains ${list}, which Abakion cannot keep. Your original file will not be ` +
              'overwritten: Save and autosave write to a new file you choose.',
          });
          return;
        }
        onOpen({ kind: gated.kind, name: gated.name, bytes: gated.bytes, handle: picked.handle });
      } catch (err) {
        toast.show(err instanceof GateError ? err.message : 'That file could not be opened.', 'error');
      } finally {
        busyRef.current = false;
        setBusy(false);
      }
    },
    [onOpen, toast],
  );

  const openKind = useCallback(
    async (kind: DocKind) => {
      preloadEditor(kind);
      try {
        await admit(await pickFile([kind]), kind);
      } catch {
        toast.show('The file picker could not be opened.', 'error');
      }
    },
    [admit, toast],
  );

  const newDoc = useCallback(
    (kind: DocKind) => {
      onOpen({
        kind,
        name: kind === 'docx' ? 'Untitled document.docx' : kind === 'xlsx' ? 'Untitled workbook.xlsx' : 'Untitled presentation.pptx',
        bytes: null,
        handle: null,
      });
    },
    [onOpen],
  );

  // The whole page is a drop target.
  useEffect(() => {
    const hasFiles = (e: DragEvent) => !!e.dataTransfer && Array.from(e.dataTransfer.types).includes('Files');
    const onEnter = (e: DragEvent) => {
      if (!hasFiles(e)) return;
      e.preventDefault();
      dragDepth.current += 1;
      setDragging(true);
    };
    const onOver = (e: DragEvent) => {
      if (!hasFiles(e)) return;
      e.preventDefault();
      if (e.dataTransfer) e.dataTransfer.dropEffect = 'copy';
    };
    const onLeave = () => {
      dragDepth.current = Math.max(0, dragDepth.current - 1);
      if (dragDepth.current === 0) setDragging(false);
    };
    const onDrop = (e: DragEvent) => {
      if (!hasFiles(e) || !e.dataTransfer) return;
      e.preventDefault();
      dragDepth.current = 0;
      setDragging(false);
      if (busyRef.current) return;
      if (e.dataTransfer.items.length > 1) {
        toast.show('Drop one file at a time.', 'error');
        return;
      }
      void fileFromDrop(e.dataTransfer).then((picked) => admit(picked));
    };
    window.addEventListener('dragenter', onEnter);
    window.addEventListener('dragover', onOver);
    window.addEventListener('dragleave', onLeave);
    window.addEventListener('drop', onDrop);
    return () => {
      window.removeEventListener('dragenter', onEnter);
      window.removeEventListener('dragover', onOver);
      window.removeEventListener('dragleave', onLeave);
      window.removeEventListener('drop', onDrop);
    };
  }, [admit, toast]);

  return (
    <div className={`hub ${dragging ? 'hub--dragging' : ''}`} aria-busy={busy}>
      <header className="hub__top">
        <div className="brand">
          <span className="brand__greek" lang="el">
            ΕΛΛΗΝΙΚΟ ΓΡΑΦΕΙΟ
          </span>
          <span className="brand__latin">GREEK OFFICE</span>
        </div>
        <span className="hub__status">
          <span className="hub__status-dot" aria-hidden="true" />
          Local · held in memory only
        </span>
      </header>
      <Meander />

      <main className="hub__main">
        <section className="hero">
          <p className="hero__eyebrow">CHOOSE YOUR INSTRUMENT</p>
          <h1 className="hero__title">Enter the Greek Office</h1>
          <p className="hero__lede">
            Edit Word, Excel and PowerPoint files right here in your browser. Nothing is uploaded. Close or reload the
            tab and the file is gone unless you saved it.
          </p>
        </section>

        <section className="cards">
          <article
            className="card card--docx"
            onMouseEnter={() => preloadEditor('docx')}
            onFocus={() => preloadEditor('docx')}
          >
            <Pediment tone="terracotta" />
            <div className="card__head">
              <div className="card__seal">
                <ScrollIcon />
              </div>
              <div>
                <h2 className="card__title">Papyros</h2>
                <span className="card__sub">
                  <span lang="el">Πάπυρος</span> · Word documents
                </span>
              </div>
            </div>
            <p className="card__text">
              Write and format .docx files: styles, tables, images, headers and page layout, saved back as a real Word
              file.
            </p>
            <div className="card__actions">
              <button type="button" className="btn btn--docx btn--lg" disabled={busy} onClick={() => void openKind('docx')}>
                Open .docx
              </button>
              <button type="button" className="btn btn--outline-docx btn--lg" disabled={busy} onClick={() => newDoc('docx')}>
                New document
              </button>
            </div>
          </article>

          <article
            className="card card--xlsx"
            onMouseEnter={() => preloadEditor('xlsx')}
            onFocus={() => preloadEditor('xlsx')}
          >
            <Pediment tone="olive" />
            <div className="card__head">
              <div className="card__seal">
                <AbacusIcon />
              </div>
              <div>
                <h2 className="card__title">Abakion</h2>
                <span className="card__sub">
                  <span lang="el">Ἀβάκιον</span> · Excel workbooks
                </span>
              </div>
            </div>
            <p className="card__text">
              Calculate in .xlsx files: formulas, multiple sheets, formatting, merged cells and frozen panes, saved back
              as a real Excel file.
            </p>
            <div className="card__actions">
              <button type="button" className="btn btn--xlsx btn--lg" disabled={busy} onClick={() => void openKind('xlsx')}>
                Open .xlsx
              </button>
              <button type="button" className="btn btn--outline-xlsx btn--lg" disabled={busy} onClick={() => newDoc('xlsx')}>
                New workbook
              </button>
            </div>
          </article>

          <article
            className="card card--pptx"
            onMouseEnter={() => preloadEditor('pptx')}
            onFocus={() => preloadEditor('pptx')}
          >
            <Pediment tone="aegean" />
            <div className="card__head">
              <div className="card__seal">
                <PinaxIcon />
              </div>
              <div>
                <h2 className="card__title">Pinax</h2>
                <span className="card__sub">
                  <span lang="el">Πίναξ</span> · PowerPoint presentations
                </span>
              </div>
            </div>
            <p className="card__text">
              Build and present .pptx slides: text, shapes, pictures, speaker notes and a full-screen slideshow, saved
              back as a real PowerPoint file.
            </p>
            <div className="card__actions">
              <button type="button" className="btn btn--pptx btn--lg" disabled={busy} onClick={() => void openKind('pptx')}>
                Open .pptx
              </button>
              <button type="button" className="btn btn--outline-pptx btn--lg" disabled={busy} onClick={() => newDoc('pptx')}>
                New presentation
              </button>
            </div>
          </article>
        </section>

        <section className="drop" aria-label="Drop a file">
          <DropIcon />
          <p className="drop__title">{dragging ? 'Release to open' : 'Drop a .docx, .xlsx or .pptx anywhere on this page'}</p>
          <p className="drop__hint">
            Only Word (.docx), Excel (.xlsx) and PowerPoint (.pptx) files are accepted. Macro files and
            password-protected files are refused.
          </p>
        </section>

        <section className="facts">
          <div className="fact">
            <MemoryIcon />
            <div>
              <h3>Memory only</h3>
              <p>Your file lives in this tab. Reloading clears it, so you are warned before leaving with unsaved work.</p>
            </div>
          </div>
          <div className="fact">
            <ClockIcon />
            <div>
              <h3>Autosave · 5 min</h3>
              <p>
                {canWriteInPlace
                  ? 'Writes back to the original file on your computer every five minutes.'
                  : 'This browser cannot write to files silently. Use Chrome or Edge for autosave; Save still downloads a copy.'}
              </p>
            </div>
          </div>
          <div className="fact">
            <ShieldIcon />
            <div>
              <h3>Checked at the gate</h3>
              <p>Every file is inspected before opening. Macros never run and nothing is sent to a server.</p>
            </div>
          </div>
        </section>
      </main>

      <Meander />
      <footer className="hub__foot">
        <span>Greek Office · runs entirely in your browser</span>
        <span>Ctrl+S save · Ctrl+Shift+S save as</span>
      </footer>
    </div>
  );
}
