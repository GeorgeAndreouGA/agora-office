import { Suspense, lazy, useCallback, useEffect, useState } from 'react';
import { ErrorBoundary } from './components/ErrorBoundary';
import { Hub } from './hub/Hub';
import { KIND_INFO, type DocKind } from './lib/config';
import { editorLoaders } from './lib/editors';
import type { OpenedDoc } from './lib/types';

// React.lazy remembers a failed import forever, so a retry has to build a fresh lazy component.
const lazyEditors = {
  docx: lazy(editorLoaders.docx),
  xlsx: lazy(editorLoaders.xlsx),
  pptx: lazy(editorLoaders.pptx),
};

type View = { view: 'hub' } | { view: 'editor'; doc: OpenedDoc; key: number };

function Loading({ kind }: { kind: DocKind }) {
  return (
    <div className="loading" role="status">
      <span className="loading__mark" aria-hidden="true" />
      <span>Opening {KIND_INFO[kind].app}…</span>
    </div>
  );
}

export function App() {
  const [view, setView] = useState<View>({ view: 'hub' });

  const open = useCallback((doc: OpenedDoc) => {
    setView({ view: 'editor', doc, key: Date.now() });
  }, []);
  // Dropping the reference releases the document from memory.
  const exit = useCallback(() => setView({ view: 'hub' }), []);
  const retryLoad = useCallback((kind: DocKind) => {
    lazyEditors[kind] = lazy(editorLoaders[kind]);
    // A new key remounts the error boundary and the editor under it.
    setView((v) => (v.view === 'editor' ? { ...v, key: v.key + 1 } : v));
  }, []);

  const inEditor = view.view === 'editor';
  // The hub accepts file drops; an editor open in the page does not by default, so a stray drop
  // would make the browser navigate to (or download) the file and throw the document away.
  // Editors that handle drops themselves call preventDefault first and are left alone.
  useEffect(() => {
    if (!inEditor) return;
    const hasFiles = (e: DragEvent) => !!e.dataTransfer && Array.from(e.dataTransfer.types).includes('Files');
    const onOver = (e: DragEvent) => {
      if (e.defaultPrevented || !hasFiles(e)) return;
      e.preventDefault();
      if (e.dataTransfer) e.dataTransfer.dropEffect = 'none';
    };
    const onDrop = (e: DragEvent) => {
      if (e.defaultPrevented || !hasFiles(e)) return;
      e.preventDefault();
    };
    window.addEventListener('dragover', onOver);
    window.addEventListener('drop', onDrop);
    return () => {
      window.removeEventListener('dragover', onOver);
      window.removeEventListener('drop', onDrop);
    };
  }, [inEditor]);

  if (view.view === 'hub') return <Hub onOpen={open} />;

  const Editor = lazyEditors[view.doc.kind];
  return (
    <Suspense fallback={<Loading kind={view.doc.kind} />}>
      <ErrorBoundary key={view.key} kind={view.doc.kind} onExit={exit} onRetry={() => retryLoad(view.doc.kind)}>
        <Editor doc={view.doc} onExit={exit} />
      </ErrorBoundary>
    </Suspense>
  );
}
