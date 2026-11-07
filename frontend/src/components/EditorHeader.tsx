import { useEffect, useState, type ReactNode } from 'react';
import { KIND_INFO, type DocKind } from '../lib/config';
import type { DocSession } from '../lib/session';
import { BackIcon } from './Ornaments';

function clock(ts: number): string {
  return new Date(ts).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
}

/** Its own ticking state, so the editor below never re-renders once a second. */
function AutosaveStatus({ session }: { session: DocSession }) {
  const [now, setNow] = useState(() => Date.now());
  useEffect(() => {
    const t = window.setInterval(() => setNow(Date.now()), 1000);
    return () => window.clearInterval(t);
  }, []);

  if (session.saving) return <span className="hdr__autosave">Saving…</span>;

  if (session.autosaveIssue === 'needs-permission') {
    return (
      <button type="button" className="hdr__chip-btn" onClick={() => void session.enableAutosave()}>
        Turn on autosave for this file
      </button>
    );
  }
  if (session.autosaveIssue === 'needs-location') {
    return (
      <button type="button" className="hdr__chip-btn" onClick={() => void session.enableAutosave()}>
        Choose where to autosave
      </button>
    );
  }
  if (session.autosaveIssue === 'unsupported') {
    return (
      <span className="hdr__autosave" title="Silent autosave needs Chrome or Edge">
        No autosave in this browser · press Save
      </span>
    );
  }

  // Nothing to save: a countdown would promise a save that will not happen.
  if (!session.dirty) {
    return (
      <span className="hdr__autosave">
        {session.lastSavedAt ? `All changes saved · ${clock(session.lastSavedAt)}` : 'Autosave on'}
      </span>
    );
  }

  const left = Math.max(0, session.nextAutosaveAt - now);
  const m = Math.floor(left / 60000);
  const s = Math.floor((left % 60000) / 1000)
    .toString()
    .padStart(2, '0');
  const last = session.lastSavedAt ? `Saved ${clock(session.lastSavedAt)} · ` : '';
  return (
    <span className="hdr__autosave">
      {last}autosave in {m}:{s}
    </span>
  );
}

interface Props {
  kind: DocKind;
  session: DocSession;
  onBack(): void;
  extraActions?: ReactNode;
}

export function EditorHeader({ kind, session, onBack, extraActions }: Props) {
  const info = KIND_INFO[kind];
  return (
    <header className={`hdr hdr--${kind}`}>
      <div className="hdr__left">
        <button type="button" className="hdr__back" onClick={onBack} aria-label="Back to the Greek Office hub">
          <BackIcon />
          <span lang="el">ΓΡΑΦΕΙΟ</span>
        </button>
        <span className="hdr__divider" aria-hidden="true" />
        <span className="hdr__app">{info.app}</span>
        <span className="hdr__file" title={session.name}>
          {session.name}
        </span>
        <span className={`hdr__pill ${session.dirty ? 'hdr__pill--dirty' : ''}`}>
          <span className="hdr__dot" aria-hidden="true" />
          {session.dirty ? 'Unsaved changes' : session.lastSavedAt ? 'All changes saved' : 'No changes'}
        </span>
      </div>
      <div className="hdr__right">
        <AutosaveStatus session={session} />
        {extraActions}
        <button type="button" className="btn btn--ghost-dark" onClick={() => void session.saveAs()} disabled={session.saving}>
          Save as…
        </button>
        <button type="button" className={`btn btn--${kind}`} onClick={() => void session.save()} disabled={session.saving}>
          Save
        </button>
      </div>
    </header>
  );
}
