import { useEffect, useId, useRef } from 'react';

interface Props {
  open: boolean;
  title: string;
  body: string;
  confirmLabel: string;
  cancelLabel?: string;
  onConfirm(): void;
  onCancel(): void;
}

/** A native modal <dialog>: focus trapping and Esc come from the browser. */
export function ConfirmDialog({ open, title, body, confirmLabel, cancelLabel = 'Stay', onConfirm, onCancel }: Props) {
  const ref = useRef<HTMLDialogElement>(null);
  // Unique per instance, so two dialogs on a page never share (or steal) a label id.
  const uid = useId();
  const titleId = `${uid}-title`;
  const bodyId = `${uid}-body`;
  useEffect(() => {
    const d = ref.current;
    if (!d) return;
    if (open && !d.open) d.showModal();
    if (!open && d.open) d.close();
  }, [open]);

  return (
    <dialog
      ref={ref}
      className="confirm"
      onCancel={(e) => {
        e.preventDefault();
        onCancel();
      }}
      aria-labelledby={titleId}
      aria-describedby={bodyId}
    >
      <h2 id={titleId}>{title}</h2>
      <p id={bodyId}>{body}</p>
      <div className="confirm__actions">
        <button type="button" className="btn btn--outline" onClick={onCancel} autoFocus>
          {cancelLabel}
        </button>
        <button type="button" className="btn btn--danger" onClick={onConfirm}>
          {confirmLabel}
        </button>
      </div>
    </dialog>
  );
}
