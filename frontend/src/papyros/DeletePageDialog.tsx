import { useEffect, useId, useRef, useState } from 'react';

interface Props {
  open: boolean;
  /** The page to offer first (the one with the cursor). */
  initialPage: number;
  totalPages: number;
  onConfirm(page: number): void;
  onCancel(): void;
}

/** Asks which page to delete. A native modal <dialog>, like ConfirmDialog. */
export function DeletePageDialog({ open, initialPage, totalPages, onConfirm, onCancel }: Props) {
  const ref = useRef<HTMLDialogElement>(null);
  const uid = useId();
  const [page, setPage] = useState(initialPage);

  useEffect(() => {
    const d = ref.current;
    if (!d) return;
    if (open && !d.open) {
      setPage(initialPage);
      d.showModal();
    }
    if (!open && d.open) d.close();
  }, [open, initialPage]);

  const valid = Number.isInteger(page) && page >= 1 && page <= totalPages;
  return (
    <dialog
      ref={ref}
      className="confirm"
      aria-labelledby={`${uid}-title`}
      aria-describedby={`${uid}-body`}
      onCancel={(e) => {
        e.preventDefault();
        onCancel();
      }}
    >
      <h2 id={`${uid}-title`}>Delete a page</h2>
      <p id={`${uid}-body`}>
        Everything on the page is removed, together with the page break or empty lines that create it, so the
        page disappears. Ctrl+Z brings it back.
      </p>
      <form
        className="delete-page__form"
        onSubmit={(e) => {
          e.preventDefault();
          if (valid) onConfirm(page);
        }}
      >
        <label htmlFor={`${uid}-page`}>Page</label>
        <input
          id={`${uid}-page`}
          type="number"
          min={1}
          max={totalPages}
          step={1}
          value={Number.isFinite(page) ? page : ''}
          onChange={(e) => setPage(e.target.valueAsNumber)}
          autoFocus
        />
        <span>of {totalPages}</span>
        <div className="confirm__actions">
          <button type="button" className="btn btn--outline" onClick={onCancel}>
            Cancel
          </button>
          <button type="submit" className="btn btn--danger" disabled={!valid}>
            Delete page {valid ? page : ''}
          </button>
        </div>
      </form>
    </dialog>
  );
}
