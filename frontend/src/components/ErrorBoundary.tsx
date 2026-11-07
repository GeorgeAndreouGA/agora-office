import { Component, type ErrorInfo, type ReactNode } from 'react';
import type { DocKind } from '../lib/config';

interface Props {
  kind: DocKind;
  onExit(): void;
  /** Called before a retry so the caller can discard a failed lazy import. */
  onRetry?(): void;
  children: ReactNode;
}

interface State {
  error: Error | null;
}

/** A failed lazy import surfaces as one of these rather than as an error in the editor's own code. */
function isChunkLoadError(err: Error): boolean {
  return (
    err.name === 'ChunkLoadError' ||
    /Failed to fetch dynamically imported module|error loading dynamically imported module|Importing a module script failed|Loading chunk .* failed/i.test(
      err.message,
    )
  );
}

/**
 * Catches a crash (or a failed editor download) so the person sees a way out instead of a blank
 * page. App keys it on the open document, so it resets whenever a different document opens.
 */
export class ErrorBoundary extends Component<Props, State> {
  state: State = { error: null };

  static getDerivedStateFromError(error: Error): State {
    return { error };
  }

  componentDidCatch(error: Error, info: ErrorInfo) {
    // Console only: nothing leaves the browser.
    console.error('Editor crashed', error, info.componentStack);
  }

  private retry = () => {
    this.props.onRetry?.();
    this.setState({ error: null });
  };

  render() {
    const { error } = this.state;
    if (!error) return this.props.children;

    const { kind, onExit } = this.props;
    const loadFailure = isChunkLoadError(error);
    // editor-failure is absolutely positioned, so it needs a positioned, full-height parent.
    return (
      <div className="editor-shell">
        <div className="editor-body">
          <div className="editor-failure" role="alert">
            <h2>Something went wrong</h2>
            <p>
              {loadFailure
                ? 'The editor could not be loaded. Check your connection and try again.'
                : error.message || 'An unexpected error stopped the editor.'}
            </p>
            {loadFailure ? (
              <button type="button" className={`btn btn--${kind}`} onClick={this.retry}>
                Try again
              </button>
            ) : null}
            <button type="button" className={`btn btn--${kind}`} onClick={onExit}>
              Back to the Greek Office
            </button>
          </div>
        </div>
      </div>
    );
  }
}
