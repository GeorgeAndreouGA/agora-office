/**
 * Full-screen slideshow. Arrows, space, Page Up/Down and clicks move; Esc leaves. Hidden
 * slides are skipped. Animations and transitions are not played (they are kept in the file).
 */
import { useCallback, useEffect, useLayoutEffect, useRef, useState } from 'react';
import type { Slide } from './model';
import { SlideView, slidePx, type RenderEnv } from './render';

export function Slideshow({ env, slides, start, onClose }: { env: RenderEnv; slides: Slide[]; start: number; onClose(): void }) {
  const shown = slides.map((s, i) => ({ s, i })).filter((x) => !x.s.hidden);
  const first = Math.max(0, shown.findIndex((x) => x.i >= start));
  const [pos, setPos] = useState(first);
  const [scale, setScale] = useState(1);
  const root = useRef<HTMLDivElement>(null);
  const { w, h } = slidePx(env.deck);
  const entered = useRef(false);
  const atEnd = pos >= shown.length;

  const next = useCallback(() => setPos((p) => Math.min(p + 1, shown.length)), [shown.length]);
  const prev = useCallback(() => setPos((p) => Math.max(p - 1, 0)), []);

  useLayoutEffect(() => {
    const fit = () => setScale(Math.min(window.innerWidth / w, window.innerHeight / h));
    fit();
    window.addEventListener('resize', fit);
    return () => window.removeEventListener('resize', fit);
  }, [w, h]);

  useEffect(() => {
    const el = root.current;
    el?.focus();
    el?.requestFullscreen?.()
      .then(() => {
        entered.current = true;
      })
      .catch(() => undefined);
    const onFs = () => {
      if (entered.current && !document.fullscreenElement) onClose();
    };
    document.addEventListener('fullscreenchange', onFs);
    return () => {
      document.removeEventListener('fullscreenchange', onFs);
      if (document.fullscreenElement) void document.exitFullscreen().catch(() => undefined);
    };
  }, [onClose]);

  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      const k = e.key;
      if (k === 'Escape') {
        e.preventDefault();
        onClose();
      } else if (['ArrowRight', 'ArrowDown', 'PageDown', ' ', 'Enter', 'n', 'N'].includes(k)) {
        e.preventDefault();
        if (atEnd) onClose();
        else next();
      } else if (['ArrowLeft', 'ArrowUp', 'PageUp', 'Backspace', 'p', 'P'].includes(k)) {
        e.preventDefault();
        prev();
      } else if (k === 'Home') {
        e.preventDefault();
        setPos(0);
      } else if (k === 'End') {
        e.preventDefault();
        setPos(Math.max(0, shown.length - 1));
      }
      e.stopPropagation();
    };
    window.addEventListener('keydown', onKey, { capture: true });
    return () => window.removeEventListener('keydown', onKey, { capture: true });
  }, [atEnd, next, prev, onClose, shown.length]);

  const cur = shown[pos];
  return (
    <div
      ref={root}
      className="px-show"
      tabIndex={-1}
      role="dialog"
      aria-modal="true"
      aria-label={atEnd ? 'End of slide show' : `Slide show, slide ${pos + 1} of ${shown.length}`}
      onClick={() => (atEnd ? onClose() : next())}
      onContextMenu={(e) => {
        e.preventDefault();
        prev();
      }}
    >
      {cur && !atEnd ? (
        <SlideView env={{ ...env, mode: 'show', editingId: null }} slide={cur.s} scale={scale} />
      ) : (
        <p className="px-show__end">End of slide show. Click or press Esc to exit.</p>
      )}
    </div>
  );
}
