/** The slide list on the left: thumbnails, drag to reorder, quick actions. */
import { useState } from 'react';
import type { Slide } from './model';
import { SlideView, slidePx, type RenderEnv } from './render';

interface Props {
  env: RenderEnv;
  slides: Slide[];
  current: number;
  onSelect(i: number): void;
  onMove(from: number, to: number): void;
  onDuplicate(i: number): void;
  onDelete(i: number): void;
  onToggleHidden(i: number): void;
}

const THUMB_W = 176;

export function Sorter({ env, slides, current, onSelect, onMove, onDuplicate, onDelete, onToggleHidden }: Props) {
  const [dragFrom, setDragFrom] = useState<number | null>(null);
  const [dropAt, setDropAt] = useState<number | null>(null);
  const scale = THUMB_W / slidePx(env.deck).w;
  const thumbEnv: RenderEnv = { ...env, mode: 'thumb', editingId: null };
  return (
    <nav className="px-sorter" aria-label="Slides">
      <ol className="px-sorter__list">
        {slides.map((s, i) => (
          <li
            key={s.id}
            className={`px-thumb ${i === current ? 'px-thumb--current' : ''} ${s.hidden ? 'px-thumb--hidden' : ''} ${dropAt === i ? 'px-thumb--drop' : ''}`}
            draggable
            onDragStart={(e) => {
              setDragFrom(i);
              e.dataTransfer.effectAllowed = 'move';
              e.dataTransfer.setData('text/x-pinax-slide', String(i));
            }}
            onDragOver={(e) => {
              if (dragFrom === null) return;
              e.preventDefault();
              e.dataTransfer.dropEffect = 'move';
              setDropAt(i);
            }}
            onDragLeave={() => setDropAt((d) => (d === i ? null : d))}
            onDrop={(e) => {
              e.preventDefault();
              if (dragFrom !== null && dragFrom !== i) onMove(dragFrom, i);
              setDragFrom(null);
              setDropAt(null);
            }}
            onDragEnd={() => {
              setDragFrom(null);
              setDropAt(null);
            }}
          >
            <span className="px-thumb__num" aria-hidden="true">
              {i + 1}
            </span>
            <button
              type="button"
              className="px-thumb__btn"
              onClick={() => onSelect(i)}
              aria-current={i === current ? 'true' : undefined}
              aria-label={`Slide ${i + 1}${s.hidden ? ' (hidden)' : ''}`}
            >
              <SlideView env={thumbEnv} slide={s} scale={scale} />
            </button>
            <span className="px-thumb__tools">
              <button type="button" title="Duplicate slide" aria-label={`Duplicate slide ${i + 1}`} onClick={() => onDuplicate(i)}>
                ⧉
              </button>
              <button type="button" title={s.hidden ? 'Show slide' : 'Hide slide'} aria-label={`${s.hidden ? 'Show' : 'Hide'} slide ${i + 1}`} onClick={() => onToggleHidden(i)}>
                {s.hidden ? '◐' : '◑'}
              </button>
              <button type="button" title="Delete slide" aria-label={`Delete slide ${i + 1}`} onClick={() => onDelete(i)} disabled={slides.length <= 1}>
                ✕
              </button>
            </span>
          </li>
        ))}
      </ol>
    </nav>
  );
}
