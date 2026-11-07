/**
 * Undo/redo over immutable snapshots. Consecutive changes with the same `group` key (typing,
 * dragging) collapse into one step, so undo goes back a gesture, not a keystroke.
 */
import { useCallback, useRef, useState } from 'react';

const LIMIT = 100;

export interface History<T> {
  present: T;
  set(next: T, group?: string): void;
  undo(): void;
  redo(): void;
  canUndo: boolean;
  canRedo: boolean;
  /** Bumps on every change (for the dirty flag). */
  version: number;
}

export function useHistory<T>(initial: T): History<T> {
  const [state, setState] = useState<{ past: T[]; present: T; future: T[]; group: string | null; version: number }>({
    past: [],
    present: initial,
    future: [],
    group: null,
    version: 0,
  });
  const ref = useRef(state);
  ref.current = state;

  const set = useCallback((next: T, group?: string) => {
    setState((s) => {
      if (next === s.present) return s;
      const merge = !!group && group === s.group;
      return {
        past: merge ? s.past : [...s.past, s.present].slice(-LIMIT),
        present: next,
        future: [],
        group: group ?? null,
        version: s.version + 1,
      };
    });
  }, []);

  const undo = useCallback(() => {
    setState((s) =>
      s.past.length
        ? { past: s.past.slice(0, -1), present: s.past[s.past.length - 1], future: [s.present, ...s.future], group: null, version: s.version + 1 }
        : s,
    );
  }, []);

  const redo = useCallback(() => {
    setState((s) =>
      s.future.length
        ? { past: [...s.past, s.present], present: s.future[0], future: s.future.slice(1), group: null, version: s.version + 1 }
        : s,
    );
  }, []);

  return {
    present: state.present,
    set,
    undo,
    redo,
    canUndo: state.past.length > 0,
    canRedo: state.future.length > 0,
    version: state.version,
  };
}
