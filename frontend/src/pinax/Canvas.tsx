/**
 * The editing canvas: the current slide fitted to the available space, with selection,
 * move / resize / rotate handles, snapping, and the in-place text editor.
 */
import { useEffect, useLayoutEffect, useRef, useState, type CSSProperties, type Ref } from 'react';
import { type ShapeEl, type Slide, type SlideElement, type TableEl, type TextBody, type Xfrm, EMU_PER_PX, emuToPx, isTextual } from './model';
import { withXfrm } from './ops';
import { SlideView, slidePx, type RenderEnv } from './render';
import { TextEditor, type TextEditorApi } from './TextEditor';

export interface Editing {
  elId: string;
  /** Table cell being edited (row, col). */
  cell?: [number, number];
}

type Handle = 'n' | 's' | 'e' | 'w' | 'ne' | 'nw' | 'se' | 'sw' | 'rot';

interface Props {
  env: RenderEnv;
  slide: Slide;
  selection: ReadonlySet<string>;
  setSelection(next: Set<string>): void;
  editing: Editing | null;
  setEditing(next: Editing | null): void;
  /** Replace the slide; `group` merges a gesture into one undo step. */
  onSlide(next: Slide, group?: string): void;
  textApiRef: Ref<TextEditorApi | null>;
  onTextSelection(): void;
}

const SNAP_PX = 6;

export function Canvas({ env, slide, selection, setSelection, editing, setEditing, onSlide, textApiRef, onTextSelection }: Props) {
  const wrap = useRef<HTMLDivElement>(null);
  const [scale, setScale] = useState(0.5);
  const { w: slideW, h: slideH } = slidePx(env.deck);
  const slideRef = useRef(slide);
  slideRef.current = slide;

  useLayoutEffect(() => {
    const el = wrap.current;
    if (!el) return;
    const fit = () => {
      const s = Math.min((el.clientWidth - 48) / slideW, (el.clientHeight - 48) / slideH);
      setScale(Math.max(0.1, Math.min(s, 2.5)));
    };
    fit();
    const ro = new ResizeObserver(fit);
    ro.observe(el);
    return () => ro.disconnect();
  }, [slideW, slideH]);

  const gesture = useRef(0);

  /** Pointer position in slide px. */
  const toSlide = (e: { clientX: number; clientY: number }) => {
    const r = wrap.current!.querySelector('.px-slide')!.getBoundingClientRect();
    return { x: (e.clientX - r.left) / scale, y: (e.clientY - r.top) / scale };
  };

  const startMove = (e: React.PointerEvent, id: string) => {
    const sel = e.shiftKey ? new Set(selection) : selection.has(id) ? new Set(selection) : new Set<string>();
    if (e.shiftKey && sel.has(id)) {
      sel.delete(id);
      setSelection(sel);
      return;
    }
    sel.add(id);
    setSelection(sel);
    const start = toSlide(e);
    const orig = new Map(slide.elements.filter((x) => sel.has(x.id)).map((x) => [x.id, x.xfrm]));
    const group = `move-${++gesture.current}`;
    let moved = false;
    const others = slide.elements.filter((x) => !sel.has(x.id)).map((x) => x.xfrm);
    const onMove = (ev: PointerEvent) => {
      const p = toSlide(ev);
      let dx = p.x - start.x;
      let dy = p.y - start.y;
      if (!moved && Math.hypot(dx, dy) < 3) return;
      moved = true;
      // Snap the first selected element's edges and centre to the slide and other elements.
      const first = orig.values().next().value as Xfrm | undefined;
      if (first && !ev.altKey) {
        const fx = emuToPx(first.x) + dx;
        const fy = emuToPx(first.y) + dy;
        const fw = emuToPx(first.w);
        const fh = emuToPx(first.h);
        const xs = [0, slideW / 2, slideW, ...others.flatMap((o) => [emuToPx(o.x), emuToPx(o.x + o.w / 2), emuToPx(o.x + o.w)])];
        const ys = [0, slideH / 2, slideH, ...others.flatMap((o) => [emuToPx(o.y), emuToPx(o.y + o.h / 2), emuToPx(o.y + o.h)])];
        const snap = (pos: number[], targets: number[]) => {
          let best: number | null = null;
          for (const p0 of pos) for (const t of targets) if (Math.abs(t - p0) * scale < SNAP_PX && (best === null || Math.abs(t - p0) < Math.abs(best))) best = t - p0;
          return best ?? 0;
        };
        dx += snap([fx, fx + fw / 2, fx + fw], xs);
        dy += snap([fy, fy + fh / 2, fy + fh], ys);
      }
      const cur = slideRef.current;
      onSlide(
        {
          ...cur,
          elements: cur.elements.map((x) => {
            const o = orig.get(x.id);
            return o ? withXfrm(x, { ...o, x: o.x + dx * EMU_PER_PX, y: o.y + dy * EMU_PER_PX }) : x;
          }),
          changed: true,
        },
        group,
      );
    };
    const onUp = () => {
      window.removeEventListener('pointermove', onMove);
      window.removeEventListener('pointerup', onUp);
    };
    window.addEventListener('pointermove', onMove);
    window.addEventListener('pointerup', onUp);
  };

  const startHandle = (e: React.PointerEvent, el: SlideElement, handle: Handle) => {
    e.stopPropagation();
    e.preventDefault();
    const o = el.xfrm;
    const cx = emuToPx(o.x + o.w / 2);
    const cy = emuToPx(o.y + o.h / 2);
    const start = toSlide(e);
    const rad = (o.rot * Math.PI) / 180;
    const group = `resize-${++gesture.current}`;
    const keepRatio = el.type === 'picture';
    const onMove = (ev: PointerEvent) => {
      const p = toSlide(ev);
      let next: Xfrm;
      if (handle === 'rot') {
        let deg = (Math.atan2(p.y - cy, p.x - cx) * 180) / Math.PI + 90;
        if (ev.shiftKey) deg = Math.round(deg / 15) * 15;
        else if (Math.abs(deg % 90) < 3 || Math.abs(deg % 90) > 87) deg = Math.round(deg / 90) * 90;
        next = { ...o, rot: ((deg % 360) + 360) % 360 };
      } else {
        // Work in the element's own (unrotated) frame, centred on its middle.
        const dx = p.x - start.x;
        const dy = p.y - start.y;
        const lx = dx * Math.cos(-rad) - dy * Math.sin(-rad);
        const ly = dx * Math.sin(-rad) + dy * Math.cos(-rad);
        const w = emuToPx(o.w);
        const h = emuToPx(o.h);
        let l = -w / 2;
        let r = w / 2;
        let t = -h / 2;
        let b = h / 2;
        if (handle.includes('w')) l += lx;
        if (handle.includes('e')) r += lx;
        if (handle.includes('n')) t += ly;
        if (handle.includes('s')) b += ly;
        if ((keepRatio || ev.shiftKey) && handle.length === 2 && w > 0 && h > 0) {
          const ratio = h / w;
          const nw = Math.max(4, r - l);
          const nh = nw * ratio;
          if (handle.includes('n')) t = b - nh;
          else b = t + nh;
        }
        const minSize = 4;
        if (r - l < minSize) {
          if (handle.includes('w')) l = r - minSize;
          else r = l + minSize;
        }
        if (b - t < minSize) {
          if (handle.includes('n')) t = b - minSize;
          else b = t + minSize;
        }
        const ncx = (l + r) / 2;
        const ncy = (t + b) / 2;
        const gx = cx + ncx * Math.cos(rad) - ncy * Math.sin(rad);
        const gy = cy + ncx * Math.sin(rad) + ncy * Math.cos(rad);
        const nw = r - l;
        const nh = b - t;
        next = { ...o, x: (gx - nw / 2) * EMU_PER_PX, y: (gy - nh / 2) * EMU_PER_PX, w: nw * EMU_PER_PX, h: nh * EMU_PER_PX };
      }
      const cur = slideRef.current;
      onSlide({ ...cur, elements: cur.elements.map((x) => (x.id === el.id ? withXfrm(x, next) : x)), changed: true }, group);
    };
    const onUp = () => {
      window.removeEventListener('pointermove', onMove);
      window.removeEventListener('pointerup', onUp);
    };
    window.addEventListener('pointermove', onMove);
    window.addEventListener('pointerup', onUp);
  };

  const onPointerDown = (e: React.PointerEvent) => {
    if (e.button !== 0) return;
    if ((e.target as HTMLElement).closest('.px-text-host')) return; // typing in the open text box
    const target = (e.target as HTMLElement).closest('[data-top]');
    if (!target) {
      if (editing) setEditing(null);
      setSelection(new Set());
      return;
    }
    const id = target.getAttribute('data-top')!;
    if (editing && editing.elId === id) return;
    if (editing) setEditing(null);
    startMove(e, id);
  };

  const onDoubleClick = (e: React.MouseEvent) => {
    const target = (e.target as HTMLElement).closest('[data-top]');
    if (!target) return;
    const id = target.getAttribute('data-top')!;
    const el = slide.elements.find((x) => x.id === id);
    if (!el) return;
    if (el.type === 'table') {
      const td = (e.target as HTMLElement).closest('td[data-cell]');
      const [r, c] = (td?.getAttribute('data-cell') ?? '0,0').split(',').map(Number);
      setSelection(new Set([id]));
      setEditing({ elId: id, cell: [r, c] });
      return;
    }
    if (isTextual(el)) {
      setSelection(new Set([id]));
      setEditing({ elId: id });
    }
  };

  // Escape leaves text editing; clicking outside the slide does too.
  useEffect(() => {
    if (!editing) return;
    if (!slide.elements.some((x) => x.id === editing.elId)) setEditing(null);
  }, [editing, slide, setEditing]);

  const selected = slide.elements.filter((x) => selection.has(x.id));
  const editingEl = editing ? slide.elements.find((x) => x.id === editing.elId) : undefined;

  return (
    <div className="px-canvas" ref={wrap} onPointerDown={onPointerDown} onDoubleClick={onDoubleClick}>
      <SlideView env={{ ...env, editingId: editing ? (editing.cell ? `${editing.elId}:${editing.cell[0]},${editing.cell[1]}` : editing.elId) : null }} slide={slide} scale={scale}>
        {selected.map((el) => (
          <SelectionBox key={el.id} el={el} scale={scale} single={selected.length === 1 && !editing} onHandle={(e, h) => startHandle(e, el, h)} />
        ))}
        {editingEl && (
          <TextOverlay
            key={`${editingEl.id}:${editing?.cell?.join(',') ?? ''}`}
            el={editingEl}
            cell={editing?.cell}
            env={env}
            scale={scale}
            onSlide={onSlide}
            slideRef={slideRef}
            onDone={() => setEditing(null)}
            textApiRef={textApiRef}
            onTextSelection={onTextSelection}
          />
        )}
      </SlideView>
    </div>
  );
}

function SelectionBox({ el, scale, single, onHandle }: { el: SlideElement; scale: number; single: boolean; onHandle(e: React.PointerEvent, h: Handle): void }) {
  const x = el.xfrm;
  const hs = 9 / scale;
  const style: CSSProperties = {
    left: emuToPx(x.x),
    top: emuToPx(x.y),
    width: emuToPx(x.w),
    height: el.type === 'table' ? undefined : emuToPx(x.h),
    transform: x.rot ? `rotate(${x.rot}deg)` : undefined,
    borderWidth: 1.5 / scale,
  };
  const handles: Handle[] = ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'];
  return (
    <div className="px-sel" style={style}>
      {single &&
        handles.map((h) => (
          <span
            key={h}
            className={`px-handle px-handle--${h}`}
            style={{ width: hs, height: hs, margin: -hs / 2, borderWidth: 1 / scale }}
            onPointerDown={(e) => onHandle(e, h)}
          />
        ))}
      {single && el.type !== 'table' && (
        <span className="px-handle px-handle--rot" style={{ width: hs * 1.2, height: hs * 1.2, marginLeft: -hs * 0.6, top: -28 / scale, borderWidth: 1 / scale }} onPointerDown={(e) => onHandle(e, 'rot')} title="Rotate" />
      )}
    </div>
  );
}

function TextOverlay({
  el,
  cell,
  env,
  scale,
  onSlide,
  slideRef,
  onDone,
  textApiRef,
  onTextSelection,
}: {
  el: SlideElement;
  cell?: [number, number];
  env: RenderEnv;
  scale: number;
  onSlide(next: Slide, group?: string): void;
  slideRef: { current: Slide };
  onDone(): void;
  textApiRef: Ref<TextEditorApi | null>;
  onTextSelection(): void;
}) {
  const group = useRef(`text-${el.id}-${Date.now()}`).current;
  const [cellBox, setCellBox] = useState<CSSProperties | null>(null);

  useLayoutEffect(() => {
    if (!cell) return;
    const slideEl = document.querySelector(`.px-canvas .px-slide`);
    const td = slideEl?.querySelector(`[data-top="${el.id}"] td[data-cell="${cell[0]},${cell[1]}"]`) as HTMLElement | null;
    if (!slideEl || !td) return;
    const sr = slideEl.getBoundingClientRect();
    const r = td.getBoundingClientRect();
    const cs = getComputedStyle(td);
    setCellBox({
      left: (r.left - sr.left) / scale,
      top: (r.top - sr.top) / scale,
      width: r.width / scale,
      minHeight: r.height / scale,
      paddingLeft: cs.paddingLeft,
      paddingRight: cs.paddingRight,
      paddingTop: cs.paddingTop,
      paddingBottom: cs.paddingBottom,
    });
  }, [cell, el.id, scale]);

  let body: TextBody | null = null;
  if (el.type === 'table' && cell) body = (el as TableEl).rows[cell[0]]?.cells[cell[1]]?.text ?? null;
  else if (el.type === 'shape') body = (el as ShapeEl).text ?? null;
  if (!body) return null;

  const onChange = (next: TextBody) => {
    const cur = slideRef.current;
    onSlide(
      {
        ...cur,
        elements: cur.elements.map((x) => {
          if (x.id !== el.id) return x;
          if (x.type === 'table' && cell) {
            const rows = x.rows.map((row, ri) =>
              ri !== cell[0] ? row : { ...row, cells: row.cells.map((c, ci) => (ci !== cell[1] ? c : { ...c, text: next, changed: true as const })) },
            );
            return { ...x, rows, changed: { ...x.changed, cells: true } };
          }
          if (x.type === 'shape') return { ...x, text: next, changed: { ...x.changed, text: true } };
          return x;
        }),
        changed: true,
      },
      group,
    );
  };

  if (cell) {
    if (!cellBox) return null;
    return (
      <div className="px-text-host" style={{ position: 'absolute', ...cellBox }}>
        <TextEditor body={body} fonts={env.fonts} onChange={onChange} onDone={onDone} apiRef={textApiRef} onSelectionChange={onTextSelection} />
      </div>
    );
  }

  const x = el.xfrm;
  const b = body.body;
  return (
    <div
      className="px-text-host"
      style={{ position: 'absolute', left: emuToPx(x.x), top: emuToPx(x.y), width: emuToPx(x.w), height: emuToPx(x.h), transform: x.rot ? `rotate(${x.rot}deg)` : undefined }}
    >
      <div
        style={{
          position: 'absolute',
          left: emuToPx(b.lIns),
          top: emuToPx(b.tIns),
          right: emuToPx(b.rIns),
          bottom: emuToPx(b.bIns),
          display: 'flex',
          flexDirection: 'column',
          justifyContent: b.anchor === 'ctr' ? 'center' : b.anchor === 'b' ? 'flex-end' : 'flex-start',
        }}
      >
        <TextEditor
          body={body}
          fonts={env.fonts}
          style={{ whiteSpace: b.wrap ? 'pre-wrap' : 'pre' }}
          onChange={onChange}
          onDone={onDone}
          apiRef={textApiRef}
          onSelectionChange={onTextSelection}
        />
      </div>
    </div>
  );
}
