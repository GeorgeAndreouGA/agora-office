/**
 * Pinax: the PowerPoint editor. Opens, edits and saves .pptx in this tab; nothing is uploaded.
 * See pptx/read.ts (file -> model), render.tsx (model -> screen), pptx/write.ts (model -> file).
 */
import { useCallback, useEffect, useMemo, useRef, useState, type ReactNode } from 'react';
import { ConfirmDialog } from '../components/ConfirmDialog';
import { EditorHeader } from '../components/EditorHeader';
import { choosePdfTarget } from '../lib/fileAccess';
import { useDocSession } from '../lib/session';
import { useToast } from '../lib/toast';
import type { EditorProps } from '../lib/types';
import { Canvas, type Editing } from './Canvas';
import { type FontRegistry, loadFonts } from './fonts';
import { useHistory } from './history';
import { MediaStore } from './media';
import { type Align, type Deck, type DeckState, type RunProps, type ShapeEl, type Slide, type SlideElement, type TextBody, type Xfrm, EMU_PER_PX, isTextual } from './model';
import { SHAPE_CHOICES, copiesOf, deleteElements, duplicateSlide, newPicture, newShape, newSlide, newTextBox, reorder, withFill, withLine, type ZMove } from './ops';
import { blankDeckBytes } from './pptx/blank';
import { DeckError, openDeck } from './pptx/read';
import { type Color, cssColor } from './pptx/theme';
import { writeDeck } from './pptx/write';
import type { RenderEnv } from './render';
import { Slideshow } from './Slideshow';
import { Sorter } from './Sorter';
import type { TextEditorApi, TextFormat } from './TextEditor';
import { runStyle, paraStyle } from './text';
import './pinax.css';

interface Loaded {
  deck: Deck;
  state: DeckState;
  media: MediaStore;
  fonts: FontRegistry;
}

export default function Pinax({ doc, onExit }: EditorProps) {
  const [loaded, setLoaded] = useState<Loaded | null>(null);
  const [failure, setFailure] = useState<string | null>(null);

  useEffect(() => {
    let cancelled = false;
    let opened: Loaded | null = null;
    (async () => {
      try {
        const { deck, state } = openDeck(doc.bytes ?? blankDeckBytes());
        const fonts = await loadFonts(deck.fonts);
        const media = new MediaStore(deck.pkg);
        opened = { deck, state, media, fonts };
        if (cancelled) {
          fonts.dispose();
          media.dispose();
          return;
        }
        setLoaded(opened);
      } catch (err) {
        if (!cancelled) setFailure(err instanceof DeckError ? err.message : 'This presentation could not be opened.');
      }
    })();
    return () => {
      cancelled = true;
      opened?.fonts.dispose();
      opened?.media.dispose();
    };
  }, [doc.bytes]);

  if (failure) {
    return (
      <div className="editor-shell">
        <div className="editor-body">
          <div className="editor-failure" role="alert">
            <h2>Pinax could not open this file</h2>
            <p>{failure}</p>
            <button type="button" className="btn btn--pptx" onClick={onExit}>
              Back to the Greek Office
            </button>
          </div>
        </div>
      </div>
    );
  }
  if (!loaded) {
    return (
      <div className="loading" role="status">
        <span className="loading__mark" aria-hidden="true" />
        <span>Opening the presentation…</span>
      </div>
    );
  }
  return <PinaxEditor doc={doc} onExit={onExit} loaded={loaded} />;
}

const PALETTE_STD = ['000000', 'FFFFFF', 'C00000', 'FF0000', 'FFC000', 'FFFF00', '92D050', '00B050', '00B0F0', '0070C0', '002060', '7030A0'];
const FONT_CHOICES = ['Calibri', 'Cambria', 'Arial', 'Times New Roman', 'Courier New', 'Century Gothic'];
const SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 40, 44, 48, 54, 60, 72, 88];
const IMAGE_TYPES: Record<string, { ext: string; magic: number[] }> = {
  'image/png': { ext: 'png', magic: [0x89, 0x50, 0x4e, 0x47] },
  'image/jpeg': { ext: 'jpeg', magic: [0xff, 0xd8, 0xff] },
  'image/gif': { ext: 'gif', magic: [0x47, 0x49, 0x46, 0x38] },
};
const MAX_IMAGE_BYTES = 10 * 1024 * 1024;

function patchRuns(body: TextBody, patch: RunProps): TextBody {
  return {
    ...body,
    paragraphs: body.paragraphs.map((p) => ({
      ...p,
      items: p.items.map((it) => (it.kind === 'br' ? it : { ...it, props: { ...it.props, ...patch } })),
      endSize: patch.size ?? p.endSize,
    })),
  };
}

function PinaxEditor({ doc, onExit, loaded }: EditorProps & { loaded: Loaded }) {
  const { deck, media, fonts } = loaded;
  const toast = useToast();
  const hist = useHistory<DeckState>(loaded.state);
  const slides = hist.present.slides;
  const stateRef = useRef(hist.present);
  stateRef.current = hist.present;

  const [cur, setCur] = useState(0);
  const [selection, setSelection] = useState<Set<string>>(new Set());
  const [editing, setEditing] = useState<Editing | null>(null);
  const [showFrom, setShowFrom] = useState<number | null>(null);
  const [busy, setBusy] = useState<string | null>(null);
  const [confirmLeave, setConfirmLeave] = useState(false);
  const [notice, setNotice] = useState(doc.notice ?? '');
  const [, setTick] = useState(0);
  const textApi = useRef<TextEditorApi | null>(null);
  const clipboard = useRef<SlideElement[]>([]);
  const pictureInput = useRef<HTMLInputElement>(null);

  const index = Math.min(cur, slides.length - 1);
  const slide = slides[index];

  const env: RenderEnv = useMemo(() => ({ deck, media, fonts, mode: 'edit' }), [deck, media, fonts]);

  const serialize = useCallback(async () => writeDeck(deck, stateRef.current, media), [deck, media]);
  const session = useDocSession({
    kind: 'pptx',
    initialName: doc.name,
    initialHandle: doc.handle,
    serialize,
    onSaved: (how) => toast.show(how === 'autosave' ? 'Autosaved.' : 'Saved.', 'success'),
  });
  const markDirty = session.markDirty;
  const firstVersion = useRef(hist.version);
  useEffect(() => {
    if (hist.version !== firstVersion.current) markDirty();
  }, [hist.version, markDirty]);

  // ---- state helpers ----
  const setSlides = useCallback(
    (next: Slide[], group?: string) => hist.set({ slides: next }, group),
    [hist],
  );
  const setSlide = useCallback(
    (next: Slide, group?: string) => {
      const all = stateRef.current.slides;
      const i = all.findIndex((s) => s.id === next.id);
      if (i < 0) return;
      setSlides(all.map((s, k) => (k === i ? next : s)), group);
    },
    [setSlides],
  );
  const selectedEls = slide ? slide.elements.filter((e) => selection.has(e.id)) : [];

  const goTo = (i: number) => {
    setEditing(null);
    setSelection(new Set());
    // Not clamped here: right after adding a slide the state below has not caught up yet.
    // Rendering clamps (see `index`).
    setCur(Math.max(0, i));
  };

  // ---- slides ----
  const addSlide = (layoutPart?: string) => {
    const preferred =
      layoutPart ??
      deck.layoutOrder.find((p) => /title and content/i.test(deck.layouts.get(p)?.name ?? '')) ??
      (slide?.layout || deck.layoutOrder[0]);
    const s = newSlide(deck, preferred);
    if (!s) return;
    const next = [...slides.slice(0, index + 1), s, ...slides.slice(index + 1)];
    setSlides(next);
    goTo(index + 1);
  };
  const dupSlide = (i: number) => {
    setSlides([...slides.slice(0, i + 1), duplicateSlide(slides[i]), ...slides.slice(i + 1)]);
    goTo(i + 1);
  };
  const delSlide = (i: number) => {
    if (slides.length <= 1) return;
    setSlides(slides.filter((_, k) => k !== i));
    goTo(Math.min(i, slides.length - 2));
  };
  const moveSlide = (from: number, to: number) => {
    const next = [...slides];
    const [s] = next.splice(from, 1);
    next.splice(to, 0, s);
    setSlides(next);
    goTo(to);
  };
  const toggleHidden = (i: number) => setSlides(slides.map((s, k) => (k === i ? { ...s, hidden: !s.hidden, changed: true } : s)));

  // ---- elements ----
  const center = (w: number, h: number): Xfrm => ({
    x: (deck.size.cx - w) / 2,
    y: (deck.size.cy - h) / 2,
    w,
    h,
    rot: 0,
    flipH: false,
    flipV: false,
  });
  const addElement = (el: SlideElement | null, edit = false) => {
    if (!el || !slide) return;
    setSlide({ ...slide, elements: [...slide.elements, el], changed: true });
    setSelection(new Set([el.id]));
    setEditing(edit ? { elId: el.id } : null);
  };
  const insertTextBox = () => addElement(newTextBox(deck, slide, center(deck.size.cx * 0.5, 914400 * 0.6)), true);
  const insertShape = (prst: string) => {
    const line = prst === 'line' || prst === 'arrowLine';
    addElement(newShape(deck, slide, prst, center(deck.size.cx * 0.2, line ? 0 : deck.size.cy * 0.2)));
  };
  const onPicture = async (file: File | undefined) => {
    if (!file || !slide) return;
    try {
      const type = IMAGE_TYPES[file.type];
      if (!type) throw new Error('Only PNG, JPEG and GIF pictures can be inserted.');
      if (file.size > MAX_IMAGE_BYTES) throw new Error('Pictures larger than 10 MB are refused.');
      const bytes = new Uint8Array(await file.arrayBuffer());
      if (!type.magic.every((b, i) => bytes[i] === b)) throw new Error('That file is not the picture its name says.');
      let n = 1;
      let part = `ppt/media/pinax${n}.${type.ext}`;
      while (deck.pkg.has(part) || media.bytes(part)) part = `ppt/media/pinax${++n}.${type.ext}`;
      media.add(part, bytes);
      const bmp = await createImageBitmap(new Blob([bytes as BlobPart], { type: file.type }));
      const maxW = deck.size.cx * 0.6;
      const maxH = deck.size.cy * 0.6;
      let w = bmp.width * EMU_PER_PX;
      let h = bmp.height * EMU_PER_PX;
      bmp.close();
      const k = Math.min(1, maxW / w, maxH / h);
      w *= k;
      h *= k;
      addElement(newPicture(deck, slide, part, center(w, h)));
    } catch (err) {
      toast.show((err as Error).message || 'That picture could not be inserted.', 'error');
    }
  };
  const removeSelected = () => {
    if (!selection.size) return;
    setSlide(deleteElements(slide, selection));
    setSelection(new Set());
    setEditing(null);
  };
  const duplicateSelected = () => {
    if (!selectedEls.length) return;
    const copies = copiesOf(selectedEls, 182880);
    setSlide({ ...slide, elements: [...slide.elements, ...copies], changed: true });
    setSelection(new Set(copies.map((c) => c.id)));
  };
  const arrange = (move: ZMove) => setSlide(reorder(slide, selection, move));
  const nudge = (dx: number, dy: number) =>
    setSlide(
      {
        ...slide,
        elements: slide.elements.map((e) =>
          selection.has(e.id) ? { ...e, xfrm: { ...e.xfrm, x: e.xfrm.x + dx, y: e.xfrm.y + dy }, changed: { ...e.changed, xfrm: true as const } } : e,
        ),
        changed: true,
      },
      'nudge',
    );

  // ---- text formatting (editor open: the selection; otherwise: whole selected shapes) ----
  const textShapes = selectedEls.filter(isTextual).filter((e) => e.text);
  const format: TextFormat | null = (() => {
    if (editing) return textApi.current?.format() ?? null;
    const s = textShapes[0] as ShapeEl | undefined;
    if (!s?.text) return null;
    const p = s.text.paragraphs[0];
    const st = paraStyle(s.text, p);
    const first = p.items.find((i) => i.kind === 'text');
    const r = runStyle(st, first && first.kind === 'text' ? first.props : {});
    return { b: r.b, i: r.i, u: r.u, strike: r.strike, size: r.size, font: r.font, color: r.color, align: st.align, bullet: st.bullet.kind, level: p.level };
  })();
  const applyRun = (patch: RunProps) => {
    if (editing) return textApi.current?.setRun(patch);
    if (!textShapes.length) return;
    setSlide({
      ...slide,
      elements: slide.elements.map((e) => (selection.has(e.id) && isTextual(e) && e.text ? { ...e, text: patchRuns(e.text, patch), changed: { ...e.changed, text: true } } : e)),
      changed: true,
    });
  };
  const toggle = (key: 'b' | 'i' | 'u' | 'strike') => {
    if (editing) return textApi.current?.toggle(key);
    if (format) applyRun({ [key]: !format[key] });
  };
  const applyParas = (fn: (p: TextBody['paragraphs'][number]) => TextBody['paragraphs'][number]) => {
    setSlide({
      ...slide,
      elements: slide.elements.map((e) =>
        selection.has(e.id) && isTextual(e) && e.text ? { ...e, text: { ...e.text, paragraphs: e.text.paragraphs.map(fn) }, changed: { ...e.changed, text: true } } : e,
      ),
      changed: true,
    });
  };
  const setAlign = (a: Align) => (editing ? textApi.current?.setAlign(a) : applyParas((p) => ({ ...p, props: { ...p.props, align: a } })));
  const setBullet = (kind: 'char' | 'num' | 'none') => {
    if (editing) return textApi.current?.setBullet(format?.bullet === kind ? 'none' : kind);
    const next = format?.bullet === kind ? 'none' : kind;
    applyParas((p) => ({
      ...p,
      props: {
        ...p.props,
        bullet:
          next === 'none'
            ? { kind: 'none' }
            : next === 'char'
              ? { kind: 'char', char: '•', font: 'Arial', color: null, size: 1 }
              : { kind: 'num', scheme: 'arabicPeriod', startAt: 1, font: null, color: null, size: 1 },
        ...(next !== 'none' && !p.props.marL ? { marL: 342900, indent: -342900 } : {}),
      },
    }));
  };
  const indent = (d: 1 | -1) => (editing ? textApi.current?.indent(d) : applyParas((p) => ({ ...p, level: Math.max(0, Math.min(8, p.level + d)) })));

  // ---- shape formatting ----
  const fillables = selectedEls.filter((e): e is ShapeEl => e.type === 'shape' && !e.connector);
  const outlinables = selectedEls.filter((e) => e.type === 'shape' || e.type === 'picture');
  const setFillColor = (c: Color | null) =>
    setSlide({
      ...slide,
      elements: slide.elements.map((e) => (selection.has(e.id) && e.type === 'shape' && !e.connector ? withFill(e, c ? { kind: 'solid', color: c } : { kind: 'none' }) : e)),
      changed: true,
    });
  const setOutline = (c: Color | null, width?: number) =>
    setSlide({
      ...slide,
      elements: slide.elements.map((e) => {
        if (!selection.has(e.id) || (e.type !== 'shape' && e.type !== 'picture')) return e;
        const base = e.line ?? { fill: { kind: 'none' as const }, width: 12700, dash: null, head: null, tail: null, cap: 'flat' as const };
        return withLine(e, c === null ? { ...base, fill: { kind: 'none' } } : { ...base, fill: { kind: 'solid', color: c }, width: width ?? (base.width || 12700) });
      }),
      changed: true,
    });

  // ---- keyboard ----
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (showFrom !== null) return;
      const t = e.target as HTMLElement | null;
      if (t && (t.closest('input, textarea, select, [contenteditable="true"]') || t.isContentEditable)) return;
      const mod = e.ctrlKey || e.metaKey;
      const k = e.key;
      if (k === 'F5') {
        e.preventDefault();
        setShowFrom(e.shiftKey ? index : 0);
      } else if (mod && k.toLowerCase() === 'z') {
        e.preventDefault();
        if (e.shiftKey) hist.redo();
        else hist.undo();
      } else if (mod && k.toLowerCase() === 'y') {
        e.preventDefault();
        hist.redo();
      } else if (mod && k.toLowerCase() === 'd') {
        e.preventDefault();
        duplicateSelected();
      } else if (mod && k.toLowerCase() === 'c' && selectedEls.length) {
        clipboard.current = selectedEls;
      } else if (mod && k.toLowerCase() === 'x' && selectedEls.length) {
        clipboard.current = selectedEls;
        removeSelected();
      } else if (mod && k.toLowerCase() === 'v' && clipboard.current.length) {
        e.preventDefault();
        const copies = copiesOf(clipboard.current, clipboard.current.some((c) => slide.elements.some((x) => x.id === c.id)) ? 182880 : 0);
        setSlide({ ...slide, elements: [...slide.elements, ...copies], changed: true });
        setSelection(new Set(copies.map((c) => c.id)));
      } else if (mod && k.toLowerCase() === 'a') {
        e.preventDefault();
        setSelection(new Set(slide.elements.map((x) => x.id)));
      } else if ((k === 'Delete' || k === 'Backspace') && selection.size) {
        e.preventDefault();
        removeSelected();
      } else if (k.startsWith('Arrow') && selection.size) {
        e.preventDefault();
        const step = (e.shiftKey ? 10 : 1) * EMU_PER_PX * 4;
        nudge(k === 'ArrowLeft' ? -step : k === 'ArrowRight' ? step : 0, k === 'ArrowUp' ? -step : k === 'ArrowDown' ? step : 0);
      } else if (k.startsWith('Arrow') || k === 'PageDown' || k === 'PageUp') {
        if (k === 'ArrowDown' || k === 'PageDown' || k === 'ArrowRight') goTo(index + 1);
        else goTo(index - 1);
        e.preventDefault();
      } else if (k === 'Enter' && selectedEls.length === 1 && isTextual(selectedEls[0])) {
        e.preventDefault();
        setEditing({ elId: selectedEls[0].id });
      } else if (k === 'Escape') {
        setSelection(new Set());
      }
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  });

  // ---- PDF ----
  const downloadPdf = async () => {
    if (busy) return;
    let target;
    try {
      target = await choosePdfTarget(session.name);
    } catch (err) {
      toast.show(`Could not save the PDF: ${(err as Error).message}`, 'error');
      return;
    }
    if (!target) return;
    setEditing(null);
    setBusy('Building PDF…');
    try {
      const { presentationToPdf } = await import('./pdf');
      const bytes = await presentationToPdf({
        deck,
        slides: stateRef.current.slides.filter((s) => !s.hidden),
        media,
        fonts,
        title: session.name.replace(/\.pptx$/i, ''),
        onProgress: (d, t) => setBusy(`Building PDF: slide ${d} of ${t}…`),
      });
      await target.write(bytes);
      toast.show('PDF saved.', 'success');
    } catch (err) {
      await target.discard();
      toast.show(`Could not build the PDF: ${(err as Error).message}`, 'error');
    } finally {
      setBusy(null);
    }
  };

  const back = () => (session.dirty ? setConfirmLeave(true) : onExit());
  const notesDisabled = slide?.notes === null && !deck.notesMaster;
  const themeColors = Object.values(slide?.ctx.theme.colors ?? {}).slice(0, 10);

  return (
    <div className="editor-shell pinax">
      <EditorHeader
        kind="pptx"
        session={session}
        onBack={back}
        extraActions={
          <>
            <button type="button" className="btn btn--ghost-dark" onClick={() => setShowFrom(index)}>
              Slideshow
            </button>
            <button type="button" className="btn btn--ghost-dark" onClick={() => void downloadPdf()} disabled={!!busy}>
              Download PDF
            </button>
          </>
        }
      />
      {session.error && (
        <div className="editor-alert" role="alert">
          {session.error}
        </div>
      )}
      {notice && (
        <div className="editor-notice" role="status">
          <span>{notice}</span>
          <button type="button" className="editor-notice__close" onClick={() => setNotice('')} aria-label="Dismiss notice">
            ×
          </button>
        </div>
      )}
      <div className="px-toolbar" role="toolbar" aria-label="Presentation tools" onMouseDown={(e) => editing && e.target instanceof HTMLButtonElement && e.preventDefault()}>
        <Group label="Slides">
          <Menu label="New slide" title="Add a slide after this one">
            {deck.layoutOrder.map((p) => (
              <button key={p} type="button" onClick={() => addSlide(p)}>
                {deck.layouts.get(p)?.name}
              </button>
            ))}
          </Menu>
          <button type="button" className="px-tb" onClick={hist.undo} disabled={!hist.canUndo} title="Undo (Ctrl+Z)" aria-label="Undo">
            ↶
          </button>
          <button type="button" className="px-tb" onClick={hist.redo} disabled={!hist.canRedo} title="Redo (Ctrl+Y)" aria-label="Redo">
            ↷
          </button>
        </Group>
        <Group label="Insert">
          <button type="button" className="px-tb px-tb--text" onClick={insertTextBox}>
            Text box
          </button>
          <Menu label="Shape" title="Insert a shape">
            {SHAPE_CHOICES.map((s) => (
              <button key={s.prst} type="button" onClick={() => insertShape(s.prst)}>
                {s.label}
              </button>
            ))}
          </Menu>
          <button type="button" className="px-tb px-tb--text" onClick={() => pictureInput.current?.click()}>
            Picture
          </button>
          <input
            ref={pictureInput}
            type="file"
            accept="image/png,image/jpeg,image/gif"
            hidden
            onChange={(e) => {
              void onPicture(e.target.files?.[0]);
              e.target.value = '';
            }}
          />
        </Group>
        <Group label="Text">
          <select
            className="px-select"
            aria-label="Font"
            disabled={!format}
            value={format ? (FONT_CHOICES.includes(format.font) ? format.font : '') : ''}
            onChange={(e) => applyRun({ font: e.target.value })}
          >
            <option value="" disabled>
              {format?.font ?? 'Font'}
            </option>
            {FONT_CHOICES.map((f) => (
              <option key={f} value={f}>
                {f}
              </option>
            ))}
          </select>
          <select
            className="px-select px-select--size"
            aria-label="Font size"
            disabled={!format}
            value={format ? String(Math.round(format.size)) : ''}
            onChange={(e) => applyRun({ size: Number(e.target.value) })}
          >
            {!format && (
              <option value="" disabled>
                Size
              </option>
            )}
            {format && !SIZES.includes(Math.round(format.size)) && <option value={String(Math.round(format.size))}>{Math.round(format.size)}</option>}
            {SIZES.map((s) => (
              <option key={s} value={String(s)}>
                {s}
              </option>
            ))}
          </select>
          <Toggle on={!!format?.b} disabled={!format} onClick={() => toggle('b')} label="Bold (Ctrl+B)">
            <b>B</b>
          </Toggle>
          <Toggle on={!!format?.i} disabled={!format} onClick={() => toggle('i')} label="Italic (Ctrl+I)">
            <i>I</i>
          </Toggle>
          <Toggle on={!!format?.u} disabled={!format} onClick={() => toggle('u')} label="Underline (Ctrl+U)">
            <u>U</u>
          </Toggle>
          <Toggle on={!!format?.strike} disabled={!format} onClick={() => toggle('strike')} label="Strikethrough">
            <s>S</s>
          </Toggle>
          <Swatches label="Text colour" disabled={!format} current={format?.color ?? null} theme={themeColors} onPick={(c) => c && applyRun({ color: c })} glyph="A" />
          {(['l', 'ctr', 'r', 'just'] as Align[]).map((a) => (
            <Toggle key={a} on={format?.align === a} disabled={!format} onClick={() => setAlign(a)} label={`Align ${{ l: 'left', ctr: 'centre', r: 'right', just: 'justified', dist: 'distributed' }[a]}`}>
              <AlignIcon a={a} />
            </Toggle>
          ))}
          <Toggle on={format?.bullet === 'char'} disabled={!format} onClick={() => setBullet('char')} label="Bullets">
            •≡
          </Toggle>
          <Toggle on={format?.bullet === 'num'} disabled={!format} onClick={() => setBullet('num')} label="Numbering">
            1≡
          </Toggle>
          <button type="button" className="px-tb" disabled={!format} onClick={() => indent(-1)} title="Decrease list level" aria-label="Decrease list level">
            ⇤
          </button>
          <button type="button" className="px-tb" disabled={!format} onClick={() => indent(1)} title="Increase list level" aria-label="Increase list level">
            ⇥
          </button>
        </Group>
        <Group label="Shape">
          <Swatches label="Fill" disabled={!fillables.length} current={null} theme={themeColors} onPick={setFillColor} allowNone glyph="◼" />
          <Swatches label="Outline" disabled={!outlinables.length} current={null} theme={themeColors} onPick={(c) => setOutline(c)} allowNone glyph="◻" />
          <Menu label="Arrange" title="Order of overlapping objects" disabled={!selection.size}>
            <button type="button" onClick={() => arrange('front')}>
              Bring to front
            </button>
            <button type="button" onClick={() => arrange('forward')}>
              Bring forward
            </button>
            <button type="button" onClick={() => arrange('backward')}>
              Send backward
            </button>
            <button type="button" onClick={() => arrange('back')}>
              Send to back
            </button>
          </Menu>
          <button type="button" className="px-tb" disabled={!selection.size} onClick={removeSelected} title="Delete (Del)" aria-label="Delete selection">
            🗑
          </button>
        </Group>
      </div>

      <div className="editor-body px-body">
        {busy && (
          <div className="print-progress" role="status">
            {busy}
          </div>
        )}
        <Sorter
          env={env}
          slides={slides}
          current={index}
          onSelect={goTo}
          onMove={moveSlide}
          onDuplicate={dupSlide}
          onDelete={delSlide}
          onToggleHidden={toggleHidden}
        />
        <div className="px-main">
          {slide && (
            <Canvas
              env={env}
              slide={slide}
              selection={selection}
              setSelection={setSelection}
              editing={editing}
              setEditing={setEditing}
              onSlide={setSlide}
              textApiRef={textApi}
              onTextSelection={() => setTick((t) => t + 1)}
            />
          )}
          <div className="px-notes">
            <label htmlFor="px-notes-text" className="px-notes__label">
              Speaker notes · slide {index + 1} of {slides.length}
              {slide?.hidden ? ' · hidden in the slideshow' : ''}
            </label>
            <textarea
              id="px-notes-text"
              className="px-notes__text"
              placeholder={notesDisabled ? 'This file has no notes page layout, so notes cannot be added.' : 'Click to add notes'}
              disabled={notesDisabled}
              value={(slide?.notes ?? []).join('\n')}
              onChange={(e) => slide && setSlide({ ...slide, notes: e.target.value.split('\n'), notesChanged: true, changed: true }, `notes-${slide.id}`)}
            />
          </div>
        </div>
      </div>

      {showFrom !== null && <Slideshow env={env} slides={slides} start={showFrom} onClose={() => setShowFrom(null)} />}

      <ConfirmDialog
        open={confirmLeave}
        title="Leave without saving?"
        body="This presentation lives only in this tab. If you go back to the Greek Office now, your unsaved changes are lost."
        confirmLabel="Discard and leave"
        onCancel={() => setConfirmLeave(false)}
        onConfirm={onExit}
      />
    </div>
  );
}

// ---------------------------------------------------------------- small UI pieces

function Group({ label, children }: { label: string; children: ReactNode }) {
  return (
    <div className="px-tb-group" role="group" aria-label={label}>
      {children}
    </div>
  );
}

function Toggle({ on, disabled, onClick, label, children }: { on: boolean; disabled?: boolean; onClick(): void; label: string; children: ReactNode }) {
  return (
    <button type="button" className={`px-tb ${on ? 'px-tb--on' : ''}`} aria-pressed={on} disabled={disabled} onClick={onClick} title={label} aria-label={label}>
      {children}
    </button>
  );
}

function Menu({ label, title, disabled, children }: { label: string; title?: string; disabled?: boolean; children: ReactNode }) {
  const [open, setOpen] = useState(false);
  const ref = useRef<HTMLDivElement>(null);
  useEffect(() => {
    if (!open) return;
    const close = (e: PointerEvent) => {
      if (!ref.current?.contains(e.target as Node)) setOpen(false);
    };
    const esc = (e: KeyboardEvent) => e.key === 'Escape' && setOpen(false);
    window.addEventListener('pointerdown', close);
    window.addEventListener('keydown', esc);
    return () => {
      window.removeEventListener('pointerdown', close);
      window.removeEventListener('keydown', esc);
    };
  }, [open]);
  return (
    <div className="px-menu" ref={ref}>
      <button type="button" className="px-tb px-tb--text" aria-haspopup="menu" aria-expanded={open} disabled={disabled} title={title} onClick={() => setOpen((o) => !o)}>
        {label} ▾
      </button>
      {open && (
        <div className="px-menu__list" role="menu" onClick={() => setOpen(false)}>
          {children}
        </div>
      )}
    </div>
  );
}

function Swatches({
  label,
  disabled,
  current,
  theme,
  onPick,
  allowNone,
  glyph,
}: {
  label: string;
  disabled?: boolean;
  current: Color | null;
  theme: string[];
  onPick(c: Color | null): void;
  allowNone?: boolean;
  glyph: string;
}) {
  const [open, setOpen] = useState(false);
  const ref = useRef<HTMLDivElement>(null);
  useEffect(() => {
    if (!open) return;
    const close = (e: PointerEvent) => {
      if (!ref.current?.contains(e.target as Node)) setOpen(false);
    };
    window.addEventListener('pointerdown', close);
    return () => window.removeEventListener('pointerdown', close);
  }, [open]);
  const pick = (hex: string | null) => {
    onPick(hex ? { hex, alpha: 1 } : null);
    setOpen(false);
  };
  return (
    <div className="px-menu" ref={ref}>
      <button type="button" className="px-tb" disabled={disabled} title={label} aria-label={label} aria-haspopup="dialog" aria-expanded={open} onClick={() => setOpen((o) => !o)}>
        <span className="px-swatch-glyph" style={{ borderBottomColor: current ? cssColor(current) : undefined }}>
          {glyph}
        </span>
      </button>
      {open && (
        <div className="px-menu__list px-swatches" role="dialog" aria-label={label}>
          <div className="px-swatches__row">
            {theme.map((hex) => (
              <button key={`t${hex}`} type="button" className="px-swatch" style={{ background: `#${hex}` }} aria-label={`#${hex}`} title={`#${hex}`} onClick={() => pick(hex)} />
            ))}
          </div>
          <div className="px-swatches__row">
            {PALETTE_STD.map((hex) => (
              <button key={hex} type="button" className="px-swatch" style={{ background: `#${hex}` }} aria-label={`#${hex}`} title={`#${hex}`} onClick={() => pick(hex)} />
            ))}
          </div>
          <div className="px-swatches__row px-swatches__more">
            <label>
              Custom <input type="color" onChange={(e) => pick(e.target.value.slice(1).toUpperCase())} />
            </label>
            {allowNone && (
              <button type="button" className="px-tb px-tb--text" onClick={() => pick(null)}>
                None
              </button>
            )}
          </div>
        </div>
      )}
    </div>
  );
}

function AlignIcon({ a }: { a: Align }) {
  const lines = a === 'l' ? [[2, 14], [2, 10], [2, 14], [2, 8]] : a === 'r' ? [[2, 14], [6, 14], [2, 14], [8, 14]] : a === 'ctr' ? [[2, 14], [4, 12], [2, 14], [5, 11]] : [[2, 14], [2, 14], [2, 14], [2, 14]];
  return (
    <svg width="16" height="16" viewBox="0 0 16 16" aria-hidden="true">
      {lines.map(([x1, x2], i) => (
        <line key={i} x1={x1} x2={x2} y1={3 + i * 3.4} y2={3 + i * 3.4} stroke="currentColor" strokeWidth="1.6" />
      ))}
    </svg>
  );
}
