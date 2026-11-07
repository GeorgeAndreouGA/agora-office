/**
 * In-place rich-text editing for one text box (or table cell), on ProseMirror.
 *
 * The document schema mirrors DrawingML text: paragraphs carrying their level and authored
 * properties (plus the original a:pPr, so unknown settings survive), and one `run` mark carrying
 * a run's authored properties and its original a:rPr. Inherited formatting comes from CSS on the
 * paragraph; a run's own properties override it, exactly as in the static renderer.
 */
import { useEffect, useImperativeHandle, useRef, type CSSProperties, type Ref } from 'react';
import { baseKeymap } from 'prosemirror-commands';
import { history, redo, undo } from 'prosemirror-history';
import { keymap } from 'prosemirror-keymap';
import { type Mark, type Node as PMNode, Schema } from 'prosemirror-model';
import { EditorState, Plugin, TextSelection, type Command, type Transaction } from 'prosemirror-state';
import { Decoration, DecorationSet, EditorView } from 'prosemirror-view';
import 'prosemirror-view/style/prosemirror.css';
import type { FontRegistry } from './fonts';
import { type Align, type Bullet, type InlineItem, type LevelProps, type Paragraph, type RunProps, type TextBody, PX_PER_PT, emuToPx } from './model';
import { cssColor, type Color } from './pptx/theme';
import { ALIGN_CSS, lineHeight, markers, paraStyle, runStyle, spacingPx } from './text';


export interface TextFormat {
  b: boolean;
  i: boolean;
  u: boolean;
  strike: boolean;
  size: number;
  font: string;
  color: Color;
  align: Align;
  bullet: Bullet['kind'];
  level: number;
}

export interface TextEditorApi {
  format(): TextFormat | null;
  setRun(patch: RunProps): void;
  toggle(key: 'b' | 'i' | 'u' | 'strike'): void;
  setAlign(a: Align): void;
  setBullet(kind: 'char' | 'num' | 'none'): void;
  indent(delta: 1 | -1): void;
  focus(): void;
}

function buildSchema(body: TextBody, fonts: FontRegistry, scale: number, reduction: number) {
  return new Schema({
    nodes: {
      doc: { content: 'paragraph+' },
      paragraph: {
        content: 'inline*',
        attrs: { level: { default: 0 }, props: { default: {} }, pPr: { default: null }, endRPr: { default: null }, endSize: { default: null } },
        parseDOM: [{ tag: 'p' }, { tag: 'div' }],
        toDOM(node) {
          const p: Paragraph = { items: [], level: node.attrs.level, props: node.attrs.props, pPr: null, endRPr: null, endSize: node.attrs.endSize };
          const st = paraStyle(body, p);
          const size = st.run.size;
          const css: CSSProperties = {
            textAlign: ALIGN_CSS[st.align] ?? 'left',
            paddingLeft: emuToPx(st.marL),
            textIndent: emuToPx(st.indent),
            marginTop: spacingPx(st.spaceBefore, size, scale),
            marginBottom: spacingPx(st.spaceAfter, size, scale),
            lineHeight: lineHeight(st, scale, reduction, size),
            fontFamily: fonts.css(st.run.font),
            fontSize: `${size * PX_PER_PT * scale}px`,
            fontWeight: st.run.b ? 700 : 400,
            fontStyle: st.run.i ? 'italic' : 'normal',
            color: cssColor(st.run.color),
            textDecoration: [st.run.u ? 'underline' : '', st.run.strike ? 'line-through' : ''].filter(Boolean).join(' ') || 'none',
          };
          return ['div', { class: 'px-p', style: styleString(css) }, 0];
        },
      },
      text: { group: 'inline' },
      hard_break: {
        inline: true,
        group: 'inline',
        selectable: false,
        attrs: { rPr: { default: null } },
        parseDOM: [{ tag: 'br' }],
        toDOM: () => ['br'],
      },
      field: {
        inline: true,
        group: 'inline',
        atom: true,
        attrs: { text: { default: '' }, xml: { default: '' }, props: { default: {} } },
        toDOM: (node) => ['span', { class: 'px-field', contenteditable: 'false' }, node.attrs.text],
      },
    },
    marks: {
      run: {
        attrs: { props: { default: {} }, rPr: { default: null }, link: { default: null } },
        parseDOM: [
          { tag: 'b', getAttrs: () => ({ props: { b: true } }) },
          { tag: 'strong', getAttrs: () => ({ props: { b: true } }) },
          { tag: 'i', getAttrs: () => ({ props: { i: true } }) },
          { tag: 'em', getAttrs: () => ({ props: { i: true } }) },
          { tag: 'u', getAttrs: () => ({ props: { u: true } }) },
        ],
        toDOM(mark) {
          const r: RunProps = mark.attrs.props;
          const css: CSSProperties = {};
          if (r.font) css.fontFamily = fonts.css(r.font);
          if (r.size !== undefined || r.baseline) css.fontSize = r.size !== undefined ? `${r.size * (r.baseline ? 0.66 : 1) * PX_PER_PT * scale}px` : '66%';
          if (r.b !== undefined) css.fontWeight = r.b ? 700 : 400;
          if (r.i !== undefined) css.fontStyle = r.i ? 'italic' : 'normal';
          if (r.color) css.color = cssColor(r.color);
          if (r.u !== undefined || r.strike !== undefined) {
            css.textDecoration = [r.u ? 'underline' : '', r.strike ? 'line-through' : ''].filter(Boolean).join(' ') || 'none';
          }
          if (r.baseline) css.verticalAlign = r.baseline > 0 ? 'super' : 'sub';
          return ['span', { style: styleString(css) }, 0];
        },
      },
    },
  });
}

function styleString(css: CSSProperties): string {
  return Object.entries(css)
    .filter(([, v]) => v !== undefined && v !== '')
    .map(([k, v]) => `${k.replace(/[A-Z]/g, (m) => `-${m.toLowerCase()}`)}:${typeof v === 'number' && !/^(fontWeight|opacity|zIndex)$/.test(k) ? `${v}px` : v}`)
    .join(';');
}

function toDoc(schema: Schema, body: TextBody): PMNode {
  const run = schema.marks.run;
  const paras = body.paragraphs.map((p) => {
    const inline: PMNode[] = [];
    for (const it of p.items) {
      if (it.kind === 'br') inline.push(schema.nodes.hard_break.create({ rPr: it.rPr }));
      else if (it.kind === 'field') inline.push(schema.nodes.field.create({ text: it.text, xml: it.xml, props: it.props }));
      else if (it.text) inline.push(schema.text(it.text, [run.create({ props: it.props, rPr: it.rPr, link: it.link })]));
    }
    return schema.nodes.paragraph.create({ level: p.level, props: p.props, pPr: p.pPr, endRPr: p.endRPr, endSize: p.endSize }, inline);
  });
  return schema.nodes.doc.create(null, paras.length ? paras : [schema.nodes.paragraph.create()]);
}

function fromDoc(doc: PMNode, body: TextBody): TextBody {
  const paragraphs: Paragraph[] = [];
  doc.forEach((pn) => {
    const items: InlineItem[] = [];
    pn.forEach((n) => {
      if (n.type.name === 'hard_break') items.push({ kind: 'br', rPr: n.attrs.rPr });
      else if (n.type.name === 'field') items.push({ kind: 'field', text: n.attrs.text, xml: n.attrs.xml, props: n.attrs.props });
      else if (n.isText) {
        const m = n.marks.find((x) => x.type.name === 'run');
        const props: RunProps = m?.attrs.props ?? {};
        const rPr: string | null = m?.attrs.rPr ?? null;
        const link: string | null = m?.attrs.link ?? null;
        const last = items[items.length - 1];
        if (last && last.kind === 'text' && last.rPr === rPr && last.link === link && JSON.stringify(last.props) === JSON.stringify(props)) last.text += n.text;
        else items.push({ kind: 'text', text: n.text ?? '', props, rPr, link });
      }
    });
    paragraphs.push({ items, level: pn.attrs.level, props: pn.attrs.props as LevelProps, pPr: pn.attrs.pPr, endRPr: pn.attrs.endRPr, endSize: pn.attrs.endSize });
  });
  return { ...body, paragraphs };
}

/** Enter: split the paragraph, the new one keeping level and properties. */
const splitParagraph: Command = (state, dispatch) => {
  const { $from } = state.selection;
  if (dispatch) {
    const tr = state.tr;
    if (!state.selection.empty) tr.deleteSelection();
    const pos = tr.selection.$from;
    const marks = state.storedMarks ?? pos.marks();
    tr.split(pos.pos, 1, [{ type: $from.parent.type, attrs: { ...$from.parent.attrs } }]);
    tr.ensureMarks(marks);
    dispatch(tr.scrollIntoView());
  }
  return true;
};

const hardBreak: Command = (state, dispatch) => {
  if (dispatch) dispatch(state.tr.replaceSelectionWith(state.schema.nodes.hard_break.create()).scrollIntoView());
  return true;
};

function paragraphsInSelection(state: EditorState, fn: (attrs: Record<string, unknown>) => Record<string, unknown>): Transaction {
  const tr = state.tr;
  const { from, to } = state.selection;
  state.doc.nodesBetween(from, to, (node, pos) => {
    if (node.type.name === 'paragraph') {
      tr.setNodeMarkup(pos, undefined, fn({ ...node.attrs }));
      return false;
    }
    return true;
  });
  return tr;
}

function runPatch(state: EditorState, patch: RunProps): Transaction {
  const run = state.schema.marks.run;
  const tr = state.tr;
  const { from, to, empty } = state.selection;
  const apply = (m: Mark | undefined) => run.create({ ...(m?.attrs ?? {}), props: { ...(m?.attrs.props ?? {}), ...patch } });
  if (empty) {
    const cur = (state.storedMarks ?? state.selection.$from.marks()).find((m) => m.type === run);
    tr.addStoredMark(apply(cur));
    return tr;
  }
  state.doc.nodesBetween(from, to, (node, pos) => {
    if (!node.isText) return true;
    const start = Math.max(from, pos);
    const end = Math.min(to, pos + node.nodeSize);
    const cur = node.marks.find((m) => m.type === run);
    tr.removeMark(start, end, run);
    tr.addMark(start, end, apply(cur));
    return false;
  });
  return tr;
}

export function TextEditor({
  body,
  fonts,
  style,
  onChange,
  onDone,
  apiRef,
  onSelectionChange,
}: {
  body: TextBody;
  fonts: FontRegistry;
  style?: CSSProperties;
  onChange(body: TextBody): void;
  onDone(): void;
  apiRef: Ref<TextEditorApi | null>;
  onSelectionChange?(): void;
}) {
  const host = useRef<HTMLDivElement>(null);
  const view = useRef<EditorView | null>(null);
  const bodyRef = useRef(body);
  const cbs = useRef({ onChange, onDone, onSelectionChange });
  cbs.current = { onChange, onDone, onSelectionChange };

  const scale = body.body.autofit === 'norm' ? body.body.fontScale : 1;
  const reduction = body.body.autofit === 'norm' ? body.body.lnSpcReduction : 0;

  useEffect(() => {
    if (!host.current) return;
    const initial = bodyRef.current;
    const schema = buildSchema(initial, fonts, scale, reduction);
    const indent = (delta: number): Command => (state, dispatch) => {
      dispatch?.(paragraphsInSelection(state, (a) => ({ ...a, level: Math.min(8, Math.max(0, Number(a.level) + delta)) })));
      return true;
    };
    const markerPlugin = new Plugin({
      props: {
        decorations(state) {
          const current = fromDoc(state.doc, bodyRef.current);
          const marks = markers(current);
          const decos: Decoration[] = [];
          let i = 0;
          state.doc.forEach((_node, offset) => {
            const m = marks[i];
            const p = current.paragraphs[i];
            i++;
            if (!m || !p) return;
            const st = paraStyle(current, p);
            const first = runStyle(st, p.items.find((x) => x.kind === 'text')?.kind === 'text' ? (p.items.find((x) => x.kind === 'text') as { props: RunProps }).props : {});
            decos.push(
              Decoration.widget(
                offset + 1,
                () => {
                  const span = document.createElement('span');
                  span.className = 'px-bullet';
                  span.textContent = m;
                  span.contentEditable = 'false';
                  const bullet = st.bullet.kind === 'none' ? null : st.bullet;
                  span.style.cssText = styleString({
                    display: 'inline-block',
                    minWidth: st.indent < 0 ? emuToPx(-st.indent) : undefined,
                    paddingRight: st.indent >= 0 ? '0.4em' : undefined,
                    textIndent: 0,
                    fontFamily: fonts.css(first.font),
                    fontSize: `${first.size * (bullet?.size ?? 1) * PX_PER_PT * scale}px`,
                    color: cssColor(bullet?.color ?? first.color),
                  });
                  return span;
                },
                { side: -1, key: `${i}:${m}` },
              ),
            );
          });
          return DecorationSet.create(state.doc, decos);
        },
      },
    });
    const state = EditorState.create({
      doc: toDoc(schema, initial),
      plugins: [
        history(),
        keymap({
          'Mod-z': undo,
          'Mod-y': redo,
          'Shift-Mod-z': redo,
          Enter: splitParagraph,
          'Shift-Enter': hardBreak,
          Tab: indent(1),
          'Shift-Tab': indent(-1),
          'Mod-b': (s, d) => (d?.(runPatch(s, { b: !formatOf(s, initial).b })), true),
          'Mod-i': (s, d) => (d?.(runPatch(s, { i: !formatOf(s, initial).i })), true),
          'Mod-u': (s, d) => (d?.(runPatch(s, { u: !formatOf(s, initial).u })), true),
          Escape: () => (cbs.current.onDone(), true),
        }),
        keymap(baseKeymap),
        markerPlugin,
      ],
    });
    // Select everything in a fresh box, the end otherwise.
    const sel = TextSelection.atEnd(state.doc);
    const v = new EditorView(host.current, {
      state: state.apply(state.tr.setSelection(sel)),
      dispatchTransaction(tr) {
        const next = v.state.apply(tr);
        v.updateState(next);
        if (tr.docChanged) {
          const nb = fromDoc(next.doc, bodyRef.current);
          bodyRef.current = nb;
          cbs.current.onChange(nb);
        }
        cbs.current.onSelectionChange?.();
      },
      attributes: { class: 'px-flow px-editing', spellcheck: 'true' },
    });
    view.current = v;
    v.focus();
    return () => {
      v.destroy();
      view.current = null;
    };
    // The editor is created once per opening; later prop changes come from the editor itself.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  useImperativeHandle(
    apiRef,
    (): TextEditorApi => ({
      format() {
        const v = view.current;
        return v ? formatOf(v.state, bodyRef.current) : null;
      },
      setRun(patch) {
        const v = view.current;
        if (!v) return;
        v.dispatch(runPatch(v.state, patch));
        v.focus();
      },
      toggle(key) {
        const v = view.current;
        if (!v) return;
        const f = formatOf(v.state, bodyRef.current);
        v.dispatch(runPatch(v.state, { [key]: !f[key] }));
        v.focus();
      },
      setAlign(a) {
        const v = view.current;
        if (!v) return;
        v.dispatch(paragraphsInSelection(v.state, (attrs) => ({ ...attrs, props: { ...(attrs.props as object), align: a } })));
        v.focus();
      },
      setBullet(kind) {
        const v = view.current;
        if (!v) return;
        const bullet: Bullet =
          kind === 'none'
            ? { kind: 'none' }
            : kind === 'char'
              ? { kind: 'char', char: '•', font: 'Arial', color: null, size: 1 }
              : { kind: 'num', scheme: 'arabicPeriod', startAt: 1, font: null, color: null, size: 1 };
        v.dispatch(
          paragraphsInSelection(v.state, (attrs) => {
            const props = { ...(attrs.props as LevelProps), bullet };
            if (kind !== 'none' && (props.marL === undefined || props.marL === 0)) {
              props.marL = 342900;
              props.indent = -342900;
            }
            if (kind === 'none') {
              props.indent = 0;
            }
            return { ...attrs, props };
          }),
        );
        v.focus();
      },
      indent(delta) {
        const v = view.current;
        if (!v) return;
        v.dispatch(paragraphsInSelection(v.state, (a) => ({ ...a, level: Math.min(8, Math.max(0, Number(a.level) + delta)) })));
        v.focus();
      },
      focus() {
        view.current?.focus();
      },
    }),
    [],
  );

  return <div ref={host} className="px-text-editor" style={style} onMouseDown={(e) => e.stopPropagation()} />;
}

/** The effective format at the selection start. */
function formatOf(state: EditorState, body: TextBody): TextFormat {
  const { $from } = state.selection;
  // Select-all puts the selection's start at document level; read the first paragraph then.
  const pn = $from.parent.type.name === 'paragraph' ? $from.parent : (state.doc.firstChild ?? $from.parent);
  const p: Paragraph = { items: [], level: pn.attrs.level ?? 0, props: pn.attrs.props ?? {}, pPr: null, endRPr: null, endSize: null };
  const st = paraStyle(body, p);
  const run = (state.storedMarks ?? $from.marks()).find((m) => m.type.name === 'run') ?? $from.nodeAfter?.marks.find((m) => m.type.name === 'run');
  const r = runStyle(st, run?.attrs.props ?? {});
  return { b: r.b, i: r.i, u: r.u, strike: r.strike, size: r.size, font: r.font, color: r.color, align: st.align, bullet: st.bullet.kind, level: p.level };
}
