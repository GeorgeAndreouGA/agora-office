/**
 * Pinax's in-memory model of a presentation.
 *
 * Plain immutable data, so undo/redo is a stack of snapshots. Every element that came from the
 * file keeps a reference to its ORIGINAL XML node (never mutated) and the part whose
 * relationships that node's r:ids point into. Saving clones that node and patches only what
 * changed, which is how everything Pinax does not model (animations, effects, extensions...)
 * survives a save untouched.
 */
import type { Color, ColorContext, Fill, Line, Theme } from './pptx/theme';
import type { Pkg } from './pptx/package';

export const EMU_PER_PX = 9525; // 96 dpi
export const EMU_PER_PT = 12700;
export const PX_PER_PT = 96 / 72;

export interface Xfrm {
  x: number;
  y: number;
  w: number;
  h: number;
  /** Degrees, clockwise. */
  rot: number;
  flipH: boolean;
  flipV: boolean;
}

// ---------------------------------------------------------------- text

export type Align = 'l' | 'ctr' | 'r' | 'just' | 'dist';

export interface Spacing {
  /** Percent of the line (1 = single) ... */
  pct?: number;
  /** ... or exact points. */
  pts?: number;
}

export type Bullet =
  | { kind: 'char'; char: string; font: string | null; color: Color | null; size: number }
  | { kind: 'num'; scheme: string; startAt: number; font: string | null; color: Color | null; size: number }
  | { kind: 'none' };

/** Run formatting. In an authored run every field is optional; resolved runs have them all. */
export interface RunStyle {
  size: number; // points
  b: boolean;
  i: boolean;
  u: boolean;
  strike: boolean;
  color: Color;
  /** Font family as the file names it (theme references already resolved). */
  font: string;
  /** Percent; +30 superscript, -25 subscript. */
  baseline: number;
}

export type RunProps = Partial<Omit<RunStyle, 'color'>> & { color?: Color };

/** One level of a text style (lvl1pPr .. lvl9pPr), fully resolved. */
export interface LevelStyle {
  marL: number; // EMU
  indent: number; // EMU
  align: Align;
  lineSpacing: Spacing;
  spaceBefore: Spacing;
  spaceAfter: Spacing;
  bullet: Bullet;
  run: RunStyle;
}

export type LevelProps = Partial<Omit<LevelStyle, 'run'>> & { run?: RunProps };

export type InlineItem =
  | { kind: 'text'; text: string; props: RunProps; rPr: string | null; link: string | null }
  | { kind: 'br'; rPr: string | null }
  | { kind: 'field'; text: string; xml: string; props: RunProps };

export interface Paragraph {
  items: InlineItem[];
  level: number; // 0..8
  props: LevelProps; // authored paragraph properties
  /** The original a:pPr, kept so unknown properties survive an edit. */
  pPr: string | null;
  endRPr: string | null;
  /** Size of the paragraph mark, for empty paragraphs. */
  endSize: number | null;
}

export interface BodyProps {
  lIns: number;
  tIns: number;
  rIns: number;
  bIns: number;
  anchor: 't' | 'ctr' | 'b';
  wrap: boolean;
  autofit: 'none' | 'norm' | 'shape';
  fontScale: number; // 1 = 100%
  lnSpcReduction: number; // 0.1 = 10%
  vert: string | null;
}

export interface TextBody {
  paragraphs: Paragraph[];
  body: BodyProps;
  /** Resolved base style per level, cascade included. */
  levels: LevelStyle[];
}

// ---------------------------------------------------------------- elements

export interface Placeholder {
  type: string; // title, body, ctrTitle, subTitle, dt, ftr, sldNum, pic, obj, ...
  idx: number;
  /** Prompt text shown while it is empty ("Click to add title"). */
  prompt: string;
}

export type Geometry =
  | { kind: 'preset'; name: string; adj: Record<string, number> }
  | { kind: 'custom'; paths: CustomPath[] };

export interface CustomPath {
  w: number;
  h: number;
  fill: boolean;
  stroke: boolean;
  /** Commands in path units: M x y, L x y, C x1 y1 x2 y2 x y, Q x1 y1 x y, A wR hR st sw, Z. */
  cmds: (string | number)[][];
}

export interface Changes {
  xfrm?: true;
  text?: true;
  fill?: true;
  line?: true;
  cells?: true;
}

interface ElementBase {
  /** Stable key for React and selection. */
  id: string;
  name: string;
  xfrm: Xfrm;
  /** The original XML node (read-only), or null for elements made in Pinax. */
  xml: Element | null;
  /** Part whose relationships the xml's r:ids refer to. */
  owner: string;
  changed: Changes;
}

export interface ShapeEl extends ElementBase {
  type: 'shape';
  geom: Geometry;
  fill: Fill | null;
  line: Line | null;
  text: TextBody | null;
  ph: Placeholder | null;
  /** A text box (txBox="1"): empty text means nothing to draw. */
  textBox: boolean;
  connector: boolean;
}

export interface PictureEl extends ElementBase {
  type: 'picture';
  /** Part name of the image, or null when it cannot be shown. */
  media: string | null;
  crop: { l: number; t: number; r: number; b: number };
  geom: Geometry;
  line: Line | null;
  ph: Placeholder | null;
  /** Video / audio poster frames get a play badge. */
  mediaKind: 'video' | 'audio' | null;
}

export interface GroupEl extends ElementBase {
  type: 'group';
  children: SlideElement[];
  /** The child coordinate space (chOff / chExt). */
  child: { x: number; y: number; w: number; h: number };
}

export interface TableCell {
  text: TextBody;
  fill: Fill | null;
  borders: { l: Line | null; r: Line | null; t: Line | null; b: Line | null };
  colSpan: number;
  rowSpan: number;
  /** Covered by a neighbour's span (hMerge / vMerge). */
  merged: boolean;
  margins: { l: number; r: number; t: number; b: number };
  anchor: 't' | 'ctr' | 'b';
  /** Text was edited. */
  changed?: true;
}

export interface TableEl extends ElementBase {
  type: 'table';
  cols: number[];
  rows: { h: number; cells: TableCell[] }[];
}

export interface FrameEl extends ElementBase {
  type: 'frame';
  kind: 'chart' | 'diagram' | 'ole' | 'other';
  label: string;
  /** A saved picture of the content, when the file has one. */
  preview: string | null;
}

export type SlideElement = ShapeEl | PictureEl | GroupEl | TableEl | FrameEl;

// ---------------------------------------------------------------- slides

export interface Context {
  color: ColorContext;
  theme: Theme;
}

export interface Slide {
  id: string;
  /** Source slide part and its parsed XML, or null for a slide made in Pinax. */
  part: string | null;
  doc: Document | null;
  layout: string;
  elements: SlideElement[];
  /** The slide's own background, or null to inherit the layout's. */
  background: Fill | null;
  hidden: boolean;
  /** Speaker notes as plain paragraphs; null when the slide has none. */
  notes: string[] | null;
  notesChanged?: true;
  ctx: Context;
  /** Draw the layout and master's decorations (logos, rules...). */
  showMasterShapes: boolean;
  changed?: true;
}

export interface Layout {
  part: string;
  name: string;
  master: string;
  /** Non-placeholder decorations, drawn under the slide. */
  decorations: SlideElement[];
  placeholders: PlaceholderDef[];
  background: Fill | null;
  ctx: Context;
  showMasterShapes: boolean;
}

export interface PlaceholderDef {
  type: string;
  idx: number;
  xfrm: Xfrm | null;
  /** The p:sp element on the layout/master. */
  sp: Element;
  owner: string;
}

export interface Master {
  part: string;
  decorations: SlideElement[];
  placeholders: PlaceholderDef[];
  background: Fill | null;
  ctx: Context;
  titleStyle: Element | null;
  bodyStyle: Element | null;
  otherStyle: Element | null;
}

export interface Deck {
  pkg: Pkg;
  presPart: string;
  size: { cx: number; cy: number };
  layouts: Map<string, Layout>;
  masters: Map<string, Master>;
  /** Layout order for the "New slide" menu. */
  layoutOrder: string[];
  defaultTextStyle: Element | null;
  notesMaster: string | null;
  tableStyles: Document | null;
  /** Every font family the deck names (for font loading). */
  fonts: Set<string>;
}

export interface DeckState {
  slides: Slide[];
}

let counter = 0;
export function newId(prefix = 'e'): string {
  counter += 1;
  return `${prefix}${Date.now().toString(36)}${counter.toString(36)}`;
}

export function emuToPx(v: number): number {
  return v / EMU_PER_PX;
}

export function isTextual(el: SlideElement): el is ShapeEl {
  return el.type === 'shape' && !el.connector;
}
