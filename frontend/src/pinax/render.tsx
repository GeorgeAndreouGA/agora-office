/**
 * Slide rendering: one DOM/SVG renderer for the editing canvas, the thumbnails, the slideshow
 * and the PDF exporter (which measures the rendered text). Units are CSS px at 96 dpi
 * (1 px = 9525 EMU); callers scale the whole slide with a transform.
 */
import { useId, type CSSProperties, type ReactNode } from 'react';
import type { FontRegistry } from './fonts';
import type { MediaStore } from './media';
import {
  type Deck,
  type FrameEl,
  type GroupEl,
  type Paragraph,
  type PictureEl,
  type ShapeEl,
  type Slide,
  type SlideElement,
  type TableEl,
  type TextBody,
  type Xfrm,
  PX_PER_PT,
  emuToPx,
} from './model';
import { arrowHead, dashPattern, endPoints, geometryPaths, isLinePreset, toSvg } from './pptx/geometry';
import { type Fill, type Line, cssColor, fillColor, fillCss } from './pptx/theme';
import { ALIGN_CSS, isEmptyBody, lineHeight, markers, paraStyle, paragraphSize, runStyle, spacingPx } from './text';

export type RenderMode = 'edit' | 'show' | 'thumb' | 'pdf';

export interface RenderEnv {
  deck: Deck;
  media: MediaStore;
  fonts: FontRegistry;
  mode: RenderMode;
  /** The element whose text is open in the editor (its static text is hidden). */
  editingId?: string | null;
  /** PDF measuring pass: no rotation, so text boxes measure upright. */
  flat?: boolean;
}

export function slidePx(deck: Deck) {
  return { w: emuToPx(deck.size.cx), h: emuToPx(deck.size.cy) };
}

function boxStyle(x: Xfrm, env: RenderEnv): CSSProperties {
  return {
    position: 'absolute',
    left: emuToPx(x.x),
    top: emuToPx(x.y),
    width: emuToPx(x.w),
    height: emuToPx(x.h),
    transform: !env.flat && x.rot ? `rotate(${x.rot}deg)` : undefined,
  };
}

function mediaOfFill(fill: Fill | null, env: RenderEnv): string | null {
  if (!fill || fill.kind !== 'image') return null;
  const rel = env.deck.pkg.rel(fill.owner, fill.rId);
  return rel && !rel.external ? env.media.url(rel.target) : null;
}

export function backgroundCss(fill: Fill | null, env: RenderEnv): CSSProperties {
  if (!fill) return { background: '#fff' };
  if (fill.kind === 'image') {
    const url = mediaOfFill(fill, env);
    return url ? { backgroundColor: '#fff', backgroundImage: `url("${url}")`, backgroundSize: '100% 100%' } : { background: '#fff' };
  }
  const css = fillCss(fill);
  return { background: css === 'transparent' ? '#fff' : css };
}

/** The whole slide at `scale`, with an optional overlay (selection handles) inside it. */
export function SlideView({
  env,
  slide,
  scale,
  children,
  className,
}: {
  env: RenderEnv;
  slide: Slide;
  scale: number;
  children?: ReactNode;
  className?: string;
}) {
  const { w, h } = slidePx(env.deck);
  const layout = env.deck.layouts.get(slide.layout);
  const master = layout ? env.deck.masters.get(layout.master) : undefined;
  const bg = slide.background ?? layout?.background ?? master?.background ?? null;
  const showMaster = slide.showMasterShapes && layout?.showMasterShapes !== false;
  return (
    <div className={`px-slide-frame ${className ?? ''}`} style={{ width: w * scale, height: h * scale }}>
      <div
        className="px-slide"
        data-slide={slide.id}
        style={{ width: w, height: h, transform: `scale(${scale})`, transformOrigin: '0 0', ...backgroundCss(bg, env) }}
      >
        {showMaster && master?.decorations.map((el) => <ElementView key={el.id} el={el} env={env} />)}
        {slide.showMasterShapes && layout?.decorations.map((el) => <ElementView key={el.id} el={el} env={env} />)}
        {slide.elements.map((el) => (
          <ElementView key={el.id} el={el} env={env} top />
        ))}
        {children}
      </div>
    </div>
  );
}

export function ElementView({ el, env, top }: { el: SlideElement; env: RenderEnv; top?: boolean }) {
  switch (el.type) {
    case 'shape':
      return <ShapeView el={el} env={env} top={top} />;
    case 'picture':
      return <PictureView el={el} env={env} top={top} />;
    case 'group':
      return <GroupView el={el} env={env} top={top} />;
    case 'table':
      return <TableView el={el} env={env} top={top} />;
    case 'frame':
      return <FrameView el={el} env={env} top={top} />;
  }
}

// ---------------------------------------------------------------- paint

interface Paint {
  value: string;
  defs: ReactNode;
}

function svgPaint(fill: Fill | null, id: string, env: RenderEnv, w: number, h: number): Paint {
  if (!fill || fill.kind === 'none') return { value: 'none', defs: null };
  if (fill.kind === 'solid') return { value: cssColor(fill.color), defs: null };
  if (fill.kind === 'pattern') return { value: cssColor(fill.fg), defs: null };
  if (fill.kind === 'gradient') {
    const stops = fill.stops.map((s, i) => (
      <stop key={i} offset={s.pos} stopColor={`#${s.color.hex}`} stopOpacity={s.color.alpha} />
    ));
    return {
      value: `url(#${id})`,
      defs:
        fill.path === 'radial' ? (
          <radialGradient id={id}>{stops}</radialGradient>
        ) : (
          <linearGradient id={id} gradientTransform={`rotate(${fill.angle}, 0.5, 0.5)`}>
            {stops}
          </linearGradient>
        ),
    };
  }
  const url = mediaOfFill(fill, env);
  if (!url) return { value: 'none', defs: null };
  return {
    value: `url(#${id})`,
    defs: (
      <pattern id={id} patternUnits="userSpaceOnUse" width={w} height={h}>
        <image href={url} width={w} height={h} preserveAspectRatio="none" />
      </pattern>
    ),
  };
}

function strokeOf(line: Line | null) {
  if (!line || line.fill.kind === 'none') return null;
  const c = line.fill.kind === 'solid' ? line.fill.color : fillColor(line.fill);
  if (!c) return null;
  const width = Math.max(emuToPx(line.width), 0.75);
  return {
    color: cssColor(c),
    width,
    dash: dashPattern(line.dash)?.map((d) => d * width).join(' '),
    cap: line.cap === 'rnd' ? ('round' as const) : line.cap === 'sq' ? ('square' as const) : ('butt' as const),
  };
}

// ---------------------------------------------------------------- shapes

function ShapeView({ el, env, top }: { el: ShapeEl; env: RenderEnv; top?: boolean }) {
  const id = useId().replace(/[^a-zA-Z0-9]/g, '');
  const w = emuToPx(el.xfrm.w);
  const h = emuToPx(el.xfrm.h);
  const paths = geometryPaths(el.geom, w, h);
  const lineOnly = el.connector || isLinePreset(el.geom);
  const paint = svgPaint(lineOnly ? null : el.fill, `f${id}`, env, w, h);
  const stroke = strokeOf(el.line);
  const ends = stroke && el.line && (el.line.head || el.line.tail) ? endPoints(paths) : null;
  const head = ends && el.line ? arrowHead(el.line.head, ends.start, ends.second, stroke!.width) : null;
  const tail = ends && el.line ? arrowHead(el.line.tail, ends.end, ends.beforeEnd, stroke!.width) : null;
  const flip = el.xfrm.flipH || el.xfrm.flipV ? `scale(${el.xfrm.flipH ? -1 : 1}, ${el.xfrm.flipV ? -1 : 1})` : undefined;
  const drawGeom = paint.value !== 'none' || !!stroke;
  return (
    <div className="px-el px-shape" data-el={el.id} data-top={top ? el.id : undefined} style={boxStyle(el.xfrm, env)}>
      {drawGeom && (
        <svg className="px-geom" width={Math.max(w, 1)} height={Math.max(h, 1)} style={{ transform: flip }} aria-hidden="true">
          {paint.defs && <defs>{paint.defs}</defs>}
          {paths.map((p, i) => (
            <path
              key={i}
              d={toSvg(p.cmds)}
              fill={p.fill ? paint.value : 'none'}
              fillRule={p.evenOdd ? 'evenodd' : undefined}
              stroke={p.stroke && stroke ? stroke.color : 'none'}
              strokeWidth={stroke?.width}
              strokeDasharray={stroke?.dash}
              strokeLinecap={stroke?.cap}
              strokeLinejoin="round"
            />
          ))}
          {head && <path d={toSvg(head)} fill={stroke!.color} />}
          {tail && <path d={toSvg(tail)} fill={stroke!.color} />}
        </svg>
      )}
      {el.text && (
        <TextBodyView body={el.text} env={env} prompt={el.ph?.prompt ?? null} hidden={env.editingId === el.id} />
      )}
    </div>
  );
}

// ---------------------------------------------------------------- text

export function TextBodyView({
  body,
  env,
  prompt,
  hidden,
  flow,
}: {
  body: TextBody;
  env: RenderEnv;
  prompt?: string | null;
  hidden?: boolean;
  /** Table cells: in normal flow instead of absolutely inset. */
  flow?: boolean;
}) {
  const b = body.body;
  const scale = b.autofit === 'norm' ? b.fontScale : 1;
  const reduction = b.autofit === 'norm' ? b.lnSpcReduction : 0;
  const showPrompt = !!prompt && env.mode === 'edit' && isEmptyBody(body);
  const marks = markers(body);
  const vertical = b.vert === 'vert' || b.vert === 'eaVert' || b.vert === 'vert270';
  const outer: CSSProperties = flow
    ? { position: 'relative' }
    : {
        position: 'absolute',
        left: emuToPx(b.lIns),
        top: emuToPx(b.tIns),
        right: emuToPx(b.rIns),
        bottom: emuToPx(b.bIns),
        display: 'flex',
        flexDirection: 'column',
        justifyContent: b.anchor === 'ctr' ? 'center' : b.anchor === 'b' ? 'flex-end' : 'flex-start',
      };
  return (
    <div className="px-text" style={{ ...outer, visibility: hidden ? 'hidden' : undefined }}>
      <div
        className="px-flow"
        style={{
          whiteSpace: b.wrap ? 'pre-wrap' : 'pre',
          overflowWrap: 'break-word',
          writingMode: vertical ? 'vertical-rl' : undefined,
          transform: b.vert === 'vert270' ? 'rotate(180deg)' : undefined,
        }}
      >
        {showPrompt ? (
          <ParagraphView body={body} p={{ ...body.paragraphs[0], items: [{ kind: 'text', text: prompt!, props: {}, rPr: null, link: null }] }} marker={null} env={env} scale={scale} reduction={reduction} prompt />
        ) : (
          body.paragraphs.map((p, i) => (
            <ParagraphView key={i} body={body} p={p} marker={marks[i]} env={env} scale={scale} reduction={reduction} />
          ))
        )}
      </div>
    </div>
  );
}

function ParagraphView({
  body,
  p,
  marker,
  env,
  scale,
  reduction,
  prompt,
}: {
  body: TextBody;
  p: Paragraph;
  marker: string | null;
  env: RenderEnv;
  scale: number;
  reduction: number;
  prompt?: boolean;
}) {
  const st = paraStyle(body, p);
  const size = paragraphSize(st, p);
  const first = p.items.find((i) => i.kind !== 'br');
  const firstRun = runStyle(st, first && 'props' in first ? first.props : {});
  const bullet = st.bullet.kind === 'none' ? null : st.bullet;
  return (
    <div
      className="px-p"
      style={{
        textAlign: ALIGN_CSS[st.align] ?? 'left',
        paddingLeft: emuToPx(st.marL),
        textIndent: emuToPx(st.indent),
        marginTop: spacingPx(st.spaceBefore, size, scale),
        marginBottom: spacingPx(st.spaceAfter, size, scale),
        lineHeight: lineHeight(st, scale, reduction, size),
        opacity: prompt ? 0.45 : undefined,
      }}
    >
      {marker && bullet && (
        <span
          className="px-bullet"
          style={{
            display: 'inline-block',
            minWidth: st.indent < 0 ? emuToPx(-st.indent) : undefined,
            paddingRight: st.indent >= 0 ? '0.4em' : undefined,
            textIndent: 0,
            fontFamily: env.fonts.css(firstRun.font),
            fontSize: `${firstRun.size * bullet.size * PX_PER_PT * scale}px`,
            fontWeight: bullet.kind === 'num' && firstRun.b ? 700 : 400,
            color: cssColor(bullet.color ?? firstRun.color),
          }}
        >
          {marker}
        </span>
      )}
      {p.items.map((it, i) => {
        if (it.kind === 'br') return <br key={i} />;
        const r = runStyle(st, it.props);
        return (
          <span key={i} style={runCss(r, env, scale)}>
            {it.text}
          </span>
        );
      })}
      {p.items.every((it) => it.kind === 'br' || !it.text) && (
        <span style={{ fontSize: `${size * PX_PER_PT * scale}px`, fontFamily: env.fonts.css(st.run.font) }}>{'​'}</span>
      )}
    </div>
  );
}

export function runCss(r: ReturnType<typeof runStyle>, env: RenderEnv, scale: number): CSSProperties {
  const deco = [r.u ? 'underline' : '', r.strike ? 'line-through' : ''].filter(Boolean).join(' ');
  return {
    fontFamily: env.fonts.css(r.font),
    fontSize: `${r.size * (r.baseline ? 0.66 : 1) * PX_PER_PT * scale}px`,
    fontWeight: r.b ? 700 : 400,
    fontStyle: r.i ? 'italic' : 'normal',
    textDecoration: deco || undefined,
    color: cssColor(r.color),
    verticalAlign: r.baseline > 0 ? 'super' : r.baseline < 0 ? 'sub' : undefined,
  };
}

// ---------------------------------------------------------------- pictures

function PictureView({ el, env, top }: { el: PictureEl; env: RenderEnv; top?: boolean }) {
  const url = env.media.url(el.media);
  const { l, t, r, b } = el.crop;
  const vw = Math.max(0.01, 1 - l - r);
  const vh = Math.max(0.01, 1 - t - b);
  const w = emuToPx(el.xfrm.w);
  const h = emuToPx(el.xfrm.h);
  const paths = geometryPaths(el.geom, w, h);
  const clip = el.geom.kind === 'preset' && el.geom.name === 'rect' ? undefined : `path('${toSvg(paths[0]?.cmds ?? [])}')`;
  const stroke = strokeOf(el.line);
  const flip = el.xfrm.flipH || el.xfrm.flipV ? `scale(${el.xfrm.flipH ? -1 : 1}, ${el.xfrm.flipV ? -1 : 1})` : undefined;
  return (
    <div className="px-el px-picture" data-el={el.id} data-top={top ? el.id : undefined} style={boxStyle(el.xfrm, env)}>
      <div className="px-picture__clip" style={{ position: 'absolute', inset: 0, overflow: 'hidden', clipPath: clip }}>
        {url ? (
          <img
            src={url}
            alt=""
            draggable={false}
            style={{
              position: 'absolute',
              left: `${(-l / vw) * 100}%`,
              top: `${(-t / vh) * 100}%`,
              width: `${100 / vw}%`,
              height: `${100 / vh}%`,
              transform: flip,
              maxWidth: 'none',
            }}
          />
        ) : (
          env.mode !== 'pdf' && (
            <div className="px-missing">
              <span>{el.media ? 'Picture format not shown' : 'Picture'}</span>
            </div>
          )
        )}
      </div>
      {stroke && (
        <svg className="px-geom" width={Math.max(w, 1)} height={Math.max(h, 1)} aria-hidden="true">
          <path d={toSvg(paths[0]?.cmds ?? [])} fill="none" stroke={stroke.color} strokeWidth={stroke.width} strokeDasharray={stroke.dash} />
        </svg>
      )}
      {el.mediaKind && env.mode !== 'pdf' && (
        <span className="px-play" aria-hidden="true">
          {el.mediaKind === 'audio' ? '♪' : '▶'}
        </span>
      )}
    </div>
  );
}

// ---------------------------------------------------------------- groups

export function mapIntoGroup(g: GroupEl, x: Xfrm): Xfrm {
  const sx = g.child.w ? g.xfrm.w / g.child.w : 1;
  const sy = g.child.h ? g.xfrm.h / g.child.h : 1;
  return { ...x, x: (x.x - g.child.x) * sx, y: (x.y - g.child.y) * sy, w: x.w * sx, h: x.h * sy };
}

function GroupView({ el, env, top }: { el: GroupEl; env: RenderEnv; top?: boolean }) {
  return (
    <div className="px-el px-group" data-el={el.id} data-top={top ? el.id : undefined} style={boxStyle(el.xfrm, env)}>
      {el.children.map((c) => (
        <ElementView key={c.id} el={{ ...c, xfrm: mapIntoGroup(el, c.xfrm) } as SlideElement} env={env} />
      ))}
    </div>
  );
}

// ---------------------------------------------------------------- tables

function borderCss(line: Line | null): string | undefined {
  const s = strokeOf(line);
  return s ? `${s.width}px ${s.dash ? 'dashed' : 'solid'} ${s.color}` : undefined;
}

function TableView({ el, env, top }: { el: TableEl; env: RenderEnv; top?: boolean }) {
  const width = el.cols.reduce((a, c) => a + emuToPx(c), 0);
  return (
    <div className="px-el px-table" data-el={el.id} data-top={top ? el.id : undefined} style={{ ...boxStyle(el.xfrm, env), height: 'auto' }}>
      <table style={{ width, borderCollapse: 'collapse', tableLayout: 'fixed' }}>
        <colgroup>
          {el.cols.map((c, i) => (
            <col key={i} style={{ width: emuToPx(c) }} />
          ))}
        </colgroup>
        <tbody>
          {el.rows.map((row, ri) => (
            <tr key={ri} style={{ height: emuToPx(row.h) }}>
              {row.cells.map((cell, ci) =>
                cell.merged ? null : (
                  <td
                    key={ci}
                    colSpan={cell.colSpan > 1 ? cell.colSpan : undefined}
                    rowSpan={cell.rowSpan > 1 ? cell.rowSpan : undefined}
                    data-cell={`${ri},${ci}`}
                    style={{
                      background: cell.fill ? fillCss(cell.fill) : undefined,
                      borderLeft: borderCss(cell.borders.l),
                      borderRight: borderCss(cell.borders.r),
                      borderTop: borderCss(cell.borders.t),
                      borderBottom: borderCss(cell.borders.b),
                      padding: `${emuToPx(cell.margins.t)}px ${emuToPx(cell.margins.r)}px ${emuToPx(cell.margins.b)}px ${emuToPx(cell.margins.l)}px`,
                      verticalAlign: cell.anchor === 'ctr' ? 'middle' : cell.anchor === 'b' ? 'bottom' : 'top',
                    }}
                  >
                    <TextBodyView body={cell.text} env={env} flow hidden={env.editingId === `${el.id}:${ri},${ci}`} />
                  </td>
                ),
              )}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

// ---------------------------------------------------------------- charts, SmartArt, objects

function FrameView({ el, env, top }: { el: FrameEl; env: RenderEnv; top?: boolean }) {
  const url = env.media.url(el.preview);
  return (
    <div className="px-el px-frame" data-el={el.id} data-top={top ? el.id : undefined} style={boxStyle(el.xfrm, env)}>
      {url ? (
        <img src={url} alt="" draggable={false} style={{ width: '100%', height: '100%', objectFit: 'fill' }} />
      ) : (
        <div className="px-frame__ph" data-kind={el.kind}>
          <span className="px-frame__label">{el.label}</span>
          {env.mode === 'edit' && <span className="px-frame__note">Kept as is · not editable here</span>}
        </div>
      )}
    </div>
  );
}
