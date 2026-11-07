/**
 * Shape geometry -> path commands, in the shape's own box (0..w, 0..h, any unit).
 *
 * The same commands feed the SVG on screen and the PDF exporter, so the two cannot disagree.
 * Arcs are converted to cubic Béziers here. About thirty common presets are drawn exactly;
 * any other preset draws as its bounding rectangle (its XML is still saved untouched).
 */
import type { CustomPath, Geometry } from '../model';

export type Cmd = ['M', number, number] | ['L', number, number] | ['C', number, number, number, number, number, number] | ['Z'];

export interface SubPath {
  cmds: Cmd[];
  fill: boolean;
  stroke: boolean;
  evenOdd?: boolean;
}

const K = 0.5522847498; // cubic approximation of a quarter circle

function rectPath(x: number, y: number, w: number, h: number): Cmd[] {
  return [['M', x, y], ['L', x + w, y], ['L', x + w, y + h], ['L', x, y + h], ['Z']];
}

function ellipsePath(cx: number, cy: number, rx: number, ry: number): Cmd[] {
  return [
    ['M', cx + rx, cy],
    ['C', cx + rx, cy + ry * K, cx + rx * K, cy + ry, cx, cy + ry],
    ['C', cx - rx * K, cy + ry, cx - rx, cy + ry * K, cx - rx, cy],
    ['C', cx - rx, cy - ry * K, cx - rx * K, cy - ry, cx, cy - ry],
    ['C', cx + rx * K, cy - ry, cx + rx, cy - ry * K, cx + rx, cy],
    ['Z'],
  ];
}

/** Rounded corners: radii per corner (tl, tr, br, bl). */
function roundedRect(w: number, h: number, [tl, tr, br, bl]: number[]): Cmd[] {
  return [
    ['M', tl, 0],
    ['L', w - tr, 0],
    ['C', w - tr + tr * K, 0, w, tr - tr * K, w, tr],
    ['L', w, h - br],
    ['C', w, h - br + br * K, w - br + br * K, h, w - br, h],
    ['L', bl, h],
    ['C', bl - bl * K, h, 0, h - bl + bl * K, 0, h - bl],
    ['L', 0, tl],
    ['C', 0, tl - tl * K, tl - tl * K, 0, tl, 0],
    ['Z'],
  ];
}

function poly(points: number[][]): Cmd[] {
  const cmds: Cmd[] = points.map(([x, y], i) => (i === 0 ? ['M', x, y] : ['L', x, y]));
  cmds.push(['Z']);
  return cmds;
}

function star(w: number, h: number, n: number, inner: number): Cmd[] {
  const pts: number[][] = [];
  for (let i = 0; i < n * 2; i++) {
    const r = i % 2 === 0 ? 1 : inner;
    const a = -Math.PI / 2 + (i * Math.PI) / n;
    pts.push([w / 2 + (Math.cos(a) * r * w) / 2, h / 2 + (Math.sin(a) * r * h) / 2]);
  }
  return poly(pts);
}

/** Arc from the current point, PowerPoint style (radii, start and sweep angles in degrees). */
function arcTo(cx0: number, cy0: number, wR: number, hR: number, stDeg: number, swDeg: number): { cmds: Cmd[]; x: number; y: number } {
  const st = (stDeg * Math.PI) / 180;
  const sw = (swDeg * Math.PI) / 180;
  const cx = cx0 - wR * Math.cos(st);
  const cy = cy0 - hR * Math.sin(st);
  const segs = Math.max(1, Math.ceil(Math.abs(sw) / (Math.PI / 2)));
  const step = sw / segs;
  const cmds: Cmd[] = [];
  let a = st;
  for (let i = 0; i < segs; i++) {
    const b = a + step;
    const t = (4 / 3) * Math.tan((b - a) / 4);
    const x1 = cx + wR * (Math.cos(a) - t * Math.sin(a));
    const y1 = cy + hR * (Math.sin(a) + t * Math.cos(a));
    const x2 = cx + wR * (Math.cos(b) + t * Math.sin(b));
    const y2 = cy + hR * (Math.sin(b) - t * Math.cos(b));
    cmds.push(['C', x1, y1, x2, y2, cx + wR * Math.cos(b), cy + hR * Math.sin(b)]);
    a = b;
  }
  return { cmds, x: cx + wR * Math.cos(st + sw), y: cy + hR * Math.sin(st + sw) };
}

function adjOf(adj: Record<string, number>, name: string, fallback: number): number {
  const v = adj[name];
  return Number.isFinite(v) ? v : fallback;
}

/** Presets drawn as lines: no fill. */
const LINE_PRESETS = new Set(['line', 'straightConnector1', 'bentConnector2', 'bentConnector3', 'bentConnector4', 'curvedConnector2', 'curvedConnector3', 'arc', 'leftBracket', 'rightBracket', 'leftBrace', 'rightBrace', 'bracketPair', 'bracePair']);

export function isLinePreset(g: Geometry): boolean {
  return g.kind === 'preset' && LINE_PRESETS.has(g.name);
}

export function presetPaths(name: string, adj: Record<string, number>, w: number, h: number): SubPath[] {
  const ss = Math.min(w, h);
  const a = (n: string, d: number) => adjOf(adj, n, d) / 100000;
  const filled = (cmds: Cmd[], evenOdd = false): SubPath[] => [{ cmds, fill: true, stroke: true, evenOdd }];
  const stroked = (cmds: Cmd[]): SubPath[] => [{ cmds, fill: false, stroke: true }];
  switch (name) {
    case 'rect':
    case 'flowChartProcess':
    case 'flowChartPredefinedProcess':
    case 'flowChartInternalStorage':
    case 'actionButtonBlank':
      return filled(rectPath(0, 0, w, h));
    case 'roundRect':
    case 'flowChartAlternateProcess': {
      const r = Math.min(ss * a('adj', 16667), ss / 2);
      return filled(roundedRect(w, h, [r, r, r, r]));
    }
    case 'round1Rect': {
      const r = ss * a('adj', 16667);
      return filled(roundedRect(w, h, [0, r, 0, 0]));
    }
    case 'round2SameRect': {
      const r1 = ss * a('adj1', 16667);
      const r2 = ss * a('adj2', 0);
      return filled(roundedRect(w, h, [r1, r1, r2, r2]));
    }
    case 'round2DiagRect': {
      const r1 = ss * a('adj1', 16667);
      const r2 = ss * a('adj2', 0);
      return filled(roundedRect(w, h, [r1, r2, r1, r2]));
    }
    case 'snip1Rect': {
      const s = ss * a('adj', 16667);
      return filled(poly([[0, 0], [w - s, 0], [w, s], [w, h], [0, h]]));
    }
    case 'snip2SameRect': {
      const s = ss * a('adj1', 16667);
      return filled(poly([[s, 0], [w - s, 0], [w, s], [w, h], [0, h], [0, s]]));
    }
    case 'ellipse':
    case 'flowChartConnector':
      return filled(ellipsePath(w / 2, h / 2, w / 2, h / 2));
    case 'flowChartTerminator': {
      const r = Math.min(w, h) / 2;
      return filled(roundedRect(w, h, [r, r, r, r]));
    }
    case 'triangle':
    case 'flowChartExtract':
      return filled(poly([[w * a('adj', 50000), 0], [w, h], [0, h]]));
    case 'flowChartMerge':
      return filled(poly([[0, 0], [w, 0], [w / 2, h]]));
    case 'rtTriangle':
      return filled(poly([[0, 0], [w, h], [0, h]]));
    case 'diamond':
    case 'flowChartDecision':
      return filled(poly([[w / 2, 0], [w, h / 2], [w / 2, h], [0, h / 2]]));
    case 'parallelogram':
    case 'flowChartInputOutput': {
      const o = name === 'parallelogram' ? ss * a('adj', 25000) : w / 5;
      return filled(poly([[o, 0], [w, 0], [w - o, h], [0, h]]));
    }
    case 'trapezoid': {
      const o = ss * a('adj', 25000);
      return filled(poly([[0, h], [o, 0], [w - o, 0], [w, h]]));
    }
    case 'pentagon':
      return filled(poly([[w / 2, 0], [w, h * 0.382], [w * 0.809, h], [w * 0.191, h], [0, h * 0.382]]));
    case 'hexagon': {
      const o = ss * a('adj', 25000);
      return filled(poly([[o, 0], [w - o, 0], [w, h / 2], [w - o, h], [o, h], [0, h / 2]]));
    }
    case 'heptagon':
    case 'decagon':
    case 'dodecagon': {
      const n = name === 'heptagon' ? 7 : name === 'decagon' ? 10 : 12;
      const pts: number[][] = [];
      for (let i = 0; i < n; i++) {
        const t = -Math.PI / 2 + (i * 2 * Math.PI) / n;
        pts.push([w / 2 + (Math.cos(t) * w) / 2, h / 2 + (Math.sin(t) * h) / 2]);
      }
      return filled(poly(pts));
    }
    case 'octagon': {
      const o = ss * a('adj', 29289);
      return filled(poly([[o, 0], [w - o, 0], [w, o], [w, h - o], [w - o, h], [o, h], [0, h - o], [0, o]]));
    }
    case 'plus':
    case 'mathPlus': {
      const o = ss * a('adj', 25000);
      return filled(
        poly([[o, 0], [w - o, 0], [w - o, o], [w, o], [w, h - o], [w - o, h - o], [w - o, h], [o, h], [o, h - o], [0, h - o], [0, o], [o, o]]),
      );
    }
    case 'star4':
      return filled(star(w, h, 4, a('adj', 12500) * 2));
    case 'star5':
      return filled(star(w, h, 5, 0.382));
    case 'star6':
      return filled(star(w, h, 6, 0.577));
    case 'star8':
      return filled(star(w, h, 8, 0.75));
    case 'rightArrow':
    case 'leftArrow': {
      const t = h * a('adj1', 50000);
      const head = ss * a('adj2', 50000);
      const y1 = h / 2 - t / 2;
      const y2 = h / 2 + t / 2;
      const x1 = w - head;
      const pts = [[0, y1], [x1, y1], [x1, 0], [w, h / 2], [x1, h], [x1, y2], [0, y2]];
      return filled(poly(name === 'rightArrow' ? pts : pts.map(([x, y]) => [w - x, y])));
    }
    case 'upArrow':
    case 'downArrow': {
      const t = w * a('adj1', 50000);
      const head = ss * a('adj2', 50000);
      const x1 = w / 2 - t / 2;
      const x2 = w / 2 + t / 2;
      const y1 = h - head;
      const pts = [[x1, 0], [x2, 0], [x2, y1], [w, y1], [w / 2, h], [0, y1], [x1, y1]];
      return filled(poly(name === 'downArrow' ? pts : pts.map(([x, y]) => [x, h - y])));
    }
    case 'leftRightArrow': {
      const t = h * a('adj1', 50000);
      const head = ss * a('adj2', 50000);
      const y1 = h / 2 - t / 2;
      const y2 = h / 2 + t / 2;
      return filled(poly([[0, h / 2], [head, 0], [head, y1], [w - head, y1], [w - head, 0], [w, h / 2], [w - head, h], [w - head, y2], [head, y2], [head, h]]));
    }
    case 'upDownArrow': {
      const t = w * a('adj1', 50000);
      const head = ss * a('adj2', 50000);
      const x1 = w / 2 - t / 2;
      const x2 = w / 2 + t / 2;
      return filled(poly([[w / 2, 0], [w, head], [x2, head], [x2, h - head], [w, h - head], [w / 2, h], [0, h - head], [x1, h - head], [x1, head], [0, head]]));
    }
    case 'chevron': {
      const o = ss * a('adj', 50000);
      return filled(poly([[0, 0], [w - o, 0], [w, h / 2], [w - o, h], [0, h], [o, h / 2]]));
    }
    case 'homePlate':
    case 'flowChartOffpageConnector': {
      const o = ss * a('adj', 50000);
      return filled(poly([[0, 0], [w - o, 0], [w, h / 2], [w - o, h], [0, h]]));
    }
    case 'donut': {
      const t = ss * a('adj', 25000);
      return filled([...ellipsePath(w / 2, h / 2, w / 2, h / 2), ...ellipsePath(w / 2, h / 2, Math.max(0, w / 2 - t), Math.max(0, h / 2 - t))], true);
    }
    case 'frame': {
      const t = ss * a('adj1', 12500);
      return filled([...rectPath(0, 0, w, h), ...rectPath(t, t, w - 2 * t, h - 2 * t)], true);
    }
    case 'heart':
      return filled([
        ['M', w / 2, h / 4],
        ['C', w / 2, 0, 0, 0, 0, h / 4],
        ['C', 0, h / 2, w / 2, (3 * h) / 4, w / 2, h],
        ['C', w / 2, (3 * h) / 4, w, h / 2, w, h / 4],
        ['C', w, 0, w / 2, 0, w / 2, h / 4],
        ['Z'],
      ]);
    case 'can': {
      const r = Math.min(h / 2, ss * a('adj', 25000) / 2);
      return [
        {
          cmds: [
            ['M', 0, r],
            ['L', 0, h - r],
            ['C', 0, h - r + r * K, w / 2 - (w / 2) * K, h, w / 2, h],
            ['C', w / 2 + (w / 2) * K, h, w, h - r + r * K, w, h - r],
            ['L', w, r],
            ['C', w, r + r * K, w / 2 + (w / 2) * K, 2 * r, w / 2, 2 * r],
            ['C', w / 2 - (w / 2) * K, 2 * r, 0, r + r * K, 0, r],
            ['Z'],
          ],
          fill: true,
          stroke: true,
        },
        { cmds: ellipsePath(w / 2, r, w / 2, r), fill: true, stroke: true },
      ];
    }
    case 'wedgeRectCallout':
    case 'wedgeRoundRectCallout':
    case 'wedgeEllipseCallout': {
      const tx = w / 2 + w * a('adj1', -20833);
      const ty = h / 2 + h * a('adj2', 62500);
      const body =
        name === 'wedgeEllipseCallout'
          ? ellipsePath(w / 2, h / 2, w / 2, h / 2)
          : name === 'wedgeRoundRectCallout'
            ? roundedRect(w, h, Array(4).fill(ss * 0.16667))
            : rectPath(0, 0, w, h);
      // The tail leaves the side facing its tip.
      const horizontal = Math.abs(tx - w / 2) / w > Math.abs(ty - h / 2) / h;
      const tail: number[][] = horizontal
        ? [[tx < w / 2 ? 0 : w, h * 0.4], [tx < w / 2 ? 0 : w, h * 0.6], [tx, ty]]
        : [[w * 0.3, ty < h / 2 ? 0 : h], [w * 0.45, ty < h / 2 ? 0 : h], [tx, ty]];
      const inset = name === 'wedgeEllipseCallout' ? 0.1 : 0.01;
      const tailPts = tail.map(([x, y], i) =>
        i < 2 ? [x + (x === 0 ? w * inset : x === w ? -w * inset : 0), y + (y === 0 ? h * inset : y === h ? -h * inset : 0)] : [x, y],
      );
      return [
        { cmds: body, fill: true, stroke: true },
        { cmds: poly(tailPts), fill: true, stroke: true },
      ];
    }
    case 'line':
    case 'straightConnector1':
      return stroked([['M', 0, 0], ['L', w, h]]);
    case 'bentConnector2':
      return stroked([['M', 0, 0], ['L', w, 0], ['L', w, h]]);
    case 'bentConnector3': {
      const x = w * a('adj1', 50000);
      return stroked([['M', 0, 0], ['L', x, 0], ['L', x, h], ['L', w, h]]);
    }
    case 'curvedConnector3':
      return stroked([['M', 0, 0], ['C', w / 2, 0, w / 2, h, w, h]]);
    case 'arc': {
      const st = adjOf(adj, 'adj1', 16200000) / 60000;
      const end = adjOf(adj, 'adj2', 0) / 60000;
      let sw = end - st;
      if (sw <= 0) sw += 360;
      const sx = w / 2 + (w / 2) * Math.cos((st * Math.PI) / 180);
      const sy = h / 2 + (h / 2) * Math.sin((st * Math.PI) / 180);
      return stroked([['M', sx, sy], ...arcTo(sx, sy, w / 2, h / 2, st, sw).cmds]);
    }
    default:
      return filled(rectPath(0, 0, w, h));
  }
}

const SUPPORTED = new Set([
  'rect', 'flowChartProcess', 'flowChartPredefinedProcess', 'flowChartInternalStorage', 'actionButtonBlank', 'roundRect',
  'flowChartAlternateProcess', 'round1Rect', 'round2SameRect', 'round2DiagRect', 'snip1Rect', 'snip2SameRect', 'ellipse',
  'flowChartConnector', 'flowChartTerminator', 'triangle', 'flowChartExtract', 'flowChartMerge', 'rtTriangle', 'diamond',
  'flowChartDecision', 'parallelogram', 'flowChartInputOutput', 'trapezoid', 'pentagon', 'hexagon', 'heptagon', 'decagon',
  'dodecagon', 'octagon', 'plus', 'mathPlus', 'star4', 'star5', 'star6', 'star8', 'rightArrow', 'leftArrow', 'upArrow',
  'downArrow', 'leftRightArrow', 'upDownArrow', 'chevron', 'homePlate', 'flowChartOffpageConnector', 'donut', 'frame', 'heart',
  'can', 'wedgeRectCallout', 'wedgeRoundRectCallout', 'wedgeEllipseCallout', 'line', 'straightConnector1', 'bentConnector2',
  'bentConnector3', 'curvedConnector3', 'arc',
]);

/** True when the preset is drawn exactly (others fall back to their bounding box). */
export function presetSupported(name: string): boolean {
  return SUPPORTED.has(name);
}

function customPaths(paths: CustomPath[], w: number, h: number): SubPath[] | null {
  const out: SubPath[] = [];
  for (const p of paths) {
    const sx = p.w ? w / p.w : 1;
    const sy = p.h ? h / p.h : 1;
    const cmds: Cmd[] = [];
    let cx = 0;
    let cy = 0;
    for (const c of p.cmds) {
      const [op, ...v] = c as [string, ...number[]];
      if (v.some((n) => !Number.isFinite(n))) return null; // formulas we do not evaluate
      if (op === 'M' || op === 'L') {
        cx = v[0] * sx;
        cy = v[1] * sy;
        cmds.push([op, cx, cy]);
      } else if (op === 'C') {
        cmds.push(['C', v[0] * sx, v[1] * sy, v[2] * sx, v[3] * sy, v[4] * sx, v[5] * sy]);
        cx = v[4] * sx;
        cy = v[5] * sy;
      } else if (op === 'Q') {
        const qx = v[0] * sx;
        const qy = v[1] * sy;
        const x = v[2] * sx;
        const y = v[3] * sy;
        cmds.push(['C', cx + (2 / 3) * (qx - cx), cy + (2 / 3) * (qy - cy), x + (2 / 3) * (qx - x), y + (2 / 3) * (qy - y), x, y]);
        cx = x;
        cy = y;
      } else if (op === 'A') {
        const arc = arcTo(cx, cy, v[0] * sx, v[1] * sy, v[2] / 60000, v[3] / 60000);
        cmds.push(...arc.cmds);
        cx = arc.x;
        cy = arc.y;
      } else if (op === 'Z') cmds.push(['Z']);
    }
    out.push({ cmds, fill: p.fill, stroke: p.stroke });
  }
  return out;
}

export function geometryPaths(g: Geometry, w: number, h: number): SubPath[] {
  if (g.kind === 'custom') return customPaths(g.paths, w, h) ?? presetPaths('rect', {}, w, h);
  return presetPaths(g.name, g.adj, w, h);
}

export function toSvg(cmds: Cmd[]): string {
  return cmds
    .map((c) => (c[0] === 'Z' ? 'Z' : `${c[0]}${(c.slice(1) as number[]).map((n) => Math.round(n * 100) / 100).join(' ')}`))
    .join('');
}

/**
 * Arrowhead polygon at the end of a line. `tip` is the line's end, `from` the previous point;
 * `width` is the line width. Returns null for 'none'.
 */
export function arrowHead(type: string | null, tip: [number, number], from: [number, number], width: number): Cmd[] | null {
  if (!type || type === 'none') return null;
  const dx = tip[0] - from[0];
  const dy = tip[1] - from[1];
  const len = Math.hypot(dx, dy) || 1;
  const ux = dx / len;
  const uy = dy / len;
  const size = Math.max(width * 3, 4);
  const back = [tip[0] - ux * size, tip[1] - uy * size];
  const nx = -uy * size * 0.5;
  const ny = ux * size * 0.5;
  if (type === 'oval') return ellipsePath(tip[0] - (ux * size) / 2, tip[1] - (uy * size) / 2, size / 2, size / 2);
  if (type === 'diamond') {
    const mid = [tip[0] - (ux * size) / 2, tip[1] - (uy * size) / 2];
    return poly([tip, [mid[0] + nx, mid[1] + ny], back, [mid[0] - nx, mid[1] - ny]]);
  }
  if (type === 'stealth') {
    const notch = [tip[0] - ux * size * 0.6, tip[1] - uy * size * 0.6];
    return poly([tip, [back[0] + nx, back[1] + ny], notch, [back[0] - nx, back[1] - ny]]);
  }
  return poly([tip, [back[0] + nx, back[1] + ny], [back[0] - nx, back[1] - ny]]);
}

/** End points of the first open subpath (for arrowheads). */
export function endPoints(paths: SubPath[]): { start: [number, number]; second: [number, number]; end: [number, number]; beforeEnd: [number, number] } | null {
  const p = paths.find((s) => !s.fill) ?? paths[0];
  if (!p) return null;
  const pts: [number, number][] = [];
  for (const c of p.cmds) {
    if (c[0] === 'M' || c[0] === 'L') pts.push([c[1], c[2]]);
    else if (c[0] === 'C') {
      pts.push([c[3], c[4]]);
      pts.push([c[5], c[6]]);
    }
  }
  if (pts.length < 2) return null;
  return { start: pts[0], second: pts[1], end: pts[pts.length - 1], beforeEnd: pts[pts.length - 2] };
}

/** Dash pattern (multiples of the line width) for a:prstDash values. */
export function dashPattern(dash: string | null): number[] | null {
  switch (dash) {
    case 'dash':
      return [4, 3];
    case 'sysDash':
      return [3, 1];
    case 'dot':
    case 'sysDot':
      return [1, 1];
    case 'dashDot':
    case 'sysDashDot':
      return [4, 3, 1, 3];
    case 'lgDash':
      return [8, 3];
    case 'lgDashDot':
      return [8, 3, 1, 3];
    case 'lgDashDotDot':
    case 'sysDashDotDot':
      return [8, 3, 1, 3, 1, 3];
    default:
      return null;
  }
}
