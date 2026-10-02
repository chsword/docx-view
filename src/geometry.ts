import type { CustomGeometry } from './types.js';

type Point = [number, number];
type ShapePath = (width: number, height: number, adjustments: Map<string, number>) => string;

function polygon(points: Point[]): string {
  return `M ${points[0]![0]} ${points[0]![1]} ${points.slice(1).map(([x, y]) => `L ${x} ${y}`).join(' ')} Z`;
}

const presets: Record<string, ShapePath> = {
  rect: (w, h) => `M 0 0 H ${w} V ${h} H 0 Z`,
  roundRect: (w, h, a) => {
    const r = Math.min(w, h) * Math.max(0.02, Math.min(0.3, a.get('adj') ?? 0.12));
    return `M ${r} 0 H ${w - r} Q ${w} 0 ${w} ${r} V ${h - r} Q ${w} ${h} ${w - r} ${h} H ${r} Q 0 ${h} 0 ${h - r} V ${r} Q 0 0 ${r} 0 Z`;
  },
  ellipse: (w, h) => `M ${w / 2} 0 A ${w / 2} ${h / 2} 0 1 1 ${w / 2} ${h} A ${w / 2} ${h / 2} 0 1 1 ${w / 2} 0 Z`,
  triangle: (w, h) => polygon([[w / 2, 0], [w, h], [0, h]]),
  rtTriangle: (w, h) => polygon([[0, 0], [0, h], [w, h]]),
  diamond: (w, h) => polygon([[w / 2, 0], [w, h / 2], [w / 2, h], [0, h / 2]]),
  parallelogram: (w, h, a) => {
    const s = w * Math.max(0.05, Math.min(0.45, a.get('adj') ?? 0.25));
    return polygon([[s, 0], [w, 0], [w - s, h], [0, h]]);
  },
  trapezoid: (w, h, a) => {
    const s = w * Math.max(0.05, Math.min(0.45, a.get('adj') ?? 0.25));
    return polygon([[s, 0], [w - s, 0], [w, h], [0, h]]);
  },
  pentagon: (w, h) => polygon([[w / 2, 0], [w, h * 0.38], [w * 0.81, h], [w * 0.19, h], [0, h * 0.38]]),
  hexagon: (w, h) => polygon([[w * 0.25, 0], [w * 0.75, 0], [w, h / 2], [w * 0.75, h], [w * 0.25, h], [0, h / 2]]),
  star5: (w, h) => {
    const points: Point[] = [];
    for (let i = 0; i < 10; i++) {
      const angle = -Math.PI / 2 + i * Math.PI / 5;
      const radius = i % 2 ? 0.22 : 0.5;
      points.push([w / 2 + Math.cos(angle) * w * radius, h / 2 + Math.sin(angle) * h * radius]);
    }
    return polygon(points);
  },
  rightArrow: (w, h, a) => {
    const neck = w * Math.max(0.35, Math.min(0.8, a.get('adj1') ?? 0.62));
    const half = h * Math.max(0.12, Math.min(0.45, a.get('adj2') ?? 0.24));
    return polygon([[0, h / 2 - half], [neck, h / 2 - half], [neck, 0], [w, h / 2], [neck, h], [neck, h / 2 + half], [0, h / 2 + half]]);
  },
  leftArrow: (w, h, a) => {
    const neck = w * Math.max(0.35, Math.min(0.8, a.get('adj1') ?? 0.62));
    const half = h * Math.max(0.12, Math.min(0.45, a.get('adj2') ?? 0.24));
    return polygon([[w, h / 2 - half], [w - neck, h / 2 - half], [w - neck, 0], [0, h / 2], [w - neck, h], [w - neck, h / 2 + half], [w, h / 2 + half]]);
  },
  upArrow: (w, h, a) => {
    const neck = h * Math.max(0.35, Math.min(0.8, a.get('adj1') ?? 0.62));
    const half = w * Math.max(0.12, Math.min(0.45, a.get('adj2') ?? 0.24));
    return polygon([[w / 2 - half, h], [w / 2 - half, neck], [0, neck], [w / 2, 0], [w, neck], [w / 2 + half, neck], [w / 2 + half, h]]);
  },
  downArrow: (w, h, a) => {
    const neck = h * Math.max(0.35, Math.min(0.8, a.get('adj1') ?? 0.62));
    const half = w * Math.max(0.12, Math.min(0.45, a.get('adj2') ?? 0.24));
    return polygon([[w / 2 - half, 0], [w / 2 + half, 0], [w / 2 + half, neck], [w, neck], [w / 2, h], [0, neck], [w / 2 - half, neck]]);
  },
  leftRightArrow: (w, h, a) => {
    const neck = w * Math.max(0.2, Math.min(0.4, a.get('adj1') ?? 0.27));
    const half = h * Math.max(0.12, Math.min(0.45, a.get('adj2') ?? 0.24));
    return polygon([[0, h / 2], [neck, 0], [neck, h / 2 - half], [w - neck, h / 2 - half], [w - neck, 0], [w, h / 2], [w - neck, h], [w - neck, h / 2 + half], [neck, h / 2 + half], [neck, h]]);
  },
  line: (w, h) => `M 0 0 L ${w} ${h}`,
  straightConnector1: (w, h) => `M 0 0 L ${w} ${h}`,
  cloudCallout: (w, h) => `M ${w * 0.12} ${h * 0.3} Q 0 ${h * 0.12} ${w * 0.18} ${h * 0.12} Q ${w * 0.2} 0 ${w * 0.4} ${h * 0.1} Q ${w * 0.58} -${h * 0.02} ${w * 0.68} ${h * 0.12} Q ${w} ${h * 0.02} ${w * 0.9} ${h * 0.3} Q ${w * 1.08} ${h * 0.52} ${w * 0.88} ${h * 0.65} Q ${w} ${h * 0.92} ${w * 0.68} ${h * 0.85} Q ${w * 0.45} ${h * 1.05} ${w * 0.3} ${h * 0.84} Q ${w * 0.02} ${h * 0.98} ${w * 0.12} ${h * 0.7} Q -${w * 0.05} ${h * 0.52} ${w * 0.12} ${h * 0.3} Z`,
};

/** 正多边形：顶点从正上方起顺时针，内接于 w×h 的椭圆。 */
function regularPolygon(w: number, h: number, sides: number): string {
  const points: Point[] = [];
  for (let i = 0; i < sides; i++) {
    const angle = -Math.PI / 2 + (i * 2 * Math.PI) / sides;
    points.push([w / 2 + (Math.cos(angle) * w) / 2, h / 2 + (Math.sin(angle) * h) / 2]);
  }
  return polygon(points);
}

/** n 角星。`inner` 是内顶点半径占外半径的比例（ECMA 的 adj 相对 50000，所以是 adj × 2）。 */
function star(w: number, h: number, points: number, inner: number): string {
  const vertices: Point[] = [];
  for (let i = 0; i < points * 2; i++) {
    const angle = -Math.PI / 2 + (i * Math.PI) / points;
    const radius = i % 2 ? inner / 2 : 0.5;
    vertices.push([w / 2 + Math.cos(angle) * w * radius, h / 2 + Math.sin(angle) * h * radius]);
  }
  return polygon(vertices);
}

const clamp = (value: number, min: number, max: number) => Math.max(min, Math.min(max, value));

/**
 * 角度调整值：ECMA 以 1/60000 度为单位，读进来时已经除过 100000，这里换回度。所以默认值也要
 * 按「原值 ÷ 100000」写：270° 是 16200000，在这里是 162。
 */
const degrees = (fraction: number) => (fraction * 100000) / 60000;

/** 椭圆上从 start 到 end（度，0° 在右、顺时针）的弧，起点用 M、不闭合。 */
function arcPath(w: number, h: number, start: number, end: number): string {
  const rx = w / 2;
  const ry = h / 2;
  const point = (deg: number): Point => [rx + rx * Math.cos((deg * Math.PI) / 180), ry + ry * Math.sin((deg * Math.PI) / 180)];
  const sweep = ((end - start) % 360 + 360) % 360 || 360;
  const [x0, y0] = point(start);
  const [x1, y1] = point(start + Math.min(sweep, 359.99));
  return `M ${x0} ${y0} A ${rx} ${ry} 0 ${sweep > 180 ? 1 : 0} 1 ${x1} ${y1}`;
}

/** 楔形标注的尾巴尖：相对中心偏 adj1·w、adj2·h（ECMA 默认 -20833 / 62500）。 */
function calloutTip(w: number, h: number, a: Map<string, number>): Point {
  return [w / 2 + (a.get('adj1') ?? -0.20833) * w, h / 2 + (a.get('adj2') ?? 0.625) * h];
}

/** 尾巴根部：在离尖端最近的那条边上取一小段。 */
function calloutBase(w: number, h: number, tip: Point): [Point, Point] {
  const dx = (tip[0] - w / 2) / w;
  const dy = (tip[1] - h / 2) / h;
  if (Math.abs(dy) >= Math.abs(dx)) {
    const y = dy > 0 ? h : 0;
    return [[w * 0.42, y], [w * 0.58, y]];
  }
  const x = dx > 0 ? w : 0;
  return [[x, h * 0.42], [x, h * 0.58]];
}

const none = new Map<string, number>();

const morePresets: Record<string, ShapePath> = {
  octagon: (w, h, a) => {
    const c = Math.min(w, h) * clamp(a.get('adj') ?? 0.29289, 0, 0.5);
    return polygon([[c, 0], [w - c, 0], [w, c], [w, h - c], [w - c, h], [c, h], [0, h - c], [0, c]]);
  },
  heptagon: (w, h) => regularPolygon(w, h, 7),
  decagon: (w, h) => regularPolygon(w, h, 10),
  dodecagon: (w, h) => regularPolygon(w, h, 12),
  star4: (w, h, a) => star(w, h, 4, clamp((a.get('adj') ?? 0.125) * 2, 0, 1)),
  star6: (w, h, a) => star(w, h, 6, clamp((a.get('adj') ?? 0.28868) * 2, 0, 1)),
  star7: (w, h, a) => star(w, h, 7, clamp((a.get('adj') ?? 0.34601) * 2, 0, 1)),
  star8: (w, h, a) => star(w, h, 8, clamp((a.get('adj') ?? 0.375) * 2, 0, 1)),
  star10: (w, h, a) => star(w, h, 10, clamp((a.get('adj') ?? 0.42533) * 2, 0, 1)),
  star12: (w, h, a) => star(w, h, 12, clamp((a.get('adj') ?? 0.375) * 2, 0, 1)),
  star16: (w, h, a) => star(w, h, 16, clamp((a.get('adj') ?? 0.375) * 2, 0, 1)),
  star24: (w, h, a) => star(w, h, 24, clamp((a.get('adj') ?? 0.375) * 2, 0, 1)),
  star32: (w, h, a) => star(w, h, 32, clamp((a.get('adj') ?? 0.375) * 2, 0, 1)),
  homePlate: (w, h, a) => {
    const x = w - Math.min(w, Math.min(w, h) * Math.max(0, a.get('adj') ?? 0.5));
    return polygon([[0, 0], [x, 0], [w, h / 2], [x, h], [0, h]]);
  },
  chevron: (w, h, a) => {
    const d = Math.min(w / 2, Math.min(w, h) * Math.max(0, a.get('adj') ?? 0.5));
    return polygon([[0, 0], [w - d, 0], [w, h / 2], [w - d, h], [0, h], [d, h / 2]]);
  },
  plus: (w, h, a) => {
    const c = Math.min(w, h) * clamp(a.get('adj') ?? 0.25, 0, 0.5);
    return polygon([[c, 0], [w - c, 0], [w - c, c], [w, c], [w, h - c], [w - c, h - c], [w - c, h], [c, h], [c, h - c], [0, h - c], [0, c], [c, c]]);
  },
  mathPlus: (w, h) => {
    const t = Math.min(w, h) * 0.1175;
    return polygon([[w / 2 - t, h * 0.13], [w / 2 + t, h * 0.13], [w / 2 + t, h / 2 - t], [w * 0.87, h / 2 - t], [w * 0.87, h / 2 + t],
      [w / 2 + t, h / 2 + t], [w / 2 + t, h * 0.87], [w / 2 - t, h * 0.87], [w / 2 - t, h / 2 + t], [w * 0.13, h / 2 + t],
      [w * 0.13, h / 2 - t], [w / 2 - t, h / 2 - t]]);
  },
  mathMinus: (w, h) => polygon([[w * 0.13, h * 0.38], [w * 0.87, h * 0.38], [w * 0.87, h * 0.62], [w * 0.13, h * 0.62]]),
  mathEqual: (w, h) => `${polygon([[w * 0.13, h * 0.24], [w * 0.87, h * 0.24], [w * 0.87, h * 0.44], [w * 0.13, h * 0.44]])} `
    + polygon([[w * 0.13, h * 0.56], [w * 0.87, h * 0.56], [w * 0.87, h * 0.76], [w * 0.13, h * 0.76]]),
  mathMultiply: (w, h) => polygon([[w * 0.1, h * 0.2], [w * 0.2, h * 0.1], [w / 2, h * 0.36], [w * 0.8, h * 0.1], [w * 0.9, h * 0.2],
    [w * 0.64, h / 2], [w * 0.9, h * 0.8], [w * 0.8, h * 0.9], [w / 2, h * 0.64], [w * 0.2, h * 0.9], [w * 0.1, h * 0.8], [w * 0.36, h / 2]]),
  mathDivide: (w, h) => {
    const r = Math.min(w, h) * 0.08;
    const dot = (cy: number) => `M ${w / 2 - r} ${cy} A ${r} ${r} 0 1 1 ${w / 2 + r} ${cy} A ${r} ${r} 0 1 1 ${w / 2 - r} ${cy} Z`;
    return `${polygon([[w * 0.13, h * 0.42], [w * 0.87, h * 0.42], [w * 0.87, h * 0.58], [w * 0.13, h * 0.58]])} ${dot(h * 0.24)} ${dot(h * 0.76)}`;
  },
  heart: (w, h) => `M ${w / 2} ${h * 0.25} C ${w * 0.5} ${h * -0.05} ${w * 1.05} ${h * 0.02} ${w * 0.98} ${h * 0.35} `
    + `C ${w * 0.93} ${h * 0.6} ${w * 0.62} ${h * 0.78} ${w / 2} ${h} C ${w * 0.38} ${h * 0.78} ${w * 0.07} ${h * 0.6} ${w * 0.02} ${h * 0.35} `
    + `C ${w * -0.05} ${h * 0.02} ${w * 0.5} ${h * -0.05} ${w / 2} ${h * 0.25} Z`,
  lightningBolt: (w, h) => polygon([[w * 0.39, 0], [w * 0.62, h * 0.29], [w * 0.53, h * 0.33], [w * 0.8, h * 0.6], [w * 0.69, h * 0.65],
    [w, h], [w * 0.48, h * 0.72], [w * 0.59, h * 0.66], [w * 0.21, h * 0.42], [w * 0.33, h * 0.37], [0, h * 0.14]]),
  moon: (w, h, a) => {
    const inner = w * clamp(a.get('adj') ?? 0.5, 0, 0.875);
    return `M ${w} 0 A ${w} ${h / 2} 0 0 0 ${w} ${h} A ${w - inner} ${h / 2} 0 0 1 ${w} 0 Z`;
  },
  donut: (w, h, a) => {
    const d = Math.min(w, h) * clamp(a.get('adj') ?? 0.25, 0, 0.5);
    const ring = (x: number, y: number, rx: number, ry: number, sweep: number) =>
      `M ${x} ${y + ry} A ${rx} ${ry} 0 1 ${sweep} ${x + 2 * rx} ${y + ry} A ${rx} ${ry} 0 1 ${sweep} ${x} ${y + ry} Z`;
    // 内圈反向绕行：nonzero 填充规则下中间是空的。
    return `${ring(0, 0, w / 2, h / 2, 1)} ${ring(d, d, w / 2 - d, h / 2 - d, 0)}`;
  },
  noSmoking: (w, h, a) => {
    const d = Math.min(w, h) * clamp(a.get('adj') ?? 0.1875, 0, 0.5);
    const t = d * 0.7;
    return `${morePresets.donut!(w, h, new Map([['adj', d / Math.min(w, h)]]))} `
      + polygon([[w * 0.2 - t / 2, h * 0.2 + t / 2], [w * 0.2 + t / 2, h * 0.2 - t / 2], [w * 0.8 + t / 2, h * 0.8 - t / 2], [w * 0.8 - t / 2, h * 0.8 + t / 2]]);
  },
  smileyFace: (w, h) => {
    const eye = (cx: number) => `M ${cx - w * 0.05} ${h * 0.37} A ${w * 0.05} ${h * 0.05} 0 1 1 ${cx + w * 0.05} ${h * 0.37} `
      + `A ${w * 0.05} ${h * 0.05} 0 1 1 ${cx - w * 0.05} ${h * 0.37} Z`;
    return `${presets.ellipse!(w, h, none)} ${eye(w * 0.35)} ${eye(w * 0.65)} M ${w * 0.27} ${h * 0.65} Q ${w / 2} ${h * 0.85} ${w * 0.73} ${h * 0.65}`;
  },
  sun: (w, h) => {
    const rays: string[] = [];
    for (let i = 0; i < 8; i++) {
      const angle = (i * Math.PI) / 4;
      const at = (offset: number, radius: number): Point =>
        [w / 2 + Math.cos(angle + offset) * w * radius, h / 2 + Math.sin(angle + offset) * h * radius];
      rays.push(polygon([at(-0.18, 0.34), at(0, 0.5), at(0.18, 0.34)]));
    }
    const rx = w * 0.25;
    const ry = h * 0.25;
    return `M ${w / 2 - rx} ${h / 2} A ${rx} ${ry} 0 1 1 ${w / 2 + rx} ${h / 2} A ${rx} ${ry} 0 1 1 ${w / 2 - rx} ${h / 2} Z ${rays.join(' ')}`;
  },
  cloud: (w, h) => presets.cloudCallout!(w, h, none),
  teardrop: (w, h, a) => {
    const k = clamp(a.get('adj') ?? 1, 0, 2);
    const tipX = w / 2 + (w / 2) * k;
    const tipY = h / 2 - (h / 2) * k;
    return `M 0 ${h / 2} A ${w / 2} ${h / 2} 0 0 1 ${w / 2} 0 Q ${tipX} 0 ${tipX} ${tipY} Q ${w} ${tipY} ${w} ${h / 2} `
      + `A ${w / 2} ${h / 2} 0 0 1 0 ${h / 2} Z`;
  },
  can: (w, h, a) => {
    const ry = (Math.min(w, h) * clamp(a.get('adj') ?? 0.25, 0, 1)) / 2;
    return `M 0 ${ry} A ${w / 2} ${ry} 0 0 1 ${w} ${ry} V ${h - ry} A ${w / 2} ${ry} 0 0 1 0 ${h - ry} Z M 0 ${ry} A ${w / 2} ${ry} 0 0 0 ${w} ${ry}`;
  },
  flowChartMagneticDisk: (w, h) => morePresets.can!(w, h, new Map([['adj', h / 3 / Math.min(w, h)]])),
  cube: (w, h, a) => {
    const d = Math.min(w, h) * clamp(a.get('adj') ?? 0.25, 0, 1);
    return `${polygon([[0, d], [w - d, d], [w - d, h], [0, h]])} ${polygon([[0, d], [d, 0], [w, 0], [w - d, d]])} `
      + polygon([[w - d, d], [w, 0], [w, h - d], [w - d, h]]);
  },
  frame: (w, h, a) => {
    const d = Math.min(w, h) * clamp(a.get('adj1') ?? 0.125, 0, 0.5);
    return `M 0 0 H ${w} V ${h} H 0 Z M ${d} ${d} V ${h - d} H ${w - d} V ${d} Z`;
  },
  bevel: (w, h, a) => {
    const d = Math.min(w, h) * clamp(a.get('adj') ?? 0.125, 0, 0.5);
    return `M 0 0 H ${w} V ${h} H 0 Z M ${d} ${d} H ${w - d} V ${h - d} H ${d} Z`;
  },
  plaque: (w, h, a) => {
    const d = Math.min(w, h) * clamp(a.get('adj') ?? 0.16667, 0, 0.5);
    return `M ${d} 0 H ${w - d} A ${d} ${d} 0 0 0 ${w} ${d} V ${h - d} A ${d} ${d} 0 0 0 ${w - d} ${h} H ${d} `
      + `A ${d} ${d} 0 0 0 0 ${h - d} V ${d} A ${d} ${d} 0 0 0 ${d} 0 Z`;
  },
  foldedCorner: (w, h, a) => {
    const d = Math.min(w, h) * clamp(a.get('adj') ?? 0.16667, 0, 0.5);
    return `M 0 0 H ${w} V ${h - d} L ${w - d} ${h} H 0 Z M ${w} ${h - d} L ${w - d * 0.8} ${h - d * 0.8} L ${w - d} ${h} Z`;
  },
  snip1Rect: (w, h, a) => {
    const d = Math.min(w, h) * clamp(a.get('adj') ?? 0.16667, 0, 0.5);
    return polygon([[0, 0], [w - d, 0], [w, d], [w, h], [0, h]]);
  },
  snip2SameRect: (w, h, a) => {
    const d = Math.min(w, h) * clamp(a.get('adj1') ?? 0.16667, 0, 0.5);
    return polygon([[d, 0], [w - d, 0], [w, d], [w, h], [0, h], [0, d]]);
  },
  snip2DiagRect: (w, h, a) => {
    const d = Math.min(w, h) * clamp(a.get('adj2') ?? 0.16667, 0, 0.5);
    return polygon([[0, 0], [w - d, 0], [w, d], [w, h], [d, h], [0, h - d]]);
  },
  round1Rect: (w, h, a) => {
    const r = Math.min(w, h) * clamp(a.get('adj') ?? 0.16667, 0, 0.5);
    return `M 0 0 H ${w - r} Q ${w} 0 ${w} ${r} V ${h} H 0 Z`;
  },
  round2SameRect: (w, h, a) => {
    const r = Math.min(w, h) * clamp(a.get('adj1') ?? 0.16667, 0, 0.5);
    return `M ${r} 0 H ${w - r} Q ${w} 0 ${w} ${r} V ${h} H 0 V ${r} Q 0 0 ${r} 0 Z`;
  },
  round2DiagRect: (w, h, a) => {
    const r = Math.min(w, h) * clamp(a.get('adj1') ?? 0.16667, 0, 0.5);
    return `M ${r} 0 H ${w} V ${h - r} Q ${w} ${h} ${w - r} ${h} H 0 V ${r} Q 0 0 ${r} 0 Z`;
  },
  upDownArrow: (w, h, a) => {
    const half = w * clamp((a.get('adj1') ?? 0.5) / 2, 0.05, 0.5);
    const head = Math.min(h / 2, Math.min(w, h) * Math.max(0, a.get('adj2') ?? 0.5));
    return polygon([[w / 2, 0], [w, head], [w / 2 + half, head], [w / 2 + half, h - head], [w, h - head], [w / 2, h],
      [0, h - head], [w / 2 - half, h - head], [w / 2 - half, head], [0, head]]);
  },
  quadArrow: (w, h) => {
    const t = Math.min(w, h) * 0.1;
    const head = Math.min(w, h) * 0.2;
    const cx = w / 2;
    const cy = h / 2;
    return polygon([[cx, 0], [cx + head, head], [cx + t, head], [cx + t, cy - t], [w - head, cy - t], [w - head, cy - head], [w, cy],
      [w - head, cy + head], [w - head, cy + t], [cx + t, cy + t], [cx + t, h - head], [cx + head, h - head], [cx, h],
      [cx - head, h - head], [cx - t, h - head], [cx - t, cy + t], [head, cy + t], [head, cy + head], [0, cy], [head, cy - head],
      [head, cy - t], [cx - t, cy - t], [cx - t, head], [cx - head, head]]);
  },
  notchedRightArrow: (w, h, a) => {
    const neck = w * clamp(a.get('adj1') ?? 0.62, 0.35, 0.8);
    const half = h * clamp(a.get('adj2') ?? 0.24, 0.12, 0.45);
    return polygon([[0, h / 2 - half], [neck, h / 2 - half], [neck, 0], [w, h / 2], [neck, h], [neck, h / 2 + half], [0, h / 2 + half], [h / 4, h / 2]]);
  },
  stripedRightArrow: (w, h, a) => {
    const neck = w * clamp(a.get('adj1') ?? 0.62, 0.35, 0.8);
    const half = h * clamp(a.get('adj2') ?? 0.24, 0.12, 0.45);
    const top = h / 2 - half;
    const bottom = h / 2 + half;
    const stripe = (x: number, width: number) => polygon([[x, top], [x + width, top], [x + width, bottom], [x, bottom]]);
    return `${stripe(0, w * 0.03)} ${stripe(w * 0.06, w * 0.06)} `
      + polygon([[w * 0.16, top], [neck, top], [neck, 0], [w, h / 2], [neck, h], [neck, bottom], [w * 0.16, bottom]]);
  },
  wedgeRectCallout: (w, h, a) => {
    const tip = calloutTip(w, h, a);
    const [b1, b2] = calloutBase(w, h, tip);
    return `M 0 0 H ${w} V ${h} H 0 Z ${polygon([b1, tip, b2])}`;
  },
  wedgeRoundRectCallout: (w, h, a) => {
    const tip = calloutTip(w, h, a);
    const [b1, b2] = calloutBase(w, h, tip);
    return `${presets.roundRect!(w, h, new Map([['adj', 0.16667]]))} ${polygon([b1, tip, b2])}`;
  },
  wedgeEllipseCallout: (w, h, a) => {
    const tip = calloutTip(w, h, a);
    const angle = Math.atan2((tip[1] - h / 2) / h, (tip[0] - w / 2) / w);
    const at = (offset: number): Point => [w / 2 + (w / 2) * Math.cos(angle + offset), h / 2 + (h / 2) * Math.sin(angle + offset)];
    return `${presets.ellipse!(w, h, none)} ${polygon([at(-0.2), tip, at(0.2)])}`;
  },
  flowChartProcess: (w, h) => presets.rect!(w, h, none),
  flowChartAlternateProcess: (w, h) => presets.roundRect!(w, h, new Map([['adj', 0.16667]])),
  flowChartDecision: (w, h) => presets.diamond!(w, h, none),
  flowChartInputOutput: (w, h) => polygon([[w / 5, 0], [w, 0], [w * 4 / 5, h], [0, h]]),
  flowChartPredefinedProcess: (w, h) => `M 0 0 H ${w} V ${h} H 0 Z M ${w / 8} 0 V ${h} M ${w * 7 / 8} 0 V ${h}`,
  flowChartInternalStorage: (w, h) => `M 0 0 H ${w} V ${h} H 0 Z M ${w / 8} 0 V ${h} M 0 ${h / 8} H ${w}`,
  flowChartDocument: (w, h) => `M 0 0 H ${w} V ${h * 0.83} C ${w * 0.75} ${h * 0.7} ${w * 0.5} ${h * 1.08} 0 ${h * 0.92} Z`,
  flowChartMultidocument: (w, h) => `M ${w * 0.1} 0 H ${w} V ${h * 0.7} H ${w * 0.9} M ${w * 0.05} ${h * 0.05} H ${w * 0.95} V ${h * 0.75} `
    + `M 0 ${h * 0.1} H ${w * 0.9} V ${h * 0.86} C ${w * 0.68} ${h * 0.75} ${w * 0.45} ${h * 1.05} 0 ${h * 0.94} Z`,
  flowChartTerminator: (w, h) => {
    const r = Math.min(h / 2, w / 2);
    return `M ${r} 0 H ${w - r} A ${r} ${h / 2} 0 0 1 ${w - r} ${h} H ${r} A ${r} ${h / 2} 0 0 1 ${r} 0 Z`;
  },
  flowChartPreparation: (w, h) => polygon([[w / 5, 0], [w * 4 / 5, 0], [w, h / 2], [w * 4 / 5, h], [w / 5, h], [0, h / 2]]),
  flowChartManualInput: (w, h) => polygon([[0, h / 5], [w, 0], [w, h], [0, h]]),
  flowChartManualOperation: (w, h) => polygon([[0, 0], [w, 0], [w * 4 / 5, h], [w / 5, h]]),
  flowChartConnector: (w, h) => presets.ellipse!(w, h, none),
  flowChartOffpageConnector: (w, h) => polygon([[0, 0], [w, 0], [w, h * 0.8], [w / 2, h], [0, h * 0.8]]),
  flowChartPunchedCard: (w, h) => polygon([[w / 5, 0], [w, 0], [w, h], [0, h], [0, h / 5]]),
  flowChartPunchedTape: (w, h) => `M 0 ${h * 0.1} Q ${w / 4} ${h * 0.3} ${w / 2} ${h * 0.1} T ${w} ${h * 0.1} V ${h * 0.9} `
    + `Q ${w * 0.75} ${h * 0.7} ${w / 2} ${h * 0.9} T 0 ${h * 0.9} Z`,
  flowChartSummingJunction: (w, h) => `${presets.ellipse!(w, h, none)} M ${w * 0.15} ${h * 0.15} L ${w * 0.85} ${h * 0.85} `
    + `M ${w * 0.85} ${h * 0.15} L ${w * 0.15} ${h * 0.85}`,
  flowChartOr: (w, h) => `${presets.ellipse!(w, h, none)} M ${w / 2} 0 V ${h} M 0 ${h / 2} H ${w}`,
  flowChartCollate: (w, h) => `${polygon([[0, 0], [w, 0], [w / 2, h / 2]])} ${polygon([[w / 2, h / 2], [w, h], [0, h]])}`,
  flowChartSort: (w, h) => `${presets.diamond!(w, h, none)} M 0 ${h / 2} H ${w}`,
  flowChartExtract: (w, h) => polygon([[w / 2, 0], [w, h], [0, h]]),
  flowChartMerge: (w, h) => polygon([[0, 0], [w, 0], [w / 2, h]]),
  flowChartDelay: (w, h) => `M 0 0 H ${w / 2} A ${w / 2} ${h / 2} 0 0 1 ${w / 2} ${h} H 0 Z`,
  flowChartDisplay: (w, h) => `M 0 ${h / 2} L ${w / 6} 0 H ${w * 5 / 6} A ${w / 6} ${h / 2} 0 0 1 ${w * 5 / 6} ${h} H ${w / 6} Z`,
  flowChartOnlineStorage: (w, h) => `M ${w / 6} 0 H ${w} A ${w / 6} ${h / 2} 0 0 0 ${w} ${h} H ${w / 6} A ${w / 6} ${h / 2} 0 0 1 ${w / 6} 0 Z`,
  flowChartMagneticTape: (w, h) => `${presets.ellipse!(w, h, none)} M ${w / 2} ${h} H ${w}`,
  pie: (w, h, a) => `M ${w / 2} ${h / 2} L ${arcPath(w, h, degrees(a.get('adj1') ?? 0), degrees(a.get('adj2') ?? 162)).slice(2)} Z`,
  chord: (w, h, a) => `${arcPath(w, h, degrees(a.get('adj1') ?? 27), degrees(a.get('adj2') ?? 162))} Z`,
  arc: (w, h, a) => arcPath(w, h, degrees(a.get('adj1') ?? 162), degrees(a.get('adj2') ?? 0)),
  blockArc: (w, h, a) => {
    const start = degrees(a.get('adj1') ?? 108);
    const end = degrees(a.get('adj2') ?? 0);
    const d = Math.min(w, h) * clamp(a.get('adj3') ?? 0.25, 0, 0.5);
    const rx = w / 2 - d;
    const ry = h / 2 - d;
    const inner = (deg: number): Point => [w / 2 + rx * Math.cos((deg * Math.PI) / 180), h / 2 + ry * Math.sin((deg * Math.PI) / 180)];
    const sweep = ((end - start) % 360 + 360) % 360 || 360;
    const [ix0, iy0] = inner(end);
    const [ix1, iy1] = inner(start);
    // 外弧顺时针走过去，内弧逆时针走回来，围成一段弧形的带子。
    return `${arcPath(w, h, start, end)} L ${ix0} ${iy0} A ${rx} ${ry} 0 ${sweep > 180 ? 1 : 0} 0 ${ix1} ${iy1} Z`;
  },
  leftBracket: (w, h, a) => {
    const y = Math.min(h / 2, Math.min(w, h) * Math.max(0, a.get('adj') ?? 0.08333));
    return `M ${w} ${h} A ${w} ${y} 0 0 1 0 ${h - y} V ${y} A ${w} ${y} 0 0 1 ${w} 0`;
  },
  rightBracket: (w, h, a) => {
    const y = Math.min(h / 2, Math.min(w, h) * Math.max(0, a.get('adj') ?? 0.08333));
    return `M 0 0 A ${w} ${y} 0 0 1 ${w} ${y} V ${h - y} A ${w} ${y} 0 0 1 0 ${h}`;
  },
  leftBrace: (w, h) => `M ${w} 0 Q ${w / 2} 0 ${w / 2} ${h * 0.1} V ${h * 0.4} Q ${w / 2} ${h / 2} 0 ${h / 2} `
    + `Q ${w / 2} ${h / 2} ${w / 2} ${h * 0.6} V ${h * 0.9} Q ${w / 2} ${h} ${w} ${h}`,
  rightBrace: (w, h) => `M 0 0 Q ${w / 2} 0 ${w / 2} ${h * 0.1} V ${h * 0.4} Q ${w / 2} ${h / 2} ${w} ${h / 2} `
    + `Q ${w / 2} ${h / 2} ${w / 2} ${h * 0.6} V ${h * 0.9} Q ${w / 2} ${h} 0 ${h}`,
  bracketPair: (w, h) => {
    const r = Math.min(w, h) * 0.16667;
    return `M ${r} 0 Q 0 0 0 ${r} V ${h - r} Q 0 ${h} ${r} ${h} M ${w - r} 0 Q ${w} 0 ${w} ${r} V ${h - r} Q ${w} ${h} ${w - r} ${h}`;
  },
  bracePair: (w, h) => {
    const r = Math.min(w, h) * 0.08333;
    return `M ${r * 2} 0 Q ${r} 0 ${r} ${r} V ${h / 2 - r} Q ${r} ${h / 2} 0 ${h / 2} Q ${r} ${h / 2} ${r} ${h / 2 + r} V ${h - r} Q ${r} ${h} ${r * 2} ${h} `
      + `M ${w - r * 2} 0 Q ${w - r} 0 ${w - r} ${r} V ${h / 2 - r} Q ${w - r} ${h / 2} ${w} ${h / 2} `
      + `Q ${w - r} ${h / 2} ${w - r} ${h / 2 + r} V ${h - r} Q ${w - r} ${h} ${w - r * 2} ${h}`;
  },
  bentConnector2: (w, h) => `M 0 0 H ${w} V ${h}`,
  bentConnector3: (w, h, a) => `M 0 0 H ${w * (a.get('adj1') ?? 0.5)} V ${h} H ${w}`,
  bentConnector4: (w, h, a) => `M 0 0 H ${w * (a.get('adj1') ?? 0.5)} V ${h * (a.get('adj2') ?? 0.5)} H ${w} V ${h}`,
  curvedConnector2: (w, h) => `M 0 0 Q ${w} 0 ${w} ${h}`,
  curvedConnector3: (w, h, a) => {
    const x = w * (a.get('adj1') ?? 0.5);
    return `M 0 0 C ${x} 0 ${x} ${h / 2} ${x} ${h / 2} S ${x} ${h} ${w} ${h}`;
  },
  wave: (w, h, a) => {
    const d = h * clamp(a.get('adj1') ?? 0.125, 0, 0.2);
    return `M 0 ${d} C ${w / 3} ${-d} ${w * 2 / 3} ${d * 3} ${w} ${d} V ${h - d} C ${w * 2 / 3} ${h + d} ${w / 3} ${h - d * 3} 0 ${h - d} Z`;
  },
  doubleWave: (w, h, a) => {
    const d = h * clamp(a.get('adj1') ?? 0.0625, 0, 0.125);
    return `M 0 ${d} C ${w / 6} ${-d} ${w / 3} ${d * 3} ${w / 2} ${d} S ${w * 5 / 6} ${-d} ${w} ${d} V ${h - d} `
      + `C ${w * 5 / 6} ${h + d} ${w * 2 / 3} ${h - d * 3} ${w / 2} ${h - d} S ${w / 6} ${h + d} 0 ${h - d} Z`;
  },
};
morePresets.cross = morePresets.plus!;
Object.assign(presets, morePresets);

/**
 * 只描边、不填充的预设：线、连接线、弧、括号。它们的路径不闭合，SVG 照样会把开放路径首尾
 * 连起来填色，所以渲染时要把填充关掉，否则括号会画成一块实心的月牙。
 */
const OPEN_PRESETS = new Set([
  'line', 'straightConnector1', 'bentConnector2', 'bentConnector3', 'bentConnector4', 'curvedConnector2',
  'curvedConnector3', 'arc', 'leftBracket', 'rightBracket', 'leftBrace', 'rightBrace', 'bracketPair', 'bracePair',
]);

export function presetGeometryIsOpen(preset: string): boolean {
  return OPEN_PRESETS.has(preset);
}

/** 支持的预设几何名。README 的清单按它核对，不手写。 */
export function supportedPresetGeometries(): string[] {
  return Object.keys(presets);
}

export function presetGeometryPath(
  preset: string,
  widthPx: number,
  heightPx: number,
  adjustments: Map<string, number> = new Map(),
): string | undefined {
  const path = presets[preset];
  if (!path || !Number.isFinite(widthPx) || !Number.isFinite(heightPx) || widthPx <= 0 || heightPx <= 0) return undefined;
  return path(widthPx, heightPx, adjustments);
}

export function customGeometryPath(geometry: CustomGeometry, widthPx: number, heightPx: number): string | undefined {
  if (!Number.isFinite(geometry.width) || !Number.isFinite(geometry.height) || geometry.width <= 0 || geometry.height <= 0 ||
      !Number.isFinite(widthPx) || !Number.isFinite(heightPx) || widthPx <= 0 || heightPx <= 0) return undefined;
  const sx = widthPx / geometry.width;
  const sy = heightPx / geometry.height;
  const n = (value: number, scale: number) => {
    if (!Number.isFinite(value)) throw new Error('Invalid custom geometry coordinate');
    return Number((value * scale).toFixed(3));
  };
  try {
    return geometry.commands.map((command) => {
      if (command.type === 'moveTo' || command.type === 'lnTo') {
        return `${command.type === 'moveTo' ? 'M' : 'L'} ${n(command.x, sx)} ${n(command.y, sy)}`;
      }
      if (command.type === 'cubicBezTo') {
        return `C ${n(command.x1, sx)} ${n(command.y1, sy)} ${n(command.x2, sx)} ${n(command.y2, sy)} ${n(command.x, sx)} ${n(command.y, sy)}`;
      }
      if (command.type === 'close') return 'Z';
      throw new Error('Unsupported custom geometry command');
    }).join(' ') || undefined;
  } catch {
    return undefined;
  }
}
