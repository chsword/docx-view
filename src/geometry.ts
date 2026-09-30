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
  wedgeRectCallout: (w, h) => `M 0 0 H ${w} V ${h} H 0 Z M ${w * 0.35} ${h} L ${w * 0.22} ${h * 1.3} L ${w * 0.52} ${h} Z`,
  cloudCallout: (w, h) => `M ${w * 0.12} ${h * 0.3} Q 0 ${h * 0.12} ${w * 0.18} ${h * 0.12} Q ${w * 0.2} 0 ${w * 0.4} ${h * 0.1} Q ${w * 0.58} -${h * 0.02} ${w * 0.68} ${h * 0.12} Q ${w} ${h * 0.02} ${w * 0.9} ${h * 0.3} Q ${w * 1.08} ${h * 0.52} ${w * 0.88} ${h * 0.65} Q ${w} ${h * 0.92} ${w * 0.68} ${h * 0.85} Q ${w * 0.45} ${h * 1.05} ${w * 0.3} ${h * 0.84} Q ${w * 0.02} ${h * 0.98} ${w * 0.12} ${h * 0.7} Q -${w * 0.05} ${h * 0.52} ${w * 0.12} ${h * 0.3} Z`,
};

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
