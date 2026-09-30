export interface ChartScale {
  min: number;
  max: number;
}

export interface Rect {
  x: number;
  y: number;
  width: number;
  height: number;
}

export interface BarOptions {
  grouping: 'clustered' | 'stacked' | 'percentStacked' | 'standard';
  gapWidth?: number;
  overlap?: number;
  direction?: 'col' | 'bar';
}

export function axisTicks(min: number, max: number, targetCount: number): { min: number; max: number; step: number; ticks: number[] } {
  const finite = [min, max].filter(Number.isFinite);
  if (!finite.length) return { min: 0, max: 1, step: 0.2, ticks: [0, 0.2, 0.4, 0.6, 0.8, 1] };
  let lo = Math.min(...finite);
  let hi = Math.max(...finite);
  if (lo === hi) {
    const delta = lo === 0 ? 1 : Math.abs(lo) * 0.1;
    lo -= delta;
    hi += delta;
  }
  const desired = Math.max(1, Number.isFinite(targetCount) ? Math.round(targetCount) : 5);
  const raw = (hi - lo) / desired;
  const power = 10 ** Math.floor(Math.log10(raw));
  const normalized = raw / power;
  const multiplier = normalized <= 1 ? 1 : normalized <= 2 ? 2 : normalized <= 5 ? 5 : 10;
  const step = multiplier * power;
  const tickMin = Math.floor(lo / step) * step;
  const tickMax = Math.ceil(hi / step) * step;
  const ticks: number[] = [];
  for (let value = tickMin, i = 0; value <= tickMax + step * 1e-9 && i < 1000; value += step, i++) {
    ticks.push(Number(value.toPrecision(12)));
  }
  return { min: tickMin, max: tickMax, step, ticks };
}

export function valueToPx(value: number, scale: ChartScale, lengthPx: number): number {
  if (!Number.isFinite(value) || !Number.isFinite(lengthPx) || lengthPx <= 0) return 0;
  if (!Number.isFinite(scale.min) || !Number.isFinite(scale.max) || scale.max === scale.min) return lengthPx / 2;
  return Math.max(0, Math.min(lengthPx, (value - scale.min) / (scale.max - scale.min) * lengthPx));
}

export function barRects(
  series: number[][],
  scale: ChartScale,
  plotWidthPx: number,
  plotHeightPx: number,
  options: BarOptions,
): Rect[][] {
  const count = Math.max(0, ...series.map((values) => values.length));
  if (!count || plotWidthPx <= 0 || plotHeightPx <= 0) return series.map(() => []);
  const horizontal = options.direction === 'bar';
  const gap = Math.max(0, Math.min(500, options.gapWidth ?? 150)) / 100;
  const categoryWidth = (horizontal ? plotHeightPx : plotWidthPx) / count;
  const groupWidth = categoryWidth * (1 - gap * 0.5);
  const grouping = options.grouping;
  const overlap = Math.max(-1, Math.min(1, (options.overlap ?? 0) / 100));
  const result = series.map(() => Array.from({ length: count }, () => ({ x: 0, y: 0, width: 0, height: 0 })));
  for (let index = 0; index < count; index++) {
    const values = series.map((values) => values[index] ?? 0);
    const positiveTotal = values.reduce((sum, value) => sum + Math.max(0, value), 0);
    const negativeTotal = values.reduce((sum, value) => sum + Math.min(0, value), 0);
    const slots = grouping === 'clustered' || grouping === 'standard' ? series.length : 1;
    const barWidth = groupWidth / Math.max(1, slots) * (1 + overlap * (slots - 1) / slots);
    let positive = 0;
    let negative = 0;
    series.forEach((_, seriesIndex) => {
      const value = values[seriesIndex] ?? 0;
      if (grouping === 'percentStacked') {
        const denominator = value >= 0 ? positiveTotal : Math.abs(negativeTotal);
        values[seriesIndex] = denominator ? value / denominator * 100 : 0;
      }
      const scaled = values[seriesIndex] ?? 0;
      let start: number;
      let end: number;
      if (grouping === 'clustered' || grouping === 'standard') {
        start = 0;
        end = scaled;
      } else if (scaled >= 0) {
        start = positive;
        positive += scaled;
        end = positive;
      } else {
        start = negative;
        negative += scaled;
        end = negative;
      }
      const x = index * categoryWidth + (categoryWidth - groupWidth) / 2 +
        (grouping === 'clustered' || grouping === 'standard' ? seriesIndex * groupWidth / series.length : 0);
      if (horizontal) {
        const x1 = valueToPx(Math.min(start, end), scale, plotWidthPx);
        const x2 = valueToPx(Math.max(start, end), scale, plotWidthPx);
        const y = index * categoryWidth + (categoryWidth - groupWidth) / 2 +
          (grouping === 'clustered' || grouping === 'standard' ? seriesIndex * groupWidth / series.length : 0);
        result[seriesIndex]![index] = { x: Math.min(x1, x2), y, width: Math.max(0, x2 - x1), height: Math.max(0, barWidth) };
      } else {
        const y1 = plotHeightPx - valueToPx(Math.max(start, end), scale, plotHeightPx);
        const y2 = plotHeightPx - valueToPx(Math.min(start, end), scale, plotHeightPx);
        result[seriesIndex]![index] = { x, y: Math.min(y1, y2), width: Math.max(0, barWidth), height: Math.max(0, Math.abs(y2 - y1)) };
      }
    });
  }
  return result;
}

export function pieSlicePath(startAngle: number, endAngle: number, cx: number, cy: number, r: number, innerR = 0): string {
  const point = (angle: number, radius: number) => `${cx + Math.cos(angle) * radius} ${cy + Math.sin(angle) * radius}`;
  const sweep = endAngle - startAngle;
  if (Math.abs(sweep) >= 2 * Math.PI) {
    const direction = Math.sign(sweep);
    const middleAngle = startAngle + direction * Math.PI;
    const outerStart = point(startAngle, r);
    const outerMiddle = point(middleAngle, r);
    if (innerR > 0) {
      const innerStart = point(startAngle, innerR);
      const innerMiddle = point(middleAngle, innerR);
      return `M ${outerStart} A ${r} ${r} 0 0 ${direction > 0 ? 1 : 0} ${outerMiddle} A ${r} ${r} 0 0 ${direction > 0 ? 1 : 0} ${outerStart} L ${innerStart} A ${innerR} ${innerR} 0 0 ${direction > 0 ? 0 : 1} ${innerMiddle} A ${innerR} ${innerR} 0 0 ${direction > 0 ? 0 : 1} ${innerStart} Z`;
    }
    return `M ${cx} ${cy} L ${outerStart} A ${r} ${r} 0 0 ${direction > 0 ? 1 : 0} ${outerMiddle} A ${r} ${r} 0 0 ${direction > 0 ? 1 : 0} ${outerStart} Z`;
  }
  const large = Math.abs(endAngle - startAngle) > Math.PI ? 1 : 0;
  const outerStart = point(startAngle, r);
  const outerEnd = point(endAngle, r);
  if (innerR > 0) {
    return `M ${outerStart} A ${r} ${r} 0 ${large} 1 ${outerEnd} L ${point(endAngle, innerR)} A ${innerR} ${innerR} 0 ${large} 0 ${point(startAngle, innerR)} Z`;
  }
  return `M ${cx} ${cy} L ${outerStart} A ${r} ${r} 0 ${large} 1 ${outerEnd} Z`;
}
