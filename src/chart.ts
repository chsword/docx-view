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

export type TrendlineType = 'linear' | 'exp' | 'log' | 'poly' | 'power' | 'movingAvg';

/**
 * 趋势线在若干 x 上的取值。点取自缓存的 (x, y)，缺失的点跳过。拟合方式与 Excel 一致：
 * - linear：最小二乘直线；poly：order 次多项式（2~6，正规方程 + 高斯消元）；
 * - exp：y = a·e^(bx)，对 ln y 做直线拟合（y ≤ 0 的点 Excel 直接拒绝，这里跳过）；
 * - log：y = a + b·ln x（x ≤ 0 跳过）；power：y = a·x^b（x、y ≤ 0 跳过）；
 * - movingAvg：period 点滑动平均，从第 period 个点起才有值。
 * 拟合不出来（点不够、矩阵奇异）时返回空数组，调用方就不画。
 */
export function trendlinePoints(
  points: Array<{ x: number; y: number }>,
  type: TrendlineType,
  options: { order?: number; period?: number; samples?: number } = {},
): Array<{ x: number; y: number }> {
  const finite = points.filter((point) => Number.isFinite(point.x) && Number.isFinite(point.y));
  if (type === 'movingAvg') {
    const period = Math.max(2, Math.min(255, Math.trunc(options.period ?? 2)));
    const result: Array<{ x: number; y: number }> = [];
    for (let index = period - 1; index < finite.length; index++) {
      const span = finite.slice(index - period + 1, index + 1);
      result.push({ x: finite[index]!.x, y: span.reduce((sum, point) => sum + point.y, 0) / period });
    }
    return result;
  }
  const usable = finite.filter((point) =>
    (type === 'exp' ? point.y > 0 : true) && (type === 'log' ? point.x > 0 : true) && (type === 'power' ? point.x > 0 && point.y > 0 : true));
  const order = type === 'poly' ? Math.max(2, Math.min(6, Math.trunc(options.order ?? 2))) : 1;
  if (usable.length < order + 1) return [];
  const transformX = (x: number) => (type === 'log' || type === 'power' ? Math.log(x) : x);
  const transformY = (y: number) => (type === 'exp' || type === 'power' ? Math.log(y) : y);
  const coefficients = polyfit(usable.map((point) => transformX(point.x)), usable.map((point) => transformY(point.y)), order);
  if (!coefficients) return [];
  const evaluate = (x: number) => {
    const t = transformX(x);
    const value = coefficients.reduce((sum, coefficient, power) => sum + coefficient * t ** power, 0);
    return type === 'exp' || type === 'power' ? Math.exp(value) : value;
  };
  const xs = usable.map((point) => point.x);
  const min = Math.min(...xs);
  const max = Math.max(...xs);
  const samples = Math.max(2, Math.min(200, options.samples ?? (type === 'linear' ? 2 : 40)));
  return Array.from({ length: samples }, (_, index) => {
    const x = min + ((max - min) * index) / (samples - 1);
    return { x, y: evaluate(x) };
  }).filter((point) => Number.isFinite(point.y));
}

/** 最小二乘多项式拟合，返回从常数项起的系数；奇异时返回 undefined。 */
function polyfit(xs: number[], ys: number[], order: number): number[] | undefined {
  const size = order + 1;
  const matrix = Array.from({ length: size }, (_, row) => Array.from({ length: size + 1 }, (_, column) => (
    column === size
      ? xs.reduce((sum, x, index) => sum + ys[index]! * x ** row, 0)
      : xs.reduce((sum, x) => sum + x ** (row + column), 0)
  )));
  for (let pivot = 0; pivot < size; pivot++) {
    let best = pivot;
    for (let row = pivot + 1; row < size; row++) if (Math.abs(matrix[row]![pivot]!) > Math.abs(matrix[best]![pivot]!)) best = row;
    if (Math.abs(matrix[best]![pivot]!) < 1e-12) return undefined;
    [matrix[pivot], matrix[best]] = [matrix[best]!, matrix[pivot]!];
    for (let row = 0; row < size; row++) {
      if (row === pivot) continue;
      const factor = matrix[row]![pivot]! / matrix[pivot]![pivot]!;
      for (let column = pivot; column <= size; column++) matrix[row]![column]! -= factor * matrix[pivot]![column]!;
    }
  }
  return matrix.map((row, index) => row[size]! / row[index]!);
}

export interface ErrorBarSpec {
  type: 'both' | 'plus' | 'minus';
  valueType: 'fixedVal' | 'percentage' | 'stdDev' | 'stdErr' | 'cust';
  value?: number;
  plus?: Array<number | null>;
  minus?: Array<number | null>;
}

/**
 * 每个点误差线的上下端（数据值，不是像素）。各取值方式按 Excel：fixedVal 是定值；percentage
 * 是该点值的百分比；stdDev 是 value 倍的**样本**标准差、以序列均值为中心（不是以各点为中心）；
 * stdErr 是标准误；cust 取缓存里每个点自己的正负值。
 */
export function errorBarRanges(values: Array<number | null>, spec: ErrorBarSpec): Array<{ low: number; high: number } | null> {
  const finite = values.filter((value): value is number => value !== null && Number.isFinite(value));
  const mean = finite.reduce((sum, value) => sum + value, 0) / Math.max(1, finite.length);
  const variance = finite.length > 1 ? finite.reduce((sum, value) => sum + (value - mean) ** 2, 0) / (finite.length - 1) : 0;
  const deviation = Math.sqrt(variance);
  return values.map((value, index) => {
    if (value === null || !Number.isFinite(value)) return null;
    const amount = spec.valueType === 'fixedVal' ? Math.abs(spec.value ?? 0)
      : spec.valueType === 'percentage' ? Math.abs((value * (spec.value ?? 0)) / 100)
        : spec.valueType === 'stdErr' ? deviation / Math.sqrt(Math.max(1, finite.length))
          : spec.valueType === 'stdDev' ? deviation * (spec.value ?? 1)
            : undefined;
    if (spec.valueType === 'stdDev') {
      const center = mean;
      return { low: spec.type === 'plus' ? value : center - (amount ?? 0), high: spec.type === 'minus' ? value : center + (amount ?? 0) };
    }
    const plus = amount ?? Math.abs(spec.plus?.[index] ?? 0);
    const minus = amount ?? Math.abs(spec.minus?.[index] ?? 0);
    return { low: spec.type === 'plus' ? value : value - minus, high: spec.type === 'minus' ? value : value + plus };
  });
}

/** 雷达图第 index 个类目的辐条角度：从正上方起顺时针均分。 */
export function radarPoint(index: number, count: number, radius: number, cx: number, cy: number): { x: number; y: number } {
  const angle = -Math.PI / 2 + (index * 2 * Math.PI) / Math.max(1, count);
  return { x: cx + Math.cos(angle) * radius, y: cy + Math.sin(angle) * radius };
}
