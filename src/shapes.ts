import type { Document, Element } from '@xmldom/xmldom';
import { emuToPx, V_NS, WP_NS, A_NS, OFFICE_REL_NS } from './drawing.js';
import type { RelationshipTarget } from './drawing.js';
import type { ChartInfo, ChartKind, ChartSeriesInfo, CustomGeometry, CustomGeometryCommand, ShapeChildInfo, ShapeInfo, ShapeKind, ShapeTextParagraph } from './types.js';
import { MC_NS, selectAlternateContentBranch } from './xml.js';
import { resolveDrawingColor, resolveDrawingThemeColor, type ThemeInfo } from './styles.js';

const WORD_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const WPS_NS = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';
const DIAGRAM_NS = 'http://schemas.openxmlformats.org/drawingml/2006/diagram';
const DSP_NS = 'http://schemas.microsoft.com/office/drawing/2008/diagram';
const CHART_NS = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const OLE_NS = 'urn:schemas-microsoft-com:office:office';

function descendants(parent: Element, namespace: string, localName: string): Element[] {
  return Array.from(parent.getElementsByTagNameNS(namespace, localName));
}

function descendantsInNamespace(parent: Document, namespace: string, localName: string): Element[] {
  return Array.from(parent.getElementsByTagNameNS(namespace, localName));
}

function first(parent: Element | undefined, namespace: string, localName: string): Element | undefined {
  return parent ? descendants(parent, namespace, localName)[0] : undefined;
}

function numberAttribute(element: Element | undefined, name: string): number {
  const value = element?.getAttribute(name);
  return value && Number.isFinite(Number(value)) && Number(value) > 0 ? Number(value) : 0;
}

function shapeKind(uri: string | undefined, hasTextContent: boolean): ShapeKind {
  if (uri === WPS_NS || hasTextContent) return 'textbox';
  if (uri === DIAGRAM_NS || uri?.includes('/diagram')) return 'smartArt';
  if (uri === CHART_NS || uri?.includes('/chart')) return 'chart';
  if (uri === OLE_NS || uri?.includes('/ole')) return 'ole';
  return 'shape';
}

function textContent(node: Element): boolean {
  return descendants(node, WORD_NS, 'txbxContent').length > 0;
}

function cssColor(value: string | undefined): string | undefined {
  if (!value) return undefined;
  if (/^#[\da-f]{3}(?:[\da-f]{3})?$/i.test(value)) return value;
  return /^(?:black|white|red|green|blue|yellow|gray|grey|silver|navy|teal|olive|maroon|purple|fuchsia|lime|aqua)$/i.test(value)
    ? value.toLowerCase() : undefined;
}

/** DrawingML 的颜色元素名，用来从填充 / 线条 / 渐变停止点里挑出那个颜色子元素。 */
const COLOR_ELEMENTS = new Set(['srgbClr', 'schemeClr', 'sysClr', 'prstClr', 'scrgbClr', 'hslClr']);

function colorChild(parent: Element | undefined): Element | undefined {
  if (!parent) return undefined;
  return Array.from(parent.childNodes).find((child) =>
    child.nodeType === 1 && (child as Element).namespaceURI === A_NS && COLOR_ELEMENTS.has((child as Element).localName ?? '')) as Element | undefined;
}

function colorWithAlpha(theme: ThemeInfo, color: Element | undefined, placeholder?: Element): string | undefined {
  const hex = resolveDrawingColor(theme, color, placeholder);
  if (!hex) return undefined;
  // 占位色上的 alpha 也算：主题格式表写 phClr，透明度可能在引用处（fillRef 的颜色）上。
  const alphaOf = (element: Element | undefined) => element ? descendants(element, A_NS, 'alpha')[0]?.getAttribute('val') : undefined;
  const alpha = alphaOf(color) ?? (color?.localName === 'schemeClr' && color.getAttribute('val') === 'phClr' ? alphaOf(placeholder) : undefined);
  if (alpha === undefined || alpha === null || !Number.isFinite(Number(alpha))) return `#${hex}`;
  const opacity = Math.max(0, Math.min(1, Number(alpha) / 100000));
  const [r, g, b] = hex.match(/../g)!.map((channel: string) => parseInt(channel, 16));
  return `rgba(${r}, ${g}, ${b}, ${opacity})`;
}

function readShapeAppearance(spPr: Element | undefined, theme: ThemeInfo, relationships: Map<string, RelationshipTarget>, style?: Element): Pick<ShapeInfo, 'fill' | 'line' | 'geometry' | 'customGeometry' | 'shadow'> {
  const presetGeom = first(spPr, A_NS, 'prstGeom');
  const geometry = presetGeom?.getAttribute('prst') ?? undefined;
  const fill = readDrawingFill(spPr, theme, relationships, style);
  const line = readLine(spPr, theme, style);
  const customGeometry = parseCustomGeometry(spPr);
  const shadow = readShadow(spPr, theme, style);
  return {
    ...(geometry ? { geometry } : {}), ...(fill ? { fill } : {}), ...(line ? { line } : {}),
    ...(customGeometry ? { customGeometry } : {}), ...(shadow ? { shadow } : {}),
  };
}

/**
 * 外阴影。`spPr` 里写了 `a:effectLst`（哪怕是空的）就以它为准——空的 effectLst 是「显式没有效果」；
 * 没写才看 `effectRef idx` 指向的主题效果样式（`effectStyleLst`，从 1 数起，0 是没有）。只画
 * `a:outerShdw`：内阴影、发光、柔化边缘、倒影、三维不画。
 */
function readShadow(spPr: Element | undefined, theme: ThemeInfo, style?: Element): ShapeInfo['shadow'] {
  const own = spPr ? Array.from(spPr.childNodes).find((node) =>
    node.nodeType === 1 && (node as Element).namespaceURI === A_NS && (node as Element).localName === 'effectLst') as Element | undefined : undefined;
  const effectRef = style ? descendants(style, A_NS, 'effectRef')[0] : undefined;
  const index = Number(effectRef?.getAttribute('idx'));
  const themed = !own && Number.isSafeInteger(index) && index >= 1 ? theme.formatScheme?.effects[index - 1] : undefined;
  const list = own ?? (themed ? first(themed, A_NS, 'effectLst') : undefined);
  const shadow = list ? first(list, A_NS, 'outerShdw') : undefined;
  if (!shadow) return undefined;
  const number = (name: string) => {
    const value = Number(shadow.getAttribute(name));
    return Number.isFinite(value) ? value : 0;
  };
  const distance = emuToPx(number('dist'));
  // 方向角单位是 1/60000 度，0° 朝右、顺时针。
  const angle = (number('dir') / 60000) * Math.PI / 180;
  const color = colorWithAlpha(theme, colorChild(shadow), colorChild(effectRef)) ?? 'rgba(0, 0, 0, 0.35)';
  return {
    dxPx: Number((Math.cos(angle) * distance).toFixed(3)),
    dyPx: Number((Math.sin(angle) * distance).toFixed(3)),
    blurPx: Number(emuToPx(number('blurRad')).toFixed(3)),
    color,
  };
}

const FILL_ELEMENTS = new Set(['noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill']);

function readDrawingFill(spPr: Element | undefined, theme: ThemeInfo, relationships: Map<string, RelationshipTarget>, style?: Element): NonNullable<ShapeInfo['fill']> | undefined {
  const fillNode = spPr && Array.from(spPr.childNodes).find((child) =>
    child.nodeType === 1 && (child as Element).namespaceURI === A_NS &&
    FILL_ELEMENTS.has((child as Element).localName ?? '')) as Element | undefined;
  if (fillNode) return fillFromNode(fillNode, theme, relationships);
  const fillRef = style && descendants(style, A_NS, 'fillRef')[0];
  if (!fillRef) return undefined;
  const placeholder = colorChild(fillRef);
  // idx 指向主题格式表：0 是「没有填充」，1..999 是 fillStyleLst，1001 起是 bgFillStyleLst。
  const index = Number(fillRef.getAttribute('idx'));
  if (index === 0) return { type: 'none' };
  const scheme = theme.formatScheme;
  const themed = !Number.isSafeInteger(index) ? undefined
    : index >= 1001 ? scheme?.backgroundFills[index - 1001] : scheme?.fills[index - 1];
  if (themed && FILL_ELEMENTS.has(themed.localName ?? '')) return fillFromNode(themed, theme, relationships, placeholder);
  // 主题里没有格式表（或 idx 越界）时退回引用处的颜色当实心填充，这是原来的行为。
  const resolved = colorWithAlpha(theme, placeholder);
  return resolved ? { type: 'solid', color: resolved } : undefined;
}

function fillFromNode(fillNode: Element, theme: ThemeInfo, relationships: Map<string, RelationshipTarget>, placeholder?: Element): NonNullable<ShapeInfo['fill']> {
  if (fillNode.localName === 'noFill') return { type: 'none' };
  if (fillNode.localName === 'solidFill') return { type: 'solid', color: colorWithAlpha(theme, colorChild(fillNode), placeholder) };
  if (fillNode.localName === 'pattFill') {
    // 图案填充画不出图案，取前景色当实心——比留白更接近原样。
    return { type: 'solid', color: colorWithAlpha(theme, colorChild(first(fillNode, A_NS, 'fgClr')), placeholder) };
  }
  if (fillNode.localName === 'gradFill') {
    const stops = descendants(fillNode, A_NS, 'gs').flatMap((stop) => {
      const position = Number(stop.getAttribute('pos'));
      const resolved = colorWithAlpha(theme, colorChild(stop), placeholder);
      return Number.isFinite(position) && resolved ? [{ position: Math.max(0, Math.min(1, position / 100000)), color: resolved }] : [];
    });
    const angle = Number(descendants(fillNode, A_NS, 'lin')[0]?.getAttribute('ang'));
    return { type: 'gradient', stops, ...(Number.isFinite(angle) ? { angle: angle / 60000 } : {}) };
  }
  const blip = descendants(fillNode, A_NS, 'blip')[0];
  const relationshipId = blip?.getAttributeNS(OFFICE_REL_NS, 'embed');
  const relationship = relationshipId ? relationships.get(relationshipId) : undefined;
  return {
    type: 'picture',
    ...(relationship?.mode !== 'External' && relationship?.partPath ? { imagePartPath: relationship.partPath } : {}),
  };
}

const LINE_DASHES: Record<string, string> = {
  dash: '6 3', dashDot: '6 3 1 3', dot: '1 3', lgDash: '10 3', lgDashDot: '10 3 1 3',
  lgDashDotDot: '10 3 1 3 1 3', sysDash: '4 2', sysDashDot: '4 2 1 2', sysDot: '1 2',
};

function readLine(spPr: Element | undefined, theme: ThemeInfo, style?: Element): ShapeInfo['line'] {
  const line = spPr && Array.from(spPr.childNodes).find((child) =>
    child.nodeType === 1 && (child as Element).namespaceURI === A_NS && (child as Element).localName === 'ln') as Element | undefined;
  const lineRef = style ? descendants(style, A_NS, 'lnRef')[0] : undefined;
  const placeholder = colorChild(lineRef);
  const index = Number(lineRef?.getAttribute('idx'));
  const themed = lineRef && Number.isSafeInteger(index) && index >= 1 ? theme.formatScheme?.lines[index - 1] : undefined;
  if (!line && !lineRef) return undefined;
  // 线条按属性逐项合并：spPr 里的 a:ln 常常只写了宽度，颜色和线型仍来自 lnRef 指向的主题线条。
  const fromLine = (element: Element | undefined): NonNullable<ShapeInfo['line']> & { none?: boolean } => {
    if (!element) return {};
    const fill = Array.from(element.childNodes).find((child) =>
      child.nodeType === 1 && FILL_ELEMENTS.has((child as Element).localName ?? '')) as Element | undefined;
    const width = Number(element.getAttribute('w'));
    const dashName = first(element, A_NS, 'prstDash')?.getAttribute('val');
    const color = fill?.localName === 'solidFill' ? colorWithAlpha(theme, colorChild(fill), placeholder)
      : fill?.localName === 'gradFill' ? colorWithAlpha(theme, colorChild(descendants(fill, A_NS, 'gs')[0]), placeholder)
        : undefined;
    return {
      ...(fill?.localName === 'noFill' ? { none: true } : {}),
      ...(color ? { color } : {}),
      ...(element.hasAttribute('w') && Number.isFinite(width) && width >= 0 ? { widthPx: emuToPx(width) } : {}),
      ...(dashName && LINE_DASHES[dashName] ? { dash: LINE_DASHES[dashName] } : {}),
    };
  };
  // lnRef idx="0"：没有主题线条。
  const base = themed ? fromLine(themed) : lineRef && index !== 0 && !line ? { ...(colorWithAlpha(theme, placeholder) ? { color: colorWithAlpha(theme, placeholder) } : {}) } : {};
  const merged = { ...base, ...fromLine(line) };
  // 显式 <a:noFill/> 的线条就是没有线：不能让主题线条的颜色从下面透上来。
  if (merged.none) return {};
  const { none: _none, ...result } = merged;
  // lnRef idx="0" 而且 spPr 里没有 a:ln：根本没有线条，和「没写」一样。
  return !line && !Object.keys(result).length ? undefined : result;
}

const TEXT_ALIGNMENTS: Record<string, ShapeTextParagraph['alignment']> = { l: 'left', ctr: 'center', r: 'right', just: 'justify', dist: 'justify' };

/**
 * DrawingML 文本体（`a:p` / `a:r` / `a:br` / `a:fld`）读成带格式的行。run 格式取 `a:rPr`：
 * `b` / `i` / `u`（`none` 以外都算下划线）/ `strike`、`sz`（百分之一磅）、`a:latin` 字体、
 * 实心填充的颜色；段落对齐取 `a:pPr/@algn`。`a:lstStyle` 与形状样式里的默认字体不展开。
 */
function readTextBodyParagraphs(txBody: Element | undefined, theme: ThemeInfo): ShapeTextParagraph[] {
  if (!txBody) return [];
  return Array.from(txBody.childNodes)
    .filter((node): node is Element => node.nodeType === 1 && (node as Element).namespaceURI === A_NS && (node as Element).localName === 'p')
    .map((paragraph) => {
      const lines: ShapeTextParagraph['lines'] = [[]];
      for (const node of Array.from(paragraph.childNodes)) {
        if (node.nodeType !== 1) continue;
        const element = node as Element;
        if (element.localName === 'br') { lines.push([]); continue; }
        if (element.localName !== 'r' && element.localName !== 'fld') continue;
        const content = first(element, A_NS, 't')?.textContent ?? '';
        if (!content) continue;
        const props = first(element, A_NS, 'rPr');
        const on = (name: string) => ['1', 'true'].includes(props?.getAttribute(name) ?? '');
        const size = Number(props?.getAttribute('sz'));
        const underline = props?.getAttribute('u');
        const strike = props?.getAttribute('strike');
        const color = colorWithAlpha(theme, colorChild(first(props, A_NS, 'solidFill')));
        const font = first(props, A_NS, 'latin')?.getAttribute('typeface') ?? first(props, A_NS, 'ea')?.getAttribute('typeface');
        lines.at(-1)!.push({
          text: content,
          ...(on('b') ? { bold: true } : {}),
          ...(on('i') ? { italic: true } : {}),
          ...(underline && underline !== 'none' ? { underline: true } : {}),
          ...(strike && strike !== 'noStrike' ? { strike: true } : {}),
          ...(Number.isFinite(size) && size > 0 ? { fontSize: size / 100 } : {}),
          // 主题字体占位（+mn-lt 之类）不是字体名，交给浏览器默认。
          ...(font && !font.startsWith('+') ? { fontFamily: font } : {}),
          ...(color ? { color } : {}),
        });
      }
      const alignment = TEXT_ALIGNMENTS[first(paragraph, A_NS, 'pPr')?.getAttribute('algn') ?? ''];
      return { ...(alignment ? { alignment } : {}), lines };
    })
    .filter((paragraph) => paragraph.lines.some((line) => line.length));
}

function direct(parent: Element | undefined, namespace: string, localName: string): Element | undefined {
  if (!parent) return undefined;
  for (let child = parent.firstChild; child; child = child.nextSibling) {
    if (child.nodeType === 1 && (child as Element).namespaceURI === namespace && (child as Element).localName === localName) return child as Element;
  }
  return undefined;
}

function cacheValues(cache: Element | undefined, numeric: boolean): Array<number | string | null> {
  if (!cache) return [];
  const points = descendants(cache, CHART_NS, 'pt');
  const count = Math.max(Number(cache.getAttribute('ptCount') ?? 0), ...points.map((point) => Number(point.getAttribute('idx')) + 1), 0);
  const values: Array<number | string | null> = Array.from({ length: count }, () => null);
  for (const point of points) {
    const index = Number(point.getAttribute('idx'));
    const value = descendants(point, CHART_NS, 'v')[0]?.textContent ?? '';
    if (!Number.isInteger(index) || index < 0) continue;
    values[index] = numeric ? (Number.isFinite(Number(value)) ? Number(value) : null) : value;
  }
  return values;
}

function chartSeriesFill(ser: Element, theme: ThemeInfo, relationships: Map<string, RelationshipTarget>): Pick<ShapeInfo, 'fill' | 'line'> {
  return readShapeAppearance(direct(ser, CHART_NS, 'spPr'), theme, relationships);
}

function chartSeriesAppearance(ser: Element, index: number, theme: ThemeInfo, relationships: Map<string, RelationshipTarget>): Pick<ShapeInfo, 'fill' | 'line'> {
  const appearance = chartSeriesFill(ser, theme, relationships);
  if (appearance.fill) return appearance;
  const color = resolveDrawingThemeColor(theme, `accent${(index % 6) + 1}`);
  return color ? { ...appearance, fill: { type: 'solid', color: color.startsWith('#') ? color : `#${color}` } } : appearance;
}

function chartPointFills(ser: Element, count: number, theme: ThemeInfo, relationships: Map<string, RelationshipTarget>): Array<ShapeInfo['fill'] | undefined> {
  return Array.from({ length: count }, (_, index) => {
    const point = descendants(ser, CHART_NS, 'dPt').find((candidate) => Number(direct(candidate, CHART_NS, 'idx')?.getAttribute('val') ?? candidate.getAttribute('idx')) === index);
    const appearance = point ? chartSeriesFill(point, theme, relationships) : {};
    if (appearance.fill) return appearance.fill;
    const color = resolveDrawingThemeColor(theme, `accent${(index % 6) + 1}`);
    return color ? { type: 'solid' as const, color: color.startsWith('#') ? color : `#${color}` } : undefined;
  });
}

function readChartInfo(
  graphicData: Element | undefined,
  relationships: Map<string, RelationshipTarget>,
  theme: ThemeInfo,
  getPartDocument?: (path: string) => Document | undefined,
): ChartInfo | undefined {
  if (!graphicData || !getPartDocument) return undefined;
  const chartId = first(graphicData, CHART_NS, 'chart')?.getAttributeNS(OFFICE_REL_NS, 'id')
    ?? first(graphicData, CHART_NS, 'chart')?.getAttribute('r:id');
  const chartPath = chartId ? relationships.get(chartId)?.partPath : undefined;
  const chartDocument = chartPath ? getPartDocument(chartPath) : undefined;
  const plotArea = chartDocument && first(chartDocument.documentElement as Element, CHART_NS, 'plotArea');
  if (!plotArea) return undefined;
  const chartTypes = Array.from(plotArea.childNodes).filter((node) =>
    node.nodeType === 1 && (node as Element).namespaceURI === CHART_NS &&
    (node as Element).localName?.endsWith('Chart')) as Element[];
  const supported = chartTypes.filter((group) => CHART_KINDS[group.localName ?? ''] !== undefined);
  const chartType = supported[0] ?? chartTypes[0];
  if (!chartType) return undefined;
  const localName = chartType.localName ?? '';
  const kind: ChartInfo['kind'] = CHART_KINDS[localName] ?? 'unsupported';
  const titleNode = direct(chartDocument.documentElement as Element, CHART_NS, 'chart')
    && direct(direct(chartDocument.documentElement as Element, CHART_NS, 'chart'), CHART_NS, 'title');
  const title = titleNode ? descendants(titleNode, A_NS, 't').map((node) => node.textContent ?? '').join('').trim() || undefined : undefined;
  const xy = (group: Element) => ['scatterChart', 'bubbleChart'].includes(group.localName ?? '');
  const firstSeries = first(chartType, CHART_NS, 'ser');
  const categoryContainer = xy(chartType) ? direct(firstSeries, CHART_NS, 'xVal') : direct(firstSeries, CHART_NS, 'cat');
  const firstCategoryRef = direct(categoryContainer, CHART_NS, 'strRef')
    ?? direct(categoryContainer, CHART_NS, 'numRef') ?? direct(categoryContainer, CHART_NS, 'strLit');
  const categories = cacheValues(
    firstCategoryRef
      ? firstCategoryRef.localName === 'strLit' ? firstCategoryRef
        : direct(firstCategoryRef, CHART_NS, firstCategoryRef.localName === 'numRef' ? 'numCache' : 'strCache')
      : undefined,
    false,
  ).map((value) => typeof value === 'string' ? value : '');

  // 次坐标轴：每个图表组用 c:axId 指向自己的一对轴。第一个组的数值轴算主轴，其余组如果指向
  // 另一根数值轴就是次坐标轴。散点 / 气泡图两根轴都是 valAx，取竖向（axPos 为 l / r）那根。
  const valueAxes = new Map(Array.from(plotArea.childNodes)
    .filter((node): node is Element => node.nodeType === 1 && (node as Element).localName === 'valAx')
    .map((axis) => [direct(axis, CHART_NS, 'axId')?.getAttribute('val') ?? '', axis] as const));
  const valueAxisOf = (group: Element): string | undefined => {
    const ids = Array.from(group.childNodes).filter((node): node is Element => node.nodeType === 1 && (node as Element).localName === 'axId')
      .map((node) => node.getAttribute('val') ?? '');
    const candidates = ids.filter((id) => valueAxes.has(id));
    return candidates.find((id) => ['l', 'r'].includes(direct(valueAxes.get(id), CHART_NS, 'axPos')?.getAttribute('val') ?? '')) ?? candidates[0];
  };
  const primaryAxis = valueAxisOf(chartType);
  let secondaryAxis: Element | undefined;

  const numbers = (container: Element | undefined): Array<number | null> | undefined => {
    const ref = container && (direct(container, CHART_NS, 'numRef') ?? direct(container, CHART_NS, 'numLit'));
    const cache = ref && (direct(ref, CHART_NS, 'numCache') ?? (ref.localName === 'numLit' ? ref : undefined));
    return cache ? cacheValues(cache, true) as Array<number | null> : undefined;
  };
  let seriesIndex = 0;
  const series = supported.flatMap((group) => {
    const groupKind = CHART_KINDS[group.localName ?? '']!;
    const axisId = valueAxisOf(group);
    const axis = axisId && primaryAxis && axisId !== primaryAxis ? 'secondary' as const : 'primary' as const;
    if (axis === 'secondary') secondaryAxis ??= valueAxes.get(axisId!);
    return descendants(group, CHART_NS, 'ser').flatMap((ser) => {
      const values = numbers(direct(ser, CHART_NS, 'val') ?? direct(ser, CHART_NS, 'yVal'));
      if (!values) return [];
      const xValues = xy(group) ? numbers(direct(ser, CHART_NS, 'xVal')) : undefined;
      const bubbleSizes = group.localName === 'bubbleChart' ? numbers(direct(ser, CHART_NS, 'bubbleSize')) : undefined;
      const nameRef = direct(direct(ser, CHART_NS, 'tx'), CHART_NS, 'strRef');
      const name = nameRef ? String(cacheValues(direct(nameRef, CHART_NS, 'strCache'), false)[0] ?? '') || undefined
        : direct(direct(ser, CHART_NS, 'tx'), CHART_NS, 'v')?.textContent ?? undefined;
      // 自动配色按 c:idx（Word 写的就是全图唯一的序号），没有才按出现顺序。
      const idx = Number(direct(ser, CHART_NS, 'idx')?.getAttribute('val'));
      const appearance = chartSeriesAppearance(ser, Number.isSafeInteger(idx) && idx >= 0 ? idx : seriesIndex, theme, relationships);
      seriesIndex++;
      const pointFills = groupKind === 'pie' || groupKind === 'doughnut'
        ? chartPointFills(ser, values.length, theme, relationships)
        : undefined;
      const trendlines = Array.from(ser.childNodes)
        .filter((node): node is Element => node.nodeType === 1 && (node as Element).localName === 'trendline')
        .flatMap((trendline) => {
          const type = direct(trendline, CHART_NS, 'trendlineType')?.getAttribute('val') ?? 'linear';
          if (!TRENDLINE_TYPES.has(type)) return [];
          const order = Number(direct(trendline, CHART_NS, 'order')?.getAttribute('val'));
          const period = Number(direct(trendline, CHART_NS, 'period')?.getAttribute('val'));
          const label = direct(trendline, CHART_NS, 'name')?.textContent?.trim();
          return [{
            type: type as NonNullable<ChartSeriesInfo['trendlines']>[number]['type'],
            ...(Number.isFinite(order) ? { order } : {}),
            ...(Number.isFinite(period) ? { period } : {}),
            ...(label ? { name: label } : {}),
          }];
        });
      const errorBars = Array.from(ser.childNodes)
        .filter((node): node is Element => node.nodeType === 1 && (node as Element).localName === 'errBars')
        .flatMap((bars) => {
          const valueType = direct(bars, CHART_NS, 'errValType')?.getAttribute('val') ?? 'fixedVal';
          if (!ERROR_VALUE_TYPES.has(valueType)) return [];
          const barType = direct(bars, CHART_NS, 'errBarType')?.getAttribute('val') ?? 'both';
          const value = Number(direct(bars, CHART_NS, 'val')?.getAttribute('val'));
          const plus = numbers(direct(bars, CHART_NS, 'plus'));
          const minus = numbers(direct(bars, CHART_NS, 'minus'));
          return [{
            direction: (direct(bars, CHART_NS, 'errDir')?.getAttribute('val') === 'x' ? 'x' : 'y') as 'x' | 'y',
            type: (['plus', 'minus'].includes(barType) ? barType : 'both') as 'both' | 'plus' | 'minus',
            valueType: valueType as NonNullable<ChartSeriesInfo['errorBars']>[number]['valueType'],
            ...(Number.isFinite(value) ? { value } : {}),
            ...(plus ? { plus } : {}),
            ...(minus ? { minus } : {}),
          }];
        });
      const entry: ChartSeriesInfo = {
        name, values, type: groupKind, axis,
        ...(xValues ? { xValues } : {}),
        ...(bubbleSizes ? { bubbleSizes } : {}),
        ...(pointFills ? { pointFills } : {}),
        ...(trendlines.length ? { trendlines } : {}),
        ...(errorBars.length ? { errorBars } : {}),
        ...appearance,
      };
      return [entry];
    });
  });
  const grouping = direct(chartType, CHART_NS, 'grouping')?.getAttribute('val') as ChartInfo['grouping'] | null;
  const barGroup = supported.find((group) => CHART_KINDS[group.localName ?? ''] === 'bar');
  const barDirection = barGroup
    ? (direct(barGroup, CHART_NS, 'barDir')?.getAttribute('val') === 'bar' ? 'bar' : 'col')
    : undefined;
  const legendPosition = descendants(chartDocument.documentElement as Element, CHART_NS, 'legendPos')[0]?.getAttribute('val');
  const legend = ['l', 'r', 't', 'b', 'tr'].includes(legendPosition ?? '') ? { position: legendPosition as 'l' | 'r' | 't' | 'b' | 'tr' } : undefined;
  const primaryValueAxis = primaryAxis ? valueAxes.get(primaryAxis) : direct(plotArea, CHART_NS, 'valAx');
  const visible = (axis: Element | undefined) => direct(axis, CHART_NS, 'delete')?.getAttribute('val') !== '1';
  const axes = {
    category: { visible: direct(plotArea, CHART_NS, 'catAx') ? visible(direct(plotArea, CHART_NS, 'catAx')) : true },
    value: {
      visible: primaryValueAxis ? visible(primaryValueAxis) : true,
      majorGridlines: Boolean(primaryValueAxis && direct(primaryValueAxis, CHART_NS, 'majorGridlines')),
    },
    ...(secondaryAxis ? { secondaryValue: { visible: visible(secondaryAxis) } } : {}),
  };
  const radarStyle = localName === 'radarChart' ? direct(chartType, CHART_NS, 'radarStyle')?.getAttribute('val') : undefined;
  return {
    kind,
    ...(title ? { title } : {}),
    categories,
    series,
    ...(barDirection ? { barDirection } : {}),
    ...(grouping ? { grouping } : {}),
    ...(legend ? { legend } : {}),
    axes,
    ...(localName.includes('3D') ? { threeD: true } : {}),
    ...(radarStyle && ['standard', 'marker', 'filled'].includes(radarStyle) ? { radarStyle: radarStyle as ChartInfo['radarStyle'] } : {}),
    ...(localName === 'stockChart' ? {
      stock: { hiLowLines: Boolean(direct(chartType, CHART_NS, 'hiLowLines')), upDownBars: Boolean(direct(chartType, CHART_NS, 'upDownBars')) },
    } : {}),
  };
}

/** 图表组元素名 → 画法。3D 图按对应的 2D 图画；曲面图（等高线）不在这里，退化为占位。 */
const CHART_KINDS: Record<string, ChartKind> = {
  barChart: 'bar', bar3DChart: 'bar', lineChart: 'line', line3DChart: 'line',
  pieChart: 'pie', pie3DChart: 'pie', ofPieChart: 'pie', doughnutChart: 'doughnut',
  areaChart: 'area', area3DChart: 'area', scatterChart: 'scatter',
  radarChart: 'radar', bubbleChart: 'bubble', stockChart: 'stock',
};
const TRENDLINE_TYPES = new Set(['linear', 'exp', 'log', 'poly', 'power', 'movingAvg']);
const ERROR_VALUE_TYPES = new Set(['fixedVal', 'percentage', 'stdDev', 'stdErr', 'cust']);

function parseCustomGeometry(spPr: Element | undefined): CustomGeometry | undefined {
  const custom = spPr && descendants(spPr, A_NS, 'custGeom')[0];
  const path = custom && descendants(custom, A_NS, 'path')[0];
  if (!path) return undefined;
  const width = Number(path.getAttribute('w'));
  const height = Number(path.getAttribute('h'));
  if (!Number.isFinite(width) || !Number.isFinite(height) || width <= 0 || height <= 0) return undefined;
  const commands: CustomGeometryCommand[] = [];
  for (let child = path.firstChild; child; child = child.nextSibling) {
    if (child.nodeType !== 1) continue;
    const command = child as Element;
    if (command.namespaceURI !== A_NS) continue;
    if (command.localName === 'close') {
      commands.push({ type: 'close' });
      continue;
    }
    if (!['moveTo', 'lnTo', 'cubicBezTo'].includes(command.localName ?? '')) return undefined;
    const points = descendants(command, A_NS, 'pt').map((point) => [Number(point.getAttribute('x')), Number(point.getAttribute('y'))]);
    if (points.some(([x, y]) => !Number.isFinite(x) || !Number.isFinite(y))) return undefined;
    if (command.localName === 'cubicBezTo' && points.length === 3) {
      commands.push({ type: 'cubicBezTo', x1: points[0]![0]!, y1: points[0]![1]!, x2: points[1]![0]!, y2: points[1]![1]!, x: points[2]![0]!, y: points[2]![1]! });
    } else if (['moveTo', 'lnTo'].includes(command.localName ?? '') && points.length === 1) {
      commands.push({ type: command.localName as 'moveTo' | 'lnTo', x: points[0]![0]!, y: points[0]![1]! });
    } else return undefined;
  }
  return commands.length ? { width, height, commands } : undefined;
}

function readSmartArtChildren(
  graphicData: Element | undefined,
  relationships: Map<string, RelationshipTarget>,
  theme: ThemeInfo,
  getPartDocument?: (path: string) => Document | undefined,
  getPartRelationships?: (path: string) => Map<string, RelationshipTarget>,
): ShapeChildInfo[] | undefined {
  if (!graphicData || !getPartDocument) return undefined;
  const relIds = first(graphicData, DIAGRAM_NS, 'relIds');
  const dataId = relIds?.getAttributeNS(OFFICE_REL_NS, 'dm') ?? relIds?.getAttribute('r:dm');
  const dataPath = dataId ? relationships.get(dataId)?.partPath : undefined;
  const dataDocument = dataPath ? getPartDocument(dataPath) : undefined;
  const dataModelExt = dataDocument
    ? descendantsInNamespace(dataDocument, DSP_NS, 'dataModelExt')[0]
    : undefined;
  const drawingId = dataModelExt?.getAttribute('relId')
    ?? dataModelExt?.getAttributeNS(OFFICE_REL_NS, 'id')
    ?? dataModelExt?.getAttribute('r:id');
  const relation = (drawingId ? relationships.get(drawingId) : undefined)
    ?? (dataPath ? [...(getPartRelationships?.(dataPath) ?? new Map()).values()].find((entry) =>
      entry.type?.endsWith('/diagramDrawing')) : undefined)
    ?? (() => {
      const candidates = [...relationships.values()].filter((entry) =>
        entry.type?.endsWith('/diagramDrawing') ||
        entry.partPath && /diagramDrawing/i.test(entry.target ?? ''));
      return candidates.length === 1 ? candidates[0] : undefined;
    })();
  const drawingPath = relation?.partPath;
  if (!drawingPath) return undefined;
  const drawing = getPartDocument(drawingPath);
  if (!drawing) return undefined;
  const drawingRelationships = getPartRelationships?.(drawingPath) ?? new Map();
  const shapes = descendants(drawing.documentElement!, DSP_NS, 'sp');
  if (!shapes.length) return undefined;
  return shapes.map((shape) => {
    const spPr = first(shape, DSP_NS, 'spPr');
    const style = first(shape, DSP_NS, 'style');
    const appearance = readShapeAppearance(spPr, theme, drawingRelationships, style);
    const xfrm = first(spPr, A_NS, 'xfrm');
    const off = first(xfrm, A_NS, 'off');
    const ext = first(xfrm, A_NS, 'ext');
    const text = descendants(shape, A_NS, 't').map((node) => node.textContent ?? '').join('');
    const paragraphs = readTextBodyParagraphs(first(shape, DSP_NS, 'txBody'), theme);
    const rotation = Number(xfrm?.getAttribute('rot'));
    return {
      offsetXPx: emuToPx(numberAttribute(off, 'x')),
      offsetYPx: emuToPx(numberAttribute(off, 'y')),
      widthPx: emuToPx(numberAttribute(ext, 'cx')),
      heightPx: emuToPx(numberAttribute(ext, 'cy')),
      ...appearance,
      ...(Number.isFinite(rotation) && rotation ? { rotation: rotation / 60000 } : {}),
      ...(xfrm?.getAttribute('flipH') === '1' || xfrm?.getAttribute('flipH') === 'true' ? { flipH: true } : {}),
      ...(xfrm?.getAttribute('flipV') === '1' || xfrm?.getAttribute('flipV') === 'true' ? { flipV: true } : {}),
      ...(text ? { text } : {}),
      ...(paragraphs.length ? { paragraphs } : {}),
    };
  });
}

function readDrawingShape(
  node: Element,
  paragraph: number,
  run: number,
  ordinal: number,
  relationships: Map<string, RelationshipTarget>,
  theme: ThemeInfo,
  getPartDocument?: (path: string) => Document | undefined,
  getPartRelationships?: (path: string) => Map<string, RelationshipTarget>,
): ShapeInfo[] {
  const result: ShapeInfo[] = [];
  const containers = [
    ...descendants(node, WP_NS, 'inline'),
    ...descendants(node, WP_NS, 'anchor'),
  ];
  for (const container of containers) {
    const placement = container.localName === 'anchor' ? 'floating' : 'inline';
    const extent = first(container, WP_NS, 'extent');
    const graphicData = first(first(container, A_NS, 'graphic'), A_NS, 'graphicData');
    const chosen = first(graphicData, MC_NS, 'Choice');
    const source = chosen ?? graphicData ?? container;
    const hasTextContent = textContent(source);
    const uri = graphicData?.getAttribute('uri') ?? undefined;
    const docPr = first(container, WP_NS, 'docPr');
    const wps = first(source, WPS_NS, 'wsp');
    const spPr = first(wps, WPS_NS, 'spPr');
    const shapeStyle = first(wps, WPS_NS, 'style');
    const xfrm = first(spPr, A_NS, 'xfrm');
    const rotation = Number(xfrm?.getAttribute('rot'));
    const appearance = readShapeAppearance(spPr, theme, relationships, shapeStyle);
    const chart = shapeKind(uri, hasTextContent) === 'chart'
      ? readChartInfo(graphicData, relationships, theme, getPartDocument)
      : undefined;
    const presetGeom = first(spPr, A_NS, 'prstGeom');
    const adjustments = first(presetGeom, A_NS, 'avLst');
    const parsedAdjustments = adjustments ? descendants(adjustments, A_NS, 'gd').flatMap((gd) => {
      const value = Number(gd.getAttribute('fmla')?.replace(/^val\s+/, ''));
      return gd.getAttribute('name') && Number.isFinite(value) ? [{ name: gd.getAttribute('name')!, value: value / 100000 }] : [];
    }) : [];
    const wrap = placement === 'floating'
      ? (first(container, WP_NS, 'wrapNone') ? 'none'
        : first(container, WP_NS, 'wrapSquare') ? 'square'
        : first(container, WP_NS, 'wrapTight') ? 'tight'
        : first(container, WP_NS, 'wrapThrough') ? 'through'
        : first(container, WP_NS, 'wrapTopAndBottom') ? 'topAndBottom' : undefined)
      : undefined;
    const id = docPr?.getAttribute('id') ?? `${paragraph}:${run}:${ordinal + result.length}`;
    const fill = appearance.fill;
    const line = appearance.line;
    const hasAppearance = fill !== undefined || line !== undefined;
    result.push({
      id,
      paragraph,
      run,
      kind: shapeKind(uri, hasTextContent),
      form: 'drawingml',
      name: docPr?.getAttribute('name') ?? undefined,
      alt: docPr?.getAttribute('descr') ?? undefined,
      title: docPr?.getAttribute('title') ?? undefined,
      widthPx: emuToPx(numberAttribute(extent, 'cx')),
      heightPx: emuToPx(numberAttribute(extent, 'cy')),
      placement,
      wrap,
      hasTextContent,
      ...appearance,
      ...(fill ? { fill } : hasAppearance ? {} : { fill: { type: 'solid' as const, color: '#f7f9fd' } }),
      ...(line ? { line } : hasAppearance ? {} : { line: { color: '#c7d3e5', widthPx: 1 } }),
      ...(Number.isFinite(rotation) && rotation ? { rotation: rotation / 60000 } : {}),
      ...(xfrm?.getAttribute('flipH') === '1' || xfrm?.getAttribute('flipH') === 'true' ? { flipH: true } : {}),
      ...(xfrm?.getAttribute('flipV') === '1' || xfrm?.getAttribute('flipV') === 'true' ? { flipV: true } : {}),
      ...(parsedAdjustments.length ? { adjustments: parsedAdjustments } : {}),
      ...(appearance.customGeometry ? { customGeometry: appearance.customGeometry } : {}),
      ...(chart ? { chart } : {}),
      ...(shapeKind(uri, hasTextContent) === 'smartArt'
        ? { children: readSmartArtChildren(graphicData, relationships, theme, getPartDocument, getPartRelationships) }
        : {}),
    });
  }
  return result;
}

function vmlPreset(type: string | undefined): string | undefined {
  if (!type) return undefined;
  const simple: Record<string, string> = { rect: 'rect', roundRect: 'roundRect', ellipse: 'ellipse', oval: 'ellipse', line: 'line', rightArrow: 'rightArrow' };
  if (simple[type]) return simple[type];
  const code = Number(type.match(/_x0000_t(\d+)/i)?.[1]);
  return ({ 1: 'rect', 2: 'roundRect', 3: 'ellipse', 4: 'diamond', 5: 'line', 202: 'wedgeRectCallout', 203: 'cloudCallout' } as Record<number, string>)[code];
}

function readVmlShape(node: Element, paragraph: number, run: number, ordinal: number): ShapeInfo[] {
  const shapes = Array.from(node.getElementsByTagNameNS(V_NS, '*'))
    .filter((shape) => ['shape', 'rect', 'oval', 'line'].includes(shape.localName ?? ''));
  return shapes.map((shape, index) => {
    const style = shape.getAttribute('style') ?? '';
    const width = style.match(/(?:^|;)width:\s*([\d.]+)pt/i);
    const height = style.match(/(?:^|;)height:\s*([\d.]+)pt/i);
    const hasTextContent = textContent(shape);
    const type = shape.getAttribute('type') || shape.localName || undefined;
    const fillColor = cssColor(shape.getAttribute('fillcolor') ?? descendants(shape, V_NS, 'fill')[0]?.getAttribute('color') ?? undefined);
    const strokeColor = cssColor(shape.getAttribute('strokecolor') ?? descendants(shape, V_NS, 'stroke')[0]?.getAttribute('color') ?? undefined);
    const strokeWeight = (shape.getAttribute('strokeweight') ?? descendants(shape, V_NS, 'stroke')[0]?.getAttribute('weight') ?? '').match(/^([\d.]+)(pt|px)?$/i);
    const customPath = descendants(shape, V_NS, 'path').length > 0;
    return {
      id: shape.getAttribute('id') ?? `${paragraph}:${run}:${ordinal + index}`,
      paragraph,
      run,
      kind: hasTextContent ? 'textbox' : 'shape',
      form: 'vml',
      name: shape.getAttribute('id') ?? undefined,
      alt: shape.getAttribute('alt') ?? undefined,
      title: shape.getAttribute('title') ?? undefined,
      widthPx: width ? Number(width[1]) * 96 / 72 : 0,
      heightPx: height ? Number(height[1]) * 96 / 72 : 0,
      placement: 'inline',
      hasTextContent,
      geometry: customPath ? undefined : vmlPreset(type),
      fill: { type: fillColor ? 'solid' : 'none', ...(fillColor ? { color: fillColor } : {}) },
      line: {
        ...(strokeColor ? { color: strokeColor } : {}),
        ...(strokeWeight ? { widthPx: Number(strokeWeight[1]) * (strokeWeight[2]?.toLowerCase() === 'pt' ? 96 / 72 : 1) } : {}),
      },
    };
  });
}

export function readRunShapes(
  runElement: Element,
  paragraph: number,
  run: number,
  sourcePartPath = '',
  relationships: Map<string, RelationshipTarget> = new Map(),
  theme: ThemeInfo = { colors: {}, fonts: {} },
  getPartDocument?: (path: string) => Document | undefined,
  getPartRelationships?: (path: string) => Map<string, RelationshipTarget>,
): ShapeInfo[] {
  const result: ShapeInfo[] = [];
  const elements: Element[] = [];
  for (let child = runElement.firstChild; child; child = child.nextSibling) {
    if (child.nodeType !== 1) continue;
    const element = child as Element;
    if (element.namespaceURI === MC_NS && element.localName === 'AlternateContent') {
      const branch = selectAlternateContentBranch(element);
      if (branch) {
        for (let branchChild = branch.firstChild; branchChild; branchChild = branchChild.nextSibling) {
          if (branchChild.nodeType === 1) elements.push(branchChild as Element);
        }
      }
      continue;
    }
    elements.push(element);
  }
  for (const element of elements) {
    if (element.namespaceURI === WORD_NS && element.localName === 'drawing') {
      result.push(...readDrawingShape(element, paragraph, run, result.length, relationships, theme, getPartDocument, getPartRelationships));
    } else if (element.namespaceURI === WORD_NS && element.localName === 'pict') {
      result.push(...readVmlShape(element, paragraph, run, result.length));
    }
  }
  return result.map((shape, index) => ({
    ...shape,
    id: `${sourcePartPath}:${paragraph}:${run}:${index}:${shape.id}`,
  }));
}

export function shapeTextElements(shape: Element): Element[] {
  return descendants(shape, WORD_NS, 'txbxContent')
    .filter((content) => {
      let ancestor = content.parentNode as Element | null;
      let inChoice = false;
      while (ancestor && ancestor !== shape) {
        if (ancestor.namespaceURI === MC_NS && ancestor.localName === 'Choice') inChoice = true;
        if (ancestor.namespaceURI === MC_NS && ancestor.localName === 'Fallback') return false;
        ancestor = ancestor.parentNode as Element | null;
      }
      return inChoice || !descendants(shape, MC_NS, 'Choice').length;
    })
    .flatMap((content) => descendants(content, WORD_NS, 'p'));
}
