import type { Element } from '@xmldom/xmldom';
import type {
  BorderSide, CellFormat, ColorSchemeMapping, CompatibilitySettings, LatentStyles, ParagraphFormat, RowFormat, RunFormat, Shading, StyleInfo, TabStop, TableFormat, ThemeFontLanguages, ThemeSettings,
} from './types.js';
import { WORD_NS, children, childrenThroughTransparent, wordValue } from './xml.js';
import { compactDefined } from './internal/elements.js';
import { parseCellFormat, parseRowFormat, parseTableFormat } from './table.js';

const DRAWINGML_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';

const DEFAULT_THEME_COLORS: Record<string, string> = {
  dk1: '000000',
  lt1: 'FFFFFF',
  dk2: '1F1F1F',
  lt2: 'EEECE1',
  accent1: '4472C4',
  accent2: 'ED7D31',
  accent3: 'A5A5A5',
  accent4: 'FFC000',
  accent5: '5B9BD5',
  accent6: '70AD47',
  hlink: '0563C1',
  folHlink: '954F72',
};

const DEFAULT_THEME_FONTS: Record<string, string> = {
  majorAscii: 'Cambria',
  majorHAnsi: 'Cambria',
  majorEastAsia: 'Cambria',
  majorBidi: 'Times New Roman',
  minorAscii: 'Calibri',
  minorHAnsi: 'Calibri',
  minorEastAsia: 'Calibri',
  minorBidi: 'Arial',
};

type StyleType = StyleInfo['type'];
type TableCondition = 'firstRow' | 'lastRow' | 'firstCol' | 'lastCol'
  | 'band1Horz' | 'band2Horz' | 'band1Vert' | 'band2Vert'
  | 'nwCell' | 'neCell' | 'swCell' | 'seCell';

export interface ThemeInfo {
  colors: Record<string, string>;
  fonts: Record<string, string>;
  colorSchemeMapping?: ColorSchemeMapping;
}

interface TableStyleLayer {
  paragraph?: ParagraphFormat;
  run?: RunFormat;
  /** `tblStylePr` 里的 `w:tcPr`：条件格式下的单元格底纹、边框等。 */
  cell?: CellFormat;
  /** `tblStylePr` 里的 `w:trPr`：条件格式下的行高、`cantSplit`、`tblHeader` 等。 */
  row?: RowFormat;
  /**
   * `tblStylePr` 里的 `w:tblPr`。条件描述的是**匹配到的区域**，不是整张表，所以这里真正有
   * 意义的是 `tblBorders` 与 `tblCellMar`——它们落到匹配到的那些单元格上。
   */
  table?: TableFormat;
}

interface TableMeta {
  rows: Element[];
  rowIndex: Map<Element, number>;
  cellIndex: WeakMap<Element, number>;
  look?: Element;
  rowBandSize: number;
  colBandSize: number;
  chain: ParsedStyle[];
}

interface TableContext {
  conditions: TableCondition[];
  chain: ParsedStyle[];
}

interface ParagraphContext {
  direct: ParagraphFormat;
  paragraphStyles: ParsedStyle[];
  tableParagraph: ParagraphFormat[];
  tableRun: RunFormat[];
  effectiveParagraph: ParagraphFormat;
}

interface ParsedStyle extends StyleInfo {
  conditions?: Partial<Record<TableCondition, TableStyleLayer>>;
  /**
   * 样式 `tblPr` 里的 `tblStyleRowBandSize` / `tblStyleColBandSize`。Word 把带宽写在**样式**
   * 里，不在表格实例上，只读实例的话真实文档里的带宽会被完全忽略（每一行都成了带）。
   */
  rowBandSize?: number;
  colBandSize?: number;
  /** 样式自身的 `w:tcPr`，相当于 `wholeTable` 层的单元格格式。 */
  cell?: CellFormat;
  /** 样式自身的 `w:trPr`，相当于 `wholeTable` 层的行格式。 */
  row?: RowFormat;
  /** 样式自身的 `w:tblPr`：全框线（「Table Grid」这类样式就在这里定）、单元格边距、底纹等。 */
  table?: TableFormat;
}

export interface StylesContext {
  docDefaults: {
    paragraph: ParagraphFormat;
    run: RunFormat;
  };
  styles: StyleInfo[];
  byId: Map<string, ParsedStyle>;
  defaults: Partial<Record<StyleType, string>>;
  theme: ThemeInfo;
  compatibilitySettings?: CompatibilitySettings;
  themeSettings: ThemeSettings;
  _tableMeta?: WeakMap<Element, TableMeta>;
  _tableContext?: WeakMap<Element, TableContext>;
  _paragraphContext?: WeakMap<Element, ParagraphContext>;
}

function wordAttr(element: Element | undefined, name: string): string | undefined {
  return element?.getAttributeNS(WORD_NS, name) ?? undefined;
}

function normalizeHex(value: string | undefined): string | undefined {
  return value && /^[0-9a-f]{6}$/i.test(value) ? value.toUpperCase() : undefined;
}

function normalizeHexOrAuto(value: string | undefined): string | undefined {
  if (!value) return undefined;
  if (value.toLowerCase() === 'auto') return 'auto';
  return normalizeHex(value);
}

function readOnOff(element: Element | undefined): boolean | undefined {
  if (!element) return undefined;
  return readOnOffValue(wordValue(element));
}

function readOnOffValue(raw: string | undefined): boolean {
  const value = (raw ?? '1').toLowerCase();
  return ['0', 'false', 'off'].includes(value) ? false : true;
}

const COMBINE_BRACKETS = ['none', 'round', 'square', 'angle', 'curly'] as const;
const FRAME_ENUMS = {
  dropCap: ['none', 'drop', 'margin'],
  heightRule: ['auto', 'exact', 'atLeast'],
  wrap: ['around', 'auto', 'none', 'notBeside', 'through', 'tight'],
  verticalAnchor: ['margin', 'page', 'text'],
  horizontalAnchor: ['margin', 'page', 'text'],
  xAlign: ['center', 'inside', 'left', 'outside', 'right'],
  yAlign: ['bottom', 'center', 'inline', 'inside', 'outside', 'top'],
} as const;

/** `w:framePr` 的全部信息都在属性上，没有 `w:val` 子元素。 */
function readParagraphFrame(element: Element | undefined): ParagraphFormat['frame'] {
  if (!element) return undefined;
  const enumValue = (name: keyof typeof FRAME_ENUMS, attribute: string = name): string | undefined => {
    const raw = wordAttr(element, attribute);
    return raw !== undefined && (FRAME_ENUMS[name] as readonly string[]).includes(raw) ? raw : undefined;
  };
  const anchorLock = wordAttr(element, 'anchorLock');
  return compactDefined({
    dropCap: enumValue('dropCap'),
    lines: readNumber(wordAttr(element, 'lines')),
    widthTwips: readNumber(wordAttr(element, 'w')),
    heightTwips: readNumber(wordAttr(element, 'h')),
    heightRule: enumValue('heightRule', 'hRule'),
    wrap: enumValue('wrap'),
    verticalAnchor: enumValue('verticalAnchor', 'vAnchor'),
    horizontalAnchor: enumValue('horizontalAnchor', 'hAnchor'),
    xTwips: readNumber(wordAttr(element, 'x')),
    yTwips: readNumber(wordAttr(element, 'y')),
    xAlign: enumValue('xAlign'),
    yAlign: enumValue('yAlign'),
    horizontalSpaceTwips: readNumber(wordAttr(element, 'hSpace')),
    verticalSpaceTwips: readNumber(wordAttr(element, 'vSpace')),
    anchorLock: anchorLock === undefined ? undefined : readOnOffValue(anchorLock),
  }) as ParagraphFormat['frame'];
}

/** `w:eastAsianLayout` 的开关放在属性上（`w:combine="true"`），不是 `w:val` 子元素。 */
function readEastAsianLayout(element: Element | undefined): RunFormat['eastAsianLayout'] {
  if (!element) return undefined;
  const brackets = wordAttr(element, 'combineBrackets');
  const flag = (name: string): boolean | undefined => {
    const raw = wordAttr(element, name);
    return raw === undefined ? undefined : readOnOffValue(raw);
  };
  return compactDefined({
    id: readNumber(wordAttr(element, 'id')),
    combine: flag('combine'),
    combineBrackets: brackets !== undefined && (COMBINE_BRACKETS as readonly string[]).includes(brackets)
      ? brackets as NonNullable<RunFormat['eastAsianLayout']>['combineBrackets'] : undefined,
    vert: flag('vert'),
    vertCompress: flag('vertCompress'),
  });
}

function readNumber(value: string | undefined): number | undefined {
  return value !== undefined && /^-?\d+$/.test(value) ? Number(value) : undefined;
}

function readShading(value: Element | undefined): Shading | undefined {
  if (!value) return undefined;
  const fill = normalizeHexOrAuto(wordAttr(value, 'fill')) ?? 'auto';
  const color = normalizeHexOrAuto(wordAttr(value, 'color'));
  return {
    pattern: wordValue(value) ?? 'clear',
    fill,
    ...(color !== undefined ? { color } : {}),
  };
}

export function readBorderSide(value: Element | undefined): BorderSide | undefined {
  if (!value) return undefined;
  const size = readNumber(wordAttr(value, 'sz'));
  const space = readNumber(wordAttr(value, 'space'));
  const color = normalizeHexOrAuto(wordAttr(value, 'color')) ?? 'auto';
  const shadow = wordAttr(value, 'shadow');
  return {
    style: wordValue(value) ?? 'none',
    size: Number.isFinite(size) && (size as number) >= 0 ? size as number : 0,
    space: Number.isFinite(space) && (space as number) >= 0 ? space as number : 0,
    color,
    ...(shadow !== undefined ? { shadow: !['0', 'false', 'off'].includes(shadow.toLowerCase()) } : {}),
  };
}

function readTabs(props: Element | undefined): TabStop[] | undefined {
  if (!props) return undefined;
  const tabsRoot = children(props, 'tabs')[0];
  const tabs = tabsRoot ? children(tabsRoot, 'tab').map((tab): TabStop => ({
    position: readNumber(wordAttr(tab, 'pos')) ?? 0,
    alignment: (wordValue(tab) ?? 'left') as TabStop['alignment'],
    ...(wordAttr(tab, 'leader') ? { leader: wordAttr(tab, 'leader') as TabStop['leader'] } : {}),
  })) : [];
  return tabs.length ? tabs : undefined;
}

function cloneParagraphFormat(format: ParagraphFormat | undefined): ParagraphFormat | undefined {
  return format ? { ...format } : undefined;
}

function cloneRunFormat(format: RunFormat | undefined): RunFormat | undefined {
  return format ? { ...format } : undefined;
}

function mergeParagraphFormats(...formats: Array<ParagraphFormat | undefined>): ParagraphFormat {
  const merged: ParagraphFormat = {};
  for (const format of formats) {
    if (!format) continue;
    for (const [key, value] of Object.entries(format) as Array<[keyof ParagraphFormat, ParagraphFormat[keyof ParagraphFormat]]>) {
      if (value !== undefined && value !== null) (merged as Record<string, unknown>)[key] = value;
    }
  }
  return merged;
}

/**
 * 合并单元格格式。`borders` 与 `margin` 要**按边**合并，不能整块替换——「整表定四边 +
 * firstRow 只定下边框」是表格样式里最常见的组合，整块替换会把其余三边抹掉。`shading` 对应
 * 单个 `w:shd` 元素，所以作为整体替换。
 */
function mergeCellFormats(...formats: Array<CellFormat | undefined>): CellFormat {
  const merged: CellFormat = {};
  const nested = new Set(['borders', 'margin']);
  for (const format of formats) {
    if (!format) continue;
    for (const [key, value] of Object.entries(format)) {
      if (value === undefined || value === null) continue;
      if (nested.has(key) && typeof value === 'object' && !Array.isArray(value)) {
        const previous = (merged as Record<string, unknown>)[key];
        const base = previous && typeof previous === 'object' ? previous as Record<string, unknown> : {};
        const next: Record<string, unknown> = { ...base };
        for (const [side, sideValue] of Object.entries(value as Record<string, unknown>)) {
          if (sideValue !== undefined && sideValue !== null) next[side] = sideValue;
        }
        (merged as Record<string, unknown>)[key] = next;
        continue;
      }
      (merged as Record<string, unknown>)[key] = value;
    }
  }
  return merged;
}

/**
 * 算上表格样式之后的表格级格式（宽度、对齐、缩进、布局、底纹、单元格边距、框线）。
 *
 * **条件不参与**：`firstRow` 之类描述的是表格里的一块区域，对表格元素自身没有意义；条件
 * `tblPr` 的边框与边距走单元格那条路（见 computeEffectiveCellFormat）。
 *
 * 样式里的 `tblStyle` 与 `tblLook` 不参与合并：前者会反过来改写表格引用的样式 id，后者会把
 * 带状 / 首行这些开关整个搅乱——它们只能来自表格自己。
 */
export function computeEffectiveTableFormat(context: StylesContext, table: Element,
  direct: TableFormat | undefined): TableFormat | undefined {
  const chain = resolveStyleChainOrDefault(context,
    wordValue(children(children(table, 'tblPr')[0] ?? table, 'tblStyle')[0]) ?? undefined, 'table');
  if (!chain.length) return direct;
  const merged: TableFormat = {};
  const apply = (format: TableFormat | undefined, skipIdentity = false) => {
    if (!format) return;
    for (const [key, value] of Object.entries(format)) {
      if (value === undefined || value === null) continue;
      if (skipIdentity && (key === 'style' || key === 'look')) continue;
      (merged as Record<string, unknown>)[key] = value;
    }
  };
  for (const style of chain) apply(style.table, true);
  apply(direct);
  return Object.keys(merged).length ? merged : undefined;
}

/**
 * 算上表格样式之后的行格式。只有由行位置决定的条件参与——`firstCol` 之类是单元格范围的，
 * 对整行没有意义。这一点靠「不给 `tableConditionsFor` 传单元格」做到：那时 `cellPosition`
 * 留在 -1，所有跟列有关的判断都不命中。曾经在这里另加过一张 `ROW_SCOPED_CONDITIONS` 过滤表，
 * 但它和上面那条是两份真相、互相兜底，结果两边的牙齿检查都没有信号，所以去掉了。
 *
 * `height` 对应单个 `w:trHeight` 元素，所以整体替换，不按字段合并。
 */
export function computeEffectiveRowFormat(context: StylesContext, row: Element,
  direct: RowFormat | undefined): RowFormat | undefined {
  const { conditions, chain } = tableConditionsFor(context, closestAncestor(row, 'tbl'), row);
  if (!chain.length) return direct;
  const merged: RowFormat = {};
  const apply = (format: RowFormat | undefined) => {
    if (!format) return;
    for (const [key, value] of Object.entries(format)) {
      if (value !== undefined && value !== null) (merged as Record<string, unknown>)[key] = value;
    }
  };
  for (const style of chain) {
    apply(style.row);
    for (const condition of conditions) apply(style.conditions?.[condition]?.row);
  }
  apply(direct);
  return Object.keys(merged).length ? merged : undefined;
}

/**
 * 算上表格样式之后的单元格格式。叠加顺序和段落 / 文字格式**用的是同一套条件与优先级**：
 * 样式链的 wholeTable 层 → 各条件层（带状 / 首行 / 角单元格……）→ 单元格自己的直接格式。
 *
 * 样式 `tblPr` 里的 `tblBorders` 当作单元格的默认边框先垫在最下层——Word 的「Table Grid」
 * 这类样式就是这么定全框线的，只看 `tcBorders` 的话那些表格一条线都没有。
 */
export function computeEffectiveCellFormat(context: StylesContext, cell: Element,
  direct: CellFormat | undefined): CellFormat | undefined {
  const { conditions, chain } = tableContextForCell(context, cell);
  if (!chain.length) return direct;
  // 每一级（样式自身、以及每个命中的条件）都是「先垫这一级 tblPr 给的默认值，再叠这一级
  // 的 tcPr」：tblPr 描述区域的默认边框与边距，tcPr 是针对单元格的，更具体。
  const fromTable = (table: TableFormat | undefined): CellFormat | undefined =>
    table && (table.borders || table.cellMargin)
      ? { ...(table.borders ? { borders: table.borders } : {}), ...(table.cellMargin ? { margin: table.cellMargin } : {}) }
      : undefined;
  const layers: Array<CellFormat | undefined> = [];
  for (const style of chain) {
    layers.push(fromTable(style.table));
    layers.push(style.cell);
    for (const condition of conditions) {
      const layer = style.conditions?.[condition];
      layers.push(fromTable(layer?.table));
      layers.push(layer?.cell);
    }
  }
  layers.push(direct);
  const merged = mergeCellFormats(...layers);
  return Object.keys(merged).length ? merged : undefined;
}

function mergeRunFormats(...formats: Array<RunFormat | undefined>): RunFormat {
  const merged: RunFormat = {};
  for (const format of formats) {
    if (!format) continue;
    for (const [key, value] of Object.entries(format) as Array<[keyof RunFormat, RunFormat[keyof RunFormat]]>) {
      if (value !== undefined && value !== null) (merged as Record<string, unknown>)[key] = value;
    }
  }
  return merged;
}

function rgbToHsl(red: number, green: number, blue: number): [number, number, number] {
  const r = red / 255;
  const g = green / 255;
  const b = blue / 255;
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const lightness = (max + min) / 2;
  if (max === min) return [0, 0, lightness];
  const delta = max - min;
  const saturation = lightness > 0.5 ? delta / (2 - max - min) : delta / (max + min);
  const hue = (
    max === r ? (g - b) / delta + (g < b ? 6 : 0)
      : max === g ? (b - r) / delta + 2
        : (r - g) / delta + 4
  ) / 6;
  return [hue, saturation, lightness];
}

function hueToRgb(low: number, high: number, hue: number): number {
  if (hue < 0) hue += 1;
  if (hue > 1) hue -= 1;
  if (hue < 1 / 6) return low + (high - low) * 6 * hue;
  if (hue < 1 / 2) return high;
  if (hue < 2 / 3) return low + (high - low) * (2 / 3 - hue) * 6;
  return low;
}

function hslToRgb(hue: number, saturation: number, lightness: number): [number, number, number] {
  if (saturation === 0) {
    const value = Math.floor(lightness * 255);
    return [value, value, value];
  }
  const high = lightness < 0.5 ? lightness * (1 + saturation) : lightness + saturation - lightness * saturation;
  const low = 2 * lightness - high;
  return [
    Math.floor(hueToRgb(low, high, hue + 1 / 3) * 255),
    Math.floor(hueToRgb(low, high, hue) * 255),
    Math.floor(hueToRgb(low, high, hue - 1 / 3) * 255),
  ];
}

function applyShadeTint(hex: string | undefined, shade: string | undefined, tint: string | undefined, percentageValues = false): string | undefined {
  const base = normalizeHex(hex);
  if (!base) return undefined;
  const transformValue = (value: string | undefined) => percentageValues
    ? value && /^\d+$/.test(value) && Number(value) <= 100000 ? Number(value) / 100000 : undefined
    : value && /^[0-9a-f]{2}$/i.test(value) ? parseInt(value, 16) / 255 : undefined;
  const shadeValue = transformValue(shade);
  const tintValue = transformValue(tint);
  if (shadeValue === undefined && tintValue === undefined) return base;
  const channels = base.match(/../g)!.map((channel) => parseInt(channel, 16));
  const [hue, saturation, lightness] = rgbToHsl(channels[0]!, channels[1]!, channels[2]!);
  let transformedLightness = lightness;
  if (shadeValue !== undefined) transformedLightness *= shadeValue;
  if (tintValue !== undefined) transformedLightness = transformedLightness * tintValue + (1 - tintValue);
  const transformed = hslToRgb(hue, saturation, Math.max(0, Math.min(1, transformedLightness)));
  return transformed.map((channel) => channel.toString(16).padStart(2, '0').toUpperCase()).join('');
}

const THEME_COLOR_ALIASES: Record<string, string> = {
  text1: 't1', tx1: 't1', background1: 'bg1', bg1: 'bg1',
  text2: 't2', tx2: 't2', background2: 'bg2', bg2: 'bg2',
};

const DEFAULT_THEME_COLOR_MAPPING: Record<string, string> = {
  t1: 'dk1',
  bg1: 'lt1',
  t2: 'dk2',
  bg2: 'lt2',
  accent1: 'accent1',
  accent2: 'accent2',
  accent3: 'accent3',
  accent4: 'accent4',
  accent5: 'accent5',
  accent6: 'accent6',
  hyperlink: 'hlink',
  followedHyperlink: 'folHlink',
};

const THEME_COLOR_NAMES: Record<string, string> = {
  dark1: 'dk1',
  light1: 'lt1',
  dark2: 'dk2',
  light2: 'lt2',
  followedHyperlink: 'folHlink',
  hyperlink: 'hlink',
};

function resolveThemeColorName(theme: ThemeInfo, name: string | undefined): string {
  const slot = THEME_COLOR_ALIASES[name ?? ''] ?? name ?? '';
  const mapped = theme.colorSchemeMapping?.[slot] ?? DEFAULT_THEME_COLOR_MAPPING[slot] ?? slot;
  return THEME_COLOR_NAMES[mapped] ?? mapped;
}

function resolveThemeValue(theme: ThemeInfo, name: string | undefined, shade: string | undefined, tint: string | undefined, percentageValues = false): string | undefined {
  return applyShadeTint(theme.colors[resolveThemeColorName(theme, name)], shade, tint, percentageValues);
}

function resolveThemeColor(theme: ThemeInfo, element: Element | undefined): string | undefined {
  if (!element) return undefined;
  const direct = normalizeHex(wordAttr(element, 'val'));
  if (direct && direct.toLowerCase() !== 'auto') return direct;
  return resolveThemeValue(theme, wordAttr(element, 'themeColor'), wordAttr(element, 'themeShade'), wordAttr(element, 'themeTint'));
}

export function resolveDrawingColor(theme: ThemeInfo, element: Element | undefined): string | undefined {
  if (!element) return undefined;
  const value = element.getAttribute('val') ?? undefined;
  const direct = normalizeHex(value);
  if (direct) return direct;
  const modifier = (name: 'shade' | 'tint') =>
    Array.from(element.getElementsByTagNameNS(DRAWINGML_NS, name))[0]?.getAttribute('val') ?? undefined;
  return resolveThemeValue(theme, value, modifier('shade'), modifier('tint'), true);
}

export function resolveDrawingThemeColor(theme: ThemeInfo, name: string): string | undefined {
  return resolveThemeValue(theme, name, undefined, undefined, true);
}

function resolveUnderlineColor(theme: ThemeInfo, element: Element | undefined): string | undefined {
  if (!element) return undefined;
  const direct = normalizeHex(wordAttr(element, 'color'));
  if (direct && direct.toLowerCase() !== 'auto') return direct;
  return resolveThemeValue(theme, wordAttr(element, 'themeColor'), wordAttr(element, 'themeShade'), wordAttr(element, 'themeTint'));
}

function resolveThemeFont(theme: ThemeInfo, value: string | undefined, fallback?: string): string | undefined {
  return value ? theme.fonts[value] ?? fallback : fallback;
}

function readFontFamily(theme: ThemeInfo, fonts: Element | undefined): { fontFamily?: string; fontFamilyEastAsia?: string } {
  if (!fonts) return {};
  const ascii = wordAttr(fonts, 'ascii') ?? wordAttr(fonts, 'hAnsi');
  const eastAsia = wordAttr(fonts, 'eastAsia');
  const fontFamily = ascii
    ?? resolveThemeFont(theme, wordAttr(fonts, 'asciiTheme') ?? wordAttr(fonts, 'hAnsiTheme'))
    ?? resolveThemeFont(theme, wordAttr(fonts, 'eastAsiaTheme'))
    ?? eastAsia;
  const fontFamilyEastAsia = eastAsia
    ?? resolveThemeFont(theme, wordAttr(fonts, 'eastAsiaTheme'), fontFamily);
  return { fontFamily, fontFamilyEastAsia };
}

export function readParagraphProperties(props: Element | undefined): ParagraphFormat {
  if (!props) return {};
  const spacing = children(props, 'spacing')[0];
  const indent = children(props, 'ind')[0];
  const alignment = wordValue(children(props, 'jc')[0]);
  const frame = readParagraphFrame(children(props, 'framePr')[0]);
  const borders = children(props, 'pBdr')[0];
  const parsedBorders = borders ? {
    top: readBorderSide(children(borders, 'top')[0]),
    left: readBorderSide(children(borders, 'left')[0]),
    bottom: readBorderSide(children(borders, 'bottom')[0]),
    right: readBorderSide(children(borders, 'right')[0]),
    between: readBorderSide(children(borders, 'between')[0]),
    bar: readBorderSide(children(borders, 'bar')[0]),
  } : undefined;
  return {
    style: wordValue(children(props, 'pStyle')[0]),
    ...(frame ? { frame } : {}),
    alignment: ['left', 'center', 'right', 'both', 'distribute'].includes(alignment ?? '')
      ? alignment as ParagraphFormat['alignment'] : undefined,
    indentLeft: readNumber(wordAttr(indent, 'left') ?? wordAttr(indent, 'start')),
    indentRight: readNumber(wordAttr(indent, 'right') ?? wordAttr(indent, 'end')),
    indentFirstLine: readNumber(wordAttr(indent, 'firstLine')),
    indentHanging: readNumber(wordAttr(indent, 'hanging')),
    spacingBefore: readNumber(wordAttr(spacing, 'before')),
    spacingAfter: readNumber(wordAttr(spacing, 'after')),
    spacingBeforeLines: readNumber(wordAttr(spacing, 'beforeLines')),
    spacingAfterLines: readNumber(wordAttr(spacing, 'afterLines')),
    spacingBeforeAuto: spacing ? readOnOffValue(wordAttr(spacing, 'beforeAutospacing')) : undefined,
    spacingAfterAuto: spacing ? readOnOffValue(wordAttr(spacing, 'afterAutospacing')) : undefined,
    contextualSpacing: readOnOff(children(props, 'contextualSpacing')[0]),
    mirrorIndents: readOnOff(children(props, 'mirrorIndents')[0]),
    lineSpacing: readNumber(wordAttr(spacing, 'line')),
    lineSpacingRule: wordAttr(spacing, 'lineRule') as ParagraphFormat['lineSpacingRule'] | undefined,
    keepNext: readOnOff(children(props, 'keepNext')[0]),
    keepLines: readOnOff(children(props, 'keepLines')[0]),
    pageBreakBefore: readOnOff(children(props, 'pageBreakBefore')[0]),
    widowControl: readOnOff(children(props, 'widowControl')[0]),
    suppressLineNumbers: readOnOff(children(props, 'suppressLineNumbers')[0]),
    suppressAutoHyphens: readOnOff(children(props, 'suppressAutoHyphens')[0]),
    kinsoku: readOnOff(children(props, 'kinsoku')[0]),
    wordWrap: readOnOff(children(props, 'wordWrap')[0]),
    overflowPunct: readOnOff(children(props, 'overflowPunct')[0]),
    topLinePunct: readOnOff(children(props, 'topLinePunct')[0]),
    autoSpaceDE: readOnOff(children(props, 'autoSpaceDE')[0]),
    autoSpaceDN: readOnOff(children(props, 'autoSpaceDN')[0]),
    bidi: readOnOff(children(props, 'bidi')[0]),
    textDirection: wordValue(children(props, 'textDirection')[0]) ?? undefined,
    outlineLevel: readNumber(wordValue(children(props, 'outlineLvl')[0])),
    tabs: readTabs(props),
    borders: parsedBorders && Object.values(parsedBorders).some((entry) => entry !== undefined) ? parsedBorders : undefined,
    shading: readShading(children(props, 'shd')[0]),
  };
}

export function readRunProperties(props: Element | undefined, theme: StylesContext['theme']): RunFormat {
  if (!props) return {};
  const underline = children(props, 'u')[0];
  const underlineValue = wordValue(underline);
  const emphasis = children(props, 'em')[0];
  const emphasisValue = (wordValue(emphasis) ?? 'dot').toLowerCase();
  const emphasisMark: RunFormat['emphasisMark'] | undefined = emphasis
    ? ['dot', 'comma', 'circle', 'underdot', 'none'].includes(emphasisValue)
      ? emphasisValue === 'underdot' ? 'underDot' : emphasisValue as NonNullable<RunFormat['emphasisMark']>
      : undefined
    : undefined;
  const size = wordValue(children(props, 'sz')[0]) ?? wordValue(children(props, 'szCs')[0]);
  const fonts = children(props, 'rFonts')[0];
  return {
    style: wordValue(children(props, 'rStyle')[0]),
    bold: readOnOff(children(props, 'b')[0]),
    italic: readOnOff(children(props, 'i')[0]),
    hidden: readOnOff(children(props, 'vanish')[0]),
    webHidden: readOnOff(children(props, 'webHidden')[0]),
    ...(emphasisMark !== undefined ? { emphasisMark } : {}),
    underline: underline ? !['none', '0', 'false'].includes((underlineValue ?? 'single').toLowerCase()) : undefined,
    underlineStyle: underline && underlineValue && !['0', 'false', 'none'].includes(underlineValue.toLowerCase()) ? underlineValue : undefined,
    underlineColor: resolveUnderlineColor(theme, underline),
    fontSize: size && Number.isFinite(Number(size)) ? Number(size) / 2 : undefined,
    ...readFontFamily(theme, fonts),
    color: resolveThemeColor(theme, children(props, 'color')[0]),
    strike: readOnOff(children(props, 'strike')[0]),
    doubleStrike: readOnOff(children(props, 'dstrike')[0]),
    rtl: readOnOff(children(props, 'rtl')[0]),
    complexScript: readOnOff(children(props, 'cs')[0]),
    verticalAlign: wordValue(children(props, 'vertAlign')[0]) as RunFormat['verticalAlign'] | undefined,
    smallCaps: readOnOff(children(props, 'smallCaps')[0]),
    allCaps: readOnOff(children(props, 'caps')[0]),
    highlight: wordValue(children(props, 'highlight')[0]) ?? undefined,
    characterSpacing: readNumber(wordValue(children(props, 'spacing')[0])),
    position: readNumber(wordValue(children(props, 'position')[0])),
    characterScale: readNumber(wordValue(children(props, 'w')[0])),
    kerning: readNumber(wordValue(children(props, 'kern')[0])),
    fitTextWidth: readNumber(wordValue(children(props, 'fitText')[0])),
    textEffect: wordValue(children(props, 'effect')[0]) ?? undefined,
    textOutline: readOnOff(children(props, 'outline')[0]),
    textShadow: readOnOff(children(props, 'shadow')[0]),
    emboss: readOnOff(children(props, 'emboss')[0]),
    imprint: readOnOff(children(props, 'imprint')[0]),
    border: readBorderSide(children(props, 'bdr')[0]),
    shading: readShading(children(props, 'shd')[0]),
    eastAsianLayout: readEastAsianLayout(children(props, 'eastAsianLayout')[0]),
  };
}

function themeColorValue(node: Element | undefined): string | undefined {
  if (!node) return undefined;
  return normalizeHex(node.getAttribute('val') ?? undefined)
    ?? normalizeHex(node.getAttribute('lastClr') ?? undefined);
}

function parseTheme(themeElement: Element | undefined, colorSchemeMapping?: ColorSchemeMapping): ThemeInfo {
  if (!themeElement) return { colors: { ...DEFAULT_THEME_COLORS }, fonts: { ...DEFAULT_THEME_FONTS }, colorSchemeMapping };
  const colorScheme = Array.from(themeElement.getElementsByTagNameNS(DRAWINGML_NS, 'clrScheme'))[0];
  const fontScheme = Array.from(themeElement.getElementsByTagNameNS(DRAWINGML_NS, 'fontScheme'))[0];
  const colors = { ...DEFAULT_THEME_COLORS };
  if (colorScheme) {
    for (const name of Object.keys(colors)) {
      const entry = Array.from(colorScheme.childNodes).find((child) =>
        child.nodeType === 1 && (child as Element).localName === name) as Element | undefined;
      const value = themeColorValue(
        entry ? Array.from(entry.childNodes).find((child) => child.nodeType === 1) as Element | undefined : undefined,
      );
      if (value) colors[name] = value;
    }
  }
  const fonts = { ...DEFAULT_THEME_FONTS };
  const fontSections = [
    ['majorFont', 'major'],
    ['minorFont', 'minor'],
  ] as const;
  for (const [sectionName, prefix] of fontSections) {
    const section = fontScheme
      ? Array.from(fontScheme.childNodes).find((child) => child.nodeType === 1 && (child as Element).localName === sectionName) as Element | undefined
      : undefined;
    const latin = section ? Array.from(section.getElementsByTagNameNS(DRAWINGML_NS, 'latin'))[0] : undefined;
    const ea = section ? Array.from(section.getElementsByTagNameNS(DRAWINGML_NS, 'ea'))[0] : undefined;
    const cs = section ? Array.from(section.getElementsByTagNameNS(DRAWINGML_NS, 'cs'))[0] : undefined;
    const latinTypeface = latin?.getAttribute('typeface') || fonts[`${prefix}Ascii`] || '';
    fonts[`${prefix}Ascii`] = latinTypeface;
    fonts[`${prefix}HAnsi`] = latinTypeface;
    fonts[`${prefix}EastAsia`] = ea?.getAttribute('typeface') || latinTypeface;
    fonts[`${prefix}Bidi`] = cs?.getAttribute('typeface') || fonts[`${prefix}Bidi`] || latinTypeface;
  }
  return { colors, fonts, colorSchemeMapping };
}

function parseStyleType(value: string | undefined): StyleType | undefined {
  return ['paragraph', 'character', 'table', 'numbering'].includes(value ?? '') ? value as StyleType : undefined;
}

const COLOR_SCHEME_SLOTS = [
  'bg1', 't1', 'bg2', 't2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6',
  'hyperlink', 'followedHyperlink',
] as const;

function readColorSchemeMapping(settingsRoot: Element | undefined): ColorSchemeMapping | undefined {
  if (!settingsRoot) return undefined;
  const mappingElement = children(settingsRoot, 'clrSchemeMapping')[0];
  if (!mappingElement) return undefined;
  const mapping: ColorSchemeMapping = {};
  for (const slot of COLOR_SCHEME_SLOTS) {
    const value = wordAttr(mappingElement, slot);
    if (value !== undefined) mapping[slot] = value;
  }
  return Object.keys(mapping).length ? mapping : undefined;
}

function readThemeFontLanguages(settingsRoot: Element | undefined): ThemeFontLanguages | undefined {
  if (!settingsRoot) return undefined;
  const element = children(settingsRoot, 'themeFontLang')[0];
  if (!element) return undefined;
  const languages: ThemeFontLanguages = {};
  const val = wordAttr(element, 'val');
  const eastAsia = wordAttr(element, 'eastAsia');
  const bidi = wordAttr(element, 'bidi');
  if (val !== undefined) languages.val = val;
  if (eastAsia !== undefined) languages.eastAsia = eastAsia;
  if (bidi !== undefined) languages.bidi = bidi;
  return Object.keys(languages).length ? languages : undefined;
}

function readLatentStyles(stylesRoot: Element | undefined): LatentStyles | undefined {
  if (!stylesRoot) return undefined;
  const element = children(stylesRoot, 'latentStyles')[0];
  if (!element) return undefined;
  const number = (name: string): number | undefined => {
    const result = readNumber(wordAttr(element, name));
    return result !== undefined && result >= 0 ? result : undefined;
  };
  const bool = (node: Element, name: string): boolean | undefined => {
    const value = wordAttr(node, name);
    return value === undefined ? undefined : readOnOffValue(value);
  };
  const result: LatentStyles = {
    exceptions: children(element, 'lsdException').flatMap((exception) => {
      const name = wordAttr(exception, 'name');
      if (name === undefined) return [];
      const uiPriority = readNumber(wordAttr(exception, 'uiPriority'));
      return [{
        name,
        ...(bool(exception, 'locked') !== undefined ? { locked: bool(exception, 'locked') } : {}),
        ...(uiPriority !== undefined && uiPriority >= 0 ? { uiPriority } : {}),
        ...(bool(exception, 'semiHidden') !== undefined ? { semiHidden: bool(exception, 'semiHidden') } : {}),
        ...(bool(exception, 'unhideWhenUsed') !== undefined ? { unhideWhenUsed: bool(exception, 'unhideWhenUsed') } : {}),
        ...(bool(exception, 'qFormat') !== undefined ? { qFormat: bool(exception, 'qFormat') } : {}),
      }];
    }),
  };
  const defaultLockedState = bool(element, 'defLockedState');
  const defaultUiPriority = number('defUIPriority');
  const defaultSemiHidden = bool(element, 'defSemiHidden');
  const defaultUnhideWhenUsed = bool(element, 'defUnhideWhenUsed');
  const defaultQFormat = bool(element, 'defQFormat');
  const count = number('count');
  if (defaultLockedState !== undefined) result.defaultLockedState = defaultLockedState;
  if (defaultUiPriority !== undefined) result.defaultUiPriority = defaultUiPriority;
  if (defaultSemiHidden !== undefined) result.defaultSemiHidden = defaultSemiHidden;
  if (defaultUnhideWhenUsed !== undefined) result.defaultUnhideWhenUsed = defaultUnhideWhenUsed;
  if (defaultQFormat !== undefined) result.defaultQFormat = defaultQFormat;
  if (count !== undefined) result.count = count;
  return result;
}

function readThemeSettings(settingsRoot: Element | undefined, stylesRoot: Element | undefined): ThemeSettings {
  const clrSchemeMapping = readColorSchemeMapping(settingsRoot);
  const themeFontLang = readThemeFontLanguages(settingsRoot);
  const latentStyles = readLatentStyles(stylesRoot);
  return {
    ...(clrSchemeMapping ? { clrSchemeMapping } : {}),
    ...(themeFontLang ? { themeFontLang } : {}),
    ...(latentStyles ? { latentStyles } : {}),
  };
}

export function parseStyles(stylesRoot: Element | undefined, themeRoot?: Element, compatibilitySettings?: CompatibilitySettings, settingsRoot?: Element): StylesContext {
  const themeSettings = readThemeSettings(settingsRoot, stylesRoot);
  const theme = parseTheme(themeRoot, themeSettings.clrSchemeMapping);
  if (!stylesRoot || stylesRoot.namespaceURI !== WORD_NS || stylesRoot.localName !== 'styles') {
    return { docDefaults: { paragraph: {}, run: {} }, styles: [], byId: new Map(), defaults: {}, theme, compatibilitySettings, themeSettings };
  }
  const docDefaults = children(stylesRoot, 'docDefaults')[0];
  const paragraphDefault = readParagraphProperties(children(children(docDefaults ?? stylesRoot, 'pPrDefault')[0] ?? stylesRoot, 'pPr')[0]);
  const runDefault = readRunProperties(children(children(docDefaults ?? stylesRoot, 'rPrDefault')[0] ?? stylesRoot, 'rPr')[0], theme);
  const styles: ParsedStyle[] = [];
  const defaults: Partial<Record<StyleType, string>> = {};
  for (const styleElement of children(stylesRoot, 'style')) {
    const type = parseStyleType(wordAttr(styleElement, 'type'));
    const id = wordAttr(styleElement, 'styleId');
    if (!type || !id) continue;
    const style: ParsedStyle = {
      id,
      name: wordValue(children(styleElement, 'name')[0]) ?? id,
      type,
      basedOn: wordValue(children(styleElement, 'basedOn')[0]) ?? undefined,
      next: wordValue(children(styleElement, 'next')[0]) ?? undefined,
      link: wordValue(children(styleElement, 'link')[0]) ?? undefined,
      aliases: (wordValue(children(styleElement, 'aliases')[0]) ?? '')
        .split(',').map((value) => value.trim()).filter(Boolean),
      isDefault: ['1', 'true', 'on'].includes((wordAttr(styleElement, 'default') ?? '').toLowerCase()) || undefined,
      uiPriority: readNumber(wordValue(children(styleElement, 'uiPriority')[0])),
      quickFormat: !!children(styleElement, 'qFormat')[0] || undefined,
      paragraph: cloneParagraphFormat(readParagraphProperties(children(styleElement, 'pPr')[0])),
      run: cloneRunFormat(readRunProperties(children(styleElement, 'rPr')[0], theme)),
      conditions: {},
      ...(() => {
        const styleTableProps = children(styleElement, 'tblPr')[0];
        const bandSize = (name: string) => {
          const value = readNumber(wordValue(children(styleTableProps ?? styleElement, name)[0]));
          return value === undefined ? undefined : Math.max(1, value);
        };
        return compactDefined({
          rowBandSize: bandSize('tblStyleRowBandSize'),
          colBandSize: bandSize('tblStyleColBandSize'),
          // 样式自身的 tcPr 相当于 wholeTable 层；tblBorders 是「Table Grid」这类样式的全框线。
          cell: parseCellFormat(children(styleElement, 'tcPr')[0]),
          row: parseRowFormat(children(styleElement, 'trPr')[0]),
          table: parseTableFormat(styleTableProps),
        });
      })(),
    };
    for (const conditionElement of children(styleElement, 'tblStylePr')) {
      const condition = wordAttr(conditionElement, 'type') as TableCondition | undefined;
      if (!condition || !['firstRow', 'lastRow', 'firstCol', 'lastCol',
        'band1Horz', 'band2Horz', 'band1Vert', 'band2Vert',
        'nwCell', 'neCell', 'swCell', 'seCell'].includes(condition)) continue;
      style.conditions![condition] = {
        paragraph: cloneParagraphFormat(readParagraphProperties(children(conditionElement, 'pPr')[0])),
        run: cloneRunFormat(readRunProperties(children(conditionElement, 'rPr')[0], theme)),
        cell: parseCellFormat(children(conditionElement, 'tcPr')[0]),
        row: parseRowFormat(children(conditionElement, 'trPr')[0]),
        table: parseTableFormat(children(conditionElement, 'tblPr')[0]),
      };
    }
    if (style.isDefault && !defaults[type]) defaults[type] = id;
    styles.push(style);
  }
  return {
    docDefaults: { paragraph: paragraphDefault, run: runDefault },
    styles,
    byId: new Map(styles.map((style) => [style.id, style])),
    defaults,
    theme,
    compatibilitySettings,
    themeSettings,
  };
}

export function resolveStyleChain(context: StylesContext, id: string | undefined, type: StyleType): ParsedStyle[] {
  const chain: ParsedStyle[] = [];
  const seen = new Set<string>();
  let currentId = id;
  while (currentId && !seen.has(currentId)) {
    seen.add(currentId);
    const style = context.byId.get(currentId);
    if (!style || style.type !== type) break;
    chain.unshift(style);
    currentId = style.basedOn;
  }
  return chain;
}

export function resolveStyleChainOrDefault(context: StylesContext, id: string | undefined, type: StyleType): ParsedStyle[] {
  const explicit = resolveStyleChain(context, id, type);
  if (explicit.length) return explicit;
  if (context.defaults[type] && context.defaults[type] !== id) return resolveStyleChain(context, context.defaults[type], type);
  return explicit;
}

function closestAncestor(element: Element, localName: string): Element | undefined {
  let current = element.parentNode;
  while (current) {
    if (current.nodeType === 1) {
      const candidate = current as Element;
      if (candidate.namespaceURI === WORD_NS && candidate.localName === localName) return candidate;
    }
    current = current.parentNode;
  }
  return undefined;
}

const TBL_LOOK_BITS: Record<string, number> = {
  firstRow: 0x0020, lastRow: 0x0040, firstColumn: 0x0080, lastColumn: 0x0100,
  noHBand: 0x0200, noVBand: 0x0400,
};

/**
 * `w:tblLook` 有两种写法：Word 2010+ 写具名属性，Word 2007 只写 `w:val` 的十六进制位掩码
 * （最常见的 `04A0` 就是 firstRow + firstColumn + noVBand）。只认具名属性的话，这类文档的
 * 表头行会被当成普通带状行。现代 Word 两种都写，所以具名属性优先，掩码兜底。
 */
function readLookFlag(look: Element | undefined, name: string, defaultValue: boolean): boolean {
  const value = wordAttr(look, name);
  if (value !== undefined) return !['0', 'false', 'off'].includes(value.toLowerCase());
  const raw = wordAttr(look, 'val');
  const bit = TBL_LOOK_BITS[name];
  if (raw !== undefined && bit !== undefined && /^[0-9a-f]{1,8}$/i.test(raw)) {
    return (Number.parseInt(raw, 16) & bit) !== 0;
  }
  return defaultValue;
}

function tableMeta(context: StylesContext, table: Element): TableMeta {
  const cache = context._tableMeta ??= new WeakMap<Element, TableMeta>();
  const existing = cache.get(table);
  if (existing) return existing;
  const rows = childrenThroughTransparent(table, 'tr');
  const rowIndex = new Map<Element, number>();
  const cellIndex = new WeakMap<Element, number>();
  rows.forEach((row, index) => {
    rowIndex.set(row, index);
    childrenThroughTransparent(row, 'tc').forEach((cell, cellPosition) => cellIndex.set(cell, cellPosition));
  });
  const tableProps = children(table, 'tblPr')[0];
  const styleId = wordValue(children(tableProps ?? table, 'tblStyle')[0]) ?? undefined;
  const chain = resolveStyleChainOrDefault(context, styleId, 'table');
  // 带宽：表格实例上的是直接格式，优先；否则取样式链里最靠近的那个（链是从基样式到最具体
  // 排的，所以从后往前找）；都没有就是 1。Word 把它写在样式里。
  const bandSize = (own: string, pick: (style: ParsedStyle) => number | undefined): number => {
    const direct = readNumber(wordValue(children(tableProps ?? table, own)[0]));
    if (direct !== undefined) return Math.max(1, direct);
    for (let index = chain.length - 1; index >= 0; index--) {
      const value = pick(chain[index]!);
      if (value !== undefined) return value;
    }
    return 1;
  };
  const meta: TableMeta = {
    rows,
    rowIndex,
    cellIndex,
    look: children(tableProps ?? table, 'tblLook')[0],
    rowBandSize: bandSize('tblStyleRowBandSize', (style) => style.rowBandSize),
    colBandSize: bandSize('tblStyleColBandSize', (style) => style.colBandSize),
    chain,
  };
  cache.set(table, meta);
  return meta;
}

function tableContext(context: StylesContext, paragraph: Element): TableContext {
  const cache = context._tableContext ??= new WeakMap<Element, TableContext>();
  const existing = cache.get(paragraph);
  if (existing) return existing;
  const resolved = tableConditionsFor(context, closestAncestor(paragraph, 'tbl'),
    closestAncestor(paragraph, 'tr'), closestAncestor(paragraph, 'tc'));
  cache.set(paragraph, resolved);
  return resolved;
}

/** 单元格自己就是 `tc`，所以不能走 closestAncestor（那会跳到外层嵌套表格的单元格上）。 */
function tableContextForCell(context: StylesContext, cell: Element): TableContext {
  const cache = context._tableContext ??= new WeakMap<Element, TableContext>();
  const existing = cache.get(cell);
  if (existing) return existing;
  const resolved = tableConditionsFor(context, closestAncestor(cell, 'tbl'), closestAncestor(cell, 'tr'), cell);
  cache.set(cell, resolved);
  return resolved;
}

/**
 * 不传 `cell` 时只算**由行位置决定**的条件（`firstRow` / `lastRow` / `band*Horz`）：
 * `cellPosition` 留在 -1，下面所有跟列有关的判断都不会命中。行格式用的就是这个子集。
 */
function tableConditionsFor(context: StylesContext, table: Element | undefined,
  row: Element | undefined, cell?: Element): TableContext {
  if (!table || !row) return { conditions: [], chain: [] };
  const meta = tableMeta(context, table);
  const rowPosition = meta.rowIndex.get(row) ?? -1;
  const cellPosition = cell ? meta.cellIndex.get(cell) ?? -1 : -1;
  const look = meta.look;
  const firstRow = readLookFlag(look, 'firstRow', false);
  const lastRow = readLookFlag(look, 'lastRow', false);
  const firstColumn = readLookFlag(look, 'firstColumn', false);
  const lastColumn = readLookFlag(look, 'lastColumn', false);
  const cellCount = childrenThroughTransparent(row, 'tc').length;
  const legacyRules = context.compatibilitySettings?.useWord2002TableStyleRules === true;
  /**
   * 这个数组的顺序就是套用顺序，后面的覆盖前面的。ECMA-376 的条件格式优先级从低到高是
   * wholeTable → band*Vert → band*Horz → firstCol/lastCol → firstRow/lastRow → 四个角单元格。
   * 原先正好反着排，于是 firstCol 压过 firstRow、band*Horz 压过 firstCol —— 而 Word 里表头行
   * 是横贯整行（含第一列）的，第一列的特殊格式也压过带状。
   */
  const conditions: TableCondition[] = [];
  const band = (position: number, count: number, size: number, skipFirst: boolean, skipLast: boolean,
    odd: TableCondition, even: TableCondition) => {
    // 首行 / 首列（以及末行 / 末列）有自己的条件格式，不参与带状计数，否则带的相位会偏一格。
    // useWord2002TableStyleRules 下 Word 把它们一起算进去。
    const start = !legacyRules && skipFirst ? 1 : 0;
    const end = !legacyRules ? count - (skipLast ? 1 : 0) : count;
    if (position < start || position >= end) return;
    conditions.push(Math.floor((position - start) / size) % 2 === 0 ? odd : even);
  };
  if (!readLookFlag(look, 'noVBand', false)) {
    band(cellPosition, cellCount, meta.colBandSize, firstColumn, lastColumn, 'band1Vert', 'band2Vert');
  }
  if (!readLookFlag(look, 'noHBand', false)) {
    band(rowPosition, meta.rows.length, meta.rowBandSize, firstRow, lastRow, 'band1Horz', 'band2Horz');
  }
  const topRow = rowPosition === 0 && firstRow;
  const bottomRow = rowPosition === meta.rows.length - 1 && lastRow;
  const leftColumn = cellPosition === 0 && firstColumn;
  // cellPosition 为 -1（没传单元格，或单元格不在索引里）时不能算成末列——空行的 cellCount
  // 是 0，`-1 === 0 - 1` 会误判为 true。
  const rightColumn = cellPosition >= 0 && cellPosition === cellCount - 1 && lastColumn;
  if (leftColumn) conditions.push('firstCol');
  if (rightColumn) conditions.push('lastCol');
  if (topRow) conditions.push('firstRow');
  if (bottomRow) conditions.push('lastRow');
  // 四个角单元格优先级最高，而且只在对应的行、列条件都开着时才生效：关掉「第一列」就没有
  // 特殊的第一列，左上角也就不该按 nwCell 画。1x1 表格四个标志全开时四条都命中，按下面的
  // 顺序最后一条胜出——Word 在这种退化情形下的行为没有明确定义，这里取确定的顺序。
  if (topRow && leftColumn) conditions.push('nwCell');
  if (topRow && rightColumn) conditions.push('neCell');
  if (bottomRow && leftColumn) conditions.push('swCell');
  if (bottomRow && rightColumn) conditions.push('seCell');
  return { conditions, chain: meta.chain };
}

function tableParagraphFormats(context: StylesContext, paragraph: Element): ParagraphFormat[] {
  const { conditions, chain } = tableContext(context, paragraph);
  const formats: ParagraphFormat[] = [];
  for (const style of chain) {
    formats.push(style.paragraph ?? {});
    for (const condition of conditions) {
      if (style.conditions?.[condition]?.paragraph) formats.push(style.conditions[condition]!.paragraph!);
    }
  }
  return formats;
}

function tableRunFormats(context: StylesContext, paragraph: Element): RunFormat[] {
  const { conditions, chain } = tableContext(context, paragraph);
  const formats: RunFormat[] = [];
  for (const style of chain) {
    formats.push(style.run ?? {});
    for (const condition of conditions) {
      if (style.conditions?.[condition]?.run) formats.push(style.conditions[condition]!.run!);
    }
  }
  return formats;
}

function paragraphContext(context: StylesContext, paragraph: Element): ParagraphContext {
  const cache = context._paragraphContext ??= new WeakMap<Element, ParagraphContext>();
  const existing = cache.get(paragraph);
  if (existing) return existing;
  const direct = readParagraphProperties(children(paragraph, 'pPr')[0]);
  const paragraphStyles = resolveStyleChainOrDefault(context, direct.style ?? undefined, 'paragraph');
  const tableParagraph = tableParagraphFormats(context, paragraph);
  const tableRun = tableRunFormats(context, paragraph);
  const resolved = {
    direct,
    paragraphStyles,
    tableParagraph,
    tableRun,
    effectiveParagraph: mergeParagraphFormats(
      context.docDefaults.paragraph,
      ...tableParagraph,
      ...paragraphStyles.map((style) => style.paragraph),
      direct,
    ),
  };
  cache.set(paragraph, resolved);
  return resolved;
}

export function computeEffectiveParagraphFormat(context: StylesContext, paragraph: Element): ParagraphFormat {
  return paragraphContext(context, paragraph).effectiveParagraph;
}

export function computeEffectiveRunFormat(context: StylesContext, paragraph: Element, run: Element): RunFormat {
  const direct = readRunProperties(children(run, 'rPr')[0], context.theme);
  const resolvedParagraph = paragraphContext(context, paragraph);
  const characterStyles = resolveStyleChainOrDefault(context, direct.style ?? undefined, 'character');
  return mergeRunFormats(
    context.docDefaults.run,
    ...resolvedParagraph.tableRun,
    ...resolvedParagraph.paragraphStyles.map((style) => style.run),
    ...characterStyles.map((style) => style.run),
    direct,
  );
}

export function cloneStyleInfo(style: StyleInfo): StyleInfo {
  return {
    ...style,
    aliases: style.aliases ? [...style.aliases] : undefined,
    paragraph: cloneParagraphFormat(style.paragraph),
    run: cloneRunFormat(style.run),
  };
}
