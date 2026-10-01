import type { Element } from '@xmldom/xmldom';
import type { BorderSide, CompatibilitySettings, ColorSchemeMapping, LatentStyles, ParagraphFormat, RunFormat, Shading, StyleInfo, TabStop, ThemeFontLanguages, ThemeSettings } from './types.js';
import { WORD_NS, children, childrenThroughTransparent, wordValue } from './xml.js';
import { compactDefined } from './internal/elements.js';

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
type TableCondition = 'firstRow' | 'lastRow' | 'firstCol' | 'lastCol' | 'band1Horz' | 'band2Horz';

export interface ThemeInfo {
  colors: Record<string, string>;
  fonts: Record<string, string>;
  colorSchemeMapping?: ColorSchemeMapping;
}

interface TableStyleLayer {
  paragraph?: ParagraphFormat;
  run?: RunFormat;
}

interface TableMeta {
  rows: Element[];
  rowIndex: Map<Element, number>;
  cellIndex: WeakMap<Element, number>;
  look?: Element;
  rowBandSize: number;
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
    };
    for (const conditionElement of children(styleElement, 'tblStylePr')) {
      const condition = wordAttr(conditionElement, 'type') as TableCondition | undefined;
      if (!condition || !['firstRow', 'lastRow', 'firstCol', 'lastCol', 'band1Horz', 'band2Horz'].includes(condition)) continue;
      style.conditions![condition] = {
        paragraph: cloneParagraphFormat(readParagraphProperties(children(conditionElement, 'pPr')[0])),
        run: cloneRunFormat(readRunProperties(children(conditionElement, 'rPr')[0], theme)),
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

function readLookFlag(look: Element | undefined, name: string, defaultValue: boolean): boolean {
  const value = wordAttr(look, name);
  if (value === undefined) return defaultValue;
  return !['0', 'false', 'off'].includes(value.toLowerCase());
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
  const rowBandSize = Math.max(1, readNumber(wordValue(children(tableProps ?? table, 'tblStyleRowBandSize')[0])) ?? 1);
  const meta = { rows, rowIndex, cellIndex, look: children(tableProps ?? table, 'tblLook')[0], rowBandSize };
  cache.set(table, meta);
  return meta;
}

function tableContext(context: StylesContext, paragraph: Element): TableContext {
  const cache = context._tableContext ??= new WeakMap<Element, TableContext>();
  const existing = cache.get(paragraph);
  if (existing) return existing;
  const table = closestAncestor(paragraph, 'tbl');
  if (!table) {
    const empty = { conditions: [], chain: [] };
    cache.set(paragraph, empty);
    return empty;
  }
  const cell = closestAncestor(paragraph, 'tc');
  const row = closestAncestor(paragraph, 'tr');
  if (!cell || !row) {
    const empty = { conditions: [], chain: [] };
    cache.set(paragraph, empty);
    return empty;
  }
  const meta = tableMeta(context, table);
  const rowPosition = meta.rowIndex.get(row) ?? -1;
  const cellPosition = meta.cellIndex.get(cell) ?? -1;
  const conditions: TableCondition[] = [];
  const look = meta.look;
  const firstRow = readLookFlag(look, 'firstRow', false);
  const lastRow = readLookFlag(look, 'lastRow', false);
  const firstColumn = readLookFlag(look, 'firstColumn', false);
  const lastColumn = readLookFlag(look, 'lastColumn', false);
  if (rowPosition === 0 && firstRow) conditions.push('firstRow');
  if (rowPosition === meta.rows.length - 1 && lastRow) conditions.push('lastRow');
  const cellCount = childrenThroughTransparent(row, 'tc').length;
  if (cellPosition === 0 && firstColumn) conditions.push('firstCol');
  if (cellPosition === cellCount - 1 && lastColumn) conditions.push('lastCol');
  if (!readLookFlag(look, 'noHBand', false)) {
  const legacyRules = context.compatibilitySettings?.useWord2002TableStyleRules === true;
  const bandStart = !legacyRules && firstRow ? 1 : 0;
  const bandEnd = !legacyRules ? meta.rows.length - (lastRow ? 1 : 0) : meta.rows.length;
    if (rowPosition >= bandStart && rowPosition < bandEnd) {
      const bandIndex = Math.floor((rowPosition - bandStart) / meta.rowBandSize);
      conditions.push(bandIndex % 2 === 0 ? 'band1Horz' : 'band2Horz');
    }
  }
  const styleId = wordValue(children(children(table, 'tblPr')[0] ?? table, 'tblStyle')[0]) ?? undefined;
  const resolved = { conditions, chain: resolveStyleChainOrDefault(context, styleId, 'table') };
  cache.set(paragraph, resolved);
  return resolved;
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
