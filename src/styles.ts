import type { Element } from '@xmldom/xmldom';
import type { ParagraphFormat, RunFormat, StyleInfo } from './types.js';
import { WORD_NS, children, wordValue } from './xml.js';

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

interface ThemeInfo {
  colors: Record<string, string>;
  fonts: Record<string, string>;
}

interface TableStyleLayer {
  paragraph?: ParagraphFormat;
  run?: RunFormat;
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
}

function wordAttr(element: Element | undefined, name: string): string | undefined {
  return element?.getAttributeNS(WORD_NS, name) ?? undefined;
}

function normalizeHex(value: string | undefined): string | undefined {
  return value && /^[0-9a-f]{6}$/i.test(value) ? value.toUpperCase() : undefined;
}

function readOnOff(element: Element | undefined): boolean | undefined {
  if (!element) return undefined;
  const value = (wordValue(element) ?? '1').toLowerCase();
  return ['0', 'false', 'off'].includes(value) ? false : true;
}

function readNumber(value: string | undefined): number | undefined {
  return value !== undefined && /^-?\d+$/.test(value) ? Number(value) : undefined;
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
      if (value !== undefined) (merged as Record<string, unknown>)[key] = value;
    }
  }
  return merged;
}

function mergeRunFormats(...formats: Array<RunFormat | undefined>): RunFormat {
  const merged: RunFormat = {};
  for (const format of formats) {
    if (!format) continue;
    for (const [key, value] of Object.entries(format) as Array<[keyof RunFormat, RunFormat[keyof RunFormat]]>) {
      if (value !== undefined) (merged as Record<string, unknown>)[key] = value;
    }
  }
  return merged;
}

function applyShadeTint(hex: string | undefined, shade: string | undefined, tint: string | undefined): string | undefined {
  const base = normalizeHex(hex);
  if (!base) return undefined;
  const shadeValue = shade && /^[0-9a-f]{2}$/i.test(shade) ? parseInt(shade, 16) / 255 : undefined;
  const tintValue = tint && /^[0-9a-f]{2}$/i.test(tint) ? parseInt(tint, 16) / 255 : undefined;
  const channels = base.match(/../g)!.map((channel) => parseInt(channel, 16));
  const transformed = channels.map((channel) => {
    let result = channel;
    if (shadeValue !== undefined) result = Math.round(result * shadeValue);
    if (tintValue !== undefined) result = Math.round(result + (255 - result) * tintValue);
    return Math.max(0, Math.min(255, result));
  });
  return transformed.map((channel) => channel.toString(16).padStart(2, '0').toUpperCase()).join('');
}

function resolveThemeColor(theme: ThemeInfo, element: Element | undefined): string | undefined {
  if (!element) return undefined;
  const direct = normalizeHex(wordAttr(element, 'val'));
  if (direct && direct.toLowerCase() !== 'auto') return direct;
  const themeColor = wordAttr(element, 'themeColor');
  return applyShadeTint(theme.colors[themeColor ?? ''], wordAttr(element, 'themeShade'), wordAttr(element, 'themeTint'));
}

function resolveUnderlineColor(theme: ThemeInfo, element: Element | undefined): string | undefined {
  if (!element) return undefined;
  const direct = normalizeHex(wordAttr(element, 'color'));
  if (direct && direct.toLowerCase() !== 'auto') return direct;
  const themeColor = wordAttr(element, 'themeColor');
  return applyShadeTint(theme.colors[themeColor ?? ''], wordAttr(element, 'themeShade'), wordAttr(element, 'themeTint'));
}

function resolveThemeFont(theme: ThemeInfo, value: string | undefined, fallback?: string): string | undefined {
  return value ? theme.fonts[value] ?? fallback : fallback;
}

function readFontFamily(theme: ThemeInfo, fonts: Element | undefined): { fontFamily?: string; fontFamilyEastAsia?: string } {
  if (!fonts) return {};
  const ascii = wordAttr(fonts, 'ascii') ?? wordAttr(fonts, 'hAnsi');
  const eastAsia = wordAttr(fonts, 'eastAsia');
  const fontFamily = ascii
    ?? resolveThemeFont(theme, wordAttr(fonts, 'asciiTheme') ?? wordAttr(fonts, 'hAnsiTheme'));
  const fontFamilyEastAsia = eastAsia
    ?? resolveThemeFont(theme, wordAttr(fonts, 'eastAsiaTheme'), fontFamily);
  return { fontFamily, fontFamilyEastAsia };
}

export function readParagraphProperties(props: Element | undefined): ParagraphFormat {
  if (!props) return {};
  const spacing = children(props, 'spacing')[0];
  const indent = children(props, 'ind')[0];
  const alignment = wordValue(children(props, 'jc')[0]);
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
    lineSpacing: readNumber(wordAttr(spacing, 'line')),
    lineSpacingRule: wordAttr(spacing, 'lineRule') as ParagraphFormat['lineSpacingRule'] | undefined,
    keepNext: readOnOff(children(props, 'keepNext')[0]),
    keepLines: readOnOff(children(props, 'keepLines')[0]),
    pageBreakBefore: readOnOff(children(props, 'pageBreakBefore')[0]),
    widowControl: readOnOff(children(props, 'widowControl')[0]),
    outlineLevel: readNumber(wordValue(children(props, 'outlineLvl')[0])),
  };
}

export function readRunProperties(props: Element | undefined, theme: StylesContext['theme']): RunFormat {
  if (!props) return {};
  const underline = children(props, 'u')[0];
  const underlineValue = wordValue(underline);
  const size = wordValue(children(props, 'sz')[0]) ?? wordValue(children(props, 'szCs')[0]);
  const fonts = children(props, 'rFonts')[0];
  return {
    style: wordValue(children(props, 'rStyle')[0]),
    bold: readOnOff(children(props, 'b')[0]),
    italic: readOnOff(children(props, 'i')[0]),
    underline: underline ? !['none', '0', 'false'].includes((underlineValue ?? 'single').toLowerCase()) : undefined,
    underlineStyle: underline && underlineValue && !['0', 'false', 'none'].includes(underlineValue.toLowerCase()) ? underlineValue : undefined,
    underlineColor: resolveUnderlineColor(theme, underline),
    fontSize: size && Number.isFinite(Number(size)) ? Number(size) / 2 : undefined,
    ...readFontFamily(theme, fonts),
    color: resolveThemeColor(theme, children(props, 'color')[0]),
    strike: readOnOff(children(props, 'strike')[0]),
    doubleStrike: readOnOff(children(props, 'dstrike')[0]),
    verticalAlign: wordValue(children(props, 'vertAlign')[0]) as RunFormat['verticalAlign'] | undefined,
    smallCaps: readOnOff(children(props, 'smallCaps')[0]),
    allCaps: readOnOff(children(props, 'caps')[0]),
    highlight: wordValue(children(props, 'highlight')[0]) ?? undefined,
    characterSpacing: readNumber(wordValue(children(props, 'spacing')[0])),
  };
}

function themeColorValue(node: Element | undefined): string | undefined {
  if (!node) return undefined;
  const srgb = node.getAttribute('val') ?? node.getAttribute('lastClr') ?? undefined;
  return normalizeHex(srgb);
}

function parseTheme(themeElement: Element | undefined): ThemeInfo {
  if (!themeElement) return { colors: { ...DEFAULT_THEME_COLORS }, fonts: { ...DEFAULT_THEME_FONTS } };
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
  return { colors, fonts };
}

function parseStyleType(value: string | undefined): StyleType | undefined {
  return ['paragraph', 'character', 'table', 'numbering'].includes(value ?? '') ? value as StyleType : undefined;
}

export function parseStyles(stylesRoot: Element | undefined, themeRoot?: Element): StylesContext {
  const theme = parseTheme(themeRoot);
  if (!stylesRoot || stylesRoot.namespaceURI !== WORD_NS || stylesRoot.localName !== 'styles') {
    return { docDefaults: { paragraph: {}, run: {} }, styles: [], byId: new Map(), defaults: {}, theme };
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
  };
}

function resolveStyleChain(context: StylesContext, id: string | undefined, type: StyleType): ParsedStyle[] {
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

function resolveStyleChainOrDefault(context: StylesContext, id: string | undefined, type: StyleType): ParsedStyle[] {
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

function tableConditions(paragraph: Element): TableCondition[] {
  const cell = closestAncestor(paragraph, 'tc');
  const row = closestAncestor(paragraph, 'tr');
  const table = closestAncestor(paragraph, 'tbl');
  if (!cell || !row || !table) return [];
  const tableProps = children(table, 'tblPr')[0];
  const look = children(tableProps ?? table, 'tblLook')[0];
  const rows = children(table, 'tr');
  const cells = children(row, 'tc');
  const rowIndex = rows.indexOf(row);
  const cellIndex = cells.indexOf(cell);
  const enabled = (name: string, defaultValue = true): boolean => {
    const value = wordAttr(look, name);
    if (value === undefined) return defaultValue;
    return !['0', 'false', 'off'].includes(value.toLowerCase());
  };
  const conditions: TableCondition[] = [];
  if (rowIndex === 0 && enabled('firstRow')) conditions.push('firstRow');
  if (rowIndex === rows.length - 1 && enabled('lastRow')) conditions.push('lastRow');
  if (cellIndex === 0 && enabled('firstColumn')) conditions.push('firstCol');
  if (cellIndex === cells.length - 1 && enabled('lastColumn')) conditions.push('lastCol');
  if (!enabled('noHBand', false)) conditions.push(((rowIndex + 1) % 2 === 1 ? 'band1Horz' : 'band2Horz'));
  return conditions;
}

function tableStyleChain(context: StylesContext, paragraph: Element): ParsedStyle[] {
  const table = closestAncestor(paragraph, 'tbl');
  const styleId = wordValue(children(children(table ?? paragraph, 'tblPr')[0] ?? paragraph, 'tblStyle')[0]) ?? undefined;
  return resolveStyleChainOrDefault(context, styleId, 'table');
}

function tableParagraphFormats(context: StylesContext, paragraph: Element): ParagraphFormat[] {
  const conditions = tableConditions(paragraph);
  const chain = tableStyleChain(context, paragraph);
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
  const conditions = tableConditions(paragraph);
  const chain = tableStyleChain(context, paragraph);
  const formats: RunFormat[] = [];
  for (const style of chain) {
    formats.push(style.run ?? {});
    for (const condition of conditions) {
      if (style.conditions?.[condition]?.run) formats.push(style.conditions[condition]!.run!);
    }
  }
  return formats;
}

export function computeEffectiveParagraphFormat(context: StylesContext, paragraph: Element): ParagraphFormat {
  const direct = readParagraphProperties(children(paragraph, 'pPr')[0]);
  const paragraphStyles = resolveStyleChainOrDefault(context, direct.style, 'paragraph');
  return mergeParagraphFormats(
    context.docDefaults.paragraph,
    ...tableParagraphFormats(context, paragraph),
    ...paragraphStyles.map((style) => style.paragraph),
    direct,
  );
}

export function computeEffectiveRunFormat(context: StylesContext, paragraph: Element, run: Element): RunFormat {
  const direct = readRunProperties(children(run, 'rPr')[0], context.theme);
  const paragraphStyleId = readParagraphProperties(children(paragraph, 'pPr')[0]).style;
  const paragraphStyles = resolveStyleChainOrDefault(context, paragraphStyleId, 'paragraph');
  const characterStyles = resolveStyleChainOrDefault(context, direct.style, 'character');
  return mergeRunFormats(
    context.docDefaults.run,
    ...tableRunFormats(context, paragraph),
    ...paragraphStyles.map((style) => style.run),
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
