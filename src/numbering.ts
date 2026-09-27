import type { Document, Element } from '@xmldom/xmldom';
import type { NumberingDefinition, NumberingInfo, NumberingLevelDefinition, RunFormat } from './types.js';
import { children, wordValue, WORD_NS } from './xml.js';

interface StyleNumberingReference {
  numId: number;
  level: number;
}

interface RawAbstractNumbering {
  abstractNumId: number;
  multiLevelType?: string;
  nsid?: string;
  tmpl?: string;
  styleLink?: string;
  numStyleLink?: string;
  levels: Map<number, NumberingLevelDefinition>;
}

interface RawLevelOverride {
  startOverride?: number;
  level?: NumberingLevelDefinition;
}

interface RawNumberingInstance {
  numId: number;
  abstractNumId: number;
  overrides: Map<number, RawLevelOverride>;
}

interface ResolvedAbstractNumbering extends Omit<NumberingDefinition, 'numId' | 'levels'> {
  levels: Map<number, NumberingLevelDefinition>;
}

interface ResolvedNumberingInstance extends NumberingDefinition {
  levelMap: Map<number, NumberingLevelDefinition>;
}

export interface NumberingModel {
  definitions: NumberingDefinition[];
  styles: Map<string, StyleNumberingReference>;
  resolveNumbering(numId: number): NumberingDefinition | undefined;
}

function parseInteger(value: string | null | undefined): number | undefined {
  if (value === undefined || value === null || value === '') return undefined;
  const parsed = Number(value);
  return Number.isSafeInteger(parsed) ? parsed : undefined;
}

function parseTwips(value: string | null | undefined): number | undefined {
  const parsed = parseInteger(value);
  return parsed === undefined ? undefined : parsed;
}

function parseOnOff(element: Element | undefined): boolean | undefined {
  if (!element) return undefined;
  const value = wordValue(element);
  return value ? !['0', 'false', 'off'].includes(value) : true;
}

function parseSuffix(value: string | undefined): 'tab' | 'space' | 'nothing' {
  return value === 'space' || value === 'nothing' ? value : 'tab';
}

function parseRunFormat(props: Element | undefined): RunFormat | undefined {
  if (!props) return undefined;
  const get = (name: string) => children(props, name)[0];
  const toggle = (name: string) => get(name) ? !['0', 'false', 'off'].includes(wordValue(get(name)) ?? '') : undefined;
  const size = wordValue(get('sz'));
  const underline = get('u');
  const color = wordValue(get('color'));
  const format: RunFormat = {};
  if (toggle('b') !== undefined) format.bold = toggle('b');
  if (toggle('i') !== undefined) format.italic = toggle('i');
  if (underline) format.underline = !['none', '0', 'false'].includes(wordValue(underline) ?? '');
  if (size && Number.isFinite(Number(size))) format.fontSize = Number(size) / 2;
  const fontFamily = get('rFonts')?.getAttributeNS(WORD_NS, 'ascii')
    ?? get('rFonts')?.getAttributeNS(WORD_NS, 'hAnsi')
    ?? undefined;
  if (fontFamily) format.fontFamily = fontFamily;
  if (color && /^[a-f\d]{6}$/i.test(color)) format.color = color;
  return Object.keys(format).length ? format : undefined;
}

function defaultLevelText(level: number, format: string): string {
  return format === 'bullet' ? '•' : format === 'none' ? '' : `%${level + 1}.`;
}

function parseLevel(level: Element, explicitLevel?: number): NumberingLevelDefinition | undefined {
  const parsedLevel = explicitLevel ?? parseInteger(level.getAttributeNS(WORD_NS, 'ilvl'));
  if (parsedLevel === undefined || parsedLevel < 0 || parsedLevel > 8) return undefined;
  const props = children(level, 'pPr')[0];
  const ind = props ? children(props, 'ind')[0] : undefined;
  const format = wordValue(children(level, 'numFmt')[0]) ?? 'decimal';
  return {
    level: parsedLevel,
    start: parseInteger(wordValue(children(level, 'start')[0])) ?? 1,
    format,
    text: wordValue(children(level, 'lvlText')[0]) ?? defaultLevelText(parsedLevel, format),
    justification: wordValue(children(level, 'lvlJc')[0]),
    suffix: parseSuffix(wordValue(children(level, 'suff')[0])),
    isLegal: parseOnOff(children(level, 'isLgl')[0]),
    restart: parseInteger(wordValue(children(level, 'lvlRestart')[0])),
    paragraphStyle: wordValue(children(level, 'pStyle')[0]),
    indentLeft: parseTwips(ind?.getAttributeNS(WORD_NS, 'left') ?? ind?.getAttributeNS(WORD_NS, 'start')),
    indentHanging: parseTwips(ind?.getAttributeNS(WORD_NS, 'hanging')),
    runFormat: parseRunFormat(children(level, 'rPr')[0]),
  };
}

function parseStyles(stylesDocument?: Document): Map<string, StyleNumberingReference> {
  const result = new Map<string, StyleNumberingReference>();
  const root = stylesDocument?.documentElement;
  if (!root || root.namespaceURI !== WORD_NS || root.localName !== 'styles') return result;
  for (const style of children(root, 'style')) {
    if (style.getAttributeNS(WORD_NS, 'type') !== 'paragraph') continue;
    const styleId = style.getAttributeNS(WORD_NS, 'styleId');
    if (!styleId) continue;
    const numPr = children(children(style, 'pPr')[0] ?? style, 'numPr')[0];
    if (!numPr) continue;
    const numId = parseInteger(wordValue(children(numPr, 'numId')[0]));
    if (numId === undefined || numId < 1) continue;
    result.set(styleId, {
      numId,
      level: parseInteger(wordValue(children(numPr, 'ilvl')[0])) ?? 0,
    });
  }
  return result;
}

function parseAbstracts(numberingDocument?: Document): Map<number, RawAbstractNumbering> {
  const result = new Map<number, RawAbstractNumbering>();
  const root = numberingDocument?.documentElement;
  if (!root || root.namespaceURI !== WORD_NS || root.localName !== 'numbering') return result;
  for (const abstract of children(root, 'abstractNum')) {
    const abstractNumId = parseInteger(abstract.getAttributeNS(WORD_NS, 'abstractNumId'));
    if (abstractNumId === undefined) continue;
    const levels = new Map<number, NumberingLevelDefinition>();
    for (const level of children(abstract, 'lvl')) {
      const parsed = parseLevel(level);
      if (parsed) levels.set(parsed.level, parsed);
    }
    result.set(abstractNumId, {
      abstractNumId,
      multiLevelType: wordValue(children(abstract, 'multiLevelType')[0]),
      nsid: wordValue(children(abstract, 'nsid')[0]),
      tmpl: wordValue(children(abstract, 'tmpl')[0]),
      styleLink: wordValue(children(abstract, 'styleLink')[0]),
      numStyleLink: wordValue(children(abstract, 'numStyleLink')[0]),
      levels,
    });
  }
  return result;
}

function parseNums(numberingDocument?: Document): Map<number, RawNumberingInstance> {
  const result = new Map<number, RawNumberingInstance>();
  const root = numberingDocument?.documentElement;
  if (!root || root.namespaceURI !== WORD_NS || root.localName !== 'numbering') return result;
  for (const num of children(root, 'num')) {
    const numId = parseInteger(num.getAttributeNS(WORD_NS, 'numId'));
    const abstractNumId = parseInteger(wordValue(children(num, 'abstractNumId')[0]));
    if (numId === undefined || abstractNumId === undefined) continue;
    const overrides = new Map<number, RawLevelOverride>();
    for (const override of children(num, 'lvlOverride')) {
      const level = parseInteger(override.getAttributeNS(WORD_NS, 'ilvl'));
      if (level === undefined || level < 0 || level > 8) continue;
      const levelElement = children(override, 'lvl')[0];
      overrides.set(level, {
        startOverride: parseInteger(wordValue(children(override, 'startOverride')[0])),
        level: levelElement ? parseLevel(levelElement, level) : undefined,
      });
    }
    result.set(numId, { numId, abstractNumId, overrides });
  }
  return result;
}

function mergeLevel(base: NumberingLevelDefinition | undefined, override: NumberingLevelDefinition | undefined): NumberingLevelDefinition | undefined {
  if (!base && !override) return undefined;
  return {
    level: override?.level ?? base?.level ?? 0,
    start: override?.start ?? base?.start ?? 1,
    format: override?.format ?? base?.format ?? 'decimal',
    text: override?.text ?? base?.text ?? defaultLevelText(override?.level ?? base?.level ?? 0, override?.format ?? base?.format ?? 'decimal'),
    justification: override?.justification ?? base?.justification,
    suffix: override?.suffix ?? base?.suffix ?? 'tab',
    isLegal: override?.isLegal ?? base?.isLegal,
    restart: override?.restart ?? base?.restart,
    paragraphStyle: override?.paragraphStyle ?? base?.paragraphStyle,
    indentLeft: override?.indentLeft ?? base?.indentLeft,
    indentHanging: override?.indentHanging ?? base?.indentHanging,
    runFormat: override?.runFormat ?? base?.runFormat,
  };
}

export function parseNumberingModel(numberingDocument?: Document, stylesDocument?: Document): NumberingModel {
  const styles = parseStyles(stylesDocument);
  const abstracts = parseAbstracts(numberingDocument);
  const nums = parseNums(numberingDocument);
  const abstractMemo = new Map<number, ResolvedAbstractNumbering | null>();
  const numMemo = new Map<number, ResolvedNumberingInstance | null>();

  function resolveStyleLink(styleId: string, seenNums: Set<number>, seenAbstracts: Set<number>, seenStyles: Set<string>): ResolvedAbstractNumbering | undefined {
    if (seenStyles.has(styleId)) return undefined;
    const reference = styles.get(styleId);
    if (!reference) return undefined;
    seenStyles.add(styleId);
    const resolved = resolveNumbering(reference.numId, seenNums, seenAbstracts, seenStyles);
    seenStyles.delete(styleId);
    if (!resolved) return undefined;
    return {
      abstractNumId: resolved.abstractNumId,
      multiLevelType: resolved.multiLevelType,
      nsid: resolved.nsid,
      tmpl: resolved.tmpl,
      styleLink: resolved.styleLink,
      numStyleLink: resolved.numStyleLink,
      levels: new Map(resolved.levelMap),
    };
  }

  function resolveAbstract(abstractNumId: number, seenNums: Set<number>, seenAbstracts: Set<number>, seenStyles: Set<string>): ResolvedAbstractNumbering | undefined {
    if (abstractMemo.has(abstractNumId)) return abstractMemo.get(abstractNumId) ?? undefined;
    if (seenAbstracts.has(abstractNumId)) return undefined;
    const raw = abstracts.get(abstractNumId);
    if (!raw) return undefined;
    seenAbstracts.add(abstractNumId);
    const linked = raw.numStyleLink
      ? resolveStyleLink(raw.numStyleLink, seenNums, seenAbstracts, seenStyles)
      : raw.styleLink
        ? resolveStyleLink(raw.styleLink, seenNums, seenAbstracts, seenStyles)
        : undefined;
    const levels = new Map<number, NumberingLevelDefinition>(linked?.levels ?? []);
    for (const [level, definition] of raw.levels) {
      levels.set(level, mergeLevel(levels.get(level), definition)!);
    }
    seenAbstracts.delete(abstractNumId);
    const resolved: ResolvedAbstractNumbering = {
      abstractNumId,
      multiLevelType: raw.multiLevelType ?? linked?.multiLevelType,
      nsid: raw.nsid ?? linked?.nsid,
      tmpl: raw.tmpl ?? linked?.tmpl,
      styleLink: raw.styleLink ?? linked?.styleLink,
      numStyleLink: raw.numStyleLink ?? linked?.numStyleLink,
      levels,
    };
    abstractMemo.set(abstractNumId, resolved);
    return resolved;
  }

  function resolveNumbering(numId: number, seenNums = new Set<number>(), seenAbstracts = new Set<number>(), seenStyles = new Set<string>()): ResolvedNumberingInstance | undefined {
    if (numMemo.has(numId)) return numMemo.get(numId) ?? undefined;
    if (seenNums.has(numId)) return undefined;
    const raw = nums.get(numId);
    if (!raw) return undefined;
    seenNums.add(numId);
    const abstract = resolveAbstract(raw.abstractNumId, seenNums, seenAbstracts, seenStyles);
    seenNums.delete(numId);
    if (!abstract) return undefined;
    const levelMap = new Map<number, NumberingLevelDefinition>(abstract.levels);
    for (const [level, override] of raw.overrides) {
      const merged = mergeLevel(levelMap.get(level), override.level);
      if (merged) {
        if (override.startOverride !== undefined) merged.start = override.startOverride;
        levelMap.set(level, merged);
      }
    }
    const definition: ResolvedNumberingInstance = {
      numId,
      abstractNumId: abstract.abstractNumId,
      multiLevelType: abstract.multiLevelType,
      nsid: abstract.nsid,
      tmpl: abstract.tmpl,
      styleLink: abstract.styleLink,
      numStyleLink: abstract.numStyleLink,
      levels: [...levelMap.values()].sort((a, b) => a.level - b.level),
      levelMap,
    };
    numMemo.set(numId, definition);
    return definition;
  }

  const definitions = [...nums.keys()]
    .sort((a, b) => a - b)
    .map(numId => resolveNumbering(numId))
    .filter((definition): definition is ResolvedNumberingInstance => Boolean(definition))
    .map(({ levelMap, ...definition }) => definition);

  return {
    definitions,
    styles,
    resolveNumbering(numId: number) {
      const resolved = resolveNumbering(numId);
      if (!resolved) return undefined;
      const { levelMap, ...definition } = resolved;
      return definition;
    },
  };
}

function effectiveParagraphNumbering(paragraph: Element, styles: Map<string, StyleNumberingReference>): StyleNumberingReference | undefined {
  const props = children(paragraph, 'pPr')[0];
  const numPr = props ? children(props, 'numPr')[0] : undefined;
  const numId = parseInteger(wordValue(numPr ? children(numPr, 'numId')[0] : undefined));
  if (numId !== undefined) {
    if (numId < 1) return undefined;
    return { numId, level: parseInteger(wordValue(children(numPr!, 'ilvl')[0])) ?? 0 };
  }
  const styleId = props ? wordValue(children(props, 'pStyle')[0]) : undefined;
  return styleId ? styles.get(styleId) : undefined;
}

function letterNumber(value: number, upper: boolean): string {
  let current = Math.max(1, value);
  let result = '';
  while (current > 0) {
    current--;
    result = String.fromCharCode((upper ? 65 : 97) + (current % 26)) + result;
    current = Math.floor(current / 26);
  }
  return result;
}

function romanNumber(value: number, upper: boolean): string {
  const numerals: Array<[number, string]> = [[1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'], [100, 'C'], [90, 'XC'], [50, 'L'], [40, 'XL'], [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I']];
  let current = Math.max(1, value);
  let result = '';
  for (const [amount, numeral] of numerals) {
    while (current >= amount) {
      result += numeral;
      current -= amount;
    }
  }
  return upper ? result : result.toLowerCase();
}

function ordinalNumber(value: number): string {
  const mod100 = value % 100;
  const mod10 = value % 10;
  const suffix = mod100 >= 11 && mod100 <= 13 ? 'th' : mod10 === 1 ? 'st' : mod10 === 2 ? 'nd' : mod10 === 3 ? 'rd' : 'th';
  return `${value}${suffix}`;
}

function enclosedCircleNumber(value: number): string {
  if (value >= 1 && value <= 20) return String.fromCodePoint(0x2460 + value - 1);
  if (value >= 21 && value <= 35) return String.fromCodePoint(0x3251 + value - 21);
  return String(value);
}

function ideographDigits(value: number): string {
  if (value === 0) return '〇';
  const digits = ['零', '一', '二', '三', '四', '五', '六', '七', '八', '九'];
  const units = ['', '十', '百', '千'];
  const groups = ['', '万', '亿'];
  let current = Math.max(0, Math.floor(value));
  let groupIndex = 0;
  let result = '';
  while (current > 0) {
    const chunk = current % 10000;
    if (chunk) {
      let chunkText = '';
      let place = 0;
      let number = chunk;
      let needsZero = false;
      while (number > 0) {
        const digit = number % 10;
        if (digit) {
          chunkText = `${digits[digit]}${units[place]}${chunkText}`;
          needsZero = true;
        } else if (needsZero && !chunkText.startsWith(digits[0]!)) {
          chunkText = `${digits[0]}${chunkText}`;
          needsZero = false;
        }
        number = Math.floor(number / 10);
        place++;
      }
      chunkText = chunkText.replace(/零+/g, '零').replace(/零$/g, '').replace(/^一十/, '十');
      result = `${chunkText}${groups[groupIndex]}${result}`;
    } else if (result && !result.startsWith(digits[0]!)) {
      result = `${digits[0]}${result}`;
    }
    current = Math.floor(current / 10000);
    groupIndex++;
  }
  return result.replace(/零+/g, '零').replace(/零$/g, '');
}

function normalizeBullet(text: string, runFormat?: RunFormat): string {
  const font = runFormat?.fontFamily?.toLowerCase() ?? '';
  const fallback = font.includes('wingdings') ? '▪' : font.includes('symbol') ? '•' : '•';
  const map = new Map<string, string>([
    ['\uF0B7', '•'],
    ['\u2022', '•'],
    ['•', '•'],
    ['\uF0A7', '◦'],
    ['◦', '◦'],
    ['\u25E6', '◦'],
    ['\uF0A8', '▪'],
    ['▪', '▪'],
    ['\u25AA', '▪'],
  ]);
  const visible = [...(text || fallback)].map(char => map.get(char) ?? (/[\uE000-\uF8FF]/u.test(char) ? fallback : char)).join('');
  return visible || fallback;
}

function formatCounter(value: number, format: string, level: NumberingLevelDefinition | undefined, legal: boolean): string {
  const effective = legal ? 'decimal' : format;
  switch (effective) {
    case 'bullet': return normalizeBullet(level?.text ?? '•', level?.runFormat);
    case 'lowerLetter': return letterNumber(value, false);
    case 'upperLetter': return letterNumber(value, true);
    case 'lowerRoman': return romanNumber(value, false);
    case 'upperRoman': return romanNumber(value, true);
    case 'none': return '';
    case 'decimalZero': return value < 10 ? `0${value}` : String(value);
    case 'ordinal': return ordinalNumber(value);
    case 'chineseCounting':
    case 'chineseCountingThousand':
    case 'ideographDigital':
    case 'japaneseCounting':
    case 'taiwaneseCounting':
      return ideographDigits(value);
    case 'decimalEnclosedCircle':
      return enclosedCircleNumber(value);
    case 'decimal':
    default:
      return String(value);
  }
}

function shouldRestart(definition: NumberingLevelDefinition | undefined, changedLevel: number, currentLevel: number): boolean {
  if (!definition || currentLevel <= changedLevel) return false;
  if (definition.restart === 0) return false;
  const restartLevel = definition.restart !== undefined ? Math.max(0, definition.restart - 1) : currentLevel - 1;
  return changedLevel <= restartLevel;
}

export function computeParagraphNumbering(paragraphs: Element[], model: NumberingModel): Map<Element, NumberingInfo> {
  const resolved = new Map<number, ResolvedNumberingInstance>();
  const states = new Map<number, number[]>();
  const result = new Map<Element, NumberingInfo>();

  const resolveConcrete = (numId: number): ResolvedNumberingInstance | undefined => {
    if (resolved.has(numId)) return resolved.get(numId);
    const definition = model.resolveNumbering(numId);
    if (!definition) return undefined;
    const concrete: ResolvedNumberingInstance = { ...definition, levelMap: new Map(definition.levels.map(level => [level.level, level])) };
    resolved.set(numId, concrete);
    return concrete;
  };

  for (const paragraph of paragraphs) {
    const reference = effectiveParagraphNumbering(paragraph, model.styles);
    if (!reference) continue;
    const definition = resolveConcrete(reference.numId);
    const level = Math.max(0, Math.min(8, reference.level ?? 0));
    const levelDefinition = definition?.levelMap.get(level);
    if (!definition || !levelDefinition) continue;
    const counts = states.get(reference.numId) ?? Array(9).fill(0);
    counts[level] = counts[level] ? counts[level]! + 1 : (levelDefinition.start ?? 1);
    for (let deeper = level + 1; deeper < counts.length; deeper++) {
      if (shouldRestart(definition.levelMap.get(deeper), level, deeper)) counts[deeper] = 0;
    }
    states.set(reference.numId, counts);
    const legal = Boolean(levelDefinition.isLegal);
    const text = levelDefinition.format === 'bullet'
      ? normalizeBullet(levelDefinition.text, levelDefinition.runFormat)
      : levelDefinition.text.replace(/%([1-9])/g, (_, placeholder: string) => {
          const targetLevel = Number(placeholder) - 1;
          const value = counts[targetLevel];
          if (!value) return '';
          const target = definition.levelMap.get(targetLevel);
          return formatCounter(value, target?.format ?? 'decimal', target, legal);
        });
    result.set(paragraph, {
      numId: reference.numId,
      level,
      format: levelDefinition.format,
      text,
      isBullet: levelDefinition.format === 'bullet',
      indentLeft: levelDefinition.indentLeft,
      indentHanging: levelDefinition.indentHanging,
      suffix: levelDefinition.suffix,
      runFormat: levelDefinition.runFormat,
    });
  }
  return result;
}
