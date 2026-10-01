import type { Element, Node } from '@xmldom/xmldom';
import type { EquationNode, FieldInfo, FieldKind, FieldSwitch } from './types.js';
import { WORD_NS, children, descendants } from './xml.js';

const PAGINATION = new Set<FieldKind>(['PAGE', 'NUMPAGES', 'PAGEREF', 'TOC', 'INDEX']);
const NEVER_EVALUATE = new Set<FieldKind>([
  'INCLUDETEXT', 'INCLUDEPICTURE', 'LINK', 'DDE', 'DDEAUTO', 'MACROBUTTON', 'GOTOBUTTON',
  'FILLIN', 'ASK', 'DATABASE', 'AUTOTEXT', 'AUTOTEXTLIST', 'MERGEFIELD',
  'FORMTEXT', 'FORMCHECKBOX', 'FORMDROPDOWN',
  // EQ 是排版域，不是取值域：它的「结果」就是 Word 缓存的那段文字，没有可重算的值。
  'EQ',
]);

/** EQ 指令的嵌套上限。和公式转换一样设一道硬界，不让文档内容决定递归多深。 */
const MAX_EQUATION_DEPTH = 16;

/**
 * 解析 `EQ` 域的指令。EQ 用的是括号加分隔符的语法（`\o\ac(\s\up 10(○),甲)`），
 * 不是「空白分词 + 反斜杠开关」那一套——按开关语法去读，`\o\ac(\s\up` 会被当成开关名，
 * 还会凭空解析出一个 argument。所以 EQ 单独走这里。
 *
 * 分隔符按 Word 的区域设置可能是 `,` 也可能是 `;`，两个都认。
 */
export function parseEquationInstruction(instruction: string): EquationNode | undefined {
  const body = instruction.replace(/^\s*EQ\b/i, '');
  if (body === instruction) return undefined;
  let cursor = 0;
  const parseParts = (depth: number): EquationNode[] => {
    const parts: EquationNode[] = [];
    let current = parseElement(depth);
    parts.push(current);
    while (cursor < body.length && (body[cursor] === ',' || body[cursor] === ';')) {
      cursor++;
      current = parseElement(depth);
      parts.push(current);
    }
    return parts;
  };
  const parseElement = (depth: number): EquationNode => {
    while (body[cursor] === ' ') cursor++;
    if (depth > MAX_EQUATION_DEPTH) {
      // 超限时把剩下的原样当文字收掉，不继续递归，也不丢内容。
      const rest = body.slice(cursor);
      cursor = body.length;
      return { text: rest };
    }
    const node: EquationNode = {};
    while (body[cursor] === '\\') {
      cursor++;
      const name = /^[A-Za-z]+/.exec(body.slice(cursor))?.[0] ?? '';
      cursor += name.length;
      if (node.switch === undefined) node.switch = name.toLowerCase();
      else (node.options ??= []).push(name.toLowerCase());
      while (body[cursor] === ' ') cursor++;
      const number = /^-?\d+(?:\.\d+)?/.exec(body.slice(cursor))?.[0];
      if (number !== undefined) {
        cursor += number.length;
        const magnitude = Number(number);
        // \do 是往下，记成负数，这样调用方只看一个数就够了。
        node.raisePoints = node.options?.includes('do') ? -magnitude : magnitude;
        while (body[cursor] === ' ') cursor++;
      }
    }
    if (body[cursor] === '(') {
      cursor++;
      node.parts = parseParts(depth + 1);
      if (body[cursor] === ')') cursor++;
    } else {
      const start = cursor;
      while (cursor < body.length && !'(),;\\'.includes(body[cursor]!)) cursor++;
      const text = body.slice(start, cursor);
      if (text) node.text = text;
    }
    return node;
  };
  const parts = parseParts(0);
  const meaningful = (node: EquationNode): boolean =>
    node.switch !== undefined || node.text !== undefined || (node.parts?.length ?? 0) > 0;
  const kept = parts.filter(meaningful);
  if (!kept.length) return undefined;
  return kept.length === 1 ? kept[0]! : { parts: kept };
}

export interface ParsedFields {
  fields: FieldInfo[];
  roles: Map<Element, { index: number; role: 'instruction' | 'result'; kind?: FieldKind; instruction?: string }>;
}

export function formatPageNumber(number: number, format = 'decimal'): string {
  if (!Number.isFinite(number) || number < 1) return '0';
  const value = Math.trunc(number);
  if (format === 'upperRoman' || format === 'lowerRoman') {
    const digits: Array<[number, string]> = [[1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'], [100, 'C'], [90, 'XC'], [50, 'L'], [40, 'XL'], [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I']];
    let result = '';
    let rest = value;
    for (const [unit, glyph] of digits) while (rest >= unit) { result += glyph; rest -= unit; }
    return format === 'lowerRoman' ? result.toLowerCase() : result;
  }
  if (format === 'upperLetter' || format === 'lowerLetter') {
    let result = '';
    let rest = value;
    while (rest > 0) { rest--; result = String.fromCharCode(65 + (rest % 26)) + result; rest = Math.floor(rest / 26); }
    return format === 'lowerLetter' ? result.toLowerCase() : result;
  }
  if (format === 'chineseCounting') {
    const digits = '〇一二三四五六七八九';
    return String(value).split('').map((digit) => digits[Number(digit)] ?? digit).join('');
  }
  return String(value);
}

export function pageFieldResult(kind: FieldKind, pageNumber: number, pageCount: number, format?: string): string | undefined {
  if (kind === 'PAGE') return formatPageNumber(pageNumber, format);
  if (kind === 'NUMPAGES') return formatPageNumber(pageCount, format);
  return undefined;
}

function visibleText(element: Element): string {
  const values: string[] = [];
  const walk = (node: Node): void => {
    for (let child = node.firstChild; child; child = child.nextSibling) {
      if (child.nodeType !== 1) continue;
      const entry = child as Element;
      if (entry.namespaceURI === WORD_NS) {
        if (entry.localName === 't' || entry.localName === 'delText') values.push(entry.textContent ?? '');
        else if (entry.localName === 'tab') values.push('\t');
        else if (entry.localName === 'br' || entry.localName === 'cr') values.push('\n');
        else if (entry.localName === 'instrText' || entry.localName === 'fldChar') continue;
      }
      walk(entry);
    }
  };
  walk(element);
  return values.join('');
}

function attr(element: Element, name: string): string | undefined {
  return element.getAttributeNS(WORD_NS, name) ?? element.getAttribute(`w:${name}`) ?? undefined;
}

function parseInstruction(instruction: string):
  { kind: FieldKind; argument?: string; switches: FieldSwitch[]; equation?: EquationNode } {
  const tokens = instruction.match(/"[^"]*"|\S+/g) ?? [];
  const rawKind = tokens.shift() ?? '';
  const known = new Set<FieldKind>([
    'SEQ', 'DATE', 'TIME', 'CREATEDATE', 'SAVEDATE', 'PRINTDATE', 'AUTHOR', 'TITLE', 'SUBJECT',
    'KEYWORDS', 'COMMENTS', 'LASTSAVEDBY', 'DOCPROPERTY', 'FILENAME', 'REF', 'PAGE', 'NUMPAGES',
    'PAGEREF', 'TOC', 'INDEX', 'INCLUDETEXT', 'INCLUDEPICTURE', 'LINK', 'DDE', 'DDEAUTO',
    'MACROBUTTON', 'GOTOBUTTON', 'FILLIN', 'ASK', 'DATABASE', 'AUTOTEXT', 'AUTOTEXTLIST',
    'HYPERLINK', 'IF', 'MERGEFIELD', 'FORMTEXT', 'FORMCHECKBOX', 'FORMDROPDOWN', 'EQ',
  ]);
  const kind = (known.has(rawKind.toUpperCase() as FieldKind) ? rawKind.toUpperCase() : 'unknown') as FieldKind;
  const switches: FieldSwitch[] = [];
  let argument: string | undefined;
  // EQ 的括号语法不是开关语法，下面这套按空白分词的读法会读出垃圾（开关名里带着半个括号，
  // 还会凭空造出一个 argument），所以它走 parseEquationInstruction，switches 留空。
  if (kind === 'EQ') return { kind, switches, equation: parseEquationInstruction(instruction) };
  for (let index = 0; index < tokens.length; index++) {
    const token = tokens[index]!;
    if (token.startsWith('\\')) {
      const name = token.slice(1);
      const next = tokens[index + 1];
      if (next && !next.startsWith('\\')) {
        switches.push({ name, value: next.replace(/^"|"$/g, '') });
        index++;
      } else switches.push({ name });
    } else if (argument === undefined) argument = token.replace(/^"|"$/g, '');
  }
  return { kind, argument, switches };
}

function fieldFlags(element: Element): { locked: boolean; dirty: boolean } {
  const value = (name: string) => attr(element, name);
  return {
    locked: ['1', 'true', 'on'].includes(value('fldLock') ?? ''),
    dirty: ['1', 'true', 'on'].includes(value('dirty') ?? ''),
  };
}

function parseFormField(begin: Element, kind: FieldKind): FieldInfo['formField'] {
  const data = children(begin, 'ffData')[0];
  if (!data || !['FORMTEXT', 'FORMCHECKBOX', 'FORMDROPDOWN'].includes(kind)) return undefined;
  const value = (parent: Element, name: string): string | undefined => {
    const node = children(parent, name)[0];
    return node ? attr(node, 'val') ?? node.textContent ?? '' : undefined;
  };
  const onOff = (parent: Element, name: string): boolean | undefined => {
    const node = children(parent, name)[0];
    if (!node) return undefined;
    const raw = attr(node, 'val');
    return raw === undefined || !['0', 'false', 'off'].includes(raw.toLowerCase());
  };
  const integer = (parent: Element, name: string): number | undefined => {
    const raw = value(parent, name);
    if (raw === undefined || raw === '') return undefined;
    const parsed = Number(raw);
    return Number.isSafeInteger(parsed) ? parsed : undefined;
  };
  const formField: NonNullable<FieldInfo['formField']> = {
    kind: kind === 'FORMTEXT' ? 'text' : kind === 'FORMCHECKBOX' ? 'checkBox' : 'dropDown',
    ...(value(data, 'name') !== undefined ? { name: value(data, 'name') } : {}),
    ...(onOff(data, 'enabled') !== undefined ? { enabled: onOff(data, 'enabled') } : {}),
    ...(value(data, 'helpText') !== undefined ? { helpText: value(data, 'helpText') } : {}),
    ...(value(data, 'statusText') !== undefined ? { statusText: value(data, 'statusText') } : {}),
    ...(value(data, 'entryMacro') !== undefined ? { entryMacro: value(data, 'entryMacro') } : {}),
    ...(value(data, 'exitMacro') !== undefined ? { exitMacro: value(data, 'exitMacro') } : {}),
  };
  if (kind === 'FORMTEXT') {
    const input = children(data, 'textInput')[0];
    if (input) {
      const defaultValue = value(input, 'default');
      const maxLength = integer(input, 'maxLength');
      const format = value(input, 'format');
      const type = value(input, 'type');
      formField.text = {
        ...(defaultValue !== undefined ? { default: defaultValue } : {}),
        ...(maxLength !== undefined ? { maxLength } : {}),
        ...(format !== undefined ? { format } : {}),
        ...(type !== undefined ? { type } : {}),
      };
    }
  } else if (kind === 'FORMCHECKBOX') {
    const checkBox = children(data, 'checkBox')[0];
    if (checkBox) {
      const defaultValue = onOff(checkBox, 'default');
      const checked = onOff(checkBox, 'checked');
      const sizeAuto = onOff(checkBox, 'sizeAuto');
      const size = integer(checkBox, 'size');
      formField.checkBox = {
        ...(defaultValue !== undefined ? { default: defaultValue } : {}),
        ...(checked !== undefined ? { checked } : {}),
        ...(sizeAuto !== undefined ? { sizeAuto } : {}),
        ...(size !== undefined ? { sizePt: size / 2 } : {}),
      };
    }
  } else {
    const dropDown = children(data, 'ddList')[0];
    if (dropDown) {
      const entries = children(dropDown, 'listEntry')
        .map((entry) => attr(entry, 'val'))
        .filter((entry): entry is string => entry !== undefined);
      // w:default 与 w:result 是 w:ddList 的子元素（<w:result w:val="1"/>），不是它的属性，
      // 所以要用读子元素 w:val 的 integer()，和上面 checkBox 的 size 一样。
      const defaultValue = integer(dropDown, 'default');
      const result = integer(dropDown, 'result');
      formField.dropDown = {
        entries,
        ...(defaultValue !== undefined ? { default: defaultValue } : {}),
        ...(result !== undefined ? { result } : {}),
      };
    }
  }
  return formField;
}

function belongsToParagraph(field: Element, paragraph: Element): boolean {
  let parent = field.parentNode;
  for (; parent && parent !== paragraph; parent = parent.parentNode) {
    if (parent.nodeType !== 1) continue;
    const element = parent as Element;
    if (element.namespaceURI === WORD_NS && ['p', 'txbxContent'].includes(element.localName ?? '')) return false;
  }
  return parent === paragraph;
}

export function parseFields(paragraphs: Element[], ownRuns: (paragraph: Element) => Element[]): ParsedFields {
  const fields: FieldInfo[] = [];
  const roles = new Map<Element, { index: number; role: 'instruction' | 'result'; kind?: FieldKind; instruction?: string }>();
  const add = (paragraph: Element, paragraphIndex: number, form: 'simple' | 'complex', instruction: string,
    runElements: Element[], resultRuns: Element[], flags: { locked: boolean; dirty: boolean },
    nestedIn?: number, formDataBegin?: Element): void => {
    const parsed = parseInstruction(instruction);
    const index = fields.length;
    const paragraphRuns = ownRuns(paragraph);
    const runIndexes = runElements.map(run => paragraphRuns.indexOf(run)).filter(run => run >= 0).sort((a, b) => a - b);
    const resultIndexes = resultRuns.map(run => paragraphRuns.indexOf(run)).filter(run => run >= 0).sort((a, b) => a - b);
    const result = resultRuns.map(visibleText).join('');
    const formField = formDataBegin ? parseFormField(formDataBegin, parsed.kind) : undefined;
    for (const run of runElements) {
      if (!roles.has(run)) roles.set(run, { index, role: resultRuns.includes(run) ? 'result' : 'instruction', kind: parsed.kind, instruction });
    }
    fields.push({
      index, paragraph: paragraphIndex, runs: [...new Set(runIndexes)], resultRuns: [...new Set(resultIndexes)], form,
      kind: parsed.kind, instruction, ...(parsed.argument !== undefined ? { argument: parsed.argument } : {}),
      ...(parsed.kind === 'MERGEFIELD' && parsed.argument !== undefined ? { mergeFieldName: parsed.argument } : {}),
      ...(formField ? { formField } : {}),
      ...(parsed.equation ? { equation: parsed.equation } : {}),
      switches: parsed.switches, result, requiresPagination: PAGINATION.has(parsed.kind),
      evaluable: !PAGINATION.has(parsed.kind) && !NEVER_EVALUATE.has(parsed.kind) &&
        ['SEQ', 'DATE', 'TIME', 'CREATEDATE', 'SAVEDATE', 'PRINTDATE', 'AUTHOR', 'TITLE', 'SUBJECT',
          'KEYWORDS', 'COMMENTS', 'LASTSAVEDBY', 'DOCPROPERTY', 'FILENAME', 'REF'].includes(parsed.kind),
      locked: flags.locked, dirty: flags.dirty, ...(nestedIn === undefined ? {} : { nestedIn }),
    });
  };
  for (const [paragraphIndex, paragraph] of paragraphs.entries()) {
    const runs = ownRuns(paragraph);
    for (const simple of descendants(paragraph, 'fldSimple').filter(field => belongsToParagraph(field, paragraph))) {
      const instruction = attr(simple, 'instr') ?? '';
      const resultRuns = runs.filter(run => {
        for (let parent = run.parentNode; parent && parent !== paragraph; parent = parent.parentNode) {
          if (parent === simple) return true;
        }
        return false;
      });
      add(paragraph, paragraphIndex, 'simple', instruction, resultRuns, resultRuns, fieldFlags(simple));
    }
    const stack: { begin: Element; instr: Element[]; runs: Element[]; resultRuns: Element[]; separated: boolean; nestedIn?: number }[] = [];
    for (const run of runs) {
      const chars = children(run, 'fldChar');
      const instr = children(run, 'instrText');
      if (instr.length && stack.length) {
        stack.at(-1)!.instr.push(...instr);
        stack.at(-1)!.runs.push(run);
      }
      for (const marker of chars) {
        const type = attr(marker, 'fldCharType');
        if (type === 'begin') {
          stack.push({ begin: marker, instr: [], runs: [run], resultRuns: [], separated: false });
        } else if (type === 'separate' && stack.length) {
          const current = stack.at(-1)!;
          current.separated = true;
          current.runs.push(run);
        } else if (type === 'end' && stack.length) {
          const current = stack.pop()!;
          current.runs.push(run);
          const instruction = current.instr.map(node => node.textContent ?? '').join('');
          const resultRuns = current.resultRuns;
          add(paragraph, paragraphIndex, 'complex', instruction, [...current.runs, ...resultRuns], resultRuns,
            fieldFlags(marker), current.nestedIn, current.begin);
        }
      }
      if (stack.length && stack.at(-1)!.separated &&
          !chars.some(marker => ['separate', 'end'].includes(attr(marker, 'fldCharType') ?? ''))) {
        stack.at(-1)!.resultRuns.push(run);
      }
    }
  }
  for (const field of fields) {
    const start = field.runs[0] ?? Number.MAX_SAFE_INTEGER;
    const end = field.runs.at(-1) ?? -1;
    const parent = fields
      .filter(candidate => candidate.index !== field.index &&
        candidate.paragraph === field.paragraph &&
        (candidate.runs[0] ?? 0) < start && (candidate.runs.at(-1) ?? -1) > end)
      .sort((a, b) => (a.runs.length - b.runs.length))[0];
    if (parent) field.nestedIn = parent.index;
  }
  return { fields, roles };
}

export function fieldKindFromInstruction(instruction: string): FieldKind {
  return parseInstruction(instruction).kind;
}

export { NEVER_EVALUATE };
