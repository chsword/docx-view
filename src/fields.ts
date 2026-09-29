import type { Element, Node } from '@xmldom/xmldom';
import type { FieldInfo, FieldKind, FieldSwitch } from './types.js';
import { WORD_NS, children, descendants } from './xml.js';

const PAGINATION = new Set<FieldKind>(['PAGE', 'NUMPAGES', 'PAGEREF', 'TOC', 'INDEX']);
const NEVER_EVALUATE = new Set<FieldKind>([
  'INCLUDETEXT', 'INCLUDEPICTURE', 'LINK', 'DDE', 'DDEAUTO', 'MACROBUTTON', 'GOTOBUTTON',
  'FILLIN', 'ASK', 'DATABASE', 'AUTOTEXT', 'AUTOTEXTLIST',
]);

export interface ParsedFields {
  fields: FieldInfo[];
  roles: Map<Element, { index: number; role: 'instruction' | 'result' }>;
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

function parseInstruction(instruction: string): { kind: FieldKind; argument?: string; switches: FieldSwitch[] } {
  const tokens = instruction.match(/"[^"]*"|\S+/g) ?? [];
  const rawKind = tokens.shift() ?? '';
  const known = new Set<FieldKind>([
    'SEQ', 'DATE', 'TIME', 'CREATEDATE', 'SAVEDATE', 'PRINTDATE', 'AUTHOR', 'TITLE', 'SUBJECT',
    'KEYWORDS', 'COMMENTS', 'LASTSAVEDBY', 'DOCPROPERTY', 'FILENAME', 'REF', 'PAGE', 'NUMPAGES',
    'PAGEREF', 'TOC', 'INDEX', 'INCLUDETEXT', 'INCLUDEPICTURE', 'LINK', 'DDE', 'DDEAUTO',
    'MACROBUTTON', 'GOTOBUTTON', 'FILLIN', 'ASK', 'DATABASE', 'AUTOTEXT', 'AUTOTEXTLIST',
    'HYPERLINK', 'IF',
  ]);
  const kind = (known.has(rawKind.toUpperCase() as FieldKind) ? rawKind.toUpperCase() : 'unknown') as FieldKind;
  const switches: FieldSwitch[] = [];
  let argument: string | undefined;
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
  const roles = new Map<Element, { index: number; role: 'instruction' | 'result' }>();
  const add = (paragraph: Element, paragraphIndex: number, form: 'simple' | 'complex', instruction: string,
    runElements: Element[], resultRuns: Element[], flags: { locked: boolean; dirty: boolean }, nestedIn?: number): void => {
    const parsed = parseInstruction(instruction);
    const index = fields.length;
    const paragraphRuns = ownRuns(paragraph);
    const runIndexes = runElements.map(run => paragraphRuns.indexOf(run)).filter(run => run >= 0).sort((a, b) => a - b);
    const resultIndexes = resultRuns.map(run => paragraphRuns.indexOf(run)).filter(run => run >= 0).sort((a, b) => a - b);
    const result = resultRuns.map(visibleText).join('');
    for (const run of runElements) {
      if (!roles.has(run)) roles.set(run, { index, role: resultRuns.includes(run) ? 'result' : 'instruction' });
    }
    fields.push({
      index, paragraph: paragraphIndex, runs: [...new Set(runIndexes)], resultRuns: [...new Set(resultIndexes)], form,
      kind: parsed.kind, instruction, ...(parsed.argument !== undefined ? { argument: parsed.argument } : {}),
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
          add(paragraph, paragraphIndex, 'complex', instruction, [...current.runs, ...resultRuns], resultRuns, fieldFlags(marker), current.nestedIn);
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
