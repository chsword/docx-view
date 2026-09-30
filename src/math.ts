import type { Element } from '@xmldom/xmldom';
import { assertText, type MathMlNode } from './types.js';

export const MATH_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
const MAX_DEPTH = 64;
export interface MathConversion {
  node: MathMlNode;
  truncated: boolean;
}
export interface LinearConversion {
  text: string;
  truncated: boolean;
}

const elementChildren = (element: Element): Element[] => Array.from(element.childNodes)
  .filter((node): node is Element => node.nodeType === 1) as Element[];
const mathChildren = (element: Element, name?: string): Element[] =>
  elementChildren(element).filter(child => child.namespaceURI === MATH_NS && (!name || child.localName === name));
const textOf = (element: Element | undefined): string =>
  element ? (element.localName === 't' ? element.textContent ?? '' :
    elementChildren(element).filter(child => child.localName === 't').map(child => child.textContent ?? '').join('')) : '';
const first = (element: Element, name: string): Element | undefined => mathChildren(element, name)[0];
const value = (element: Element | undefined, name: string, fallback: string): string =>
  element?.getAttributeNS(MATH_NS, name) ?? element?.getAttribute(`m:${name}`) ?? fallback;
const prop = (parent: Element | undefined, name: string, fallback: string): string => {
  if (!parent) return fallback;
  const child = first(parent, name);
  return value(child, 'val', textOf(child) || fallback);
};
const directText = (element: Element | undefined, fallback: string): string =>
  value(element, 'val', textOf(element) || fallback);

function leaf(text: string): MathMlNode[] {
  const nodes: MathMlNode[] = [];
  let index = 0;
  const operator = (char: string): boolean => /[\p{Sm}\p{So}]/u.test(char) ||
    '≠∈∉⊂⊆∪∩∀∃→←↔⇒⇔∞∂∇'.includes(char) ||
    /^[+\-−×÷=<>≤≥±*/(),.;:|]$/.test(char);
  while (index < text.length) {
    const char = text[index]!;
    if (/\d/.test(char)) {
      const start = index++;
      while (index < text.length && /[\d.,]/.test(text[index]!)) index++;
      nodes.push({ tag: 'mn', text: text.slice(start, index) });
    } else if (/\s/.test(char)) {
      index++;
    } else if (operator(char)) {
      nodes.push({ tag: 'mo', text: char });
      index++;
    } else {
      const start = index++;
      while (index < text.length && !/[\d\s]/.test(text[index]!) && !operator(text[index]!)) index++;
      nodes.push({ tag: 'mi', text: text.slice(start, index) });
    }
  }
  return nodes;
}

function childrenAsRow(element: Element, depth: number, context: { maxDepth: number; truncated: boolean }): MathMlNode {
  const children = elementChildren(element).filter(child => child.namespaceURI === MATH_NS)
    .flatMap(child => convert(child, depth + 1, context));
  return children.length === 1 ? children[0]! : { tag: 'mrow', children };
}

function convert(element: Element, depth: number, context: { maxDepth: number; truncated: boolean }): MathMlNode[] {
  if (depth > context.maxDepth) {
    context.truncated = true;
    return [{ tag: 'mtext', text: '…' }];
  }
  const name = element.localName;
  if (name === 't') return leaf(element.textContent ?? '');
  if (name === 'r') {
    const text = textOf(first(element, 't'));
    if (!text) return [];
    const nodes = leaf(text);
    const sty = first(first(element, 'rPr') ?? element, 'sty');
    const style = sty ? directText(sty, '') : '';
    const scr = first(first(element, 'rPr') ?? element, 'scr');
    const script = scr ? directText(scr, '') : '';
    const variants: Record<string, string> = {
      p: 'normal', b: 'bold', i: 'italic', bi: 'bold-italic',
      'double-struck': 'double-struck', script: 'script', fraktur: 'fraktur',
      monospace: 'monospace', 'sans-serif': 'sans-serif',
    };
    const variant = variants[script] ?? variants[style];
    for (const node of nodes) {
      if (variant) node.attrs = { mathvariant: variant };
      else if (node.tag === 'mi') node.attrs = { mathvariant: 'italic' };
    }
    return nodes;
  }
  if (name === 'oMath' || name === 'oMathPara') {
    const content = mathChildren(element).flatMap(child =>
      name === 'oMathPara' && child.localName === 'oMath'
        ? mathChildren(child).flatMap(nested => convert(nested, depth + 1, context))
        : convert(child, depth + 1, context));
    const attrs: Record<string, string> = { display: name === 'oMathPara' ? 'block' : 'inline' };
    const alignment = name === 'oMathPara' ? directText(first(element, 'jc'), '') : '';
    if (alignment) attrs.style = `text-align:${alignment === 'centerGroup' ? 'center' : alignment}`;
    return [{ tag: 'math', attrs, children: content }];
  }
  if (name === 'e' || name === 'num' || name === 'den' || name === 'sub' || name === 'sup' || name === 'deg' ||
      name === 'fName' || name === 'fName' || name === 'lim' || name === 'mr') {
    return [childrenAsRow(element, depth, context)];
  }
  if (name === 'f') {
    const num = first(element, 'num');
    const den = first(element, 'den');
    const numerator = num ? childrenAsRow(num, depth, context) : { tag: 'mrow', children: [] };
    const denominator = den ? childrenAsRow(den, depth, context) : { tag: 'mrow', children: [] };
    const fractionType = prop(first(element, 'fPr'), 'type', '');
    if (fractionType === 'lin') {
      return [{ tag: 'mrow', children: [numerator, { tag: 'mo', text: '/' }, denominator] }];
    }
    const attrs = fractionType === 'noBar' ? { linethickness: '0' } : undefined;
    return [{ tag: 'mfrac', ...(attrs ? { attrs } : {}), children: [numerator, denominator] }];
  }
  if (name === 'sSup' || name === 'sSub' || name === 'sSubSup') {
    const baseElement = first(element, 'e');
    const base = baseElement ? childrenAsRow(baseElement, depth, context) : { tag: 'mrow', children: [] };
    const sub = first(element, 'sub');
    const sup = first(element, 'sup');
    if (name === 'sSup') return [{ tag: 'msup', children: [base, sup ? childrenAsRow(sup, depth, context) : { tag: 'mrow', children: [] }] }];
    if (name === 'sSub') return [{ tag: 'msub', children: [base, sub ? childrenAsRow(sub, depth, context) : { tag: 'mrow', children: [] }] }];
    return [{ tag: 'msubsup', children: [base, sub ? childrenAsRow(sub, depth, context) : { tag: 'mrow', children: [] }, sup ? childrenAsRow(sup, depth, context) : { tag: 'mrow', children: [] }] }];
  }
  if (name === 'sPre') {
    const base = first(element, 'e');
    const sub = first(element, 'sub');
    const sup = first(element, 'sup');
    return [{ tag: 'mmultiscripts', children: [
      base ? childrenAsRow(base, depth, context) : { tag: 'mrow', children: [] },
      { tag: 'mprescripts' },
      sub ? childrenAsRow(sub, depth, context) : { tag: 'mrow', children: [] },
      sup ? childrenAsRow(sup, depth, context) : { tag: 'mrow', children: [] },
    ] }];
  }
  if (name === 'rad') {
    const base = first(element, 'e');
    const degree = first(element, 'deg');
    return [degree ? { tag: 'mroot', children: [base ? childrenAsRow(base, depth, context) : { tag: 'mrow', children: [] }, childrenAsRow(degree, depth, context)] }
      : { tag: 'msqrt', children: [base ? childrenAsRow(base, depth, context) : { tag: 'mrow', children: [] }] }];
  }
  if (name === 'nary') {
    const props = first(element, 'naryPr');
    const symbol = directText(first(props ?? element, 'chr'), '∑');
    const op: MathMlNode = { tag: 'mo', text: symbol };
    const base = first(element, 'e');
    const sub = first(element, 'sub');
    const sup = first(element, 'sup');
    const body = base ? childrenAsRow(base, depth, context) : { tag: 'mrow', children: [] };
    const loc = prop(props ?? element, 'limLoc', 'undOvr');
    const operator = loc === 'subSup'
      ? { tag: 'msubsup', children: [op, sub ? childrenAsRow(sub, depth, context) : { tag: 'mrow', children: [] }, sup ? childrenAsRow(sup, depth, context) : { tag: 'mrow', children: [] }] }
      : { tag: 'munderover', children: [op, sub ? childrenAsRow(sub, depth, context) : { tag: 'mrow', children: [] }, sup ? childrenAsRow(sup, depth, context) : { tag: 'mrow', children: [] }] };
    return [{ tag: 'mrow', children: [operator, body] }];
  }
  if (name === 'd') {
    const props = first(element, 'dPr');
    const beg = directText(first(props ?? element, 'begChr'), '(');
    const end = directText(first(props ?? element, 'endChr'), ')');
    const sep = directText(first(props ?? element, 'sepChr'), ',');
    const content = mathChildren(element, 'e').flatMap((entry, index) => [
      ...(index ? [{ tag: 'mo', attrs: { stretchy: 'true' }, text: sep }] : []),
      childrenAsRow(entry, depth, context),
    ]);
    return [{ tag: 'mrow', children: [{ tag: 'mo', attrs: { stretchy: 'true' }, text: beg }, ...content, { tag: 'mo', attrs: { stretchy: 'true' }, text: end }] }];
  }
  if (name === 'func') {
    const fname = first(element, 'fName');
    const expr = first(element, 'e');
    return [{ tag: 'mrow', children: [fname ? childrenAsRow(fname, depth, context) : { tag: 'mi', text: '' }, expr ? childrenAsRow(expr, depth, context) : { tag: 'mrow', children: [] }] }];
  }
  if (name === 'limLow' || name === 'limUpp') {
    const base = first(element, 'e');
    const lim = first(element, 'lim');
    return [{ tag: name === 'limLow' ? 'munder' : 'mover', children: [base ? childrenAsRow(base, depth, context) : { tag: 'mrow', children: [] }, lim ? childrenAsRow(lim, depth, context) : { tag: 'mrow', children: [] }] }];
  }
  if (name === 'acc' || name === 'bar' || name === 'groupChr') {
    const props = first(element, `${name}Pr`);
    const base = first(element, 'e');
    const chr = directText(first(props ?? element, 'chr'), name === 'acc' ? '̂' : name === 'groupChr' ? '⏞' : '¯');
    const pos = prop(props ?? element, 'pos', 'top');
    return [{ tag: name === 'acc' || pos === 'top' ? 'mover' : 'munder', children: [base ? childrenAsRow(base, depth, context) : { tag: 'mrow', children: [] }, { tag: 'mo', ...(name === 'acc' ? { attrs: { accent: 'true' } } : {}), text: chr }] }];
  }
  if (name === 'm' || name === 'eqArr') {
    const rows = name === 'm' ? mathChildren(element, 'mr') : mathChildren(element, 'e');
    return [{ tag: 'mtable', children: rows.map(row => ({ tag: 'mtr', children: [({ tag: 'mtd', children: [childrenAsRow(row, depth, context)] })] })) }];
  }
  if (name === 'borderBox') return [{ tag: 'menclose', children: [childrenAsRow(element, depth, context)] }];
  if (name === 'phant') return [{ tag: 'mphantom', children: [childrenAsRow(element, depth, context)] }];
  if (name === 'box') return [childrenAsRow(element, depth, context)];
  return [{ tag: 'mrow', children: mathChildren(element).flatMap(child => convert(child, depth + 1, context)) }];
}

export function ommlToMathMlWithInfo(oMath: Element, options?: { maxDepth?: number }): MathConversion {
  const context = { maxDepth: options?.maxDepth !== undefined && Number.isSafeInteger(options.maxDepth) && options.maxDepth >= 0
    ? options.maxDepth : MAX_DEPTH, truncated: false };
  return { node: convert(oMath, 0, context)[0] ?? { tag: 'mtext', text: '…' }, truncated: context.truncated };
}

export function ommlToMathMl(oMath: Element, options?: { maxDepth?: number }): MathMlNode {
  return ommlToMathMlWithInfo(oMath, options).node;
}

function linear(node: MathMlNode): string {
  if (node.text !== undefined) return node.text;
  const c = (node.children ?? []).map(linear);
  switch (node.tag) {
    case 'mfrac': return `${node.children?.[0]?.tag === 'mrow' ? `(${c[0] ?? ''})` : c[0] ?? ''}/${c[1] ?? ''}`;
    case 'msup': return `${c[0] ?? ''}^${c[1] ?? ''}`;
    case 'msub': return `${c[0] ?? ''}_${c[1] ?? ''}`;
    case 'msubsup': return `${c[0] ?? ''}_${c[1] ?? ''}^${c[2] ?? ''}`;
    case 'mmultiscripts': {
      const pre = `${c[2] ? `_${c[2]}` : ''}${c[3] ? `^${c[3]}` : ''}`;
      return `${pre}${c[0] ?? ''}`;
    }
    case 'msqrt': return `√(${c[0] ?? ''})`;
    case 'mroot': return `${c[1] ?? ''}√(${c[0] ?? ''})`;
    case 'munderover': return `${c[0] ?? ''}_(${c[1] ?? ''})^(${c[2] ?? ''})`;
    case 'mrow': return c.join('');
    case 'math': return c.join('');
    default: return c.join('');
  }
}

export function ommlToLinearText(oMath: Element, options?: { maxDepth?: number }): string {
  return linear(ommlToMathMl(oMath, options));
}

export function ommlToLinearTextWithInfo(oMath: Element, options?: { maxDepth?: number }): LinearConversion {
  const result = ommlToMathMlWithInfo(oMath, options);
  return { text: linear(result.node), truncated: result.truncated };
}

const MATHML_TAGS = new Set([
  'math', 'mrow', 'mi', 'mn', 'mo', 'mfrac', 'msqrt', 'mroot', 'msup', 'msub', 'msubsup',
  'munder', 'mover', 'munderover', 'mtable', 'mtr', 'mtd', 'menclose', 'mphantom', 'mtext',
  'mmultiscripts', 'mprescripts',
]);
const MATHML_ATTRS = new Set(['mathvariant', 'display', 'linethickness', 'stretchy', 'accent', 'encoding']);
const LINEAR_SUPPORT = 'a/b, a^b, a_b, a_b^c, √(a), sqrt(a), (…), and ∑_(a)^(b) c / ∫_(a)^(b) c';
const MATH_VARIANTS: Record<string, string> = {
  normal: 'p',
  bold: 'b',
  italic: 'i',
  'bold-italic': 'bi',
  'double-struck': 'double-struck',
  script: 'script',
  fraktur: 'fraktur',
  monospace: 'monospace',
  'sans-serif': 'sans-serif',
};

interface LinearParser {
  source: string;
  index: number;
  maxDepth: number;
}

function parsedRow(children: MathMlNode[]): MathMlNode {
  return children.length === 1 ? children[0]! : { tag: 'mrow', children };
}

function mathToken(source: string, index: number): { node: MathMlNode; end: number } | undefined {
  const char = source[index]!;
  if (/\d/.test(char)) {
    let end = index + 1;
    while (end < source.length && /[\d.,]/.test(source[end]!)) end++;
    return { node: { tag: 'mn', text: source.slice(index, end) }, end };
  }
  if (/\s/.test(char)) return { node: { tag: 'mtext', text: char }, end: index + 1 };
  if (char === ')' || char === '/' || char === '^' || char === '_') return undefined;
  if (/[\p{Sm}\p{So}]/u.test(char) || '≠∈∉⊂⊆∪∩∀∃→←↔⇒⇔∞∂∇+-−×÷=<>≤≥±*,.;:|'.includes(char)) {
    return { node: { tag: 'mo', text: char }, end: index + 1 };
  }
  if (/[\\{}[\]]/.test(char)) return undefined;
  let end = index + 1;
  while (end < source.length && !/[\d\s]/.test(source[end]!) &&
      !(/[\p{Sm}\p{So}]/u.test(source[end]!) || '≠∈∉⊂⊆∪∩∀∃→←↔⇒⇔∞∂∇+-−×÷=<>≤≥±*/,;:|()_^√'.includes(source[end]!))) {
    if (/[\\{}[\]]/.test(source[end]!)) break;
    end++;
  }
  return { node: { tag: 'mi', attrs: { mathvariant: 'italic' }, text: source.slice(index, end) }, end };
}

function parseLinearExpression(parser: LinearParser, depth: number, stopAtParen = false): MathMlNode[] {
  if (depth > parser.maxDepth) throw new Error(`Linear math exceeds the maximum depth of ${parser.maxDepth}.`);
  const nodes: MathMlNode[] = [];
  while (parser.index < parser.source.length) {
    while (/\s/.test(parser.source[parser.index] ?? '')) parser.index++;
    if (parser.index >= parser.source.length || (stopAtParen && parser.source[parser.index] === ')')) break;

    const char = parser.source[parser.index]!;
    if (char === '(') {
      parser.index++;
      const inside = parseLinearExpression(parser, depth + 1, true);
      if (parser.source[parser.index] !== ')') throw new Error('Unclosed group in linear math.');
      parser.index++;
      nodes.push({ tag: 'mrow', children: [
        { tag: 'mo', text: '(' }, ...inside, { tag: 'mo', text: ')' },
      ] });
      continue;
    }
    if (char === '√' || parser.source.startsWith('sqrt(', parser.index)) {
      if (char === '√') parser.index++;
      else parser.index += 4;
      if (parser.source[parser.index] !== '(') throw new Error('Square roots must use √(a) or sqrt(a).');
      parser.index++;
      const inside = parseLinearExpression(parser, depth + 1, true);
      if (parser.source[parser.index] !== ')') throw new Error('Unclosed square root in linear math.');
      parser.index++;
      nodes.push({ tag: 'msqrt', children: [parsedRow(inside)] });
      continue;
    }
    if (char === '∑' || char === '∫') {
      const operator = char;
      parser.index++;
      if (parser.source[parser.index] !== '_' || parser.source[parser.index + 1] !== '(') {
        throw new Error('Summation and integral forms must use ∑_(a)^(b) c or ∫_(a)^(b) c.');
      }
      parser.index += 2;
      const sub = parseLinearExpression(parser, depth + 1, true);
      if (parser.source[parser.index] !== ')') throw new Error('Unclosed lower limit in linear math.');
      parser.index++;
      if (parser.source[parser.index] !== '^' || parser.source[parser.index + 1] !== '(') {
        throw new Error('Summation and integral forms must use ∑_(a)^(b) c or ∫_(a)^(b) c.');
      }
      parser.index += 2;
      const sup = parseLinearExpression(parser, depth + 1, true);
      if (parser.source[parser.index] !== ')') throw new Error('Unclosed upper limit in linear math.');
      parser.index++;
      const body = parseLinearFactor(parser, depth + 1);
      if (!body) throw new Error('Summation and integral forms must use ∑_(a)^(b) c or ∫_(a)^(b) c.');
      nodes.push({ tag: 'mrow', children: [
        { tag: 'munderover', children: [
          { tag: 'mo', text: operator }, parsedRow(sub), parsedRow(sup),
        ] },
        body,
      ] });
      continue;
    }
    const factor = parseLinearFactor(parser, depth + 1);
    if (!factor) throw new Error(`Unsupported linear math. Supported forms: ${LINEAR_SUPPORT}.`);
    nodes.push(factor);
  }
  return nodes;
}

function parseLinearFactor(parser: LinearParser, depth: number, allowScripts = true): MathMlNode | undefined {
  if (depth > parser.maxDepth) throw new Error(`Linear math exceeds the maximum depth of ${parser.maxDepth}.`);
  while (/\s/.test(parser.source[parser.index] ?? '')) parser.index++;
  let base: MathMlNode | undefined;
  if (parser.source[parser.index] === '(') {
    parser.index++;
    const inside = parseLinearExpression(parser, depth + 1, true);
    if (parser.source[parser.index] !== ')') throw new Error('Unclosed group in linear math.');
    parser.index++;
    base = { tag: 'mrow', children: [{ tag: 'mo', text: '(' }, ...inside, { tag: 'mo', text: ')' }] };
  } else if (parser.source[parser.index] === '√' || parser.source.startsWith('sqrt(', parser.index)) {
    if (parser.source[parser.index] === '√') parser.index++;
    else parser.index += 4;
    if (parser.source[parser.index] !== '(') throw new Error('Square roots must use √(a) or sqrt(a).');
    parser.index++;
    const inside = parseLinearExpression(parser, depth + 1, true);
    if (parser.source[parser.index] !== ')') throw new Error('Unclosed square root in linear math.');
    parser.index++;
    base = { tag: 'msqrt', children: [parsedRow(inside)] };
  } else {
    const token = mathToken(parser.source, parser.index);
    if (!token) return undefined;
    base = token.node;
    parser.index = token.end;
  }

  if (parser.source[parser.index] === '/') {
    parser.index++;
    const denominator = parseLinearFactor(parser, depth + 1);
    if (!denominator) throw new Error('Fractions must use a/b.');
    base = { tag: 'mfrac', children: [base, denominator] };
  }
  if (allowScripts && (parser.source[parser.index] === '_' || parser.source[parser.index] === '^')) {
    let sub: MathMlNode | undefined;
    let sup: MathMlNode | undefined;
    for (let count = 0; count < 2; count++) {
      const marker = parser.source[parser.index];
      if (marker !== '_' && marker !== '^') break;
      parser.index++;
      const script = parseLinearFactor(parser, depth + 1, false);
      if (!script) throw new Error('Scripts must use a_b, a^b, or a_b^c.');
      if (marker === '_') {
        if (sub) throw new Error('Only one subscript is supported per base.');
        sub = script;
      } else {
        if (sup) throw new Error('Only one superscript is supported per base.');
        sup = script;
      }
    }
    base = sub && sup ? { tag: 'msubsup', children: [base, sub, sup] }
      : sub ? { tag: 'msub', children: [base, sub] }
        : { tag: 'msup', children: [base, sup!] };
  }
  return base;
}

export function linearToMathMl(source: string, options?: { maxDepth?: number }): MathMlNode {
  assertText(source, 'linear math');
  const maxDepth = options?.maxDepth ?? MAX_DEPTH;
  if (!Number.isSafeInteger(maxDepth) || maxDepth < 0) throw new Error('maxDepth must be a non-negative safe integer.');
  const parser: LinearParser = { source, index: 0, maxDepth };
  const children = parseLinearExpression(parser, 0);
  while (/\s/.test(parser.source[parser.index] ?? '')) parser.index++;
  if (!children.length || parser.index !== source.length) {
    throw new Error(`Unsupported linear math. Supported forms: ${LINEAR_SUPPORT}.`);
  }
  return { tag: 'math', attrs: { display: 'inline' }, children };
}

function createOmmlElement(doc: import('@xmldom/xmldom').Document, name: string): Element {
  return doc.createElementNS(MATH_NS, `m:${name}`);
}

function setOmmlValue(element: Element, value: string): void {
  element.setAttributeNS(MATH_NS, 'm:val', value);
}

function appendOmmlTextRun(parent: Element, doc: import('@xmldom/xmldom').Document, text: string, variant?: string): void {
  assertText(text, 'MathML text');
  const run = createOmmlElement(doc, 'r');
  if (variant && !(variant === 'italic')) {
    const mapped = MATH_VARIANTS[variant];
    if (!mapped) throw new Error(`Unsupported MathML mathvariant: ${variant}.`);
    const properties = createOmmlElement(doc, 'rPr');
    const style = createOmmlElement(doc, 'sty');
    setOmmlValue(style, mapped);
    properties.appendChild(style);
    run.appendChild(properties);
  }
  const content = createOmmlElement(doc, 't');
  if (/^\s|\s$/u.test(text)) content.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
  content.appendChild(doc.createTextNode(text));
  run.appendChild(content);
  parent.appendChild(run);
}

function elementText(node: MathMlNode): string {
  if (node.text !== undefined) {
    assertText(node.text, 'MathML text');
    return node.text;
  }
  if (node.children?.length) throw new Error(`${node.tag} must contain text, not child elements.`);
  return '';
}

function mathMlChildren(node: MathMlNode): MathMlNode[] {
  if (node.children !== undefined && !Array.isArray(node.children)) throw new Error(`${node.tag}.children must be an array.`);
  const children = node.children ?? [];
  if (node.text !== undefined && children.length) throw new Error(`${node.tag} cannot contain both text and child elements.`);
  return children;
}

function mathMlToOmmlNode(node: MathMlNode, doc: import('@xmldom/xmldom').Document, depth: number, maxDepth: number): Element {
  if (!node || typeof node !== 'object' || Array.isArray(node)) throw new Error('MathML nodes must be objects.');
  if (typeof node.tag !== 'string' || !MATHML_TAGS.has(node.tag)) throw new Error(`Unsupported MathML tag: ${String(node.tag)}.`);
  if (depth > maxDepth) throw new Error(`MathML exceeds the maximum depth of ${maxDepth}.`);
  if (node.attrs !== undefined) {
    if (!node.attrs || typeof node.attrs !== 'object' || Array.isArray(node.attrs)) throw new Error(`${node.tag}.attrs must be an object.`);
    for (const [name, value] of Object.entries(node.attrs)) {
      if (!MATHML_ATTRS.has(name)) throw new Error(`Unsupported MathML attribute: ${name}.`);
      assertText(value, `MathML ${name}`);
    }
  }
  const attrs = node.attrs ?? {};
  const children = mathMlChildren(node);
  if (node.text !== undefined && !['mi', 'mn', 'mo', 'mtext'].includes(node.tag)) {
    throw new Error(`${node.tag} cannot contain text.`);
  }
  if (node.tag === 'math' && attrs.display !== undefined && attrs.display !== 'inline' && attrs.display !== 'block') {
    throw new Error('MathML display must be inline or block.');
  }
  if (node.tag === 'math') {
    const root = createOmmlElement(doc, 'oMath');
    for (const child of children) root.appendChild(mathMlToOmmlNode(child, doc, depth + 1, maxDepth));
    return root;
  }
  if (node.tag === 'mi' || node.tag === 'mn' || node.tag === 'mo' || node.tag === 'mtext') {
    if (children.length) throw new Error(`${node.tag} cannot contain child elements.`);
    const text = elementText(node);
    const container = createOmmlElement(doc, 'oMath');
    appendOmmlTextRun(container, doc, text, node.tag === 'mi' && attrs.mathvariant === 'italic'
      ? undefined : attrs.mathvariant);
    return container.firstChild as Element;
  }
  if (node.tag === 'mprescripts') {
    if (children.length || node.text !== undefined) throw new Error('mprescripts must be empty.');
    return createOmmlElement(doc, 'sPre');
  }
  if (node.tag === 'mrow' && children.length >= 2 &&
      children[0]!.tag === 'mo' && children.at(-1)!.tag === 'mo' &&
      children[0]!.attrs?.stretchy === 'true' && children.at(-1)!.attrs?.stretchy === 'true') {
    const delimiters = createOmmlElement(doc, 'd');
    const properties = createOmmlElement(doc, 'dPr');
    const begin = createOmmlElement(doc, 'begChr');
    setOmmlValue(begin, elementText(children[0]!));
    const end = createOmmlElement(doc, 'endChr');
    setOmmlValue(end, elementText(children.at(-1)!));
    properties.appendChild(begin);
    const segments: MathMlNode[][] = [[]];
    let separator: string | undefined;
    for (const child of children.slice(1, -1)) {
      if (child.tag === 'mo' && child.attrs?.stretchy === 'true') {
        const nextSeparator = elementText(child);
        if (separator !== undefined && separator !== nextSeparator) throw new Error('A delimiter group can use only one separator.');
        separator = nextSeparator;
        segments.push([]);
      } else segments[segments.length - 1]!.push(child);
    }
    if (segments.length > 1) {
      const sep = createOmmlElement(doc, 'sepChr');
      setOmmlValue(sep, separator ?? ',');
      properties.appendChild(sep);
    }
    properties.appendChild(end);
    delimiters.appendChild(properties);
    for (const segment of segments) {
      const entry = createOmmlElement(doc, 'e');
      for (const child of segment) entry.appendChild(mathMlToOmmlNode(child, doc, depth + 1, maxDepth));
      delimiters.appendChild(entry);
    }
    return delimiters;
  }
  if (node.tag === 'mrow') {
    const row = createOmmlElement(doc, 'box');
    const entry = createOmmlElement(doc, 'e');
    for (let index = 0; index < children.length; index++) {
      const child = children[index]!;
      if (child.tag === 'munderover' && children[index + 1]) {
        entry.appendChild(buildNary([child, children[++index]!], doc, depth, maxDepth));
      } else entry.appendChild(mathMlToOmmlNode(child, doc, depth + 1, maxDepth));
    }
    row.appendChild(entry);
    return row;
  }
  if (node.tag === 'mfrac') {
    if (children.length !== 2) throw new Error('mfrac must contain a numerator and denominator.');
    const fraction = createOmmlElement(doc, 'f');
    if (attrs.linethickness === '0') {
      const properties = createOmmlElement(doc, 'fPr');
      const type = createOmmlElement(doc, 'type');
      setOmmlValue(type, 'noBar');
      properties.appendChild(type);
      fraction.appendChild(properties);
    } else if (attrs.linethickness !== undefined) throw new Error('Only linethickness="0" is supported.');
    for (const [name, child] of [['num', children[0]!], ['den', children[1]!]] as const) {
      const entry = createOmmlElement(doc, name);
      entry.appendChild(mathMlToOmmlNode(child, doc, depth + 1, maxDepth));
      fraction.appendChild(entry);
    }
    return fraction;
  }
  if (node.tag === 'msup' || node.tag === 'msub' || node.tag === 'msubsup') {
    const expected = node.tag === 'msubsup' ? 3 : 2;
    if (children.length !== expected) throw new Error(`${node.tag} must contain ${expected} child elements.`);
    const script = createOmmlElement(doc, node.tag === 'msup' ? 'sSup' : node.tag === 'msub' ? 'sSub' : 'sSubSup');
    const base = createOmmlElement(doc, 'e');
    base.appendChild(mathMlToOmmlNode(children[0]!, doc, depth + 1, maxDepth));
    script.appendChild(base);
    const scriptChildren = node.tag === 'msup' ? [['sup', children[1]!]] :
      node.tag === 'msub' ? [['sub', children[1]!]] : [['sub', children[1]!], ['sup', children[2]!]];
    for (const [name, child] of scriptChildren as [string, MathMlNode][]) {
      const entry = createOmmlElement(doc, name);
      entry.appendChild(mathMlToOmmlNode(child, doc, depth + 1, maxDepth));
      script.appendChild(entry);
    }
    return script;
  }
  if (node.tag === 'mmultiscripts') {
    const mark = children.findIndex(child => child.tag === 'mprescripts');
    if (children.length < 2 || mark < 1 || children.length - mark !== 3) {
      throw new Error('mmultiscripts must contain a base, mprescripts, and a presubscript/presuperscript pair.');
    }
    const script = createOmmlElement(doc, 'sPre');
    const base = createOmmlElement(doc, 'e');
    base.appendChild(mathMlToOmmlNode(children[0]!, doc, depth + 1, maxDepth));
    script.appendChild(base);
    for (const [name, child] of [['sub', children[mark + 1]!], ['sup', children[mark + 2]!]] as const) {
      const entry = createOmmlElement(doc, name);
      entry.appendChild(mathMlToOmmlNode(child, doc, depth + 1, maxDepth));
      script.appendChild(entry);
    }
    return script;
  }
  if (node.tag === 'msqrt' || node.tag === 'mroot') {
    const expected = node.tag === 'msqrt' ? 1 : 2;
    if (children.length !== expected) throw new Error(`${node.tag} must contain ${expected} child elements.`);
    const radical = createOmmlElement(doc, 'rad');
    if (node.tag === 'mroot') {
      const properties = createOmmlElement(doc, 'radPr');
      const degreeHide = createOmmlElement(doc, 'degHide');
      setOmmlValue(degreeHide, '0');
      properties.appendChild(degreeHide);
      radical.appendChild(properties);
    }
    if (node.tag === 'mroot') {
      const degree = createOmmlElement(doc, 'deg');
      degree.appendChild(mathMlToOmmlNode(children[1]!, doc, depth + 1, maxDepth));
      radical.appendChild(degree);
    }
    const base = createOmmlElement(doc, 'e');
    base.appendChild(mathMlToOmmlNode(children[0]!, doc, depth + 1, maxDepth));
    radical.appendChild(base);
    return radical;
  }
  if (node.tag === 'munder' || node.tag === 'mover') {
    if (children.length !== 2) throw new Error(`${node.tag} must contain a base and a limit.`);
    const accent = node.tag === 'mover' && children[1]!.tag === 'mo' && children[1]!.attrs?.accent === 'true';
    const limit = createOmmlElement(doc, accent ? 'acc' : node.tag === 'munder' ? 'limLow' : 'limUpp');
    if (accent) {
      const properties = createOmmlElement(doc, 'accPr');
      const char = createOmmlElement(doc, 'chr');
      setOmmlValue(char, elementText(children[1]!));
      properties.appendChild(char);
      limit.appendChild(properties);
    }
    const base = createOmmlElement(doc, 'e');
    base.appendChild(mathMlToOmmlNode(children[0]!, doc, depth + 1, maxDepth));
    limit.appendChild(base);
    if (!accent) {
      const script = createOmmlElement(doc, 'lim');
      script.appendChild(mathMlToOmmlNode(children[1]!, doc, depth + 1, maxDepth));
      limit.appendChild(script);
    }
    return limit;
  }
  if (node.tag === 'munderover') return buildNary([node, { tag: 'mrow', children: [] }], doc, depth, maxDepth);
  if (node.tag === 'mtable') {
    const matrix = createOmmlElement(doc, 'm');
    for (const rowNode of children) {
      if (rowNode.tag !== 'mtr') throw new Error('mtable children must be mtr nodes.');
      const row = createOmmlElement(doc, 'mr');
      for (const cellNode of mathMlChildren(rowNode)) {
        if (cellNode.tag !== 'mtd') throw new Error('mtr children must be mtd nodes.');
        const cell = createOmmlElement(doc, 'e');
        for (const content of mathMlChildren(cellNode)) cell.appendChild(mathMlToOmmlNode(content, doc, depth + 2, maxDepth));
        row.appendChild(cell);
      }
      matrix.appendChild(row);
    }
    return matrix;
  }
  if (node.tag === 'mtr' || node.tag === 'mtd') {
    const wrapper = createOmmlElement(doc, node.tag === 'mtr' ? 'mr' : 'e');
    for (const child of children) wrapper.appendChild(mathMlToOmmlNode(child, doc, depth + 1, maxDepth));
    return wrapper;
  }
  if (node.tag === 'menclose' || node.tag === 'mphantom') {
    const wrapper = createOmmlElement(doc, node.tag === 'menclose' ? 'borderBox' : 'phant');
    const entry = createOmmlElement(doc, 'e');
    for (const child of children) entry.appendChild(mathMlToOmmlNode(child, doc, depth + 1, maxDepth));
    wrapper.appendChild(entry);
    return wrapper;
  }
  throw new Error(`Unsupported MathML tag: ${node.tag}.`);
}

function buildNary(nodes: MathMlNode[], doc: import('@xmldom/xmldom').Document, depth: number, maxDepth: number): Element {
  const [operator, body] = nodes;
  const parts = operator?.children ?? [];
  if (operator?.tag !== 'munderover' || parts.length !== 3 || parts[0]!.tag !== 'mo') {
    throw new Error('N-ary operators must contain a symbol, lower limit, upper limit, and body.');
  }
  const nary = createOmmlElement(doc, 'nary');
  const properties = createOmmlElement(doc, 'naryPr');
  const char = createOmmlElement(doc, 'chr');
  setOmmlValue(char, elementText(parts[0]!));
  properties.appendChild(char);
  nary.appendChild(properties);
  for (const [name, child] of [['sub', parts[1]!], ['sup', parts[2]!], ['e', body!]] as const) {
    const entry = createOmmlElement(doc, name);
    entry.appendChild(mathMlToOmmlNode(child, doc, depth + 1, maxDepth));
    nary.appendChild(entry);
  }
  return nary;
}

export function mathMlToOmml(node: MathMlNode, doc: import('@xmldom/xmldom').Document, options?: { maxDepth?: number }): Element {
  const maxDepth = options?.maxDepth ?? MAX_DEPTH;
  if (!Number.isSafeInteger(maxDepth) || maxDepth < 0) throw new Error('maxDepth must be a non-negative safe integer.');
  if (node?.tag === 'math') {
    const math = mathMlToOmmlNode(node, doc, 0, maxDepth);
    if (node.attrs?.display !== 'block') return math;
    const paragraph = createOmmlElement(doc, 'oMathPara');
    paragraph.appendChild(math);
    return paragraph;
  }
  const root = createOmmlElement(doc, 'oMath');
  root.appendChild(mathMlToOmmlNode(node, doc, 0, maxDepth));
  return root;
}
