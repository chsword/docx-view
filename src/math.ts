import type { Element } from '@xmldom/xmldom';
import type { MathMlNode } from './types.js';

export const MATH_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
const MAX_DEPTH = 64;

let conversionLimit = MAX_DEPTH;

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

function leaf(text: string): MathMlNode {
  if (/^\d+(?:[.,]\d+)*$/.test(text)) return { tag: 'mn', text };
  if (/^[+\-−×÷=<>≤≥±*/(),.;:|]$/.test(text)) return { tag: 'mo', text };
  return { tag: 'mi', text };
}

function childrenAsRow(element: Element, depth: number): MathMlNode {
  const children = elementChildren(element).filter(child => child.namespaceURI === MATH_NS)
    .flatMap(child => convert(child, depth + 1));
  return children.length === 1 ? children[0]! : { tag: 'mrow', children };
}

function convert(element: Element, depth: number): MathMlNode[] {
  if (depth > conversionLimit) return [{ tag: 'mtext', text: '…' }];
  const name = element.localName;
  if (name === 't') return [leaf(element.textContent ?? '')];
  if (name === 'r') {
    const text = textOf(first(element, 't'));
    if (!text) return [];
    const node = leaf(text);
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
    if (variant) node.attrs = { mathvariant: variant };
    else if (node.tag === 'mi') node.attrs = { mathvariant: 'italic' };
    return [node];
  }
  if (name === 'oMath' || name === 'oMathPara') {
    const content = mathChildren(element).flatMap(child =>
      name === 'oMathPara' && child.localName === 'oMath'
        ? mathChildren(child).flatMap(nested => convert(nested, depth + 1))
        : convert(child, depth + 1));
    return [{ tag: 'math', attrs: { display: name === 'oMathPara' ? 'block' : 'inline' }, children: content }];
  }
  if (name === 'e' || name === 'num' || name === 'den' || name === 'sub' || name === 'sup' || name === 'deg' ||
      name === 'fName' || name === 'fName' || name === 'lim' || name === 'mr') {
    return [childrenAsRow(element, depth)];
  }
  if (name === 'f') {
    const num = first(element, 'num');
    const den = first(element, 'den');
    const numerator = num ? childrenAsRow(num, depth) : { tag: 'mrow', children: [] };
    const denominator = den ? childrenAsRow(den, depth) : { tag: 'mrow', children: [] };
    const fractionType = prop(first(element, 'fPr'), 'type', '');
    if (fractionType === 'lin') {
      return [{ tag: 'mrow', children: [numerator, { tag: 'mo', text: '/' }, denominator] }];
    }
    const attrs = fractionType === 'noBar' ? { linethickness: '0' } : undefined;
    return [{ tag: 'mfrac', ...(attrs ? { attrs } : {}), children: [numerator, denominator] }];
  }
  if (name === 'sSup' || name === 'sSub' || name === 'sSubSup') {
    const baseElement = first(element, 'e');
    const base = baseElement ? childrenAsRow(baseElement, depth) : { tag: 'mrow', children: [] };
    const sub = first(element, 'sub');
    const sup = first(element, 'sup');
    if (name === 'sSup') return [{ tag: 'msup', children: [base, sup ? childrenAsRow(sup, depth) : { tag: 'mrow', children: [] }] }];
    if (name === 'sSub') return [{ tag: 'msub', children: [base, sub ? childrenAsRow(sub, depth) : { tag: 'mrow', children: [] }] }];
    return [{ tag: 'msubsup', children: [base, sub ? childrenAsRow(sub, depth) : { tag: 'mrow', children: [] }, sup ? childrenAsRow(sup, depth) : { tag: 'mrow', children: [] }] }];
  }
  if (name === 'rad') {
    const base = first(element, 'e');
    const degree = first(element, 'deg');
    return [degree ? { tag: 'mroot', children: [base ? childrenAsRow(base, depth) : { tag: 'mrow', children: [] }, childrenAsRow(degree, depth)] }
      : { tag: 'msqrt', children: [base ? childrenAsRow(base, depth) : { tag: 'mrow', children: [] }] }];
  }
  if (name === 'nary') {
    const props = first(element, 'naryPr');
    const symbol = directText(first(props ?? element, 'chr'), '∑');
    const op: MathMlNode = { tag: 'mo', text: symbol };
    const base = first(element, 'e');
    const sub = first(element, 'sub');
    const sup = first(element, 'sup');
    const children = [op, base ? childrenAsRow(base, depth) : { tag: 'mrow', children: [] }];
    const loc = prop(props ?? element, 'limLoc', 'undOvr');
    return [loc === 'subSup' ? { tag: 'msubsup', children: [children[0]!, sub ? childrenAsRow(sub, depth) : { tag: 'mrow', children: [] }, sup ? childrenAsRow(sup, depth) : { tag: 'mrow', children: [] }] }
      : { tag: 'munderover', children: [children[0]!, sub ? childrenAsRow(sub, depth) : { tag: 'mrow', children: [] }, sup ? childrenAsRow(sup, depth) : { tag: 'mrow', children: [] }] }];
  }
  if (name === 'd') {
    const props = first(element, 'dPr');
    const beg = directText(first(props ?? element, 'begChr'), '(');
    const end = directText(first(props ?? element, 'endChr'), ')');
    const sep = directText(first(props ?? element, 'sepChr'), ',');
    const content = mathChildren(element, 'e').flatMap((entry, index) => [
      ...(index ? [{ tag: 'mo', attrs: { stretchy: 'true' }, text: sep }] : []),
      childrenAsRow(entry, depth),
    ]);
    return [{ tag: 'mrow', children: [{ tag: 'mo', attrs: { stretchy: 'true' }, text: beg }, ...content, { tag: 'mo', attrs: { stretchy: 'true' }, text: end }] }];
  }
  if (name === 'func') {
    const fname = first(element, 'fName');
    const expr = first(element, 'e');
    return [{ tag: 'mrow', children: [fname ? childrenAsRow(fname, depth) : { tag: 'mi', text: '' }, expr ? childrenAsRow(expr, depth) : { tag: 'mrow', children: [] }] }];
  }
  if (name === 'limLow' || name === 'limUpp') {
    const base = first(element, 'e');
    const lim = first(element, 'lim');
    return [{ tag: name === 'limLow' ? 'munder' : 'mover', children: [base ? childrenAsRow(base, depth) : { tag: 'mrow', children: [] }, lim ? childrenAsRow(lim, depth) : { tag: 'mrow', children: [] }] }];
  }
  if (name === 'acc' || name === 'bar' || name === 'groupChr') {
    const props = first(element, `${name}Pr`);
    const base = first(element, 'e');
    const chr = directText(first(props ?? element, 'chr'), name === 'acc' ? '̂' : name === 'groupChr' ? '⏞' : '¯');
    const pos = prop(props ?? element, 'pos', 'top');
    return [{ tag: name === 'acc' || pos === 'top' ? 'mover' : 'munder', children: [base ? childrenAsRow(base, depth) : { tag: 'mrow', children: [] }, { tag: 'mo', ...(name === 'acc' ? { attrs: { accent: 'true' } } : {}), text: chr }] }];
  }
  if (name === 'm' || name === 'eqArr') {
    const rows = name === 'm' ? mathChildren(element, 'mr') : mathChildren(element, 'e');
    return [{ tag: 'mtable', children: rows.map(row => ({ tag: 'mtr', children: [({ tag: 'mtd', children: [childrenAsRow(row, depth)] })] })) }];
  }
  if (name === 'borderBox') return [{ tag: 'menclose', children: [childrenAsRow(element, depth)] }];
  if (name === 'phant') return [{ tag: 'mphantom', children: [childrenAsRow(element, depth)] }];
  if (name === 'box') return [{ tag: 'mrow', children: [childrenAsRow(element, depth)] }];
  return [{ tag: 'mrow', children: mathChildren(element).flatMap(child => convert(child, depth + 1)) }];
}

export function ommlToMathMl(oMath: Element, options?: { maxDepth?: number }): MathMlNode {
  const previous = conversionLimit;
  conversionLimit = options?.maxDepth !== undefined && Number.isSafeInteger(options.maxDepth) && options.maxDepth >= 0
    ? options.maxDepth : MAX_DEPTH;
  try {
    return convert(oMath, 0)[0] ?? { tag: 'mtext', text: '…' };
  } finally {
    conversionLimit = previous;
  }
}

function linear(node: MathMlNode): string {
  if (node.text !== undefined) return node.text;
  const c = (node.children ?? []).map(linear);
  switch (node.tag) {
    case 'mfrac': return `(${c[0] ?? ''})/${c[1] ?? ''}`;
    case 'msup': return `${c[0] ?? ''}^${c[1] ?? ''}`;
    case 'msub': return `${c[0] ?? ''}_${c[1] ?? ''}`;
    case 'msubsup': return `${c[0] ?? ''}_${c[1] ?? ''}^${c[2] ?? ''}`;
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
