import type { Document, Element, Node } from '@xmldom/xmldom';
import type { ReviewerAuthorKind, ReviewerFilterAuthor } from '../types.js';
import { children, descendants, isValidXmlCharCode, WORD_NS, wordElement } from '../xml.js';

const encoder = new TextEncoder();
const REVIEWER_FILTER_BUCKET_KEYS = new Set(['unattributed', 'empty', 'blank']);

export const PARAGRAPH_LOOKUP_CACHE = new WeakMap<Document, Element[]>();

export function assertIndex(value: unknown): asserts value is number {
  if (!Number.isSafeInteger(value) || (value as number) < 0) {
    throw new Error('Index/revision must be a non-negative safe integer.');
  }
}

export function decodeXml(bytes: Uint8Array): string {
  const utf16le = (bytes[0] === 0xff && bytes[1] === 0xfe) || (bytes[0] === 0x3c && bytes[1] === 0);
  const utf16be = (bytes[0] === 0xfe && bytes[1] === 0xff) || (bytes[0] === 0 && bytes[1] === 0x3c);
  return new TextDecoder(utf16le ? 'utf-16le' : utf16be ? 'utf-16be' : 'utf-8', { fatal: true }).decode(bytes);
}

export function encodeXml(xml: string): Uint8Array {
  return encoder.encode(xml.replace(/^(<\?xml\b[^?]*\bencoding\s*=\s*)(["'])[^"']*\2/i, '$1"UTF-8"'));
}

export function dirname(path: string): string {
  const slash = path.lastIndexOf('/');
  return slash === -1 ? '' : path.slice(0, slash);
}

export function basename(path: string): string {
  return path.slice(path.lastIndexOf('/') + 1);
}

export function bodyOf(document: Document): Element {
  const root = document.documentElement;
  if (!root || root.namespaceURI !== WORD_NS || root.localName !== 'document') {
    throw new Error('Only transitional WordprocessingML documents are supported.');
  }
  const body = children(root, 'body');
  if (body.length !== 1) throw new Error('Main document must have exactly one w:body.');
  return body[0]!;
}

export function mainParagraphElements(body: Element): Element[] {
  return descendants(body, 'p').filter((paragraph) => {
    let ancestor = paragraph.parentNode as Element | null;
    while (ancestor && ancestor !== body) {
      if (ancestor.namespaceURI === WORD_NS && ancestor.localName === 'txbxContent') return false;
      ancestor = ancestor.parentNode as Element | null;
    }
    return true;
  });
}

export function compactDefined<T extends object>(value: T): T {
  return Object.fromEntries(Object.entries(value).filter(([, entry]) => entry !== undefined)) as T;
}

export function partDirectory(path: string): string {
  const index = path.lastIndexOf('/');
  return index === -1 ? '' : path.slice(0, index);
}

export function relsPath(partPath: string): string {
  const directory = partDirectory(partPath);
  const name = partPath.slice(partPath.lastIndexOf('/') + 1);
  return `${directory ? `${directory}/` : ''}_rels/${name}.rels`;
}

export function paragraphAt(document: Document, index: number): Element {
  assertIndex(index);
  const cached = PARAGRAPH_LOOKUP_CACHE.get(document);
  if (cached) {
    const paragraph = cached[index];
    if (!paragraph) throw new Error(`Paragraph ${index} does not exist.`);
    return paragraph;
  }
  let remaining = index;
  const stack: (Node | null)[] = [bodyOf(document).firstChild];
  while (stack.length) {
    const current = stack[stack.length - 1]!;
    if (!current) {
      stack.pop();
      continue;
    }
    stack[stack.length - 1] = current.nextSibling;
    if (current.nodeType === 1) {
      const element = current as Element;
      if (element.namespaceURI === WORD_NS && element.localName === 'txbxContent') {
        continue;
      }
      if (element.namespaceURI === WORD_NS && element.localName === 'p') {
        if (remaining === 0) return element;
        remaining--;
      }
    }
    if (current.firstChild) stack.push(current.firstChild);
  }
  throw new Error(`Paragraph ${index} does not exist.`);
}

export function blockContainerOf(document: Document): Element {
  const root = document.documentElement;
  if (!root || root.namespaceURI !== WORD_NS) throw new Error('Unsupported WordprocessingML part.');
  if (root.localName === 'document') return bodyOf(document);
  if (['hdr', 'ftr', 'footnotes', 'endnotes'].includes(root.localName ?? '')) return root;
  throw new Error('Part does not contain block-level WordprocessingML content.');
}

export function appendText(parent: Element, text: string, before: Node | null = null): void {
  const document = parent.ownerDocument!;
  for (const chunk of text.split(/(\t|\r\n|\r|\n)/)) {
    if (!chunk) continue;
    const element = wordElement(document, chunk === '\t' ? 'tab' : /^[\r\n]+$/.test(chunk) ? 'br' : 't');
    if (element.localName === 't') {
      element.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
      element.appendChild(document.createTextNode(chunk));
    }
    parent.insertBefore(element, before);
  }
}

export function properties(element: Element, name: 'pPr' | 'rPr'): Element {
  let result = children(element, name)[0];
  if (!result) {
    result = wordElement(element.ownerDocument!, name);
    element.insertBefore(result, element.firstChild);
  }
  return result;
}

// Known property order keeps generated pPr/rPr conformant without dropping unknown properties.
export const PROPERTY_ORDER = {
  pPr: ['pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl', 'numPr',
    'suppressLineNumbers', 'pBdr', 'shd', 'tabs', 'suppressAutoHyphens', 'kinsoku', 'wordWrap',
    'overflowPunct', 'topLinePunct', 'autoSpaceDE', 'autoSpaceDN', 'bidi', 'adjustRightInd',
    'snapToGrid', 'spacing', 'ind', 'contextualSpacing', 'mirrorIndents', 'suppressOverlap',
    'jc', 'textDirection', 'textAlignment', 'textboxTightWrap', 'outlineLvl', 'divId',
    'cnfStyle', 'rPr', 'sectPr', 'pPrChange'],
  rPr: ['rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike', 'dstrike',
    'outline', 'shadow', 'emboss', 'imprint', 'noProof', 'snapToGrid', 'vanish', 'webHidden',
    'color', 'spacing', 'w', 'kern', 'position', 'sz', 'szCs', 'highlight', 'u', 'effect',
    'bdr', 'shd', 'fitText', 'vertAlign', 'rtl', 'cs', 'em', 'lang', 'eastAsianLayout',
    'specVanish', 'oMath', 'rPrChange'],
  paraRPr: ['ins', 'del', 'moveFrom', 'moveTo', 'rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs',
    'caps', 'smallCaps', 'strike', 'dstrike', 'outline', 'shadow', 'emboss', 'imprint',
    'noProof', 'snapToGrid', 'vanish', 'webHidden', 'color', 'spacing', 'w', 'kern', 'position',
    'sz', 'szCs', 'highlight', 'u', 'effect', 'bdr', 'shd', 'fitText', 'vertAlign', 'rtl', 'cs',
    'em', 'lang', 'eastAsianLayout', 'specVanish', 'oMath', 'rPrChange'],
  style: ['name', 'aliases', 'basedOn', 'next', 'link', 'autoRedefine', 'hidden', 'uiPriority',
    'semiHidden', 'unhideWhenUsed', 'qFormat', 'locked', 'personal', 'personalCompose',
    'personalReply', 'rsid', 'pPr', 'rPr', 'tblPr', 'trPr', 'tcPr', 'tblStylePr', 'extLst'],
  tblPr: ['tblStyle', 'tblpPr', 'tblOverlap', 'bidiVisual', 'tblStyleRowBandSize', 'tblStyleColBandSize',
    'tblW', 'jc', 'tblCellSpacing', 'tblInd', 'tblBorders', 'shd', 'tblLayout', 'tblCellMar',
    'tblLook', 'tblCaption', 'tblDescription', 'tblPrChange'],
  // CT_TblPrEx：tblPr 的子集，相对顺序与 tblPr 一致。
  tblPrEx: ['tblW', 'jc', 'tblCellSpacing', 'tblInd', 'tblBorders', 'shd', 'tblLayout', 'tblCellMar', 'tblLook',
    'tblPrExChange'],
  trPr: ['cnfStyle', 'divId', 'gridBefore', 'gridAfter', 'wBefore', 'wAfter', 'cantSplit', 'trHeight',
    'tblHeader', 'tblCellSpacing', 'jc', 'hidden', 'ins', 'del', 'trPrChange'],
  tcPr: ['cnfStyle', 'tcW', 'gridSpan', 'hMerge', 'vMerge', 'tcBorders', 'shd', 'noWrap', 'tcMar',
    'textDirection', 'tcFitText', 'vAlign', 'hideMark', 'headers', 'cellIns', 'cellDel', 'cellMerge', 'tcPrChange'],
  tblBorders: ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'],
  tcBorders: ['top', 'left', 'bottom', 'right', 'insideH', 'insideV', 'tl2br', 'tr2bl'],
  tblCellMar: ['top', 'left', 'bottom', 'right'],
  tcMar: ['top', 'left', 'bottom', 'right'],
};

/**
 * 按 schema 顺序取（没有就建）一个属性子元素。
 *
 * **名字不在 `PROPERTY_ORDER` 里就追加到末尾**，而且没有任何提示——写出来的 XML 是顺序
 * 不合法的，Word 可能直接忽略那条属性。所以新增写入点时必须同时把元素名补进上面的表里；
 * `test/table.test.mjs` 的 `every generated table property name is in PROPERTY_ORDER`
 * 就是盯着这件事的（`w:trPr` 原先漏了 `w:tblCellSpacing`）。
 */
export function property(parent: Element, name: string): Element {
  let result = children(parent, name)[0];
  if (!result) {
    result = wordElement(parent.ownerDocument!, name);
    const order = PROPERTY_ORDER[parent.localName as keyof typeof PROPERTY_ORDER] ?? [];
    const position = order.indexOf(name);
    const following = position === -1 ? undefined : children(parent).find(child => {
      const childPosition = order.indexOf(child.localName!);
      return childPosition > position;
    });
    parent.insertBefore(result, following ?? null);
  }
  return result;
}

const BODY_TEXT_NAMES = new Set(['t', 'tab', 'br', 'cr', 'noBreakHyphen', 'softHyphen', 'sym']);
const DELETED_TEXT_NAMES = new Set(['t', 'delText', 'tab', 'br', 'cr', 'noBreakHyphen', 'softHyphen', 'sym']);

function isDeletedWrapper(element: Element): boolean {
  return element.namespaceURI === WORD_NS && ['del', 'moveFrom'].includes(element.localName ?? '');
}

/**
 * 段落正文由哪些元素承载。`revisions.ts` 与 `content-control.ts` 原先各有一份一模一样的
 * 遍历(连 `w:sym` 的写法都抄了两遍,只有有效性校验那一处悄悄不同),规则必须只有一处,
 * 否则下一个人改了一边就会走偏。
 *
 * `w:ruby` 的 `w:rt` 是**注音、不是正文**:它排在基字符上方,不进阅读顺序。收进来的话
 * `paragraph.text` 会变成「读hàn漢字」,而 `setParagraphText()` 把文本按 `w:t` 顺序铺回去时
 * 会把注音当成可写的槽位——实测改一次段落文本就会留下 `<w:rt><w:r/></w:rt>` 的空壳,
 * 注音和基字符一起没了。只有 `w:rubyBase` 里的才算正文。
 */
export function collectTextElements(element: Element, mode: 'visible' | 'deleted',
  excludedRuns?: ReadonlySet<Element>): Element[] {
  const result: Element[] = [];
  const walk = (node: Node, deletedDepth = 0): void => {
    for (let child = node.firstChild; child; child = child.nextSibling) {
      if (child.nodeType !== 1) continue;
      const current = child as Element;
      if (excludedRuns?.has(current)) continue;
      if (current.namespaceURI === WORD_NS) {
        const localName = current.localName ?? '';
        if (localName === 'p' || localName === 'rt') continue;
        const inDeleted = deletedDepth > 0 || localName === 'delText';
        if (mode === 'visible' && BODY_TEXT_NAMES.has(localName) && !inDeleted) {
          result.push(current);
          continue;
        }
        if (mode === 'deleted' && DELETED_TEXT_NAMES.has(localName) && inDeleted) {
          result.push(current);
          continue;
        }
      }
      walk(current, deletedDepth + (isDeletedWrapper(current) ? 1 : 0));
    }
  };
  walk(element);
  return result;
}

/** `w:sym` 的字符码来自文档内容,无效时退回 U+FFFD,绝不原样放进字符串。 */
export function elementText(element: Element): string {
  if (['t', 'delText'].includes(element.localName ?? '')) return element.textContent ?? '';
  if (element.localName === 'tab') return '\t';
  if (element.localName === 'noBreakHyphen') return '\u2011';
  if (element.localName === 'softHyphen') return '\u00ad';
  if (element.localName === 'sym') {
    const value = element.getAttributeNS(WORD_NS, 'char') ?? element.getAttribute('w:char');
    if (!value || !/^[a-f0-9]{1,4}$/i.test(value)) return '';
    const code = Number.parseInt(value, 16);
    return Number.isFinite(code) && isValidXmlCharCode(code) ? String.fromCharCode(code) : '\ufffd';
  }
  return '\n';
}

export function ownRuns(paragraph: Element): Element[] {
  return descendants(paragraph, 'r').filter(run => {
    let parent = run.parentNode;
    while (parent && parent !== paragraph) {
      if (parent.nodeType === 1 && (parent as Element).namespaceURI === WORD_NS) {
        const localName = (parent as Element).localName;
        if (localName === 'p') return false;
        // w:ruby 把注音和基字符各放在一个 w:r 里,而这些 w:r 又在外层 w:r 之内。
        // 后代搜索会把同一批字符数三遍(外层 run、w:rt 的 run、w:rubyBase 的 run),
        // 下标因此全不可靠,formatRun() 落到内层 run 上还会静默无效。只认最外层那个。
        if (localName === 'r') return false;
      }
      parent = parent.parentNode;
    }

    return true;
  });
}

export function isDescendantOfWithin(node: Node, ancestor: Node, stopAt: Node): boolean {
  let current: Node | null = node;
  while (current && current !== stopAt) {
    if (current === ancestor) return true;
    current = current.parentNode;
  }
  return false;
}

export function removeWordAttribute(element: Element, name: string): void {
  element.removeAttributeNS(WORD_NS, name);
  element.removeAttribute(`w:${name}`);
}

export function reviewerBucketOf(author: string | undefined): ReviewerFilterAuthor {
  if (author === undefined) return { kind: 'unattributed' };
  if (author === '') return { kind: 'empty', author: '' };
  const normalized = author.trim();
  if (!normalized) return { kind: 'blank', author };
  return { kind: 'named', author: normalized };
}

export function reviewerBucketKey(author: { kind: ReviewerAuthorKind; author?: string }): string {
  return author.kind === 'named' ? `named:${author.author ?? ''}` : author.kind;
}

export function reviewerFilterKeyOf(author: string): string {
  if (REVIEWER_FILTER_BUCKET_KEYS.has(author) || author.startsWith('named:')) return author;
  return reviewerBucketKey(reviewerBucketOf(author));
}

export function normalizeReviewerFilterAuthors(authors: string[]): Set<string> {
  return new Set(authors.map((author) => reviewerFilterKeyOf(author)));
}

export function nearestParagraph(node: Node | null): Element | null {
  let current = node;
  while (current) {
    if (current.nodeType === 1 && (current as Element).namespaceURI === WORD_NS && (current as Element).localName === 'p') {
      return current as Element;
    }
    current = current.parentNode;
  }
  return null;
}

export function preOrderElements(root: Element): Element[] {
  const result: Element[] = [];
  const walk = (node: Node): void => {
    for (let child = node.firstChild; child; child = child.nextSibling) {
      if (child.nodeType !== 1) continue;
      const element = child as Element;
      result.push(element);
      walk(element);
    }
  };
  walk(root);
  return result;
}
