import type { Document, Element } from '@xmldom/xmldom';
import type { BookmarkInfo } from './types.js';
import type { PartAccess } from './internal/context.js';
import { assertIndex } from './operations.js';
import { WORD_NS, assertText, children, descendants, wordElement } from './xml.js';
import { bodyOf, mainParagraphElements, nearestParagraph, paragraphAt, preOrderElements } from './internal/elements.js';

type Bound<F> = F extends (ctx: BookmarkContext, ...args: infer A) => infer R ? (...args: A) => R : never;

export interface BookmarkContext extends Pick<PartAccess, 'mainPath' | 'getPartDocument' | 'updatePartXml'> {
  getBookmarks: Bound<typeof getBookmarks>;
}

export function isInternalBookmark(name: string): boolean {
  return /^_GoBack$/i.test(name) || /^_Toc/i.test(name);
}

export function getBookmarks(ctx: BookmarkContext, options: { includeInternal?: boolean } = {}): BookmarkInfo[] {
    const body = bodyOf(ctx.getPartDocument(ctx.mainPath));
    const paragraphs = mainParagraphElements(body);
    const paragraphIndex = new Map(paragraphs.map((paragraph, index) => [paragraph, index]));
    const order = preOrderElements(body);
    const paragraphFromNode = (node: Element): number | undefined => {
      const paragraph = nearestParagraph(node);
      if (paragraph) return paragraphIndex.get(paragraph);
      const offset = order.indexOf(node);
      if (offset === -1) return undefined;
      for (let index = offset + 1; index < order.length; index++) {
        const candidate = order[index];
        if (candidate?.localName === 'p' && candidate.namespaceURI === WORD_NS) {
          return paragraphIndex.get(candidate);
        }
      }
      return paragraphs.length ? paragraphs.length - 1 : 0;
    };
    const starts = new Map<number, { name: string; paragraph: number }>();
    const ends = new Map<number, number>();
    for (const start of descendants(body, 'bookmarkStart')) {
      const id = Number(start.getAttributeNS(WORD_NS, 'id') ?? start.getAttribute('w:id'));
      const name = start.getAttributeNS(WORD_NS, 'name') ?? start.getAttribute('w:name') ?? '';
      const startParagraph = paragraphFromNode(start);
      if (!Number.isSafeInteger(id) || startParagraph === undefined || !name) continue;
      starts.set(id, { name, paragraph: startParagraph });
    }
    for (const end of descendants(body, 'bookmarkEnd')) {
      const id = Number(end.getAttributeNS(WORD_NS, 'id') ?? end.getAttribute('w:id'));
      const endParagraph = paragraphFromNode(end);
      if (!Number.isSafeInteger(id) || endParagraph === undefined) continue;
      ends.set(id, endParagraph);
    }
    return [...starts.entries()]
      .map(([id, start]) => ({
        id,
        name: start.name,
        startParagraph: start.paragraph,
        endParagraph: ends.get(id) ?? start.paragraph,
        isInternal: isInternalBookmark(start.name),
      }))
      .filter(bookmark => options.includeInternal || !bookmark.isInternal);
}

export function insertBookmark(
  ctx: BookmarkContext,
  name: string,
  range: { startParagraph: number; endParagraph?: number },
): BookmarkInfo {
    assertText(name, 'name');
    assertIndex(range.startParagraph);
    if (range.endParagraph !== undefined) assertIndex(range.endParagraph);
    if (range.endParagraph !== undefined && range.endParagraph < range.startParagraph) {
      throw new Error('range.endParagraph must be >= range.startParagraph.');
    }
    if (!name) throw new Error('name must not be empty.');
    if (ctx.getBookmarks({ includeInternal: true }).some(bookmark => bookmark.name === name)) {
      throw new Error(`Bookmark "${name}" already exists.`);
    }
    let bookmark: BookmarkInfo | undefined;
    ctx.updatePartXml(ctx.mainPath, document => {
      const start = paragraphAt(document, range.startParagraph);
      const end = paragraphAt(document, range.endParagraph ?? range.startParagraph);
      const ids = descendants(bodyOf(document), 'bookmarkStart')
        .map(item => Number(item.getAttributeNS(WORD_NS, 'id') ?? item.getAttribute('w:id')))
        .filter(value => Number.isSafeInteger(value));
      const id = (ids.length ? Math.max(...ids) : 0) + 1;
      const startMark = wordElement(document, 'bookmarkStart');
      startMark.setAttributeNS(WORD_NS, 'w:id', String(id));
      startMark.setAttributeNS(WORD_NS, 'w:name', name);
      start.insertBefore(startMark, children(start, 'pPr')[0]?.nextSibling ?? start.firstChild);
      const endMark = wordElement(document, 'bookmarkEnd');
      endMark.setAttributeNS(WORD_NS, 'w:id', String(id));
      end.appendChild(endMark);
      bookmark = {
        id,
        name,
        startParagraph: range.startParagraph,
        endParagraph: range.endParagraph ?? range.startParagraph,
        isInternal: isInternalBookmark(name),
      };
    });
    return bookmark!;
}

export function deleteBookmark(ctx: BookmarkContext, name: string): void {
    assertText(name, 'name');
    ctx.updatePartXml(ctx.mainPath, document => {
      const starts = descendants(bodyOf(document), 'bookmarkStart')
        .filter(start => (start.getAttributeNS(WORD_NS, 'name') ?? start.getAttribute('w:name')) === name);
      if (!starts.length) throw new Error(`Bookmark "${name}" does not exist.`);
      const ids = starts.map(start => start.getAttributeNS(WORD_NS, 'id') ?? start.getAttribute('w:id'));
      for (const start of starts) start.parentNode!.removeChild(start);
      for (const end of descendants(bodyOf(document), 'bookmarkEnd')) {
        const id = end.getAttributeNS(WORD_NS, 'id') ?? end.getAttribute('w:id');
        if (ids.includes(id)) end.parentNode!.removeChild(end);
      }
    });
}

export function ensureTocBookmark(
  document: Document,
  paragraphIndex: number,
  namesByParagraph: Map<number, string>,
  usedNames: Set<string>,
  nextId: { value: number },
): string | undefined {
    const existing = namesByParagraph.get(paragraphIndex);
    if (existing) return existing;
    const paragraph = mainParagraphElements(bodyOf(document))[paragraphIndex];
    if (!paragraph) return undefined;
    let name = `_Toc${paragraphIndex + 1}`;
    let suffix = 1;
    while (usedNames.has(name)) name = `_Toc${paragraphIndex + 1}_${suffix++}`;
    usedNames.add(name);
    namesByParagraph.set(paragraphIndex, name);
    const start = wordElement(document, 'bookmarkStart');
    const end = wordElement(document, 'bookmarkEnd');
    const id = String(nextId.value++);
    start.setAttributeNS(WORD_NS, 'w:id', id);
    start.setAttributeNS(WORD_NS, 'w:name', name);
    end.setAttributeNS(WORD_NS, 'w:id', id);
    const properties = children(paragraph, 'pPr')[0];
    paragraph.insertBefore(start, properties?.nextSibling ?? paragraph.firstChild);
    paragraph.appendChild(end);
    return name;
}
