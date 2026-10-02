import type { Document, Element, Node } from '@xmldom/xmldom';
import type { ParagraphFormat, RevisionInfo, RevisionMark, ReviewerAuthorKind, RunFormat } from './types.js';
import { readParagraphProperties, readRunProperties } from './styles.js';
import { WORD_NS, assertText, children, descendants, wordElement } from './xml.js';
import { collectTextElements, elementText, PROPERTY_ORDER, reviewerBucketKey, reviewerBucketOf } from './internal/elements.js';

export { reviewerBucketKey, reviewerBucketOf } from './internal/elements.js';

const REVISION_NAMES = ['ins', 'del', 'moveFrom', 'moveTo', 'rPrChange', 'pPrChange', 'tblPrChange', 'trPrChange', 'tcPrChange', 'sectPrChange', 'cellIns', 'cellDel'] as const;
const DEFAULT_REVISION_AUTHOR = 'docx-view';
const WRAPPER_KIND = {
  ins: 'insertion',
  moveTo: 'move',
  del: 'deletion',
  moveFrom: 'move',
} as const satisfies Partial<Record<string, RevisionMark['kind']>>;
const CHANGE_KIND = {
  rPrChange: 'runFormatChange',
  pPrChange: 'paragraphFormatChange',
  tblPrChange: 'tableFormatChange',
  trPrChange: 'rowFormatChange',
  tcPrChange: 'cellFormatChange',
  sectPrChange: 'sectionFormatChange',
} as const satisfies Partial<Record<string, RevisionMark['kind']>>;
const REVIEWER_PLACEHOLDERS: Record<Exclude<ReviewerAuthorKind, 'named'>, string> = {
  unattributed: '(unattributed)',
  empty: '(empty author)',
  blank: '(blank author)',
};
function revisionAttribute(element: Element, name: string): string | undefined {
  return element.getAttributeNS(WORD_NS, name) ?? element.getAttribute(`w:${name}`) ?? undefined;
}

function revisionIdOf(element: Element): number | undefined {
  const raw = revisionAttribute(element, 'id');
  if (!raw || !/^\d+$/.test(raw)) return undefined;
  const value = Number(raw);
  return Number.isSafeInteger(value) ? value : undefined;
}

export function revisionAuthorOf(element: Element): string | undefined {
  return revisionAttribute(element, 'author');
}

function revisionDateOf(element: Element): string | undefined {
  const date = revisionAttribute(element, 'date')?.trim();
  return date && Number.isFinite(Date.parse(date)) ? date : undefined;
}

function moveNameOf(element: Element): string {
  const name = revisionAttribute(element, 'name');
  return name?.trim() ?? '';
}

function moveSideOf(element: Element): 'from' | 'to' | undefined {
  if (element.localName === 'moveFrom') return 'from';
  if (element.localName === 'moveTo') return 'to';
  return undefined;
}

function previousFormatOf(element: Element, theme: Parameters<typeof readRunProperties>[1]): RunFormat | ParagraphFormat | undefined {
  const compact = <T extends object>(value: T): T =>
    Object.fromEntries(Object.entries(value).filter(([, entry]) => entry !== undefined)) as T;
  switch (element.localName) {
    case 'rPrChange': {
      const format = compact(readRunProperties(children(element, 'rPr')[0], theme));
      return Object.values(format).some((value) => value !== undefined) ? format : undefined;
    }
    case 'pPrChange': {
      const format = compact(readParagraphProperties(children(element, 'pPr')[0]));
      return Object.values(format).some((value) => value !== undefined) ? format : undefined;
    }
    default:
      return undefined;
  }
}

function orderedRevisionChild(parent: Element, name: string, orderName = parent.localName as keyof typeof PROPERTY_ORDER): Element {
  let result = children(parent, name)[0];
  if (result) return result;
  result = wordElement(parent.ownerDocument!, name);
  const order: string[] = [...(PROPERTY_ORDER[orderName] ?? [])];
  const position = order.indexOf(name);
  const following = position === -1 ? undefined : children(parent).find((child) => order.indexOf(child.localName ?? '') > position);
  parent.insertBefore(result, following ?? null);
  return result;
}

export function hasRevisionMarkup(root: Document | Element): boolean {
  return REVISION_NAMES.some((name) => descendants(root as Document | Element, name).length > 0);
}

export function visibleTextOf(element: Element, excludedRuns?: ReadonlySet<Element>): string {
  return collectTextElements(element, 'visible', excludedRuns).map(elementText).join('');
}

export function deletedTextOf(element: Element): string {
  return collectTextElements(element, 'deleted').map(elementText).join('');
}

export function reviewerBucketLabel(author: { kind: ReviewerAuthorKind; author?: string }): string {
  return author.kind === 'named' ? (author.author ?? '') : REVIEWER_PLACEHOLDERS[author.kind];
}

export function readRevisionMark(element: Element, kind: RevisionMark['kind'], theme?: Parameters<typeof readRunProperties>[1]):
  (RevisionMark & Pick<RevisionInfo, 'previousFormat'>) | undefined {
  const id = revisionIdOf(element);
  if (id === undefined) return undefined;
  return {
    id,
    kind,
    author: revisionAuthorOf(element),
    date: revisionDateOf(element),
    ...(kind === 'move' ? {
      move: {
        name: moveNameOf(element),
        side: moveSideOf(element) ?? 'from',
      },
    } : {}),
    previousFormat: theme ? previousFormatOf(element, theme) : undefined,
  };
}

export function readRunRevisionMarks(run: Element, paragraph: Element, theme: Parameters<typeof readRunProperties>[1]):
  (RevisionMark & Pick<RevisionInfo, 'previousFormat'>)[] {
  const result: (RevisionMark & Pick<RevisionInfo, 'previousFormat'>)[] = [];
  const push = (element: Element | undefined, kind: RevisionMark['kind'] | undefined): void => {
    if (!element || !kind) return;
    const mark = readRevisionMark(element, kind, theme);
    if (mark) result.push(mark);
  };
  for (let parent = run.parentNode; parent && parent !== paragraph; parent = parent.parentNode) {
    if (parent.nodeType !== 1) continue;
    const element = parent as Element;
    if (element.namespaceURI !== WORD_NS) continue;
    push(element, WRAPPER_KIND[element.localName as keyof typeof WRAPPER_KIND]);
  }
  const props = children(run, 'rPr')[0];
  if (!props) return result;
  push(children(props, 'ins')[0], 'insertion');
  push(children(props, 'del')[0], 'deletion');
  push(children(props, 'rPrChange')[0], 'runFormatChange');
  return result;
}

export function readParagraphRevisionMark(paragraph: Element, theme: Parameters<typeof readRunProperties>[1]):
  (RevisionMark & Pick<RevisionInfo, 'previousFormat'>) | undefined {
  const props = children(paragraph, 'pPr')[0];
  const change = props ? children(props, 'pPrChange')[0] : undefined;
  return change ? readRevisionMark(change, 'paragraphFormatChange', theme) : undefined;
}

export function createRevisionWrapper(document: Document, kind: 'ins' | 'del' | 'cellIns' | 'cellDel' | 'cellMerge', author?: string, date?: string): Element {
  return applyRevisionMetadata(wordElement(document, kind), author, date);
}

export function markRevision(
  parent: Element,
  kind: 'ins' | 'del' | 'cellIns' | 'cellDel' | 'cellMerge',
  author?: string,
  date?: string,
  orderName?: keyof typeof PROPERTY_ORDER,
): Element {
  const marker = orderedRevisionChild(parent, kind, orderName);
  return applyRevisionMetadata(marker, author, date);
}

export function markFormatRevision(
  parent: Element,
  kind: 'rPrChange' | 'pPrChange' | 'tblPrChange' | 'trPrChange' | 'tcPrChange',
  snapshotName: 'rPr' | 'pPr' | 'tblPr' | 'trPr' | 'tcPr',
  author?: string,
  date?: string,
): Element {
  const marker = orderedRevisionChild(parent, kind);
  applyRevisionMetadata(marker, author, date);
  for (const child of children(marker)) {
    if (child.localName !== snapshotName) marker.removeChild(child);
  }
  let snapshot = children(marker, snapshotName)[0];
  if (!snapshot) {
    snapshot = wordElement(parent.ownerDocument!, snapshotName);
    marker.appendChild(snapshot);
    return snapshot;
  }
  return snapshot;
}

function applyRevisionMetadata(marker: Element, author?: string, date?: string): Element {
  if (author !== undefined) assertText(author, 'author');
  if (date !== undefined) assertText(date, 'date');
  const document = marker.ownerDocument!;
  const used = Array.from(document.getElementsByTagNameNS(WORD_NS, '*'))
    .map((element) => revisionIdOf(element))
    .filter((value): value is number => value !== undefined);
  const resolvedAuthor = author?.trim() || DEFAULT_REVISION_AUTHOR;
  marker.setAttributeNS(WORD_NS, 'w:id', String((used.length ? Math.max(...used) : 0) + 1));
  marker.setAttributeNS(WORD_NS, 'w:author', resolvedAuthor);
  if (date?.trim()) marker.setAttributeNS(WORD_NS, 'w:date', date.trim());
  else { marker.removeAttributeNS(WORD_NS, 'date'); marker.removeAttribute('w:date'); }
  marker.removeAttributeNS(WORD_NS, 'val');
  marker.removeAttribute('w:val');
  return marker;
}
