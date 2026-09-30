import type { Document, Element, Node } from '@xmldom/xmldom';
import type { CommentAnchor, CommentInfo, DocumentBlock, DocumentRange, TextRange } from './types.js';
import type { NoteKind } from './notes.js';
import type { StylesContext } from './styles.js';
import type { CacheBundle, HistoryRecorder, PartAccess } from './internal/context.js';
import { reviewerBucketKey, reviewerBucketOf } from './revisions.js';
import { resolveTargetPath } from './drawing.js';
import { assertIndex } from './operations.js';
import { WORD_NS, REL_NS, CONTENT_TYPES_NS, assertText, children, descendants, setWordValue, wordElement } from './xml.js';

export interface CommentLocation {
  sourcePartPath: string;
  paragraph?: number;
  runs?: number[];
  startParagraph?: number;
  endParagraph?: number;
  order: number;
}

export interface CommentPartBinding {
  sourcePartPath: string;
  commentsPath?: string;
  commentsExtendedPath?: string;
}

type Bound<F> = F extends (ctx: CommentContext, ...args: infer A) => infer R ? (...args: A) => R : never;

export interface CommentContext extends PartAccess, HistoryRecorder<CommentContext> {
  caches: CacheBundle & {
    commentStateCache?: { revision: number; comments: CommentInfo[] };
    commentBindingsCache?: { revision: number; bindings: CommentPartBinding[] };
  };
  mayContainComments: Bound<typeof mayContainComments>;
  commentBindings: Bound<typeof commentBindings>;
  collectCommentLocations: Bound<typeof collectCommentLocations>;
  readCommentBody: Bound<typeof readCommentBody>;
  getAllComments: Bound<typeof getAllComments>;
  getCommentsPartPath: Bound<typeof getCommentsPartPath>;
  getCommentsExtendedPartPath: Bound<typeof getCommentsExtendedPartPath>;
  ensureCommentsParts: Bound<typeof ensureCommentsParts>;
  nextCommentId: Bound<typeof nextCommentId>;
  nextCommentParaId: Bound<typeof nextCommentParaId>;
  locateComment: Bound<typeof locateComment>;
  ensureCommentParaId: Bound<typeof ensureCommentParaId>;
  writeCommentBody: Bound<typeof writeCommentBody>;
  upsertCommentEx: Bound<typeof upsertCommentEx>;
  removeRelationshipTarget: Bound<typeof removeRelationshipTarget>;
  relationshipCount: Bound<typeof relationshipCount>;
  removePartOverride: Bound<typeof removePartOverride>;
  removePartAndOverride: Bound<typeof removePartAndOverride>;
  cleanupCommentParts: Bound<typeof cleanupCommentParts>;
  addCommentDirect: Bound<typeof addCommentDirect>;
  replyCommentDirect: Bound<typeof replyCommentDirect>;
  setCommentResolvedDirect: Bound<typeof setCommentResolvedDirect>;
  setCommentTextDirect: Bound<typeof setCommentTextDirect>;
  deleteCommentDirect: Bound<typeof deleteCommentDirect>;
  boundaryRun(paragraph: Element, offset: number): Element | null;
  buildBlocksFromElement(element: Element, styles: StylesContext, noteNumber: (kind: NoteKind, id: number) => { number: number; marker: string } | null, path: string): DocumentBlock[];
  getStylesContext(): StylesContext;
  normalizeRangeOn(document: Document, range: TextRange): { paragraph: Element; start: number; end: number };
  normalizeDocumentRange(document: Document, range: DocumentRange): DocumentRange;
  splitRunAtOffset(paragraph: Element, offset: number): void;
  blockContainerOf(document: Document): Element;
  preOrderElements(root: Element): Element[];
  ownRuns(paragraph: Element): Element[];
  nearestParagraph(node: Node | null): Element | null;
  paragraphAt(document: Document, index: number): Element;
  properties(parent: Element, name: string): Element;
  property(parent: Element, name: string): Element;
  appendText(parent: Element, text: string, before?: Node | null): void;
  basename(path: string): string;
  dirname(path: string): string;
  relsPath(path: string): string;
  encodeXml(xml: string): Uint8Array;
  decodeXml(bytes: Uint8Array): string;
  normalizeReviewerFilterAuthors(authors: string[]): Set<string>;
}

function commentReferenceInRun(run: Element): number | null {
  const reference = children(run, 'commentReference')[0];
  if (!reference) return null;
  const id = Number(reference.getAttributeNS(WORD_NS, 'id') ?? reference.getAttribute('w:id'));
  return Number.isSafeInteger(id) && id >= 0 ? id : null;
}

function cloneCommentInfo(comment: CommentInfo): CommentInfo {
  return JSON.parse(JSON.stringify(comment)) as CommentInfo;
}

function normalizeCommentDate(value: string | undefined): string | undefined {
  if (!value) return undefined;
  return Number.isFinite(Date.parse(value)) ? value : undefined;
}

export const COMMENTS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments';
export const COMMENTS_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml';
export const COMMENTS_EXTENDED_REL = 'http://schemas.microsoft.com/office/2011/relationships/commentsExtended';
export const COMMENTS_EXTENDED_TYPE = 'application/vnd.ms-word.commentsExtended+xml';
export const W14_NS = 'http://schemas.microsoft.com/office/word/2010/wordml';
export const W15_NS = 'http://schemas.microsoft.com/office/word/2012/wordml';

export interface ParsedCommentEntry {
  id: number;
  element: Element;
  paraId?: string;
}

export interface ParsedCommentExEntry {
  paraId: string;
  paraIdParent?: string;
  done?: boolean;
  element: Element;
}

export function defaultCommentsXml(): string {
  return `<w:comments xmlns:w="${WORD_NS}" xmlns:w14="${W14_NS}" xmlns:w15="${W15_NS}"/>`;
}

export function defaultCommentsExtendedXml(): string {
  return `<w15:commentsEx xmlns:w15="${W15_NS}"/>`;
}

export function commentReferenceStyle(): 'CommentReference' {
  return 'CommentReference';
}

export function commentParagraphStyle(): 'CommentText' {
  return 'CommentText';
}

export function parseCommentEntries(document: Document | null): ParsedCommentEntry[] {
  if (!document?.documentElement || document.documentElement.namespaceURI !== WORD_NS || document.documentElement.localName !== 'comments') {
    return [];
  }
  return children(document.documentElement, 'comment')
    .map((element): ParsedCommentEntry | null => {
      const id = Number(element.getAttributeNS(WORD_NS, 'id') ?? element.getAttribute('w:id'));
      if (!Number.isSafeInteger(id) || id < 0) return null;
      const paragraphs = children(element, 'p');
      const lastParagraph = paragraphs.at(-1);
      const paraId = lastParagraph?.getAttributeNS(W14_NS, 'paraId') ?? lastParagraph?.getAttribute('w14:paraId') ?? undefined;
      return { id, element, paraId };
    })
    .filter((entry): entry is ParsedCommentEntry => Boolean(entry));
}

export function parseCommentExEntries(document: Document | null): ParsedCommentExEntry[] {
  if (!document?.documentElement || document.documentElement.namespaceURI !== W15_NS || document.documentElement.localName !== 'commentsEx') {
    return [];
  }
  return children(document.documentElement, 'commentEx', W15_NS)
    .map((element): ParsedCommentExEntry | null => {
      const paraId = element.getAttributeNS(W15_NS, 'paraId') ?? element.getAttribute('w15:paraId') ?? '';
      if (!/^[0-9a-f]{8}$/i.test(paraId)) return null;
      const parent = element.getAttributeNS(W15_NS, 'paraIdParent') ?? element.getAttribute('w15:paraIdParent') ?? undefined;
      const doneValue = element.getAttributeNS(W15_NS, 'done') ?? element.getAttribute('w15:done') ?? undefined;
      return {
        paraId: paraId.toUpperCase(),
        paraIdParent: /^[0-9a-f]{8}$/i.test(parent ?? '') ? parent!.toUpperCase() : undefined,
        done: doneValue === undefined ? undefined : ['1', 'true', 'on'].includes(doneValue),
        element,
      };
    })
    .filter((entry): entry is ParsedCommentExEntry => Boolean(entry));
}

export function ensureCommentParagraphParaId(paragraph: Element, paraId: string): void {
  paragraph.setAttributeNS(W14_NS, 'w14:paraId', paraId.toUpperCase());
}

export function makeCommentAnnotationRun(document: Document): Element {
  const run = wordElement(document, 'r');
  const props = wordElement(document, 'rPr');
  const style = wordElement(document, 'rStyle');
  style.setAttributeNS(WORD_NS, 'w:val', commentReferenceStyle());
  props.appendChild(style);
  run.appendChild(props);
  run.appendChild(wordElement(document, 'annotationRef'));
  return run;
}

export function mayContainComments(ctx: CommentContext, path: string): boolean {
    const bytes = ctx.readPart(path);
    if (!bytes) return false;
    try {
      const xml = ctx.decodeXml(bytes);
      return /commentRange(Start|End)|commentReference|<w:comment\b/i.test(xml);
    } catch {
      return false;
    }
}

export function commentBindings(ctx: CommentContext): CommentPartBinding[] {
    if (ctx.caches.commentBindingsCache?.revision === ctx.revision) return ctx.caches.commentBindingsCache.bindings.map((binding) => ({ ...binding }));
    const bindings = ctx.contentPartPaths()
      .map((sourcePartPath) => ({
        sourcePartPath,
        commentsPath: ctx.relatedPartPathFor(sourcePartPath, COMMENTS_REL),
        commentsExtendedPath: ctx.relatedPartPathFor(sourcePartPath, COMMENTS_EXTENDED_REL),
      }))
      .filter((binding) => binding.commentsPath || binding.commentsExtendedPath || ctx.mayContainComments(binding.sourcePartPath));
    ctx.caches.commentBindingsCache = { revision: ctx.revision, bindings };
    return bindings.map((binding) => ({ ...binding }));
}

export function collectCommentLocations(ctx: CommentContext, sourcePartPath: string): Map<number, CommentLocation> {
    const document = ctx.partDocumentOrUndefined(sourcePartPath);
    if (!document) return new Map();
    let container: Element;
    try {
      container = ctx.blockContainerOf(document);
    } catch {
      return new Map();
    }
    const paragraphs = descendants(container, 'p');
    const paragraphIndex = new Map(paragraphs.map((paragraph, index) => [paragraph, index]));
    const order = ctx.preOrderElements(container);
    const orderIndex = new Map(order.map((element, index) => [element, index]));
    const starts = new Map<number, Element>();
    const ends = new Map<number, Element>();
    const refs = new Map<number, { paragraph: number; run: number }[]>();
    for (const start of descendants(container, 'commentRangeStart')) {
      const id = Number(start.getAttributeNS(WORD_NS, 'id') ?? start.getAttribute('w:id'));
      if (Number.isSafeInteger(id) && id >= 0 && !starts.has(id)) starts.set(id, start);
    }
    for (const end of descendants(container, 'commentRangeEnd')) {
      const id = Number(end.getAttributeNS(WORD_NS, 'id') ?? end.getAttribute('w:id'));
      if (Number.isSafeInteger(id) && id >= 0) ends.set(id, end);
    }
    for (const [paragraph, paragraphNumber] of paragraphIndex.entries()) {
      for (const [runNumber, run] of ctx.ownRuns(paragraph).entries()) {
        const id = commentReferenceInRun(run);
        if (id === null) continue;
        const list = refs.get(id) ?? [];
        list.push({ paragraph: paragraphNumber, run: runNumber });
        refs.set(id, list);
      }
    }
    const result = new Map<number, CommentLocation>();
    const ids = new Set<number>([...starts.keys(), ...ends.keys(), ...refs.keys()]);
    for (const id of ids) {
      const start = starts.get(id);
      const end = ends.get(id);
      if (start && end) {
        const startParagraph = ctx.nearestParagraph(start);
        const endParagraph = ctx.nearestParagraph(end);
        const startParagraphNumber = startParagraph ? paragraphIndex.get(startParagraph) : undefined;
        const endParagraphNumber = endParagraph ? paragraphIndex.get(endParagraph) : undefined;
        const startOrder = orderIndex.get(start) ?? Number.MAX_SAFE_INTEGER;
        const endOrder = orderIndex.get(end) ?? Number.MAX_SAFE_INTEGER;
        if (startParagraph && endParagraph && startParagraphNumber !== undefined && endParagraphNumber !== undefined) {
          if (startParagraph === endParagraph) {
            const runs = ctx.ownRuns(startParagraph).flatMap((run, runIndex) => {
              const position = orderIndex.get(run) ?? -1;
              return position > startOrder && position < endOrder ? [runIndex] : [];
            });
            if (runs.length) {
              result.set(id, { sourcePartPath, paragraph: startParagraphNumber, runs, order: startOrder });
              continue;
            }
          }
          result.set(id, {
            sourcePartPath,
            startParagraph: startParagraphNumber,
            endParagraph: endParagraphNumber,
            order: startOrder,
          });
          continue;
        }
      }
      const references = refs.get(id)?.slice().sort((a, b) => a.paragraph - b.paragraph || a.run - b.run) ?? [];
      if (!references.length) continue;
      const first = references[0]!;
      const last = references[references.length - 1]!;
      if (references.every((entry) => entry.paragraph === first.paragraph)) {
        result.set(id, {
          sourcePartPath,
          paragraph: first.paragraph,
          runs: [...new Set(references.map((entry) => entry.run))],
          order: first.paragraph * 10_000 + first.run,
        });
      } else {
        result.set(id, {
          sourcePartPath,
          startParagraph: first.paragraph,
          endParagraph: last.paragraph,
          order: first.paragraph * 10_000 + first.run,
        });
      }
    }
    return result;
}

export function readCommentBody(ctx: CommentContext, entry: ReturnType<typeof parseCommentEntries>[number], commentsPath: string): { blocks: DocumentBlock[]; text: string } {
    const blocks = ctx.buildBlocksFromElement(entry.element, ctx.getStylesContext(), () => null, commentsPath);
    const text = blocks
      .flatMap((block) => block.type === 'paragraph' ? [block.paragraph.text] : [])
      .join('\n')
      .trim();
    return { blocks, text };
}

export function getAllComments(ctx: CommentContext): CommentInfo[] {
    if (ctx.caches.commentStateCache?.revision === ctx.revision) return ctx.caches.commentStateCache.comments;
    const bindings = ctx.commentBindings();
    if (!bindings.length) {
      ctx.caches.commentStateCache = { revision: ctx.revision, comments: [] };
      return [];
    }
    const locations = new Map<number, CommentLocation>();
    for (const binding of bindings) {
      for (const [id, location] of ctx.collectCommentLocations(binding.sourcePartPath)) {
        const previous = locations.get(id);
        if (!previous || location.order < previous.order) locations.set(id, location);
      }
    }
    const entries = new Map<number, { entry: ReturnType<typeof parseCommentEntries>[number]; commentsPath: string }>();
    const exEntries = new Map<string, ReturnType<typeof parseCommentExEntries>[number]>();
    const parsedCommentsPaths = new Set<string>();
    const parsedExtendedPaths = new Set<string>();
    for (const binding of bindings) {
      if (binding.commentsPath && !parsedCommentsPaths.has(binding.commentsPath) && ctx.hasPart(binding.commentsPath)) {
        parsedCommentsPaths.add(binding.commentsPath);
        for (const entry of parseCommentEntries(ctx.partDocumentOrUndefined(binding.commentsPath) ?? null)) {
          if (!entries.has(entry.id)) entries.set(entry.id, { entry, commentsPath: binding.commentsPath });
        }
      }
      if (binding.commentsExtendedPath && !parsedExtendedPaths.has(binding.commentsExtendedPath) && ctx.hasPart(binding.commentsExtendedPath)) {
        parsedExtendedPaths.add(binding.commentsExtendedPath);
        for (const entry of parseCommentExEntries(ctx.partDocumentOrUndefined(binding.commentsExtendedPath) ?? null)) exEntries.set(entry.paraId, entry);
      }
    }
    const paraToId = new Map<string, number>();
    for (const [id, { entry }] of entries) {
      if (entry.paraId) paraToId.set(entry.paraId.toUpperCase(), id);
    }
    const comments = new Map<number, CommentInfo>();
    const ids = new Set<number>([...locations.keys(), ...entries.keys()]);
    for (const id of ids) {
      const location = locations.get(id);
      const entryInfo = entries.get(id);
      const entry = entryInfo?.entry;
      const extended = entry?.paraId ? exEntries.get(entry.paraId.toUpperCase()) : undefined;
      const body = entry && entryInfo ? ctx.readCommentBody(entry, entryInfo.commentsPath) : undefined;
      const anchor = location
        ? (location.runs
            ? { sourcePartPath: location.sourcePartPath, paragraph: location.paragraph!, runs: [...location.runs] }
            : { sourcePartPath: location.sourcePartPath, startParagraph: location.startParagraph!, endParagraph: location.endParagraph! }) satisfies CommentAnchor
        : undefined;
      comments.set(id, {
        id,
        author: entry?.element.getAttributeNS(WORD_NS, 'author') ?? entry?.element.getAttribute('w:author') ?? undefined,
        initials: entry?.element.getAttributeNS(WORD_NS, 'initials') ?? entry?.element.getAttribute('w:initials') ?? undefined,
        date: normalizeCommentDate(entry?.element.getAttributeNS(WORD_NS, 'date') ?? entry?.element.getAttribute('w:date') ?? undefined),
        text: body?.text ?? '',
        blocks: body?.blocks,
        anchor,
        parentId: extended?.paraIdParent ? paraToId.get(extended.paraIdParent) : undefined,
        resolved: extended?.done,
        isOrphan: !entry || !anchor,
      });
    }
    for (const comment of comments.values()) {
      if (!comment.anchor && comment.parentId !== undefined) {
        const parent = comments.get(comment.parentId);
        if (parent?.anchor) {
          comment.anchor = cloneCommentInfo({ ...parent, blocks: undefined }).anchor;
          comment.isOrphan = false;
        }
      }
    }
    const ordered = [...comments.values()].sort((a, b) => {
      const aLocation = locations.get(a.id)?.order ?? Number.MAX_SAFE_INTEGER;
      const bLocation = locations.get(b.id)?.order ?? Number.MAX_SAFE_INTEGER;
      return aLocation - bLocation || a.id - b.id;
    });
    ctx.caches.commentStateCache = { revision: ctx.revision, comments: ordered };
    return ordered;
}

export function getCommentsPartPath(ctx: CommentContext, sourcePartPath = ctx.mainPath): string | undefined {
    const conventional = `${ctx.dirname(sourcePartPath) ? `${ctx.dirname(sourcePartPath)}/` : ''}comments.xml`;
    return ctx.relatedPartPathFor(sourcePartPath, COMMENTS_REL) ?? (ctx.hasPart(conventional) ? conventional : undefined);
}

export function getCommentsExtendedPartPath(ctx: CommentContext, sourcePartPath = ctx.mainPath): string | undefined {
    const conventional = `${ctx.dirname(sourcePartPath) ? `${ctx.dirname(sourcePartPath)}/` : ''}commentsExtended.xml`;
    return ctx.relatedPartPathFor(sourcePartPath, COMMENTS_EXTENDED_REL) ?? (ctx.hasPart(conventional) ? conventional : undefined);
}

export function getComments(ctx: CommentContext, filter: { authors?: string[]; resolved?: boolean } = {}): CommentInfo[] {
    if (!filter || typeof filter !== 'object' || Array.isArray(filter)) throw new Error('filter must be an object.');
    if ('authors' in filter && filter.authors !== undefined) {
      if (!Array.isArray(filter.authors)) throw new Error('filter.authors must be an array.');
      for (const author of filter.authors) assertText(author, 'filter.authors[]');
    }
    if ('resolved' in filter && filter.resolved !== undefined && typeof filter.resolved !== 'boolean') {
      throw new Error('filter.resolved must be boolean.');
    }
    let comments = ctx.getAllComments();
    if (filter.authors?.length) {
      const authors = ctx.normalizeReviewerFilterAuthors(filter.authors);
      comments = comments.filter((comment) => authors.has(reviewerBucketKey(reviewerBucketOf(comment.author))));
    }
    if (filter.resolved !== undefined) comments = comments.filter((comment) => comment.resolved === filter.resolved);
    return comments.map(cloneCommentInfo);
}

export function ensureCommentsParts(ctx: CommentContext, sourcePartPath: string, includeExtended = false): { commentsPath: string; commentsExtendedPath?: string } {
    const sourceBase = ctx.basename(sourcePartPath).replace(/\.xml$/i, '');
    let commentsPath = ctx.getCommentsPartPath(sourcePartPath);
    if (!commentsPath) {
      // Keep new comment parts next to the source part; non-main fallback names stay source-specific
      // so future cleanup/interoperability logic can reconstruct the relationship target deterministically.
      commentsPath = sourcePartPath === ctx.mainPath
        ? `${ctx.dirname(sourcePartPath) ? `${ctx.dirname(sourcePartPath)}/` : ''}comments.xml`
        : `${ctx.dirname(sourcePartPath) ? `${ctx.dirname(sourcePartPath)}/` : ''}${sourceBase}-comments.xml`;
    }
    if (!ctx.hasPart(commentsPath)) ctx.addPart(commentsPath, ctx.encodeXml(defaultCommentsXml()), COMMENTS_TYPE);
    ctx.ensurePartRelationship(sourcePartPath, COMMENTS_REL, commentsPath);
    let commentsExtendedPath: string | undefined;
    if (includeExtended) {
      commentsExtendedPath = ctx.getCommentsExtendedPartPath(sourcePartPath);
      if (!commentsExtendedPath) {
        commentsExtendedPath = sourcePartPath === ctx.mainPath
          ? `${ctx.dirname(sourcePartPath) ? `${ctx.dirname(sourcePartPath)}/` : ''}commentsExtended.xml`
          : `${ctx.dirname(sourcePartPath) ? `${ctx.dirname(sourcePartPath)}/` : ''}${sourceBase}-commentsExtended.xml`;
      }
      if (!ctx.hasPart(commentsExtendedPath)) {
        ctx.addPart(commentsExtendedPath, ctx.encodeXml(defaultCommentsExtendedXml()), COMMENTS_EXTENDED_TYPE);
      }
      ctx.ensurePartRelationship(sourcePartPath, COMMENTS_EXTENDED_REL, commentsExtendedPath);
    }
    return { commentsPath, commentsExtendedPath };
}

export function nextCommentId(ctx: CommentContext): number {
    const max = ctx.getAllComments().reduce((current, comment) => Math.max(current, comment.id), 0);
    return max + 1;
}

export function nextCommentParaId(ctx: CommentContext): string {
    const used = new Set<number>();
    for (const binding of ctx.commentBindings()) {
      if (binding.commentsPath && ctx.hasPart(binding.commentsPath)) {
        for (const entry of parseCommentEntries(ctx.partDocumentOrUndefined(binding.commentsPath) ?? null)) {
          if (entry.paraId && /^[0-9a-f]{8}$/i.test(entry.paraId)) used.add(Number.parseInt(entry.paraId, 16));
        }
      }
      if (binding.commentsExtendedPath && ctx.hasPart(binding.commentsExtendedPath)) {
        for (const entry of parseCommentExEntries(ctx.partDocumentOrUndefined(binding.commentsExtendedPath) ?? null)) {
          used.add(Number.parseInt(entry.paraId, 16));
          if (entry.paraIdParent) used.add(Number.parseInt(entry.paraIdParent, 16));
        }
      }
    }
    let next = 1;
    while (used.has(next)) next++;
    return next.toString(16).toUpperCase().padStart(8, '0');
}

export function locateComment(ctx: CommentContext, id: number): { sourcePartPath: string; commentsPath: string; commentsExtendedPath?: string; entry: ReturnType<typeof parseCommentEntries>[number] } {
    assertIndex(id);
    const parsed = new Set<string>();
    for (const binding of ctx.commentBindings()) {
      if (!binding.commentsPath || parsed.has(binding.commentsPath) || !ctx.hasPart(binding.commentsPath)) continue;
      parsed.add(binding.commentsPath);
      const entry = parseCommentEntries(ctx.partDocumentOrUndefined(binding.commentsPath) ?? null).find((item) => item.id === id);
      if (entry) return { sourcePartPath: binding.sourcePartPath, commentsPath: binding.commentsPath, commentsExtendedPath: binding.commentsExtendedPath, entry };
    }
    throw new Error(`Comment ${id} does not exist.`);
}

export function ensureCommentParaId(ctx: CommentContext, commentsPath: string, id: number): string {
    let resolved = '';
    ctx.updatePartXml(commentsPath, (document) => {
      const entry = parseCommentEntries(document).find((item) => item.id === id);
      if (!entry) throw new Error(`Comment ${id} does not exist.`);
      if (entry.paraId) {
        resolved = entry.paraId.toUpperCase();
        return;
      }
      const paragraphs = children(entry.element, 'p');
      const paragraph = paragraphs.at(-1) ?? wordElement(document, 'p');
      if (!paragraph.parentNode) entry.element.appendChild(paragraph);
      resolved = ctx.nextCommentParaId();
      ensureCommentParagraphParaId(paragraph, resolved);
    });
    return resolved;
}

export function writeCommentBody(ctx: CommentContext, commentElement: Element, paraId: string, text: string): void {
    const document = commentElement.ownerDocument!;
    while (commentElement.firstChild) commentElement.removeChild(commentElement.firstChild);
    const paragraph = wordElement(document, 'p');
    ensureCommentParagraphParaId(paragraph, paraId);
    const paragraphProps = ctx.properties(paragraph, 'pPr');
    setWordValue(ctx.property(paragraphProps, 'pStyle'), commentParagraphStyle());
    paragraph.appendChild(makeCommentAnnotationRun(document));
    if (text) {
      const run = wordElement(document, 'r');
      const props = ctx.properties(run, 'rPr');
      setWordValue(ctx.property(props, 'rStyle'), commentParagraphStyle());
      ctx.appendText(run, text);
      paragraph.appendChild(run);
    }
    commentElement.appendChild(paragraph);
}

export function upsertCommentEx(ctx: CommentContext, commentsExtendedPath: string, paraId: string, patch: { parentParaId?: string | null; done?: boolean }): void {
    ctx.updatePartXml(commentsExtendedPath, (document) => {
      const root = document.documentElement!;
      let entry = parseCommentExEntries(document).find((item) => item.paraId === paraId)?.element;
      if (!entry) {
        entry = document.createElementNS(W15_NS, 'w15:commentEx');
        entry.setAttributeNS(W15_NS, 'w15:paraId', paraId);
        root.appendChild(entry);
      }
      if (patch.parentParaId === null) {
        entry.removeAttributeNS(W15_NS, 'paraIdParent');
        entry.removeAttribute('w15:paraIdParent');
      } else if (patch.parentParaId) {
        entry.setAttributeNS(W15_NS, 'w15:paraIdParent', patch.parentParaId);
      }
      if (patch.done !== undefined) entry.setAttributeNS(W15_NS, 'w15:done', patch.done ? '1' : '0');
    });
}

export function removeRelationshipTarget(ctx: CommentContext, sourcePartPath: string, relationType: string, targetPath: string): void {
    const relationshipPath = ctx.relsPath(sourcePartPath);
    if (!ctx.hasPart(relationshipPath)) return;
    ctx.updatePartXml(relationshipPath, (document) => {
      for (const relation of children(document.documentElement!, 'Relationship', REL_NS)) {
        if (relation.getAttribute('Type') !== relationType || relation.getAttribute('TargetMode') === 'External') continue;
        const target = relation.getAttribute('Target');
        if (!target) continue;
        let resolved: string | undefined;
        try {
          resolved = resolveTargetPath(sourcePartPath, decodeURIComponent(target));
        } catch {
          resolved = undefined;
        }
        if (resolved === targetPath) relation.parentNode?.removeChild(relation);
      }
    });
}

export function relationshipCount(ctx: CommentContext, targetPath: string, relationType: string): number {
    let count = 0;
    for (const sourcePartPath of ctx.contentPartPaths()) {
      const relationshipPath = ctx.relsPath(sourcePartPath);
      if (!ctx.hasPart(relationshipPath)) continue;
      const document = ctx.partDocumentOrUndefined(relationshipPath)?.documentElement;
      if (!document) continue;
      for (const relation of children(document, 'Relationship', REL_NS)) {
        if (relation.getAttribute('Type') !== relationType || relation.getAttribute('TargetMode') === 'External') continue;
        const target = relation.getAttribute('Target');
        if (!target) continue;
        let resolved: string | undefined;
        try {
          resolved = resolveTargetPath(sourcePartPath, decodeURIComponent(target));
        } catch {
          resolved = undefined;
        }
        if (resolved === targetPath) count++;
      }
    }
    return count;
}

export function cleanupCommentParts(ctx: CommentContext, commentsPath: string, commentsExtendedPath?: string): void {
    const hasComments = ctx.hasPart(commentsPath) && parseCommentEntries(ctx.partDocumentOrUndefined(commentsPath) ?? null).length > 0;
    if (!hasComments) {
      for (const binding of ctx.commentBindings()) {
        ctx.removeRelationshipTarget(binding.sourcePartPath, COMMENTS_REL, commentsPath);
      }
      if (ctx.relationshipCount(commentsPath, COMMENTS_REL) === 0) ctx.removePartAndOverride(commentsPath);
    }
    if (commentsExtendedPath && ctx.hasPart(commentsExtendedPath) && parseCommentExEntries(ctx.partDocumentOrUndefined(commentsExtendedPath) ?? null).length === 0) {
      for (const binding of ctx.commentBindings()) {
        ctx.removeRelationshipTarget(binding.sourcePartPath, COMMENTS_EXTENDED_REL, commentsExtendedPath);
      }
      if (ctx.relationshipCount(commentsExtendedPath, COMMENTS_EXTENDED_REL) === 0) ctx.removePartAndOverride(commentsExtendedPath);
    }
}

export function addComment(ctx: CommentContext, range: TextRange | DocumentRange, comment: { author?: string; initials?: string; text: string }): number {
    // TextRange/DocumentRange use the main-document paragraph namespace in this API surface.
    // Header/footer/footnote/endnote comment anchors are readable via getComments(), but write APIs
    // currently insert anchors only into the main document part.
    return ctx.withDraft((draft) => draft.addCommentDirect(range, comment));
}

export function addCommentDirect(ctx: CommentContext, range: TextRange | DocumentRange, comment: { author?: string; initials?: string; text: string }): number {
    if (comment.author !== undefined) assertText(comment.author, 'comment.author');
    if (comment.initials !== undefined) assertText(comment.initials, 'comment.initials');
    assertText(comment.text, 'comment.text');
    const preview = ctx.getPartDocument(ctx.mainPath);
    const normalized = 'paragraph' in range
      ? (() => {
          const current = ctx.normalizeRangeOn(preview, range);
          return {
            startParagraph: range.paragraph,
            startOffset: current.start,
            endParagraph: range.paragraph,
            endOffset: current.end,
          };
        })()
      : (() => {
          const current = ctx.normalizeDocumentRange(preview, range);
          return {
            startParagraph: current.start.paragraph,
            startOffset: current.start.offset,
            endParagraph: current.end.paragraph,
            endOffset: current.end.offset,
          };
        })();
    const id = ctx.nextCommentId();
    const paraId = ctx.nextCommentParaId();
    const date = new Date().toISOString();
    ctx.updatePartXml(ctx.mainPath, (document) => {
      const startParagraph = ctx.paragraphAt(document, normalized.startParagraph);
      const endParagraph = ctx.paragraphAt(document, normalized.endParagraph);
      const collapsed = normalized.startParagraph === normalized.endParagraph && normalized.startOffset === normalized.endOffset;
      if (startParagraph === endParagraph) {
        ctx.splitRunAtOffset(startParagraph, normalized.endOffset);
        ctx.splitRunAtOffset(startParagraph, normalized.startOffset);
      } else {
        ctx.splitRunAtOffset(endParagraph, normalized.endOffset);
        ctx.splitRunAtOffset(startParagraph, normalized.startOffset);
      }
      const startMarker = wordElement(document, 'commentRangeStart');
      startMarker.setAttributeNS(WORD_NS, 'w:id', String(id));
      const startBoundary = ctx.boundaryRun(startParagraph, normalized.startOffset);
      if (startBoundary?.parentNode) startBoundary.parentNode.insertBefore(startMarker, startBoundary);
      else startParagraph.insertBefore(startMarker, startParagraph.firstChild);

      const endMarker = wordElement(document, 'commentRangeEnd');
      endMarker.setAttributeNS(WORD_NS, 'w:id', String(id));
      const referenceRun = wordElement(document, 'r');
      const props = ctx.properties(referenceRun, 'rPr');
      setWordValue(ctx.property(props, 'rStyle'), commentReferenceStyle());
      const reference = wordElement(document, 'commentReference');
      reference.setAttributeNS(WORD_NS, 'w:id', String(id));
      referenceRun.appendChild(reference);
      if (collapsed) {
        const parent = startMarker.parentNode;
        if (!parent) throw new Error('Comment anchor parent is missing.');
        if (startMarker.nextSibling) parent.insertBefore(endMarker, startMarker.nextSibling);
        else parent.appendChild(endMarker);
        if (endMarker.nextSibling) parent.insertBefore(referenceRun, endMarker.nextSibling);
        else parent.appendChild(referenceRun);
      } else {
        const endBoundary = ctx.boundaryRun(endParagraph, normalized.endOffset);
        if (endBoundary?.parentNode) endBoundary.parentNode.insertBefore(endMarker, endBoundary);
        else endParagraph.appendChild(endMarker);
        if (endBoundary?.parentNode) endBoundary.parentNode.insertBefore(referenceRun, endBoundary);
        else endParagraph.appendChild(referenceRun);
      }
    });
    const { commentsPath, commentsExtendedPath } = ctx.ensureCommentsParts(ctx.mainPath, true);
    ctx.updatePartXml(commentsPath, (document) => {
      const root = document.documentElement!;
      const commentElement = wordElement(document, 'comment');
      commentElement.setAttributeNS(WORD_NS, 'w:id', String(id));
      if (comment.author) commentElement.setAttributeNS(WORD_NS, 'w:author', comment.author);
      if (comment.initials) commentElement.setAttributeNS(WORD_NS, 'w:initials', comment.initials);
      commentElement.setAttributeNS(WORD_NS, 'w:date', date);
      ctx.writeCommentBody(commentElement, paraId, comment.text);
      root.appendChild(commentElement);
    });
    ctx.upsertCommentEx(commentsExtendedPath!, paraId, { done: false, parentParaId: null });
    return id;
}

export function replyComment(ctx: CommentContext, parentId: number, comment: { author?: string; initials?: string; text: string }): number {
    return ctx.withDraft((draft) => draft.replyCommentDirect(parentId, comment));
}

export function replyCommentDirect(ctx: CommentContext, parentId: number, comment: { author?: string; initials?: string; text: string }): number {
    if (comment.author !== undefined) assertText(comment.author, 'comment.author');
    if (comment.initials !== undefined) assertText(comment.initials, 'comment.initials');
    assertText(comment.text, 'comment.text');
    const parent = ctx.locateComment(parentId);
    const parentParaId = ctx.ensureCommentParaId(parent.commentsPath, parentId);
    const id = ctx.nextCommentId();
    const paraId = ctx.nextCommentParaId();
    const { commentsPath, commentsExtendedPath } = ctx.ensureCommentsParts(parent.sourcePartPath, true);
    ctx.updatePartXml(commentsPath, (document) => {
      const root = document.documentElement!;
      const commentElement = wordElement(document, 'comment');
      commentElement.setAttributeNS(WORD_NS, 'w:id', String(id));
      if (comment.author) commentElement.setAttributeNS(WORD_NS, 'w:author', comment.author);
      if (comment.initials) commentElement.setAttributeNS(WORD_NS, 'w:initials', comment.initials);
      commentElement.setAttributeNS(WORD_NS, 'w:date', new Date().toISOString());
      ctx.writeCommentBody(commentElement, paraId, comment.text);
      root.appendChild(commentElement);
    });
    ctx.upsertCommentEx(commentsExtendedPath!, paraId, { parentParaId, done: false });
    return id;
}

export function setCommentResolved(ctx: CommentContext, id: number, resolved: boolean): void {
    ctx.withDraft((draft) => draft.setCommentResolvedDirect(id, resolved));
}

export function setCommentResolvedDirect(ctx: CommentContext, id: number, resolved: boolean): void {
    assertIndex(id);
    if (typeof resolved !== 'boolean') throw new Error('resolved must be boolean.');
    const target = ctx.locateComment(id);
    const paraId = ctx.ensureCommentParaId(target.commentsPath, id);
    const { commentsExtendedPath } = ctx.ensureCommentsParts(target.sourcePartPath, true);
    ctx.upsertCommentEx(commentsExtendedPath!, paraId, { done: resolved });
}

export function setCommentText(ctx: CommentContext, id: number, text: string): void {
    ctx.withDraft((draft) => draft.setCommentTextDirect(id, text));
}

export function setCommentTextDirect(ctx: CommentContext, id: number, text: string): void {
    assertIndex(id);
    assertText(text, 'text');
    const target = ctx.locateComment(id);
    const paraId = ctx.ensureCommentParaId(target.commentsPath, id);
    ctx.updatePartXml(target.commentsPath, (document) => {
      const entry = parseCommentEntries(document).find((item) => item.id === id);
      if (!entry) throw new Error(`Comment ${id} does not exist.`);
      ctx.writeCommentBody(entry.element, paraId, text);
    });
}

export function deleteComment(ctx: CommentContext, id: number, options: { withReplies?: boolean } = {}): void {
    ctx.withDraft((draft) => draft.deleteCommentDirect(id, options));
}

export function deleteCommentDirect(ctx: CommentContext, id: number, options: { withReplies?: boolean } = {}): void {
    assertIndex(id);
    const withReplies = options.withReplies ?? true;
    const allComments = ctx.getAllComments();
    if (!allComments.some((comment) => comment.id === id)) throw new Error(`Comment ${id} does not exist.`);
    const deleteIds = new Set<number>([id]);
    if (withReplies) {
      let changed = true;
      while (changed) {
        changed = false;
        for (const comment of allComments) {
          if (comment.parentId !== undefined && deleteIds.has(comment.parentId) && !deleteIds.has(comment.id)) {
            deleteIds.add(comment.id);
            changed = true;
          }
        }
      }
    }
    const byCommentsPath = new Map<string, number[]>();
    const paraIds = new Map<number, { commentsPath: string; commentsExtendedPath?: string; paraId: string; sourcePartPath: string }>();
    for (const commentId of deleteIds) {
      const target = ctx.locateComment(commentId);
      const paraId = ctx.ensureCommentParaId(target.commentsPath, commentId);
      const list = byCommentsPath.get(target.commentsPath) ?? [];
      list.push(commentId);
      byCommentsPath.set(target.commentsPath, list);
      paraIds.set(commentId, { commentsPath: target.commentsPath, commentsExtendedPath: target.commentsExtendedPath, paraId, sourcePartPath: target.sourcePartPath });
    }
    for (const [commentsPath, ids] of byCommentsPath) {
      ctx.updatePartXml(commentsPath, (document) => {
        for (const entry of parseCommentEntries(document)) {
          if (ids.includes(entry.id)) entry.element.parentNode?.removeChild(entry.element);
        }
      });
    }
    const extendedPaths = new Set<string>();
    for (const info of paraIds.values()) {
      if (!info.commentsExtendedPath || extendedPaths.has(info.commentsExtendedPath) || !ctx.hasPart(info.commentsExtendedPath)) continue;
      extendedPaths.add(info.commentsExtendedPath);
      const deletedParaIds = new Set([...paraIds.values()]
        .filter((entry) => entry.commentsExtendedPath === info.commentsExtendedPath)
        .map((entry) => entry.paraId));
      ctx.updatePartXml(info.commentsExtendedPath, (document) => {
        for (const entry of parseCommentExEntries(document)) {
          if (deletedParaIds.has(entry.paraId)) entry.element.parentNode?.removeChild(entry.element);
        }
      });
    }
    const affectedSourceParts = new Set([...paraIds.values()].map((info) => info.sourcePartPath));
    for (const sourcePartPath of affectedSourceParts) {
      ctx.updatePartXml(sourcePartPath, (document) => {
        const container = ctx.blockContainerOf(document);
        for (const nodeName of ['commentRangeStart', 'commentRangeEnd'] as const) {
          for (const node of descendants(container, nodeName)) {
            const commentId = Number(node.getAttributeNS(WORD_NS, 'id') ?? node.getAttribute('w:id'));
            if (deleteIds.has(commentId)) node.parentNode?.removeChild(node);
          }
        }
        for (const paragraph of descendants(container, 'p')) {
          for (const run of ctx.ownRuns(paragraph)) {
            const commentId = commentReferenceInRun(run);
            if (commentId === null || !deleteIds.has(commentId)) continue;
            const reference = children(run, 'commentReference')[0];
            if (reference) run.removeChild(reference);
            const hasMeaningfulContent = Array.from(run.childNodes).some((child) => {
              if (child.nodeType === 1) return (child as Element).localName !== 'rPr';
              return (child.textContent ?? '').trim().length > 0;
            });
            if (!hasMeaningfulContent) run.parentNode?.removeChild(run);
          }
        }
      });
    }
    for (const info of paraIds.values()) ctx.cleanupCommentParts(info.commentsPath, info.commentsExtendedPath);
}

export function removePartOverride(ctx: CommentContext, path: string): void {
    ctx.updatePartXml('[Content_Types].xml', (document) => {
      for (const override of children(document.documentElement!, 'Override', CONTENT_TYPES_NS)) {
        if (override.getAttribute('PartName') === `/${path}`) override.parentNode?.removeChild(override);
      }
    });
}

export function removePartAndOverride(ctx: CommentContext, path: string): void {
    if (!ctx.hasPart(path)) return;
    ctx.deletePart(path);
    ctx.forgetPartDocument(path);
    ctx.forgetDirtyPartXml(path);
    ctx.forgetDirtyPartSize(path);
    ctx.removePartOverride(path);
}
