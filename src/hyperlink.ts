import type { Document, Element, Node } from '@xmldom/xmldom';
import type { HyperlinkInfo } from './types.js';
import type { RelationshipTarget } from './drawing.js';
import type { HistoryRecorder, PartAccess } from './internal/context.js';
import { assertIndex } from './operations.js';
import { OFFICE_REL_NS, REL_NS, WORD_NS, assertText, children, descendants, parseXml, serializeXml, setWordValue, wordElement } from './xml.js';
import { resolveRelationshipsPath } from './drawing.js';
import { textOf } from './content-control.js';
import { bodyOf, encodeXml, isDescendantOfWithin, mainParagraphElements, ownRuns, paragraphAt, property, properties, removeWordAttribute } from './internal/elements.js';

const MAX_LINK_LENGTH = 8_192;
const MAX_ANCHOR_LENGTH = 256;
const MAX_TOOLTIP_LENGTH = 2_048;
const ALLOWED_URL_SCHEMES = new Set(['http', 'https', 'mailto']);
const BLOCKED_URL_SCHEMES = new Set(['javascript', 'data', 'vbscript', 'file']);
export const HYPERLINK_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';

type Bound<F> = F extends (ctx: HyperlinkContext, ...args: infer A) => infer R ? (...args: A) => R : never;

export interface HyperlinkContext
  extends Pick<PartAccess,
    'mainPath' | 'hasPart' | 'addPart' | 'partDocumentCopy' | 'updatePartXmlFromCopy' | 'writePartXml'>,
    HistoryRecorder<HyperlinkContext> {
  relationshipsFor(partPath: string): Map<string, RelationshipTarget>;
  nextRelationshipId(rels: Document): string;
  splitRunAtOffset(paragraph: Element, offset: number): void;
  getHyperlinks: Bound<typeof getHyperlinks>;
  hyperlinkNode: Bound<typeof hyperlinkNode>;
  resolveHyperlink: Bound<typeof resolveHyperlink>;
}

export function relationshipIdOf(hyperlink: Element): string | undefined {
  return hyperlink.getAttributeNS(OFFICE_REL_NS, 'id') ?? hyperlink.getAttribute('r:id') ?? undefined;
}

function fieldInstruction(link: { url?: string; anchor?: string }): string {
  const escapeFieldText = (value: string) => value.replace(/\\/g, '\\\\').replace(/"/g, '\\"');
  const chunks = ['HYPERLINK'];
  if (link.url) chunks.push(`"${escapeFieldText(link.url)}"`);
  if (link.anchor) chunks.push(`\\l "${escapeFieldText(link.anchor)}"`);
  return chunks.join(' ');
}

export function textRangeLength(paragraph: Element, start: number, end: number): void {
  const size = textOf(paragraph).length;
  if (!Number.isSafeInteger(start) || !Number.isSafeInteger(end) || start < 0 || end < start || end > size) {
    throw new Error(`Range [${start}, ${end}) is out of bounds for paragraph text length ${size}.`);
  }
}

export interface HyperlinkTarget {
  url?: string;
  anchor?: string;
  tooltip?: string;
}

export function relationshipPath(partPath: string): string {
  const slash = partPath.lastIndexOf('/');
  const dir = slash === -1 ? '' : `${partPath.slice(0, slash + 1)}`;
  const file = slash === -1 ? partPath : partPath.slice(slash + 1);
  return `${dir}_rels/${file}.rels`;
}

function urlScheme(url: string): string | null {
  const trimmed = url.trim();
  if (!trimmed) return null;
  const match = trimmed.match(/^([a-z][a-z\d+\-.]*):/i);
  return match ? match[1]!.toLowerCase() : null;
}

export function isSafeHyperlinkUrl(url: string): boolean {
  const scheme = urlScheme(url);
  return scheme !== null && ALLOWED_URL_SCHEMES.has(scheme);
}

export function isUnsafeHyperlinkUrl(url: string): boolean {
  const scheme = urlScheme(url);
  return scheme !== null && BLOCKED_URL_SCHEMES.has(scheme);
}

export function isUnsafeHyperlink(link: { url?: string; anchor?: string }): boolean {
  if (!link.url) return false;
  return !isSafeHyperlinkUrl(link.url);
}

export function assertHyperlinkInput(link: HyperlinkTarget): void {
  if (!link || typeof link !== 'object') throw new Error('link must be an object.');
  const { url, anchor, tooltip } = link;
  if (url !== undefined && typeof url !== 'string') throw new Error('link.url must be a string.');
  if (anchor !== undefined && typeof anchor !== 'string') throw new Error('link.anchor must be a string.');
  if (tooltip !== undefined && typeof tooltip !== 'string') throw new Error('link.tooltip must be a string.');
  if (!url && !anchor) throw new Error('link.url or link.anchor is required.');
  if (url) {
    assertText(url, 'link.url');
    if (url.length > MAX_LINK_LENGTH) throw new Error(`link.url must be at most ${MAX_LINK_LENGTH} characters.`);
    if (!isSafeHyperlinkUrl(url)) throw new Error('link.url must use http, https or mailto.');
  }
  if (anchor) {
    assertText(anchor, 'link.anchor');
    if (anchor.length > MAX_ANCHOR_LENGTH) throw new Error(`link.anchor must be at most ${MAX_ANCHOR_LENGTH} characters.`);
  }
  if (tooltip) {
    assertText(tooltip, 'link.tooltip');
    if (tooltip.length > MAX_TOOLTIP_LENGTH) throw new Error(`link.tooltip must be at most ${MAX_TOOLTIP_LENGTH} characters.`);
  }
}

export function parseFldSimpleHyperlink(instruction: string): HyperlinkTarget | null {
  const normalized = instruction.replace(/\s+/g, ' ').trim();
  if (!/^HYPERLINK\b/i.test(normalized)) return null;
  const parseQuoted = (pattern: RegExp): string | undefined => {
    const value = normalized.match(pattern)?.[1];
    return value ? value.replace(/\\(["\\])/g, '$1') : undefined;
  };
  const url = parseQuoted(/HYPERLINK\s+"((?:\\.|[^"\\])+)"/i);
  const anchor = parseQuoted(/\\l\s+"((?:\\.|[^"\\])+)"/i);
  if (!url && !anchor) return null;
  return { url: url ?? undefined, anchor: anchor ?? undefined };
}

export function getHyperlinks(ctx: HyperlinkContext): HyperlinkInfo[] {
    const main = ctx.partDocumentCopy(ctx.mainPath);
    const body = bodyOf(main);
    const paragraphs = mainParagraphElements(body);
    const relationships = ctx.relationshipsFor(ctx.mainPath);
    const result: HyperlinkInfo[] = [];
    const paragraphIndex = new Map(paragraphs.map((paragraph, index) => [paragraph, index]));
    for (const paragraph of paragraphs) {
      const runs = ownRuns(paragraph);
      for (const hyperlink of descendants(paragraph, 'hyperlink')) {
        const linkedRuns = runs.flatMap((run, index) =>
          isDescendantOfWithin(run, hyperlink, paragraph) ? [index] : []);
        if (!linkedRuns.length) continue;
        const relationshipId = relationshipIdOf(hyperlink);
        const relation = relationshipId ? relationships.get(relationshipId) : undefined;
        const url = relation?.mode === 'External' ? relation.target : undefined;
        const anchor = hyperlink.getAttributeNS(WORD_NS, 'anchor') ?? hyperlink.getAttribute('w:anchor') ?? undefined;
        const tooltip = hyperlink.getAttributeNS(WORD_NS, 'tooltip') ?? hyperlink.getAttribute('w:tooltip') ?? undefined;
        result.push({
          paragraph: paragraphIndex.get(paragraph)!,
          runs: linkedRuns,
          text: textOf(hyperlink),
          url,
          anchor,
          tooltip,
          isExternal: Boolean(url),
          unsafe: isUnsafeHyperlink({ url, anchor }),
          relationshipId,
        });
      }
      for (const field of descendants(paragraph, 'fldSimple')) {
        const instruction = field.getAttributeNS(WORD_NS, 'instr') ?? field.getAttribute('w:instr') ?? '';
        const parsed = parseFldSimpleHyperlink(instruction);
        if (!parsed) continue;
        const tooltip = field.getAttributeNS(WORD_NS, 'tooltip') ?? field.getAttribute('w:tooltip') ?? undefined;
        const linkedRuns = runs.flatMap((run, index) =>
          isDescendantOfWithin(run, field, paragraph) ? [index] : []);
        if (!linkedRuns.length) continue;
        result.push({
          paragraph: paragraphIndex.get(paragraph)!,
          runs: linkedRuns,
          text: textOf(field),
          ...parsed,
          tooltip,
          isExternal: Boolean(parsed.url),
          unsafe: isUnsafeHyperlink(parsed),
        });
      }
    }
    return result;
}

export function hyperlinkNode(ctx: HyperlinkContext, hyperlink: HyperlinkInfo, document: Document): Element {
    const paragraph = paragraphAt(document, hyperlink.paragraph);
    const run = ownRuns(paragraph)[hyperlink.runs[0]!] ?? null;
    let node: Node | null = run;
    while (node && node !== paragraph) {
      if (node.nodeType === 1 && (node as Element).namespaceURI === WORD_NS &&
          ['hyperlink', 'fldSimple'].includes((node as Element).localName ?? '')) {
        return node as Element;
      }
      node = node.parentNode;
    }
    throw new Error('Hyperlink node was not found.');
}

export function resolveHyperlink(
  ctx: HyperlinkContext,
  reference: HyperlinkInfo | number | { paragraph: number; runs: number[]; text: string },
): HyperlinkInfo {
    const hyperlinks = ctx.getHyperlinks();
    if (typeof reference === 'number') {
      const hyperlink = hyperlinks[reference];
      if (!hyperlink) throw new Error('Hyperlink does not exist.');
      return hyperlink;
    }
    if (!reference || typeof reference !== 'object' || !Number.isSafeInteger(reference.paragraph) ||
        !Array.isArray(reference.runs) || !reference.runs.every(run => Number.isSafeInteger(run) && run >= 0) ||
        typeof reference.text !== 'string') {
      throw new Error('hyperlink must be a hyperlink index or object with paragraph, runs and text.');
    }
    const hyperlink = hyperlinks.find(item => item.paragraph === reference.paragraph &&
      item.runs[0] === reference.runs[0] && item.text === reference.text);
    if (!hyperlink) throw new Error('Hyperlink does not exist.');
    return hyperlink;
}

export function insertHyperlink(
  ctx: HyperlinkContext,
  target: { paragraph: number; start: number; end: number },
  link: { url?: string; anchor?: string; tooltip?: string },
): HyperlinkInfo {
    assertIndex(target.paragraph);
    assertIndex(target.start);
    assertIndex(target.end);
    assertHyperlinkInput(link);
    return ctx.withDraft(draft => {
      let createdId: string | undefined;
      let createdMarker: { paragraph: number; run: number; text: string } | undefined;
      let relationshipPath: string | undefined;
      let relationships: Document | undefined;
      if (link.url) {
        relationshipPath = resolveRelationshipsPath(draft.mainPath);
        relationships = draft.hasPart(relationshipPath)
          ? draft.partDocumentCopy(relationshipPath)
          : parseXml(`<Relationships xmlns="${REL_NS}"/>`);
        createdId = draft.nextRelationshipId(relationships);
        const relationship = relationships.createElementNS(REL_NS, 'Relationship');
        relationship.setAttribute('Id', createdId);
        relationship.setAttribute('Type', HYPERLINK_REL);
        relationship.setAttribute('Target', link.url);
        relationship.setAttribute('TargetMode', 'External');
        relationships.documentElement!.appendChild(relationship);
      }
      draft.updatePartXmlFromCopy(draft.mainPath, document => {
        const setHyperlinkAttributes = (node: Element): void => {
          if (createdId) node.setAttributeNS(OFFICE_REL_NS, 'r:id', createdId);
          else { node.removeAttributeNS(OFFICE_REL_NS, 'id'); node.removeAttribute('r:id'); }
          if (link.anchor) node.setAttributeNS(WORD_NS, 'w:anchor', link.anchor);
          else removeWordAttribute(node, 'anchor');
          if (link.tooltip) node.setAttributeNS(WORD_NS, 'w:tooltip', link.tooltip);
          else removeWordAttribute(node, 'tooltip');
        };
        const paragraph = paragraphAt(document, target.paragraph);
        textRangeLength(paragraph, target.start, target.end);
        draft.splitRunAtOffset(paragraph, target.end);
        draft.splitRunAtOffset(paragraph, target.start);
        const runs = ownRuns(paragraph);
        let cursor = 0;
        const selected: Element[] = [];
        for (const run of runs) {
          const text = textOf(run);
          const next = cursor + text.length;
          if (target.start < next && target.end > cursor) selected.push(run);
          cursor = next;
        }
        if (!selected.length) throw new Error('Hyperlink range must include text.');
        const firstParent = selected[0]!.parentNode as Element;
        if (selected.some(run => run.parentNode !== firstParent)) {
          throw new Error('Hyperlink range cannot cross different run containers.');
        }
        let hyperlink: Element;
        if (firstParent.namespaceURI === WORD_NS && firstParent.localName === 'hyperlink') {
          const original = firstParent;
          const selectedSet = new Set<Node>(selected);
          const beforeNodes: Node[] = [];
          const selectedNodes: Node[] = [];
          const afterNodes: Node[] = [];
          let phase: 'before' | 'selected' | 'after' = 'before';
          for (const child of Array.from(original.childNodes)) {
            if (selectedSet.has(child)) {
              if (phase === 'after') {
                throw new Error('Hyperlink range must map to a contiguous segment inside an existing hyperlink.');
              }
              phase = 'selected';
              selectedNodes.push(child);
            } else if (phase === 'before') {
              beforeNodes.push(child);
            } else {
              phase = 'after';
              afterNodes.push(child);
            }
          }
          if (!selectedNodes.length) throw new Error('Hyperlink range must include text.');
          if (!beforeNodes.length && !afterNodes.length) {
            hyperlink = original;
            setHyperlinkAttributes(hyperlink);
          } else {
            const parent = original.parentNode as Element | null;
            if (!parent) throw new Error('Hyperlink container is detached.');
            const selectedLink = original.cloneNode(false) as Element;
            setHyperlinkAttributes(selectedLink);
            for (const node of selectedNodes) selectedLink.appendChild(node);
            const beforeLink = beforeNodes.length ? original.cloneNode(false) as Element : undefined;
            const afterLink = afterNodes.length ? original.cloneNode(false) as Element : undefined;
            if (beforeLink) for (const node of beforeNodes) beforeLink.appendChild(node);
            if (afterLink) for (const node of afterNodes) afterLink.appendChild(node);
            parent.insertBefore(beforeLink ?? selectedLink, original);
            if (beforeLink) parent.insertBefore(selectedLink, original);
            if (afterLink) parent.insertBefore(afterLink, original);
            parent.removeChild(original);
            hyperlink = selectedLink;
          }
        } else {
          hyperlink = wordElement(document, 'hyperlink');
          setHyperlinkAttributes(hyperlink);
          firstParent.insertBefore(hyperlink, selected[0]!);
          for (const run of selected) hyperlink.appendChild(run);
        }
        const firstHyperlinkRun = descendants(hyperlink, 'r')[0];
        if (!firstHyperlinkRun) throw new Error('Inserted hyperlink does not contain runs.');
        const props = properties(firstHyperlinkRun, 'rPr');
        if (!children(props, 'rStyle').length) setWordValue(property(props, 'rStyle'), 'Hyperlink');
        if (!children(props, 'color').length) setWordValue(property(props, 'color'), '0563C1');
        if (!children(props, 'u').length) setWordValue(property(props, 'u'), 'single');
        const firstRun = ownRuns(paragraph).findIndex(run => run === firstHyperlinkRun);
        if (firstRun < 0) throw new Error('Inserted hyperlink run index could not be resolved.');
        createdMarker = { paragraph: target.paragraph, run: firstRun, text: textOf(hyperlink) };
      });
      if (relationshipPath && relationships) {
        if (draft.hasPart(relationshipPath)) {
          draft.writePartXml(relationshipPath, serializeXml(relationships));
        } else {
          draft.addPart(relationshipPath, encodeXml(serializeXml(relationships)), 'application/vnd.openxmlformats-package.relationships+xml');
        }
      }
      const created = draft.getHyperlinks().find(item => item.paragraph === createdMarker?.paragraph &&
        item.runs[0] === createdMarker?.run && item.text === createdMarker?.text &&
        item.url === link.url && item.anchor === link.anchor);
      if (!created) throw new Error('Inserted hyperlink could not be resolved.');
      return created;
    });
}

export function updateHyperlink(
  ctx: HyperlinkContext,
  hyperlink: HyperlinkInfo | number | { paragraph: number; runs: number[]; text: string },
  link: { url?: string; anchor?: string; tooltip?: string },
): void {
    assertHyperlinkInput(link);
    const current = ctx.resolveHyperlink(hyperlink);
    const currentNodeName = ctx.hyperlinkNode(current, ctx.partDocumentCopy(ctx.mainPath)).localName;
    ctx.withDraft((draft) => {
      let nextRelationshipId = current.relationshipId;
      const currentRelationshipUsers = current.relationshipId
        ? draft.getHyperlinks().filter(item => item.relationshipId === current.relationshipId).length
        : 0;
      if (currentNodeName !== 'fldSimple' && link.url && link.url !== current.url) {
        const relPath = resolveRelationshipsPath(draft.mainPath);
        const rels = draft.hasPart(relPath) ? draft.partDocumentCopy(relPath) : parseXml(`<Relationships xmlns="${REL_NS}"/>`);
        const canReuse = Boolean(current.relationshipId && currentRelationshipUsers <= 1);
        nextRelationshipId = canReuse ? current.relationshipId : draft.nextRelationshipId(rels);
        const existing = canReuse
          ? children(rels.documentElement!, 'Relationship', REL_NS)
            .find(item => item.getAttribute('Id') === nextRelationshipId)
          : undefined;
        const relationship = existing ?? rels.createElementNS(REL_NS, 'Relationship');
        relationship.setAttribute('Id', nextRelationshipId!);
        relationship.setAttribute('Type', HYPERLINK_REL);
        relationship.setAttribute('Target', link.url);
        relationship.setAttribute('TargetMode', 'External');
        if (!existing) rels.documentElement!.appendChild(relationship);
        if (draft.hasPart(relPath)) {
          draft.writePartXml(relPath, serializeXml(rels));
        } else {
          draft.addPart(relPath, encodeXml(serializeXml(rels)), 'application/vnd.openxmlformats-package.relationships+xml');
        }
      }
      draft.updatePartXmlFromCopy(draft.mainPath, document => {
        const node = draft.hyperlinkNode(current, document);
        if (node.localName === 'fldSimple') {
          node.setAttributeNS(WORD_NS, 'w:instr', fieldInstruction(link));
          if (link.tooltip) node.setAttributeNS(WORD_NS, 'w:tooltip', link.tooltip);
          else removeWordAttribute(node, 'tooltip');
        } else {
          if (link.url) node.setAttributeNS(OFFICE_REL_NS, 'r:id', nextRelationshipId!);
          else { node.removeAttributeNS(OFFICE_REL_NS, 'id'); node.removeAttribute('r:id'); }
          if (link.anchor) node.setAttributeNS(WORD_NS, 'w:anchor', link.anchor);
          else removeWordAttribute(node, 'anchor');
          if (link.tooltip) node.setAttributeNS(WORD_NS, 'w:tooltip', link.tooltip);
          else removeWordAttribute(node, 'tooltip');
        }
      });
      if (current.relationshipId && (!link.url || current.relationshipId !== nextRelationshipId)) {
        const stillUsed = draft.getHyperlinks().some(item => item.relationshipId === current.relationshipId);
        if (!stillUsed) {
          const relPath = resolveRelationshipsPath(draft.mainPath);
          if (draft.hasPart(relPath)) {
            draft.updatePartXmlFromCopy(relPath, rels => {
              for (const relationship of children(rels.documentElement!, 'Relationship', REL_NS)) {
                if (relationship.getAttribute('Id') === current.relationshipId) {
                  rels.documentElement!.removeChild(relationship);
                  break;
                }
              }
            });
          }
        }
      }
    });
}

export function removeHyperlink(
  ctx: HyperlinkContext,
  hyperlink: HyperlinkInfo | number | { paragraph: number; runs: number[]; text: string },
  options: { keepText?: boolean } = {},
): void {
    const link = ctx.resolveHyperlink(hyperlink);
    ctx.withDraft((draft) => {
      draft.updatePartXmlFromCopy(draft.mainPath, document => {
        const node = draft.hyperlinkNode(link, document);
        const parent = node.parentNode as Element;
        if (options.keepText === false) parent.removeChild(node);
        else {
          while (node.firstChild) parent.insertBefore(node.firstChild, node);
          parent.removeChild(node);
        }
      });
      if (link.relationshipId) {
        const stillUsed = draft.getHyperlinks().some(item => item.relationshipId === link.relationshipId);
        if (!stillUsed) {
          const relPath = resolveRelationshipsPath(draft.mainPath);
          if (draft.hasPart(relPath)) {
            draft.updatePartXmlFromCopy(relPath, rels => {
              for (const relationship of children(rels.documentElement!, 'Relationship', REL_NS)) {
                if (relationship.getAttribute('Id') === link.relationshipId) {
                  rels.documentElement!.removeChild(relationship);
                  break;
                }
              }
            });
          }
        }
      }
    });
}

export function readExternalRelationship(relationships: Element | undefined, id: string): string | undefined {
  if (!relationships) return undefined;
  for (const rel of Array.from(relationships.getElementsByTagNameNS(REL_NS, 'Relationship'))) {
    if (rel.getAttribute('Id') !== id) continue;
    if (rel.getAttribute('TargetMode') !== 'External') return undefined;
    return rel.getAttribute('Target') ?? undefined;
  }
  return undefined;
}
