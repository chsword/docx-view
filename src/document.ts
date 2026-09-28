import JSZip from 'jszip';
import type { Document, Element, Node } from '@xmldom/xmldom';
import type {
  AgentRequest, BookmarkInfo, CellFormat, ClipboardBlock, ClipboardFragment, ClipboardImage, ClipboardParagraph, ClipboardRun,
  CommentAnchor, CommentInfo, DocumentBlock, DocumentRange, DocumentSnapshot,
  DocumentProperties, DocumentProtection, HistoryEntry, HyperlinkInfo, ImageInfo, NoteInfo, NoteSettings, NoteSettingsValue, NumberingDefinition,
  NumberingInfo, OutlineNode, PageSetup, ParagraphFormat, ParagraphInfo, RevisionInfo, ReviewerInfo, RowFormat, RunFormat, RunInfo, SectionInfo,
  SectionType, Shading, StyleInfo, TabStop, TableFormat, TableInfo, TextRange,
  TableCellLocation,
} from './types.js';
import {
  APP_PROPERTY_KEYS,
  APP_PROPS_REL,
  APP_PROPS_TYPE,
  CORE_PROPERTY_KEYS,
  CORE_PROPS_REL,
  CORE_PROPS_TYPE,
  assertDocumentPropertiesPatch,
  defaultAppPropertiesXml,
  defaultCorePropertiesXml,
  parseDocumentProperties,
  setAppDocumentPropertiesOn,
  setCoreDocumentPropertiesOn,
} from './docprops.js';
import type { NumberingModel } from './numbering.js';
import { computeParagraphNumbering, parseNumberingModel } from './numbering.js';
import {
  A_NS, dataUrlForBytes, decodeBase64, detectImageContentType, detectImageSize, emuToPx, extensionForContentType, IMAGE_REL,
  isBrowserRenderableContentType, OFFICE_REL_NS, PIC_NS, placeholderDataUrl, ptToEmu, pxToEmu, readRunImages,
  resolveRelationshipsPath, resolveTargetPath, V_NS, WP_NS,
} from './drawing.js';
import type { RelationshipTarget } from './drawing.js';
import {
  assertText, children, childrenThroughTransparent, CONTENT_TYPES_NS, descendants, isTransparentWordWrapper,
  isValidXmlCharCode, OFFICE_DOCUMENT_REL, parseXml, REL_NS, sanitizeText, serializeXml, setWordValue, validatePath,
  WORD_NS, wordElement, wordValue,
} from './xml.js';
import {
  assertIndex,
  validateBorderSide,
  validateDocShading,
  validateParagraphBorders,
  validateParagraphFormat,
  validateRequest,
  validateRows,
  validateRunFormat,
  validateTabs,
} from './operations.js';
import { collectSections, readSections, SECTION_ORDER } from './section.js';
import {
  cloneStyleInfo,
  computeEffectiveParagraphFormat,
  computeEffectiveRunFormat,
  parseStyles,
  readParagraphProperties,
  readRunProperties,
  resolveStyleChainOrDefault,
  type StylesContext,
} from './styles.js';
import { cellSpan, parseCellFormat, parseTableFormat, readTable, rowCells, tableGrid } from './table.js';
import {
  assertHyperlinkInput,
  isInternalBookmark,
  isUnsafeHyperlink,
  parseFldSimpleHyperlink,
} from './hyperlink.js';
import {
  assertDocumentProtection, defaultNotePartXml, formatNoteMarker, noteContentType, noteRefName, noteReferenceName, noteReferenceStyle,
  noteRelationshipType, parseCustomMark, parseDocumentNoteSettings, parseDocumentProtection, parseNoteEntries, parseSectionNoteSettings,
  setDocumentProtectionOn, setNoteSettingsOn, setTrackChangesOn, type NoteKind,
} from './notes.js';
import {
  COMMENTS_EXTENDED_REL,
  COMMENTS_EXTENDED_TYPE,
  COMMENTS_REL,
  COMMENTS_TYPE,
  W14_NS,
  W15_NS,
  commentParagraphStyle,
  commentReferenceStyle,
  defaultCommentsExtendedXml,
  defaultCommentsXml,
  ensureCommentParagraphParaId,
  makeCommentAnnotationRun,
  parseCommentEntries,
  parseCommentExEntries,
} from './comments.js';
import {
  createRevisionWrapper,
  deletedTextOf,
  hasRevisionMarkup,
  markFormatRevision,
  markRevision,
  revisionAuthorOf,
  reviewerBucketKey,
  reviewerBucketOf,
  readParagraphRevisionMark,
  readRevisionMark,
  readRunRevisionMarks,
  visibleTextOf,
} from './revisions.js';

const MAX_ARCHIVE = 50 * 1024 * 1024;
const MAX_PART = 16 * 1024 * 1024;
const MAX_TOTAL = 64 * 1024 * 1024;
const MAX_PARTS = 2048;
const DOCX_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
const MAIN_TYPE = `${DOCX_TYPE}.main+xml`;
const SETTINGS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings';
const HEADER_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header';
const FOOTER_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer';
const RELS_CONTENT_TYPE = 'application/vnd.openxmlformats-package.relationships+xml';
const HEADER_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml';
const FOOTER_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml';
const SETTINGS_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml';
const DEFAULT_HEADER_FOOTER_KIND: 'default' | 'first' | 'even' = 'default';
const NUMBERING_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml';
const NUMBERING_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering';
const STYLES_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles';
const STYLES_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml';
const THEME_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme';
const encoder = new TextEncoder();
const IMAGE_LIMIT = 16 * 1024 * 1024;
const HYPERLINK_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';
const STRUCTURE_PARTS = new Set(['[Content_Types].xml', '_rels/.rels']);
const HISTORY_MAX_ENTRIES = 50;
const HISTORY_MAX_BYTES = 64 * 1024 * 1024;
const HISTORY_MERGE_WINDOW_MS = 500;
const CLIPBOARD_MAX_PARAGRAPHS = 1_000;
const CLIPBOARD_MAX_RUNS = 10_000;
const CLIPBOARD_MAX_RUN_TEXT_LENGTH = 1_000_000;
const CLIPBOARD_MAX_IMAGES = 200;
const REVIEWER_FILTER_BUCKET_KEYS = new Set(['unattributed', 'empty', 'blank']);
const REVISION_FILTER_MAX_AUTHORS = 1_000;
const COMPARE_MAX_PARAGRAPHS = 1_000;
const COMPARE_PARAGRAPH_PAIR_THRESHOLD = 0.5;
const REVISION_ELEMENT_NAMES = new Set([
  'ins', 'del', 'moveFrom', 'moveTo', 'rPrChange', 'pPrChange', 'tblPrChange', 'trPrChange', 'tcPrChange', 'cellIns', 'cellDel',
]);

type HistoryAction =
  | { kind: 'setParagraphText'; paragraph: number }
  | { kind: 'transaction' }
  | { kind: 'group' }
  | { kind: 'undo' }
  | { kind: 'redo' }
  | { kind: 'other' };

interface HistoryState {
  parts: Map<string, Uint8Array>;
  bytes: number;
  entry: HistoryEntry;
  action: HistoryAction;
}

function elementChildren(node: Node, namespace?: string, localName?: string): Element[] {
  const result: Element[] = [];
  for (let child = node.firstChild; child; child = child.nextSibling) {
    if (child.nodeType !== 1) continue;
    const element = child as Element;
    if ((!namespace || element.namespaceURI === namespace) && (!localName || element.localName === localName)) {
      result.push(element);
    }
  }
  return result;
}

interface NumberingContext {
  revision: number;
  mainPath: string;
  numberingPath?: string;
  stylesPath?: string;
  model: NumberingModel;
}

function decodeXml(bytes: Uint8Array): string {
  const utf16le = (bytes[0] === 0xff && bytes[1] === 0xfe) || (bytes[0] === 0x3c && bytes[1] === 0);
  const utf16be = (bytes[0] === 0xfe && bytes[1] === 0xff) || (bytes[0] === 0 && bytes[1] === 0x3c);
  return new TextDecoder(utf16le ? 'utf-16le' : utf16be ? 'utf-16be' : 'utf-8', { fatal: true }).decode(bytes);
}

function encodeXml(xml: string): Uint8Array {
  return encoder.encode(xml.replace(/^(<\?xml\b[^?]*\bencoding\s*=\s*)(["'])[^"']*\2/i, '$1"UTF-8"'));
}

function encodeBase64(bytes: Uint8Array): string {
  const NodeBuffer = (globalThis as { Buffer?: { from(bytes: ArrayBufferLike, byteOffset?: number, length?: number): { toString(encoding: 'base64'): string } } }).Buffer;
  if (NodeBuffer) {
    return NodeBuffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength).toString('base64');
  }
  let binary = '';
  const chunkSize = 0x8000;
  for (let index = 0; index < bytes.length; index += chunkSize) {
    binary += String.fromCharCode(...bytes.subarray(index, index + chunkSize));
  }
  return btoa(binary);
}

function equalBytes(a: Uint8Array, b: Uint8Array): boolean {
  if (a === b) return true;
  if (a.byteLength !== b.byteLength) return false;
  for (let index = 0; index < a.byteLength; index++) {
    if (a[index] !== b[index]) return false;
  }
  return true;
}

function equalPartMap(a: Map<string, Uint8Array>, b: Map<string, Uint8Array>): boolean {
  if (a === b) return true;
  if (a.size !== b.size) return false;
  for (const [path, bytes] of a) {
    const next = b.get(path);
    if (!next || !equalBytes(bytes, next)) return false;
  }
  return true;
}

function cloneTableCellLocation(location: TableCellLocation): TableCellLocation {
  return { ...location };
}

function dirname(path: string): string {
  const slash = path.lastIndexOf('/');
  return slash === -1 ? '' : path.slice(0, slash);
}

function basename(path: string): string {
  return path.slice(path.lastIndexOf('/') + 1);
}

function nextRelationshipId(root: Element): string {
  const used = new Set(children(root, 'Relationship', REL_NS).map((relation) => relation.getAttribute('Id')).filter(Boolean));
  let index = 1;
  while (used.has(`rId${index}`)) index++;
  return `rId${index}`;
}

function bodyOf(document: Document): Element {
  const root = document.documentElement;
  if (!root || root.namespaceURI !== WORD_NS || root.localName !== 'document') {
    throw new Error('Only transitional WordprocessingML documents are supported.');
  }
  const body = children(root, 'body');
  if (body.length !== 1) throw new Error('Main document must have exactly one w:body.');
  return body[0]!;
}

function partDirectory(path: string): string {
  const index = path.lastIndexOf('/');
  return index === -1 ? '' : path.slice(0, index);
}

function normalizePath(path: string): string {
  const stack: string[] = [];
  for (const part of path.split('/')) {
    if (!part || part === '.') continue;
    if (part === '..') {
      stack.pop();
      continue;
    }
    stack.push(part);
  }
  return stack.join('/');
}

function resolvePartPath(basePart: string, target: string): string {
  let decoded: string;
  try {
    decoded = decodeURIComponent(target);
  } catch {
    return '';
  }
  decoded = decoded.replace(/^\/+/, '');
  return normalizePath(target.startsWith('/') ? decoded : `${partDirectory(basePart)}/${decoded}`);
}

function decodePackageTarget(target: string): string {
  try {
    return decodeURIComponent(target).replace(/^\/+/, '');
  } catch {
    return '';
  }
}

function relativeTarget(fromPart: string, toPart: string): string {
  const from = partDirectory(fromPart).split('/').filter(Boolean);
  const to = toPart.split('/').filter(Boolean);
  while (from.length && to.length && from[0] === to[0]) {
    from.shift();
    to.shift();
  }
  return `${'../'.repeat(from.length)}${to.join('/')}` || toPart;
}

function relsPath(partPath: string): string {
  const directory = partDirectory(partPath);
  const name = partPath.slice(partPath.lastIndexOf('/') + 1);
  return `${directory ? `${directory}/` : ''}_rels/${name}.rels`;
}

function paragraphAt(document: Document, index: number): Element {
  assertIndex(index);
  const paragraph = descendants(bodyOf(document), 'p')[index];
  if (!paragraph) throw new Error(`Paragraph ${index} does not exist.`);
  return paragraph;
}

function blockContainerOf(document: Document): Element {
  const root = document.documentElement;
  if (!root || root.namespaceURI !== WORD_NS) throw new Error('Unsupported WordprocessingML part.');
  if (root.localName === 'document') return bodyOf(document);
  if (['hdr', 'ftr', 'footnotes', 'endnotes'].includes(root.localName ?? '')) return root;
  throw new Error('Part does not contain block-level WordprocessingML content.');
}

function isTransparentWrapper(element: Element): boolean {
  return isTransparentWordWrapper(element);
}

function blockPositions(parent: Element): { block: Element; parent: Element }[] {
  return children(parent).flatMap(child => {
    if (isTransparentWrapper(child)) return blockPositions(child);
    return ['p', 'tbl'].includes(child.localName ?? '') ? [{ block: child, parent }] : [];
  });
}

function blockElements(parent: Element): Element[] {
  return blockPositions(parent).map(position => position.block);
}

function paragraphContainer(paragraph: Element): Element {
  let parent = paragraph.parentNode;
  while (parent && parent.nodeType === 1) {
    const element = parent as Element;
    if (element.namespaceURI === WORD_NS && ['body', 'tc'].includes(element.localName ?? '')) return element;
    if (!isTransparentWrapper(element)) break;
    parent = element.parentNode;
  }
  throw new Error('Paragraph is not inside a body or table cell container.');
}

function clearParagraphContent(paragraph: Element): void {
  for (const child of [...children(paragraph)]) {
    if (child.localName !== 'pPr') paragraph.removeChild(child);
  }
}

function deleteParagraphElement(paragraph: Element): void {
  const parent = paragraph.parentNode as Element;
  let container: Element | undefined;
  try { container = paragraphContainer(paragraph); } catch {}
  // A cell must end with a paragraph, and section properties must not be silently lost.
  if (children(paragraph, 'pPr').some(props => children(props, 'sectPr').length)) {
    throw new Error('Cannot delete a section-break paragraph; edit its XML explicitly.');
  }
  if (container) {
    const blocks = blockElements(container);
    const remaining = blocks.filter(block => block !== paragraph);
    const indexInContainer = blocks.indexOf(paragraph);
    const mustKeepParagraph =
      (container.localName === 'body' && (
        !remaining.length ||
        remaining.at(-1)?.localName === 'tbl' ||
        (indexInContainer > 0 &&
          indexInContainer < blocks.length - 1 &&
          blocks[indexInContainer - 1]?.localName === 'tbl' &&
          blocks[indexInContainer + 1]?.localName === 'tbl')
      )) ||
      (container.localName === 'tc' && remaining.at(-1)?.localName !== 'p');
    if (mustKeepParagraph) {
      clearParagraphContent(paragraph);
      return;
    }
  }
  parent.removeChild(paragraph);
}

function isParagraphAnchorMarker(element: Element): boolean {
  return [
    'bookmarkStart',
    'bookmarkEnd',
    'commentRangeStart',
    'commentRangeEnd',
    'proofErr',
    'permStart',
    'permEnd',
    'moveFromRangeStart',
    'moveFromRangeEnd',
    'moveToRangeStart',
    'moveToRangeEnd',
    'customXmlInsRangeStart',
    'customXmlInsRangeEnd',
    'customXmlDelRangeStart',
    'customXmlDelRangeEnd',
    'customXmlMoveFromRangeStart',
    'customXmlMoveFromRangeEnd',
    'customXmlMoveToRangeStart',
    'customXmlMoveToRangeEnd',
  ].includes(element.localName ?? '');
}

function textElements(element: Element): Element[] {
  const result: Element[] = [];
  function walk(node: Node, deletedDepth = 0): void {
    for (let child = node.firstChild; child; child = child.nextSibling) {
      if (child.nodeType !== 1) continue;
      const element = child as Element;
      if (element.namespaceURI === WORD_NS) {
        const localName = element.localName ?? '';
        if (element.localName === 'p') continue;
        const inDeleted = deletedDepth > 0 || localName === 'delText';
        if (!inDeleted && ['t', 'tab', 'br', 'cr', 'noBreakHyphen', 'softHyphen', 'sym'].includes(localName)) {
          result.push(element);
          continue;
        }
      }
      walk(element, deletedDepth + (element.namespaceURI === WORD_NS && ['del', 'moveFrom'].includes(element.localName ?? '') ? 1 : 0));
    }
  }
  walk(element);
  return result;
}

function numberingProperty(parent: Element, name: 'ilvl' | 'numId'): Element {
  let result = children(parent, name)[0];
  if (!result) {
    result = wordElement(parent.ownerDocument!, name);
    if (name === 'ilvl') parent.insertBefore(result, children(parent, 'numId')[0] ?? null);
    else parent.appendChild(result);
  }
  return result;
}

function elementText(element: Element): string {
  if (['t', 'delText'].includes(element.localName ?? '')) return element.textContent ?? '';
  if (element.localName === 'tab') return '\t';
  if (element.localName === 'noBreakHyphen') return '\u2011';
  if (element.localName === 'softHyphen') return '\u00ad';
  if (element.localName === 'sym') {
    const value = element.getAttributeNS(WORD_NS, 'char') ?? element.getAttribute('w:char');
    if (!value || !/^[a-f0-9]{1,4}$/i.test(value)) return '';
    const code = Number.parseInt(value, 16);
    return Number.isFinite(code) && isValidXmlCharCode(code) ? String.fromCharCode(code) : '�';
  }
  return '\n';
}

function fieldPlaceholder(root: Element): string | undefined {
  const instructions = [
    ...children(root, 'fldSimple').map(node => (node.getAttributeNS(WORD_NS, 'instr') ?? '').toUpperCase()),
    ...descendants(root, 'instrText').map(node => (node.textContent ?? '').toUpperCase()),
  ];
  if (instructions.some(instruction => instruction.includes('NUMPAGES'))) return '?';
  if (instructions.some(instruction => instruction.includes('PAGE'))) return '1';
  return undefined;
}

function textOf(element: Element): string {
  const text = visibleTextOf(element) || textElements(element).map(elementText).join('');
  if (text) return text;
  return fieldPlaceholder(element) ?? '';
}

function compactDefined<T extends object>(value: T): T {
  return Object.fromEntries(Object.entries(value).filter(([, entry]) => entry !== undefined)) as T;
}

function appendText(parent: Element, text: string, before: Node | null = null): void {
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

function newParagraph(document: Document, text: string): Element {
  const paragraph = wordElement(document, 'p');
  const run = wordElement(document, 'r');
  appendText(run, text);
  paragraph.appendChild(run);
  return paragraph;
}

function isRunVisuallyEmpty(run: Element): boolean {
  return !textOf(run) && !Array.from(run.childNodes).some((child) => child.nodeType === 1 &&
    (child as Element).namespaceURI === WORD_NS && ['drawing', 'pict', 'object'].includes((child as Element).localName ?? ''));
}

function isEmptyRun(run: Element): boolean {
  return !Array.from(run.childNodes).some((child) => {
    if (child.nodeType !== 1) return child.textContent?.trim().length;
    const element = child as Element;
    return !(element.namespaceURI === WORD_NS && ['rPr'].includes(element.localName ?? ''));
  });
}

// Edit only text-bearing nodes. Drawings, bookmarks, field codes and other XML survive.
function replaceSpan(paragraph: Element, start: number, end: number, replacement: string): void {
  const elements = textElements(paragraph);
  if (elements.length === 0) {
    const run = children(paragraph, 'r')[0] ?? wordElement(paragraph.ownerDocument!, 'r');
    if (!run.parentNode) paragraph.appendChild(run);
    appendText(run, replacement);
    return;
  }
  let offset = 0;
  let inserted = false;
  for (const element of elements) {
    const oldText = elementText(element);
    const next = offset + oldText.length;
    const insertHere = !inserted && start >= offset && (start === end ? start <= next : start < next);
    const overlaps = offset < end && next > start;
    if (insertHere || overlaps) {
      if (insertHere && !overlaps && element.localName !== 't') {
        appendText(element.parentNode as Element, replacement, start === offset ? element : element.nextSibling);
        inserted = true;
        offset = next;
        continue;
      }
      const prefix = oldText.slice(0, Math.max(0, Math.min(oldText.length, start - offset)));
      const suffix = oldText.slice(Math.max(0, Math.min(oldText.length, end - offset)));
      const text = prefix + (insertHere ? replacement : '') + suffix;
      const parent = element.parentNode as Element;
      appendText(parent, text, element);
      parent.removeChild(element);
      if (insertHere) inserted = true;
    }
    offset = next;
  }
}

function revisionTextElement(run: Element, name: 't' | 'delText'): void {
  const document = run.ownerDocument!;
  for (const element of descendants(run, 't')) {
    const replacement = wordElement(document, name);
    for (let index = 0; index < element.attributes.length; index++) {
      const attribute = element.attributes.item(index);
      if (!attribute) continue;
      replacement.setAttributeNS(attribute.namespaceURI, attribute.name, attribute.value);
    }
    while (element.firstChild) replacement.appendChild(element.firstChild);
    element.parentNode?.replaceChild(replacement, element);
  }
}

function wrapRunsWithRevision(
  paragraph: Element,
  runs: Element[],
  kind: 'ins' | 'del',
  author?: string,
  date?: string,
  options: { convertText?: boolean } = {},
): void {
  const selected = new Set(runs);
  for (let index = 0; index < runs.length; index++) {
    const run = runs[index]!;
    if (!selected.has(run)) continue;
    const parent = run.parentNode?.nodeType === 1 ? run.parentNode as Element : undefined;
    if (!parent) continue;
    const group: Element[] = [run];
    let cursor = run.nextSibling;
    while (cursor?.nodeType === 1 && selected.has(cursor as Element)) {
      group.push(cursor as Element);
      cursor = cursor.nextSibling;
    }
    const wrapper = createRevisionWrapper(parent.ownerDocument!, kind, author, date);
    parent.insertBefore(wrapper, group[0]!);
    for (const entry of group) {
      wrapper.appendChild(entry);
      selected.delete(entry);
    }
  }
  if (kind === 'del' && options.convertText !== false) {
    for (const run of runs) revisionTextElement(run, 'delText');
  }
}

function properties(element: Element, name: 'pPr' | 'rPr'): Element {
  let result = children(element, name)[0];
  if (!result) {
    result = wordElement(element.ownerDocument!, name);
    element.insertBefore(result, element.firstChild);
  }
  return result;
}

// Known property order keeps generated pPr/rPr conformant without dropping unknown properties.
const PROPERTY_ORDER = {
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
  style: ['name', 'aliases', 'basedOn', 'next', 'link', 'autoRedefine', 'hidden', 'uiPriority',
    'semiHidden', 'unhideWhenUsed', 'qFormat', 'locked', 'personal', 'personalCompose',
    'personalReply', 'rsid', 'pPr', 'rPr', 'tblPr', 'trPr', 'tcPr', 'tblStylePr', 'extLst'],
  tblPr: ['tblStyle', 'tblpPr', 'tblOverlap', 'bidiVisual', 'tblStyleRowBandSize', 'tblStyleColBandSize',
    'tblW', 'jc', 'tblCellSpacing', 'tblInd', 'tblBorders', 'shd', 'tblLayout', 'tblCellMar',
    'tblLook', 'tblCaption', 'tblDescription', 'tblPrChange'],
  trPr: ['cnfStyle', 'divId', 'gridBefore', 'gridAfter', 'wBefore', 'wAfter', 'cantSplit', 'trHeight',
    'tblHeader', 'jc', 'hidden', 'ins', 'del', 'trPrChange'],
  tcPr: ['cnfStyle', 'tcW', 'gridSpan', 'hMerge', 'vMerge', 'tcBorders', 'shd', 'noWrap', 'tcMar',
    'textDirection', 'tcFitText', 'vAlign', 'hideMark', 'headers', 'cellIns', 'cellDel', 'cellMerge', 'tcPrChange'],
  tblBorders: ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'],
  tcBorders: ['top', 'left', 'bottom', 'right', 'insideH', 'insideV', 'tl2br', 'tr2bl'],
  tblCellMar: ['top', 'left', 'bottom', 'right'],
  tcMar: ['top', 'left', 'bottom', 'right'],
};

function property(parent: Element, name: string): Element {
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

function insertPropertyChild(parent: Element, child: Element): void {
  const order = PROPERTY_ORDER[parent.localName as keyof typeof PROPERTY_ORDER] ?? [];
  const position = order.indexOf(child.localName ?? '');
  if (position === -1) {
    parent.appendChild(child);
    return;
  }
  const following = children(parent).find((entry) => {
    const entryPosition = order.indexOf(entry.localName ?? '');
    return entryPosition > position;
  });
  parent.insertBefore(child, following ?? null);
}

function nearestNonTransparentAncestor(element: Element): Element {
  let result = element;
  while (isTransparentWordWrapper(result) && result.parentNode?.nodeType === 1) {
    result = result.parentNode as Element;
  }
  return result;
}

function sectionProperty(parent: Element, name: string): Element {
  let result = children(parent, name)[0];
  if (!result) {
    result = wordElement(parent.ownerDocument!, name);
    const index = SECTION_ORDER.indexOf(name as typeof SECTION_ORDER[number]);
    const following = children(parent).find(child => SECTION_ORDER.indexOf(child.localName as typeof SECTION_ORDER[number]) > index);
    parent.insertBefore(result, following ?? null);
  }
  return result;
}

function ownRuns(paragraph: Element): Element[] {
  return descendants(paragraph, 'r').filter(run => {
    let parent = run.parentNode;
    while (parent && parent !== paragraph) {
      if (parent.nodeType === 1 && (parent as Element).namespaceURI === WORD_NS &&
          (parent as Element).localName === 'p') return false;
      parent = parent.parentNode;
    }
    return true;
  });
}

function relationshipIdOf(hyperlink: Element): string | undefined {
  return hyperlink.getAttributeNS(OFFICE_REL_NS, 'id') ?? hyperlink.getAttribute('r:id') ?? undefined;
}

function isDescendantOfWithin(node: Node, ancestor: Node, stopAt: Node): boolean {
  let current: Node | null = node;
  while (current && current !== stopAt) {
    if (current === ancestor) return true;
    current = current.parentNode;
  }
  return false;
}

function nearestWordAncestor(node: Node | null, localName: string): Element | undefined {
  for (let current = node; current; current = current.parentNode) {
    if (current.nodeType !== 1) continue;
    const element = current as Element;
    if (element.namespaceURI === WORD_NS && element.localName === localName) return element;
  }
  return undefined;
}

function revisionIdOf(element: Element): number | undefined {
  const raw = element.getAttributeNS(WORD_NS, 'id') ?? element.getAttribute('w:id');
  if (!raw || !/^\d+$/.test(raw)) return undefined;
  const id = Number(raw);
  return Number.isSafeInteger(id) ? id : undefined;
}

function revisionNameOf(element: Element): string | undefined {
  const value = element.getAttributeNS(WORD_NS, 'name') ?? element.getAttribute('w:name') ?? undefined;
  const name = value?.trim();
  return name ? name : undefined;
}

function reviewerFilterKeyOf(author: string): string {
  if (REVIEWER_FILTER_BUCKET_KEYS.has(author) || author.startsWith('named:')) return author;
  return reviewerBucketKey(reviewerBucketOf(author));
}

function normalizeReviewerFilterAuthors(authors: string[]): Set<string> {
  return new Set(authors.map((author) => reviewerFilterKeyOf(author)));
}

function moveRevisionSideOf(element: Element): 'from' | 'to' | undefined {
  if (element.namespaceURI !== WORD_NS) return undefined;
  if (element.localName === 'moveFrom') return 'from';
  if (element.localName === 'moveTo') return 'to';
  return undefined;
}

function isMoveRevisionMarker(element: Element): boolean {
  return moveRevisionSideOf(element) !== undefined;
}

function rangeMarkerIdOf(element: Element): string | undefined {
  const value = element.getAttributeNS(WORD_NS, 'id') ?? element.getAttribute('w:id') ?? undefined;
  return value && /^\d+$/.test(value) ? value : undefined;
}

function activeMoveRangeId(paragraph: Element, marker: Element, side: 'from' | 'to'): string | undefined {
  const startName = side === 'from' ? 'moveFromRangeStart' : 'moveToRangeStart';
  const endName = side === 'from' ? 'moveFromRangeEnd' : 'moveToRangeEnd';
  const active: string[] = [];
  let found: string | undefined;
  const walk = (node: Node): boolean => {
    for (let child = node.firstChild; child; child = child.nextSibling) {
      if (child.nodeType !== 1) continue;
      const element = child as Element;
      if (element.namespaceURI === WORD_NS && element.localName === startName) {
        const id = rangeMarkerIdOf(element);
        if (id) active.push(id);
      }
      if (element === marker) {
        found = active.at(-1);
        return true;
      }
      if (walk(element)) return true;
      if (element.namespaceURI === WORD_NS && element.localName === endName) {
        const id = rangeMarkerIdOf(element);
        if (id) {
          const index = active.lastIndexOf(id);
          if (index >= 0) active.splice(index, 1);
        }
      }
    }
    return false;
  };
  walk(paragraph);
  return found;
}

interface MoveMarkerPairing {
  pairs: Map<Element, Element>;
  stable: Set<Element>;
}

function pairMoveRevisionMarkers(markers: Element[]): MoveMarkerPairing {
  const pairs = new Map<Element, Element>();
  const stable = new Set<Element>();
  interface MoveEntry {
    marker: Element;
    index: number;
    side: 'from' | 'to';
    name: string;
    rangeId?: string;
  }
  const entries: MoveEntry[] = [];
  for (const [index, marker] of markers.entries()) {
    const side = moveRevisionSideOf(marker);
    const name = revisionNameOf(marker);
    if (!side || !name) continue;
    const paragraph = revisionParagraphAnchor(marker);
    const rangeId = paragraph ? activeMoveRangeId(paragraph, marker, side) : undefined;
    entries.push({ marker, index, side, name, ...(rangeId ? { rangeId } : {}) });
  }
  const countsByName = new Map<string, { from: number; to: number }>();
  for (const entry of entries) {
    const count = countsByName.get(entry.name) ?? { from: 0, to: 0 };
    count[entry.side]++;
    countsByName.set(entry.name, count);
  }
  const pairByKey = (keyOf: (entry: MoveEntry) => string | undefined, markStable: (entry: MoveEntry) => boolean): void => {
    const pending = new Map<string, { from: MoveEntry[]; to: MoveEntry[] }>();
    for (const entry of entries) {
      if (pairs.has(entry.marker)) continue;
      const key = keyOf(entry);
      if (!key) continue;
      const bucket = pending.get(key) ?? { from: [], to: [] };
      pending.set(key, bucket);
      const opposite = entry.side === 'from' ? bucket.to : bucket.from;
      if (opposite.length) {
        const mate = opposite.shift()!;
        if (mate.side === entry.side) continue;
        pairs.set(entry.marker, mate.marker);
        pairs.set(mate.marker, entry.marker);
        if (markStable(entry) && markStable(mate)) {
          stable.add(entry.marker);
          stable.add(mate.marker);
        }
      } else {
        (entry.side === 'from' ? bucket.from : bucket.to).push(entry);
      }
    }
  };
  pairByKey((entry) => entry.rangeId ? `${entry.name}\u0000${entry.rangeId}` : undefined, () => true);
  pairByKey((entry) => entry.name, (entry) => {
    const count = countsByName.get(entry.name);
    return count?.from === 1 && count?.to === 1;
  });
  return { pairs, stable };
}

function unwrapNode(element: Element): void {
  const parent = element.parentNode;
  if (!parent) return;
  while (element.firstChild) parent.insertBefore(element.firstChild, element);
  parent.removeChild(element);
}

function replaceRevisionTextNodes(root: Element): void {
  const document = root.ownerDocument!;
  for (const from of descendants(root, 'delText')) {
    const to = wordElement(document, 't');
    for (let index = 0; index < from.attributes.length; index++) {
      const attribute = from.attributes.item(index);
      if (!attribute) continue;
      to.setAttributeNS(attribute.namespaceURI, attribute.name, attribute.value);
    }
    while (from.firstChild) to.appendChild(from.firstChild);
    from.parentNode?.replaceChild(to, from);
  }
}

function revisionParagraphAnchor(element: Element): Element | undefined {
  const ancestorParagraph = nearestWordAncestor(element, 'p');
  if (ancestorParagraph) return ancestorParagraph;
  const descendantParagraph = descendants(element, 'p')[0];
  if (descendantParagraph) return descendantParagraph;
  for (let current = element.parentNode; current; current = current.parentNode) {
    if (current.nodeType !== 1) continue;
    const paragraph = descendants(current as Element, 'p')[0];
    if (paragraph) return paragraph;
  }
  return undefined;
}

function revisionRunAnchor(paragraph: Element, element: Element): Element | undefined {
  const ancestorRun = nearestWordAncestor(element, 'r');
  if (ancestorRun && isDescendantOfWithin(ancestorRun, paragraph, paragraph.parentNode ?? paragraph)) return ancestorRun;
  return ownRuns(paragraph).find((run) => isDescendantOfWithin(run, element, paragraph));
}

function runHyperlinkInfo(run: Element, paragraph: Element, relationships?: Map<string, RelationshipTarget>): RunInfo['hyperlink'] {
  let parent = run.parentNode;
  while (parent && parent !== paragraph) {
    if (parent.nodeType === 1 && (parent as Element).namespaceURI === WORD_NS &&
        (parent as Element).localName === 'hyperlink') {
      const hyperlink = parent as Element;
      const relationshipId = relationshipIdOf(hyperlink);
      const relation = relationshipId ? relationships?.get(relationshipId) : undefined;
      const url = relation?.mode === 'External' ? relation.target : undefined;
      const anchor = hyperlink.getAttributeNS(WORD_NS, 'anchor') ?? hyperlink.getAttribute('w:anchor') ?? undefined;
      const tooltip = hyperlink.getAttributeNS(WORD_NS, 'tooltip') ?? hyperlink.getAttribute('w:tooltip') ?? undefined;
      return { url, anchor, tooltip, unsafe: isUnsafeHyperlink({ url, anchor }) };
    }
    if (parent.nodeType === 1 && (parent as Element).namespaceURI === WORD_NS &&
        (parent as Element).localName === 'fldSimple') {
      const fld = parent as Element;
      const instruction = fld.getAttributeNS(WORD_NS, 'instr') ?? fld.getAttribute('w:instr') ?? '';
      const parsed = parseFldSimpleHyperlink(instruction);
      if (parsed) return { ...parsed, unsafe: isUnsafeHyperlink(parsed) };
    }
    parent = parent.parentNode;
  }
  return undefined;
}

function noteReferenceInRun(run: Element): { kind: NoteKind; id: number; customMarkFollows: boolean } | null {
  for (const kind of ['footnote', 'endnote'] as const) {
    const reference = children(run, noteReferenceName(kind))[0];
    if (!reference) continue;
    const id = Number(reference.getAttributeNS(WORD_NS, 'id'));
    if (!Number.isSafeInteger(id)) return null;
    return {
      kind,
      id,
      customMarkFollows: ['1', 'true', 'on'].includes(reference.getAttributeNS(WORD_NS, 'customMarkFollows') ?? ''),
    };
  }
  return null;
}

function commentReferenceInRun(run: Element): number | null {
  const reference = children(run, 'commentReference')[0];
  if (!reference) return null;
  const id = Number(reference.getAttributeNS(WORD_NS, 'id') ?? reference.getAttribute('w:id'));
  return Number.isSafeInteger(id) && id >= 0 ? id : null;
}

interface NoteReferenceRecord {
  kind: NoteKind;
  id: number;
  customMarkFollows: boolean;
  paragraph: number;
  run: number;
  section: number;
}

interface NoteState {
  byKind: Record<NoteKind, Map<number, { number: number; marker: string }>>;
  refs: NoteReferenceRecord[];
  entries: Record<NoteKind, Map<number, ReturnType<typeof parseNoteEntries>[number]>>;
}

interface CommentLocation {
  sourcePartPath: string;
  paragraph?: number;
  runs?: number[];
  startParagraph?: number;
  endParagraph?: number;
  order: number;
}

interface CommentPartBinding {
  sourcePartPath: string;
  commentsPath?: string;
  commentsExtendedPath?: string;
}

function normalizedNoteSettingsPatch(input: Partial<NoteSettings>): Partial<NoteSettings> {
  const result: Partial<NoteSettings> = {};
  for (const kind of ['footnote', 'endnote'] as const) {
    if (!(kind in input)) continue;
    const value = input[kind];
    if (value === undefined) continue;
    if (!value || typeof value !== 'object' || Array.isArray(value)) throw new Error(`${kind} settings must be an object.`);
    const known = new Set(['pos', 'numFmt', 'numStart', 'numRestart']);
    for (const key of Object.keys(value)) {
      if (!known.has(key)) throw new Error(`Unknown note setting property: ${kind}.${key}`);
    }
    const patch: NoteSettingsValue = {};
    if ('pos' in value && value.pos !== undefined) patch.pos = value.pos;
    if ('numFmt' in value && value.numFmt !== undefined) patch.numFmt = value.numFmt;
    if ('numStart' in value && value.numStart !== undefined) patch.numStart = value.numStart;
    if ('numRestart' in value && value.numRestart !== undefined) patch.numRestart = value.numRestart;
    if (Object.keys(patch).length) result[kind] = patch;
  }
  for (const key of Object.keys(input)) {
    if (!['footnote', 'endnote'].includes(key)) throw new Error(`Unknown note settings key: ${key}`);
  }
  return result;
}

function bodyChildren(parent: Element): Element[] {
  return children(parent).flatMap((child) => ['sdt', 'sdtContent', 'customXml'].includes(child.localName ?? '') ? bodyChildren(child) : [child]);
}

function sectionOfParagraphs(body: Element): Map<Element, number> {
  const mapping = new Map<Element, number>();
  let current = 0;
  for (const child of bodyChildren(body)) {
    if (child.localName === 'p') {
      mapping.set(child, current);
      if (children(children(child, 'pPr')[0] ?? child, 'sectPr').length) current++;
      continue;
    }
    if (child.localName === 'tbl') {
      for (const p of descendants(child, 'p')) mapping.set(p, current);
    }
  }
  return mapping;
}

function sectionSettings(body: Element, base: NoteSettings): Map<number, NoteSettings> {
  const result = new Map<number, NoteSettings>();
  let section = 0;
  let current = base;
  result.set(section, current);
  for (const item of bodyChildren(body)) {
    if (item.localName !== 'p') continue;
    const sectPr = children(children(item, 'pPr')[0] ?? item, 'sectPr')[0];
    if (!sectPr) continue;
    current = parseSectionNoteSettings(sectPr, current);
    result.set(section, current);
    section++;
    result.set(section, current);
  }
  const bodySectPr = children(body, 'sectPr')[0];
  if (bodySectPr) {
    current = parseSectionNoteSettings(bodySectPr, current);
    result.set(section, current);
  }
  return result;
}

function referenceRecords(body: Element): NoteReferenceRecord[] {
  const sectionMap = sectionOfParagraphs(body);
  const paragraphs = descendants(body, 'p');
  const records: NoteReferenceRecord[] = [];
  for (const [paragraphIndex, paragraph] of paragraphs.entries()) {
    for (const [runIndex, run] of ownRuns(paragraph).entries()) {
      const reference = noteReferenceInRun(run);
      if (!reference) continue;
      records.push({
        kind: reference.kind,
        id: reference.id,
        customMarkFollows: reference.customMarkFollows,
        paragraph: paragraphIndex,
        run: runIndex,
        section: sectionMap.get(paragraph) ?? 0,
      });
    }
  }
  return records;
}

interface ImageReadContext {
  relationships: Map<string, RelationshipTarget>;
  getContentType: (path: string) => string | undefined;
  sourcePartPath: string;
}

function readRun(run: Element, index: number, styles: StylesContext, paragraph: Element,
  paragraphIndex: number, imageContext?: ImageReadContext,
  noteNumber?: (kind: NoteKind, id: number) => { number: number; marker: string } | null): RunInfo {
  const direct = readRunProperties(children(run, 'rPr')[0], styles.theme);
  const revisions = readRunRevisionMarks(run, paragraph, styles.theme);
  const images = imageContext
    ? readRunImages(run, paragraphIndex, index, imageContext.relationships, imageContext.getContentType, imageContext.sourcePartPath)
    : [];
  const note = noteReferenceInRun(run);
  const resolved = note && noteNumber ? noteNumber(note.kind, note.id) : null;
  return {
    index,
    text: textOf(run),
    ...direct,
    revisions: revisions.length ? revisions.map(({ id, kind, author, date, move }) => ({
      id,
      kind,
      ...(author !== undefined ? { author } : {}),
      ...(date !== undefined ? { date } : {}),
      ...(move !== undefined ? { move } : {}),
    })) : undefined,
    effective: computeEffectiveRunFormat(styles, paragraph, run),
    hyperlink: runHyperlinkInfo(run, paragraph, imageContext?.relationships),
    images,
    image: images[0],
    noteReference: note && resolved ? { kind: note.kind, id: note.id, number: resolved.number, marker: resolved.marker } : undefined,
  };
}

function readParagraph(paragraph: Element, index: number, styles: StylesContext, numbering?: NumberingInfo,
  imageContext?: ImageReadContext,
  noteNumber?: (kind: NoteKind, id: number) => { number: number; marker: string } | null): ParagraphInfo {
  const direct = readParagraphProperties(children(paragraph, 'pPr')[0]);
  const paragraphRevision = readParagraphRevisionMark(paragraph, styles.theme);
  const runElements = ownRuns(paragraph);
  const runs = runElements.map((run, runIndex) => readRun(run, runIndex, styles, paragraph, index, imageContext, noteNumber));
  return {
    index,
    text: textOf(paragraph),
    ...direct,
    runs,
    paragraphRevision: paragraphRevision ? {
      id: paragraphRevision.id,
      kind: paragraphRevision.kind,
      ...(paragraphRevision.author !== undefined ? { author: paragraphRevision.author } : {}),
      ...(paragraphRevision.date !== undefined ? { date: paragraphRevision.date } : {}),
    } : undefined,
    effective: computeEffectiveParagraphFormat(styles, paragraph),
    numbering,
    images: runs.flatMap(run => run.images ?? []),
  };
}

function removeProperty(parent: Element, name: string): void {
  for (const child of children(parent, name)) parent.removeChild(child);
}

function removeWordAttribute(element: Element, name: string): void {
  element.removeAttributeNS(WORD_NS, name);
  element.removeAttribute(`w:${name}`);
}

function removeWordAttributes(element: Element, ...names: string[]): void {
  for (const name of names) removeWordAttribute(element, name);
}

function removeIfEmpty(element: Element | undefined): void {
  if (!element) return;
  if (!element.attributes.length && !element.firstChild) element.parentNode?.removeChild(element);
}

function removeEmptyParagraphWrappers(paragraph: Element): void {
  const removable = new Set(['hyperlink', 'sdt', 'sdtContent', 'smartTag', 'ins', 'del', 'moveFrom', 'moveTo', 'customXml']);
  const hasMeaningfulText = (element: Element): boolean => {
    for (let child = element.firstChild; child; child = child.nextSibling) {
      if (child.nodeType !== 3) continue;
      if ((child.nodeValue ?? '').trim().length) return true;
    }
    return false;
  };
  const walk = (node: Element): void => {
    for (const child of [...children(node)]) walk(child);
    if (node.namespaceURI !== WORD_NS) return;
    if (node.localName === 'pPr') {
      if (!node.attributes.length && !children(node).length && !hasMeaningfulText(node)) node.parentNode?.removeChild(node);
      return;
    }
    const localName = node.localName ?? '';
    if (!removable.has(localName)) return;
    if (localName === 'sdt') {
      if (!children(node, 'sdtContent')[0]) node.parentNode?.removeChild(node);
      return;
    }
    if (!children(node).length && !hasMeaningfulText(node)) node.parentNode?.removeChild(node);
  };
  walk(paragraph);
}

function setOnOff(parent: Element, name: string, value: boolean, onValue = '1', offValue = '0'): void {
  setWordValue(property(parent, name), value ? onValue : offValue);
}

function normalizeHexOrAuto(value: string): string {
  return value.toLowerCase() === 'auto' ? 'auto' : value.toUpperCase();
}

function writeBorderSide(element: Element, border: NonNullable<RunFormat['border']>): void {
  setWordValue(element, border.style);
  element.setAttributeNS(WORD_NS, 'w:sz', String(Math.max(0, Math.min(border.size, 2048))));
  element.setAttributeNS(WORD_NS, 'w:space', String(Math.max(0, border.space)));
  element.setAttributeNS(WORD_NS, 'w:color', normalizeHexOrAuto(border.color));
  if (border.shadow !== undefined) element.setAttributeNS(WORD_NS, 'w:shadow', border.shadow ? '1' : '0');
  else removeWordAttribute(element, 'shadow');
}

function writeShading(element: Element, shading: Shading): void {
  setWordValue(element, shading.pattern);
  element.setAttributeNS(WORD_NS, 'w:fill', normalizeHexOrAuto(shading.fill));
  if (shading.color !== undefined) element.setAttributeNS(WORD_NS, 'w:color', normalizeHexOrAuto(shading.color));
  else removeWordAttribute(element, 'color');
}

function normalizeTabsForWrite(tabs: TabStop[]): TabStop[] {
  const byPosition = new Map<number, TabStop>();
  for (const tab of tabs) {
    if (!Number.isFinite(tab.position)) continue;
    const position = Math.floor(tab.position);
    if (tab.alignment === 'clear') {
      byPosition.delete(position);
      continue;
    }
    byPosition.set(position, { ...tab, position });
  }
  return [...byPosition.values()].sort((a, b) => a.position - b.position);
}

function applyParagraphFormatTo(props: Element, format: ParagraphFormat): void {
  if ('style' in format) {
    if (format.style === null) removeProperty(props, 'pStyle');
    else if (format.style !== undefined) setWordValue(property(props, 'pStyle'), format.style);
  }
  if ('alignment' in format) {
    if (format.alignment === null) removeProperty(props, 'jc');
    else if (format.alignment !== undefined) setWordValue(property(props, 'jc'), format.alignment);
  }
  for (const [key, tag] of [
    ['keepNext', 'keepNext'],
    ['keepLines', 'keepLines'],
    ['pageBreakBefore', 'pageBreakBefore'],
    ['widowControl', 'widowControl'],
    ['suppressLineNumbers', 'suppressLineNumbers'],
    ['suppressAutoHyphens', 'suppressAutoHyphens'],
  ] as const) {
    if (!(key in format)) continue;
    if (format[key] === null) removeProperty(props, tag);
    else if (format[key] !== undefined) setOnOff(props, tag, format[key]!);
  }
  if ('tabs' in format) {
    if (format.tabs === null || (Array.isArray(format.tabs) && format.tabs.length === 0)) {
      removeProperty(props, 'tabs');
    } else if (format.tabs !== undefined) {
      const tabsElement = property(props, 'tabs');
      for (const child of children(tabsElement, 'tab')) tabsElement.removeChild(child);
      for (const tab of normalizeTabsForWrite(format.tabs)) {
        const entry = wordElement(tabsElement.ownerDocument!, 'tab');
        setWordValue(entry, tab.alignment);
        entry.setAttributeNS(WORD_NS, 'w:pos', String(tab.position));
        if (tab.leader !== undefined) entry.setAttributeNS(WORD_NS, 'w:leader', tab.leader);
        else removeWordAttribute(entry, 'leader');
        tabsElement.appendChild(entry);
      }
      removeIfEmpty(tabsElement);
    }
  }
  if ('borders' in format) {
    if (format.borders === null) {
      removeProperty(props, 'pBdr');
    } else if (format.borders !== undefined) {
      const borderElement = children(props, 'pBdr')[0] ?? property(props, 'pBdr');
      for (const child of children(borderElement)) borderElement.removeChild(child);
      for (const side of ['top', 'left', 'bottom', 'right', 'between', 'bar'] as const) {
        const border = format.borders[side];
        if (!border) continue;
        const child = property(borderElement, side);
        writeBorderSide(child, border);
      }
      removeIfEmpty(borderElement);
    }
  }
  if ('shading' in format) {
    if (format.shading === null) removeProperty(props, 'shd');
    else if (format.shading !== undefined) writeShading(property(props, 'shd'), format.shading);
  }
  if (['indentLeft', 'indentRight', 'indentFirstLine', 'indentHanging'].some(key => key in format)) {
    const indent = children(props, 'ind')[0] ?? property(props, 'ind');
    if ('indentLeft' in format) {
      if (format.indentLeft === null) removeWordAttributes(indent, 'left', 'start');
      else if (format.indentLeft !== undefined) {
        removeWordAttribute(indent, 'start');
        indent.setAttributeNS(WORD_NS, 'w:left', String(format.indentLeft));
      }
    }
    if ('indentRight' in format) {
      if (format.indentRight === null) removeWordAttributes(indent, 'right', 'end');
      else if (format.indentRight !== undefined) {
        removeWordAttribute(indent, 'end');
        indent.setAttributeNS(WORD_NS, 'w:right', String(format.indentRight));
      }
    }
    if ('indentFirstLine' in format) {
      if (format.indentFirstLine === null) removeWordAttribute(indent, 'firstLine');
      else if (format.indentFirstLine !== undefined) indent.setAttributeNS(WORD_NS, 'w:firstLine', String(format.indentFirstLine));
    }
    if ('indentHanging' in format) {
      if (format.indentHanging === null) removeWordAttribute(indent, 'hanging');
      else if (format.indentHanging !== undefined) indent.setAttributeNS(WORD_NS, 'w:hanging', String(format.indentHanging));
    }
    removeIfEmpty(indent);
  }
  if (['spacingBefore', 'spacingAfter', 'lineSpacing', 'lineSpacingRule'].some(key => key in format)) {
    const spacing = children(props, 'spacing')[0] ?? property(props, 'spacing');
    if ('spacingBefore' in format) {
      if (format.spacingBefore === null) removeWordAttribute(spacing, 'before');
      else if (format.spacingBefore !== undefined) spacing.setAttributeNS(WORD_NS, 'w:before', String(format.spacingBefore));
    }
    if ('spacingAfter' in format) {
      if (format.spacingAfter === null) removeWordAttribute(spacing, 'after');
      else if (format.spacingAfter !== undefined) spacing.setAttributeNS(WORD_NS, 'w:after', String(format.spacingAfter));
    }
    if ('lineSpacing' in format) {
      if (format.lineSpacing === null) removeWordAttribute(spacing, 'line');
      else if (format.lineSpacing !== undefined) spacing.setAttributeNS(WORD_NS, 'w:line', String(format.lineSpacing));
    }
    if ('lineSpacingRule' in format) {
      if (format.lineSpacingRule === null) removeWordAttribute(spacing, 'lineRule');
      else if (format.lineSpacingRule !== undefined) spacing.setAttributeNS(WORD_NS, 'w:lineRule', format.lineSpacingRule);
    }
    removeIfEmpty(spacing);
  }
  if ('outlineLevel' in format) {
    if (format.outlineLevel === null) removeProperty(props, 'outlineLvl');
    else if (format.outlineLevel !== undefined) setWordValue(property(props, 'outlineLvl'), String(format.outlineLevel));
  }
}

function applyRunFormatTo(props: Element, format: RunFormat): void {
  if ('style' in format) {
    if (format.style === null) removeProperty(props, 'rStyle');
    else if (format.style !== undefined) setWordValue(property(props, 'rStyle'), format.style);
  }
  for (const [key, tag] of [
    ['bold', 'b'],
    ['italic', 'i'],
    ['strike', 'strike'],
    ['doubleStrike', 'dstrike'],
    ['smallCaps', 'smallCaps'],
    ['allCaps', 'caps'],
  ] as const) {
    if (!(key in format)) continue;
    if (format[key] === null) removeProperty(props, tag);
    else if (format[key] !== undefined) setOnOff(props, tag, format[key]!);
  }
  if ('underline' in format || 'underlineStyle' in format || 'underlineColor' in format) {
    if (format.underline === null || (format.underlineStyle === null &&
        format.underline === undefined && format.underlineColor === undefined)) {
      removeProperty(props, 'u');
    } else {
      const underline = children(props, 'u')[0] ?? property(props, 'u');
      if ('underlineColor' in format) {
        if (format.underlineColor === null) removeWordAttribute(underline, 'color');
        else if (format.underlineColor !== undefined) underline.setAttributeNS(WORD_NS, 'w:color', format.underlineColor);
      }
      if ('underlineStyle' in format && format.underlineStyle === null) removeWordAttribute(underline, 'val');
      if (format.underline !== undefined || format.underlineStyle !== undefined) {
        const value = format.underlineStyle ?? (format.underline ? 'single' : format.underline === false ? 'none' : undefined);
        if (value !== undefined && value !== null) setWordValue(underline, value);
      }
      removeIfEmpty(underline);
    }
  }
  if ('fontSize' in format) {
    if (format.fontSize === null) {
      removeProperty(props, 'sz');
      removeProperty(props, 'szCs');
    } else if (format.fontSize !== undefined) {
      setWordValue(property(props, 'sz'), String(format.fontSize * 2));
      setWordValue(property(props, 'szCs'), String(format.fontSize * 2));
    }
  }
  if ('color' in format) {
    if (format.color === null) removeProperty(props, 'color');
    else if (format.color !== undefined) setWordValue(property(props, 'color'), format.color);
  }
  if ('fontFamily' in format || 'fontFamilyEastAsia' in format) {
    const fonts = children(props, 'rFonts')[0] ?? property(props, 'rFonts');
    if ('fontFamily' in format) {
      if (format.fontFamily === null) {
        for (const name of ['ascii', 'hAnsi', 'cs', 'eastAsia']) removeWordAttribute(fonts, name);
      } else if (format.fontFamily !== undefined) {
        for (const name of ['ascii', 'hAnsi', 'cs']) fonts.setAttributeNS(WORD_NS, `w:${name}`, format.fontFamily);
        if (format.fontFamilyEastAsia === undefined) fonts.setAttributeNS(WORD_NS, 'w:eastAsia', format.fontFamily);
      }
    }
    if ('fontFamilyEastAsia' in format) {
      if (format.fontFamilyEastAsia === null) removeWordAttribute(fonts, 'eastAsia');
      else if (format.fontFamilyEastAsia !== undefined) fonts.setAttributeNS(WORD_NS, 'w:eastAsia', format.fontFamilyEastAsia);
    }
    removeIfEmpty(fonts);
  }
  if ('verticalAlign' in format) {
    if (format.verticalAlign === null || format.verticalAlign === 'baseline') removeProperty(props, 'vertAlign');
    else if (format.verticalAlign !== undefined) setWordValue(property(props, 'vertAlign'), format.verticalAlign);
  }
  if ('highlight' in format) {
    if (format.highlight === null || format.highlight === 'none') removeProperty(props, 'highlight');
    else if (format.highlight !== undefined) setWordValue(property(props, 'highlight'), format.highlight);
  }
  if ('characterSpacing' in format) {
    if (format.characterSpacing === null) removeProperty(props, 'spacing');
    else if (format.characterSpacing !== undefined) setWordValue(property(props, 'spacing'), String(format.characterSpacing));
  }
  if ('border' in format) {
    if (format.border === null) removeProperty(props, 'bdr');
    else if (format.border !== undefined) writeBorderSide(property(props, 'bdr'), format.border);
  }
  if ('shading' in format) {
    if (format.shading === null) removeProperty(props, 'shd');
    else if (format.shading !== undefined) writeShading(property(props, 'shd'), format.shading);
  }
  removeIfEmpty(props);
}

function rejectNullFormatValues(format: ParagraphFormat | RunFormat, label: string): void {
  for (const [key, value] of Object.entries(format)) {
    if (value === null) throw new Error(`${label}.${key} cannot be null in defineStyle().`);
  }
}

function setWordAttr(element: Element, name: string, value: string | number): void {
  element.setAttributeNS(WORD_NS, `w:${name}`, String(value));
}

function appendWordValueElement(parent: Element, name: string, value?: string | number): Element {
  const element = wordElement(parent.ownerDocument!, name);
  if (value !== undefined) setWordValue(element, String(value));
  parent.appendChild(element);
  return element;
}

function defaultNumberingDefinition(kind: 'bullet' | 'decimal' | 'multilevel'): Omit<NumberingDefinition, 'numId' | 'abstractNumId'> {
  const makeLevel = (level: number, format: string, text: string, fontFamily?: string): NumberingDefinition['levels'][number] => ({
    level,
    start: 1,
    format,
    text,
    suffix: 'tab',
    justification: level === 0 ? 'left' : undefined,
    indentLeft: 720 * (level + 1),
    indentHanging: 360,
    runFormat: fontFamily ? { fontFamily } : undefined,
  });
  if (kind === 'bullet') {
    return {
      multiLevelType: 'hybridMultilevel',
      levels: Array.from({ length: 9 }, (_, level) => makeLevel(level, 'bullet', level === 0 ? '•' : level % 2 ? '◦' : '▪', 'Symbol')),
    };
  }
  if (kind === 'decimal') {
    return {
      multiLevelType: 'multilevel',
      levels: Array.from({ length: 9 }, (_, level) => makeLevel(level, 'decimal', `%${level + 1}.`)),
    };
  }
  return {
    multiLevelType: 'multilevel',
    levels: Array.from({ length: 9 }, (_, level) => makeLevel(level, 'decimal', Array.from({ length: level + 1 }, (_, index) => `%${index + 1}`).join('.') + '.')),
  };
}

function appendRunProperties(parent: Element, format: RunFormat | undefined): void {
  if (!format) return;
  const props = wordElement(parent.ownerDocument!, 'rPr');
  if (format.fontFamily) {
    const fonts = wordElement(parent.ownerDocument!, 'rFonts');
    for (const name of ['ascii', 'hAnsi', 'eastAsia', 'cs']) fonts.setAttributeNS(WORD_NS, `w:${name}`, format.fontFamily);
    props.appendChild(fonts);
  }
  if (format.bold !== undefined) setWordValue(appendWordValueElement(props, 'b'), format.bold ? '1' : '0');
  if (format.italic !== undefined) setWordValue(appendWordValueElement(props, 'i'), format.italic ? '1' : '0');
  if (format.color) appendWordValueElement(props, 'color', format.color);
  if (format.fontSize !== undefined && format.fontSize !== null) {
    appendWordValueElement(props, 'sz', format.fontSize * 2);
    appendWordValueElement(props, 'szCs', format.fontSize * 2);
  }
  if (format.underline !== undefined) appendWordValueElement(props, 'u', format.underline ? 'single' : 'none');
  if (props.childNodes.length) parent.appendChild(props);
}

function buildLevelElement(document: Document, definition: NumberingDefinition['levels'][number]): Element {
  const level = wordElement(document, 'lvl');
  setWordAttr(level, 'ilvl', definition.level);
  appendWordValueElement(level, 'start', definition.start ?? 1);
  appendWordValueElement(level, 'numFmt', definition.format);
  if (definition.restart !== undefined) appendWordValueElement(level, 'lvlRestart', definition.restart);
  if (definition.paragraphStyle) appendWordValueElement(level, 'pStyle', definition.paragraphStyle);
  if (definition.isLegal) appendWordValueElement(level, 'isLgl');
  appendWordValueElement(level, 'suff', definition.suffix);
  appendWordValueElement(level, 'lvlText', definition.text);
  if (definition.justification) appendWordValueElement(level, 'lvlJc', definition.justification);
  if (definition.indentLeft !== undefined || definition.indentHanging !== undefined) {
    const props = wordElement(document, 'pPr');
    const ind = wordElement(document, 'ind');
    if (definition.indentLeft !== undefined) setWordAttr(ind, 'left', definition.indentLeft);
    if (definition.indentHanging !== undefined) setWordAttr(ind, 'hanging', definition.indentHanging);
    props.appendChild(ind);
    level.appendChild(props);
  }
  appendRunProperties(level, definition.runFormat);
  return level;
}

function insertNumberingNode(root: Element, node: Element): void {
  const childrenInRoot = children(root);
  if (node.localName === 'abstractNum') {
    root.insertBefore(node, childrenInRoot.find(child => ['num', 'numIdMacAtCleanup'].includes(child.localName ?? '')) ?? null);
    return;
  }
  if (node.localName === 'num') {
    root.insertBefore(node, childrenInRoot.find(child => child.localName === 'numIdMacAtCleanup') ?? null);
    return;
  }
  root.appendChild(node);
}

function bodyBlocks(body: Element): Element[] {
  return children(body).filter(child => ['p', 'tbl'].includes(child.localName ?? ''));
}

function tableRows(table: Element): Element[] {
  return childrenThroughTransparent(table, 'tr');
}

function tableAt(document: Document, index: number): Element {
  assertIndex(index);
  const table = childrenThroughTransparent(bodyOf(document), 'tbl')[index];
  if (!table) throw new Error(`Table ${index} does not exist.`);
  return table;
}

function bodyBlockAt(document: Document, index: number): Element {
  assertIndex(index);
  const block = bodyBlocks(bodyOf(document))[index];
  if (!block) throw new Error(`Block ${index} does not exist.`);
  return block;
}

function tableProperty(parent: Element, name: 'tblPr' | 'trPr' | 'tcPr'): Element {
  let result = children(parent, name)[0];
  if (!result) {
    result = wordElement(parent.ownerDocument!, name);
    parent.insertBefore(result, parent.firstChild);
  }
  return result;
}

function ensureTableGrid(table: Element): Element {
  let grid = children(table, 'tblGrid')[0];
  if (!grid) {
    grid = wordElement(table.ownerDocument!, 'tblGrid');
    const tableProps = children(table, 'tblPr')[0];
    table.insertBefore(grid, tableProps?.nextSibling ?? tableRows(table)[0] ?? null);
    for (const width of tableGrid(table)) grid.appendChild(gridCol(table.ownerDocument!, width));
  } else if (!children(grid, 'gridCol').length) {
    for (const width of tableGrid(table)) grid.appendChild(gridCol(table.ownerDocument!, width));
  }
  return grid;
}

function removeWordChildren(parent: Element | undefined, ...names: string[]): void {
  if (!parent) return;
  for (const child of children(parent)) {
    if (names.includes(child.localName!)) parent.removeChild(child);
  }
}

function ensureCellParagraph(cell: Element): void {
  if (!childrenThroughTransparent(cell, 'p').length) cell.appendChild(newParagraph(cell.ownerDocument!, ''));
}

function blankCell(document: Document): Element {
  const cell = wordElement(document, 'tc');
  cell.appendChild(newParagraph(document, ''));
  return cell;
}

function gridCol(document: Document, width = 2250): Element {
  const column = wordElement(document, 'gridCol');
  column.setAttributeNS(WORD_NS, 'w:w', String(width));
  return column;
}

function widthValue(parent: Element, name: string, value: { type: 'auto' | 'dxa' | 'pct'; value: number } | undefined): void {
  removeWordChildren(parent, name);
  if (!value) return;
  const width = property(parent, name);
  width.setAttributeNS(WORD_NS, 'w:type', value.type);
  width.setAttributeNS(WORD_NS, 'w:w', String(value.value));
}

function boolValue(parent: Element, name: string, value: boolean | undefined): void {
  removeWordChildren(parent, name);
  if (value === undefined) return;
  const element = property(parent, name);
  if (!value) setWordValue(element, '0');
}

function valueElement(parent: Element, name: string, value: string | undefined): void {
  removeWordChildren(parent, name);
  if (value === undefined) return;
  const element = property(parent, name);
  setWordValue(element, value);
}

function mergeElement(parent: Element, name: 'gridSpan' | 'vMerge' | 'hMerge', value: number | 'restart' | 'continue' | undefined): void {
  removeWordChildren(parent, name);
  if (value === undefined) return;
  const element = property(parent, name);
  if (typeof value === 'number') setWordValue(element, String(value));
  else if (value === 'restart') setWordValue(element, value);
}

function setBorders(parent: Element, name: 'tblBorders' | 'tcBorders', borders: TableFormat['borders'] | CellFormat['borders'] | undefined): void {
  removeWordChildren(parent, name);
  if (!borders) return;
  const element = property(parent, name);
  for (const side of ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'] as const) {
    const border = borders[side];
    if (!border) continue;
    const child = property(element, side);
    if (border.style) setWordValue(child, border.style);
    if (border.size !== undefined) child.setAttributeNS(WORD_NS, 'w:sz', String(border.size));
    if (border.space !== undefined) child.setAttributeNS(WORD_NS, 'w:space', String(border.space));
    if (border.color) child.setAttributeNS(WORD_NS, 'w:color', border.color);
    if (border.none && !border.style) setWordValue(child, 'nil');
  }
}

function setShading(parent: Element, shading: TableFormat['shading'] | CellFormat['shading'] | undefined): void {
  removeWordChildren(parent, 'shd');
  if (!shading) return;
  const element = property(parent, 'shd');
  if (shading.fill) element.setAttributeNS(WORD_NS, 'w:fill', shading.fill);
  if (shading.color) element.setAttributeNS(WORD_NS, 'w:color', shading.color);
  if (shading.value) setWordValue(element, shading.value);
}

function setMargins(parent: Element, name: 'tblCellMar' | 'tcMar', margin: TableFormat['cellMargin'] | CellFormat['margin'] | undefined): void {
  removeWordChildren(parent, name);
  if (!margin) return;
  const element = property(parent, name);
  for (const side of ['top', 'left', 'bottom', 'right'] as const) {
    const value = margin[side];
    if (!value) continue;
    const child = property(element, side);
    child.setAttributeNS(WORD_NS, 'w:type', value.type);
    child.setAttributeNS(WORD_NS, 'w:w', String(value.value));
  }
}

function setTableFormat(tbl: Element, format: TableFormat): void {
  const props = tableProperty(tbl, 'tblPr');
  if (format.width !== undefined) widthValue(props, 'tblW', format.width);
  if (format.alignment !== undefined) valueElement(props, 'jc', format.alignment);
  if (format.indent !== undefined) {
    removeWordChildren(props, 'tblInd');
    const ind = property(props, 'tblInd');
    ind.setAttributeNS(WORD_NS, 'w:w', String(format.indent));
    ind.setAttributeNS(WORD_NS, 'w:type', 'dxa');
  }
  if (format.borders !== undefined) setBorders(props, 'tblBorders', format.borders);
  if (format.shading !== undefined) setShading(props, format.shading);
  if (format.cellMargin !== undefined) setMargins(props, 'tblCellMar', format.cellMargin);
  if (format.layout !== undefined) valueElement(props, 'tblLayout', format.layout);
  if (format.style !== undefined) valueElement(props, 'tblStyle', format.style);
  if (format.look !== undefined) valueElement(props, 'tblLook', format.look);
  if (format.caption !== undefined) valueElement(props, 'tblCaption', format.caption);
  if (format.description !== undefined) valueElement(props, 'tblDescription', format.description);
}

function setRowFormat(row: Element, format: RowFormat): void {
  const props = tableProperty(row, 'trPr');
  if (format.height !== undefined) {
    removeWordChildren(props, 'trHeight');
    const height = property(props, 'trHeight');
    height.setAttributeNS(WORD_NS, 'w:val', String(format.height.value));
    if (format.height.rule) height.setAttributeNS(WORD_NS, 'w:hRule', format.height.rule);
  }
  if (format.cantSplit !== undefined) boolValue(props, 'cantSplit', format.cantSplit);
  if (format.header !== undefined) boolValue(props, 'tblHeader', format.header);
  if (format.alignment !== undefined) valueElement(props, 'jc', format.alignment);
  if (format.deleted !== undefined) {
    removeWordChildren(props, 'del');
    if (format.deleted) markRevision(props, 'del', format.revision?.author, format.revision?.date);
  }
  if (format.inserted !== undefined) {
    removeWordChildren(props, 'ins');
    if (format.inserted) markRevision(props, 'ins', format.revision?.author, format.revision?.date);
  }
}

function setCellFormat(cell: Element, format: CellFormat): void {
  const props = tableProperty(cell, 'tcPr');
  if (format.width !== undefined) widthValue(props, 'tcW', format.width);
  if (format.borders !== undefined) setBorders(props, 'tcBorders', format.borders);
  if (format.shading !== undefined) setShading(props, format.shading);
  if (format.margin !== undefined) setMargins(props, 'tcMar', format.margin);
  if (format.verticalAlign !== undefined) valueElement(props, 'vAlign', format.verticalAlign);
  if (format.textDirection !== undefined) valueElement(props, 'textDirection', format.textDirection);
  if (format.noWrap !== undefined) boolValue(props, 'noWrap', format.noWrap);
  if (format.hideMark !== undefined) boolValue(props, 'hideMark', format.hideMark);
  if (format.hMerge !== undefined) mergeElement(props, 'hMerge', format.hMerge);
  if (format.vMerge !== undefined) mergeElement(props, 'vMerge', format.vMerge);
}

type XmlCellRef = { cell: Element; row: number; start: number; colSpan: number; rowSpan: number; isContinuation: boolean };

function tableModel(table: Element): { rows: Element[]; grid: number[]; matrix: XmlCellRef[][]; refs: XmlCellRef[] } {
  const rows = tableRows(table);
  const grid = tableGrid(table);
  const matrix: XmlCellRef[][] = Array.from({ length: rows.length }, () => []);
  const refs: XmlCellRef[] = [];
  const active = new Map<number, { end: number; master: XmlCellRef }>();
  rows.forEach((row, rowIndex) => {
    const nextActive = new Map<number, { end: number; master: XmlCellRef }>();
    for (const position of rowCells(row)) {
      const ref: XmlCellRef = {
        cell: position.cell,
        row: rowIndex,
        start: position.start,
        colSpan: position.span,
        rowSpan: 1,
        isContinuation: false,
      };
      const activeMerge = active.get(position.start);
      const sameMerge = activeMerge && activeMerge.end === position.start + position.span;
      if (position.vMerge === 'continue' && sameMerge) {
        activeMerge.master.rowSpan += 1;
        ref.rowSpan = 0;
        ref.isContinuation = true;
        nextActive.set(position.start, activeMerge);
        for (let col = position.start; col < position.start + position.span; col++) matrix[rowIndex]![col] = activeMerge.master;
      } else {
        refs.push(ref);
        if (position.vMerge === 'restart') nextActive.set(position.start, { end: position.start + position.span, master: ref });
        for (let col = position.start; col < position.start + position.span; col++) matrix[rowIndex]![col] = ref;
      }
    }
    active.clear();
    for (const [start, merge] of nextActive) active.set(start, merge);
  });
  return { rows, grid, matrix, refs };
}

function cellAt(table: Element, row: number, col: number): XmlCellRef {
  assertIndex(row); assertIndex(col);
  const model = tableModel(table);
  const cell = model.matrix[row]?.[col];
  if (!cell) throw new Error(`Cell ${row},${col} does not exist.`);
  return cell;
}

function clearCellContent(cell: Element): void {
  for (let child = cell.firstChild; child;) {
    const next = child.nextSibling;
    if (!(child.nodeType === 1 && (child as Element).namespaceURI === WORD_NS && (child as Element).localName === 'tcPr')) {
      cell.removeChild(child);
    }
    child = next;
  }
  ensureCellParagraph(cell);
}

function appendCellContent(target: Element, source: Element): void {
  for (const child of children(source).filter(node => node.localName !== 'tcPr')) target.appendChild(child);
  ensureCellParagraph(target);
}

function buildTable(document: Document, rows: number, cols: number, format?: TableFormat, texts?: string[][]): Element {
  const table = wordElement(document, 'tbl');
  const tableProps = wordElement(document, 'tblPr');
  const grid = wordElement(document, 'tblGrid');
  for (let i = 0; i < cols; i++) grid.appendChild(gridCol(document, Math.floor(9000 / cols)));
  table.appendChild(tableProps);
  table.appendChild(grid);
  if (format) setTableFormat(table, format);
  for (let rowIndex = 0; rowIndex < rows; rowIndex++) {
    const tr = wordElement(document, 'tr');
    for (let colIndex = 0; colIndex < cols; colIndex++) {
      const cell = blankCell(document);
      const paragraph = children(cell, 'p')[0]!;
      replaceSpan(paragraph, 0, 0, texts?.[rowIndex]?.[colIndex] ?? '');
      tr.appendChild(cell);
    }
    table.appendChild(tr);
  }
  return table;
}

function repairVerticalMerges(table: Element): void {
  const active = new Map<number, { end: number }>();
  for (const row of tableRows(table)) {
    const nextActive = new Map<number, { end: number }>();
    for (const position of rowCells(row)) {
      const props = tableProperty(position.cell, 'tcPr');
      const merge = children(props, 'vMerge')[0];
      const continuing = merge && (wordValue(merge) ?? 'continue') === 'continue';
      const sameMerge = active.get(position.start)?.end === position.start + position.span;
      if (continuing && !sameMerge) {
        mergeElement(props, 'vMerge', 'restart');
        nextActive.set(position.start, { end: position.start + position.span });
      } else if (merge && (wordValue(merge) === 'restart' || (continuing && sameMerge))) {
        nextActive.set(position.start, { end: position.start + position.span });
      }
    }
    active.clear();
    for (const [start, merge] of nextActive) active.set(start, merge);
  }


}

function deleteTableRowElement(table: Element, at: number): void {
  const rows = tableRows(table);
  const row = rows[at];
  if (!row) throw new Error(`Row ${at} does not exist.`);
  if (rows.length <= 1) throw new Error('Cannot delete the only table row.');
  const next = rows[at + 1];
  if (next) {
    const nextPositions = new Map(rowCells(next).map(position => [position.start, position]));
    for (const position of rowCells(row)) {
      if (position.vMerge === 'restart') {
        const continuation = nextPositions.get(position.start);
        if (continuation?.vMerge === 'continue') {
          const props = tableProperty(continuation.cell, 'tcPr');
          mergeElement(props, 'vMerge', 'restart');
        }
      }
    }
  }
  row.parentNode!.removeChild(row);
  repairVerticalMerges(table);
}

function deleteTableColumnElement(table: Element, at: number): void {
  const grid = ensureTableGrid(table);
  const columns = children(grid, 'gridCol');
  if (!columns[at]) throw new Error(`Column ${at} does not exist.`);
  if (columns.length <= 1) throw new Error('Cannot delete the only table column.');
  const model = tableModel(table);
  grid.removeChild(columns[at]!);
  const handled = new Set<Element>();
  for (const [rowIndex, row] of tableRows(table).entries()) {
    const covering = model.matrix[rowIndex]?.[at];
    if (covering && covering.rowSpan > 1) {
      if (handled.has(covering.cell)) continue;
      handled.add(covering.cell);
      for (let index = covering.row; index < covering.row + covering.rowSpan; index++) {
        const chain = rowCells(model.rows[index]!).find(position => position.start === covering.start);
        if (!chain) continue;
        if (chain.span > 1) {
          const props = tableProperty(chain.cell, 'tcPr');
          mergeElement(props, 'gridSpan', chain.span - 1 > 1 ? chain.span - 1 : undefined);
        } else {
          model.rows[index]!.removeChild(chain.cell);
        }
      }
      continue;
    }
    const position = rowCells(row).find(cell => at >= cell.start && at < cell.start + cell.span);
    if (!position) throw new Error(`Column ${at} does not exist.`);
    if (position.span > 1) {
      const props = tableProperty(position.cell, 'tcPr');
      mergeElement(props, 'gridSpan', position.span - 1 > 1 ? position.span - 1 : undefined);
    } else {
      row.removeChild(position.cell);
    }
  }
}

function setOptionalAttribute(element: Element, name: string, value: string | undefined): void {
  if (value === undefined || value === '') element.removeAttribute(name);
  else element.setAttribute(name, value);
}

function imageElementForRun(run: Element, relationshipId: string, ordinal = 0): Element | undefined {
  let index = 0;
  for (let child = run.firstChild; child; child = child.nextSibling) {
    if (child.nodeType !== 1) continue;
    const element = child as Element;
    const ids = element.namespaceURI === WORD_NS && element.localName === 'drawing'
      ? Array.from(element.getElementsByTagNameNS(A_NS, 'blip')).map((blip) =>
        blip.getAttributeNS(OFFICE_REL_NS, 'embed') ?? blip.getAttributeNS(OFFICE_REL_NS, 'link') ??
        blip.getAttribute('r:embed') ?? blip.getAttribute('r:link'))
      : element.namespaceURI === WORD_NS && element.localName === 'pict'
        ? Array.from(element.getElementsByTagNameNS(V_NS, 'imagedata')).map((node) =>
          node.getAttributeNS(OFFICE_REL_NS, 'id') ?? node.getAttribute('r:id'))
        : [];
    for (const id of ids) {
      if (index === ordinal && id === relationshipId) return element;
      index++;
    }
  }
  return undefined;
}

function paragraphDirectChild(paragraph: Element, node: Node): Node {
  let current: Node | null = node;
  while (current?.parentNode && current.parentNode !== paragraph) current = current.parentNode;
  if (!current || current.parentNode !== paragraph) throw new Error('Target run is not inside the requested paragraph.');
  return current;
}

function cloneRunShell(run: Element): Element {
  const clone = wordElement(run.ownerDocument!, 'r');
  const props = children(run, 'rPr')[0];
  if (props) clone.appendChild(props.cloneNode(true));
  return clone;
}

function isolateRunChild(run: Element, child: Element): Element {
  const content = [...run.childNodes].filter((node) => !(node.nodeType === 1 && (node as Element).localName === 'rPr'));
  const targetIndex = content.indexOf(child);
  if (targetIndex === -1) throw new Error('Target child is not inside the requested run.');
  if (content.length === 1) return run;
  const parent = run.parentNode;
  if (!parent) throw new Error('Run is detached.');
  const isolated = cloneRunShell(run);
  if (targetIndex > 0) {
    parent.insertBefore(isolated, run.nextSibling);
    if (targetIndex < content.length - 1) {
      const after = cloneRunShell(run);
      for (const node of content.slice(targetIndex + 1)) after.appendChild(node);
      parent.insertBefore(after, isolated.nextSibling);
    }
  } else {
    parent.insertBefore(isolated, run);
  }
  isolated.appendChild(child);
  return isolated;
}

function nearestParagraph(node: Node | null): Element | null {
  let current = node;
  while (current) {
    if (current.nodeType === 1 && (current as Element).namespaceURI === WORD_NS && (current as Element).localName === 'p') {
      return current as Element;
    }
    current = current.parentNode;
  }
  return null;
}

function textRangeLength(paragraph: Element, start: number, end: number): void {
  const size = textOf(paragraph).length;
  if (!Number.isSafeInteger(start) || !Number.isSafeInteger(end) || start < 0 || end < start || end > size) {
    throw new Error(`Range [${start}, ${end}) is out of bounds for paragraph text length ${size}.`);
  }
}

const RUN_FORMAT_FIELDS = [
  'style', 'bold', 'italic', 'underline', 'underlineStyle', 'underlineColor', 'fontSize', 'fontFamily',
  'fontFamilyEastAsia', 'color', 'strike', 'doubleStrike', 'verticalAlign', 'smallCaps', 'allCaps',
  'highlight', 'characterSpacing', 'border', 'shading',
] as const satisfies readonly (keyof RunFormat)[];
const PARAGRAPH_FORMAT_FIELDS = [
  'alignment', 'style', 'indentLeft', 'indentRight', 'indentFirstLine', 'indentHanging',
  'spacingBefore', 'spacingAfter', 'lineSpacing', 'lineSpacingRule', 'keepNext', 'keepLines',
  'pageBreakBefore', 'widowControl', 'suppressLineNumbers', 'suppressAutoHyphens',
  'outlineLevel', 'tabs', 'borders', 'shading',
] as const satisfies readonly (keyof ParagraphFormat)[];

interface CompareBlockInfo {
  kind: 'paragraph' | 'table';
  text: string;
  opaque: boolean;
  exactKey: string;
  element: Element;
}

type CompareStep =
  | { kind: 'equal'; baseIndex: number; revisedIndex: number }
  | { kind: 'modify'; baseIndex: number; revisedIndex: number }
  | { kind: 'delete'; baseIndex: number }
  | { kind: 'insert'; revisedIndex: number };

const PARAGRAPH_DIRECT_FIELDS = [
  'alignment', 'style', 'indentLeft', 'indentRight', 'indentFirstLine', 'indentHanging',
  'spacingBefore', 'spacingAfter', 'lineSpacing', 'lineSpacingRule', 'keepNext', 'keepLines',
  'pageBreakBefore', 'widowControl', 'suppressLineNumbers', 'suppressAutoHyphens',
  'outlineLevel', 'tabs', 'borders', 'shading',
] as const satisfies readonly (keyof ParagraphFormat)[];

const OUTLINE_MAX_LEVEL = 8;

function cloneRunFormatValue<T>(value: T): T {
  if (!value || typeof value !== 'object') return value;
  return JSON.parse(JSON.stringify(value)) as T;
}

function runFormatValueEqual(a: unknown, b: unknown): boolean {
  if (a === b) return true;
  if (!a || !b || typeof a !== 'object' || typeof b !== 'object') return false;
  return JSON.stringify(a) === JSON.stringify(b);
}

function headingOutlineLevel(name: string | undefined): number | undefined {
  if (!name) return undefined;
  const match = /^heading\s+([1-9])$/i.exec(name.trim());
  if (!match) return undefined;
  return Number(match[1]) - 1;
}

function cloneOutlineNode(node: OutlineNode): OutlineNode {
  return {
    paragraph: node.paragraph,
    level: node.level,
    text: node.text,
    ...(node.styleId !== undefined ? { styleId: node.styleId } : {}),
    children: node.children.map(cloneOutlineNode),
  };
}

function directParagraphFormatOf(paragraph: Element): ParagraphFormat {
  return compactDefined(readParagraphProperties(children(paragraph, 'pPr')[0]));
}

function directRunFormatOf(run: Element, theme: StylesContext['theme']): RunFormat {
  return compactDefined(readRunProperties(children(run, 'rPr')[0], theme));
}

function buildFormatPatch<T extends object>(
  source: Partial<T>,
  target: Partial<T>,
  fields: readonly (keyof T)[],
): Partial<Record<keyof T, unknown>> {
  const patch: Partial<Record<keyof T, unknown>> = {};
  for (const field of fields) {
    const sourceValue = source[field];
    const targetValue = target[field];
    if (runFormatValueEqual(sourceValue, targetValue)) continue;
    patch[field] = targetValue === undefined ? null : cloneRunFormatValue(targetValue);
  }
  return patch;
}

function hasOpaqueParagraphContent(paragraph: Element): boolean {
  return descendants(paragraph, 'drawing').length > 0 ||
    descendants(paragraph, 'pict').length > 0 ||
    descendants(paragraph, 'object').length > 0;
}

function hasUnalignableCompareRuns(paragraph: Element): boolean {
  return ownRuns(paragraph).some((run) => {
    if (textOf(run).length !== 0) return false;
    return children(run, 'rPr').length > 0 || children(run).some((child) => child.localName !== 'rPr');
  });
}

function preservesParagraphMarkers(node: Element): boolean {
  if (isParagraphAnchorMarker(node)) return true;
  if (node.namespaceURI !== WORD_NS) return false;
  if (node.localName === 'sdt') {
    return children(node).every((child) => ['sdtPr', 'sdtEndPr'].includes(child.localName ?? '') || preservesParagraphMarkers(child));
  }
  if (!isTransparentWordWrapper(node)) return false;
  return children(node).every((child) => preservesParagraphMarkers(child));
}

function buildCompareBlocks(body: Element): CompareBlockInfo[] {
  return blockElements(body).map((element) => ({
    kind: element.localName === 'tbl' ? 'table' : 'paragraph',
    text: element.localName === 'p' ? textOf(element) : '',
    opaque: element.localName === 'tbl' || hasOpaqueParagraphContent(element) || hasUnalignableCompareRuns(element),
    exactKey: element.toString(),
    element,
  }));
}

function sharedTextLength(a: string, b: string): number {
  let prefix = 0;
  while (prefix < a.length && prefix < b.length && a[prefix] === b[prefix]) prefix++;
  let suffix = 0;
  while (
    suffix < a.length - prefix &&
    suffix < b.length - prefix &&
    a[a.length - 1 - suffix] === b[b.length - 1 - suffix]
  ) suffix++;
  return prefix + suffix;
}

function canPairCompareBlocks(base: CompareBlockInfo, revised: CompareBlockInfo): boolean {
  if (base.kind !== 'paragraph' || revised.kind !== 'paragraph') return false;
  if (base.opaque || revised.opaque) return false;
  if (base.text === revised.text) return true;
  const longest = Math.max(base.text.length, revised.text.length, 1);
  return sharedTextLength(base.text, revised.text) / longest >= COMPARE_PARAGRAPH_PAIR_THRESHOLD;
}

function alignCompareBlocks(base: CompareBlockInfo[], revised: CompareBlockInfo[]): CompareStep[] {
  const rows = base.length + 1;
  const cols = revised.length + 1;
  const costs = new Uint32Array(rows * cols);
  const moves = new Uint8Array(rows * cols);
  const at = (row: number, col: number): number => row * cols + col;
  for (let row = 1; row < rows; row++) {
    costs[at(row, 0)] = row;
    moves[at(row, 0)] = 2;
  }
  for (let col = 1; col < cols; col++) {
    costs[at(0, col)] = col;
    moves[at(0, col)] = 3;
  }
  for (let row = 1; row < rows; row++) {
    for (let col = 1; col < cols; col++) {
      const baseBlock = base[row - 1]!;
      const revisedBlock = revised[col - 1]!;
      const exact = baseBlock.exactKey === revisedBlock.exactKey;
      const pairable = exact || canPairCompareBlocks(baseBlock, revisedBlock);
      const diagonal = pairable ? costs[at(row - 1, col - 1)]! + (exact ? 0 : 1) : Number.MAX_SAFE_INTEGER;
      const deletion = costs[at(row - 1, col)]! + 1;
      const insertion = costs[at(row, col - 1)]! + 1;
      let best = diagonal;
      let move = exact ? 0 : 1;
      if (deletion < best) {
        best = deletion;
        move = 2;
      }
      if (insertion < best) {
        best = insertion;
        move = 3;
      }
      costs[at(row, col)] = best;
      moves[at(row, col)] = move;
    }
  }
  const steps: CompareStep[] = [];
  let row = base.length;
  let col = revised.length;
  while (row > 0 || col > 0) {
    const move = moves[at(row, col)];
    if ((move === 0 || move === 1) && row > 0 && col > 0) {
      steps.push(move === 0
        ? { kind: 'equal', baseIndex: row - 1, revisedIndex: col - 1 }
        : { kind: 'modify', baseIndex: row - 1, revisedIndex: col - 1 });
      row--;
      col--;
      continue;
    }
    if (move === 2 && row > 0) {
      steps.push({ kind: 'delete', baseIndex: row - 1 });
      row--;
      continue;
    }
    if (col > 0) {
      steps.push({ kind: 'insert', revisedIndex: col - 1 });
      col--;
      continue;
    }
    steps.push({ kind: 'delete', baseIndex: row - 1 });
    row--;
  }
  return steps.reverse();
}

function cloneNodeIntoDocument(document: Document, node: Node): Node {
  switch (node.nodeType) {
    case 1: {
      const source = node as Element;
      const clone = document.createElementNS(source.namespaceURI, source.tagName);
      for (let index = 0; index < source.attributes.length; index++) {
        const attribute = source.attributes.item(index);
        if (!attribute) continue;
        clone.setAttributeNS(attribute.namespaceURI, attribute.name, attribute.value);
      }
      for (let child = source.firstChild; child; child = child.nextSibling) {
        clone.appendChild(cloneNodeIntoDocument(document, child));
      }
      return clone;
    }
    case 3:
      return document.createTextNode(node.nodeValue ?? '');
    case 4:
      return document.createCDATASection(node.nodeValue ?? '');
    case 7: {
      const instruction = node as Node & { target?: string; data?: string };
      return document.createProcessingInstruction(instruction.target ?? '', instruction.data ?? node.nodeValue ?? '');
    }
    case 8:
      return document.createComment(node.nodeValue ?? '');
    default:
      return document.createTextNode(node.nodeValue ?? '');
  }
}

function blockInsertionReference(body: Element, blockIndex: number): Node | null {
  const blocks = blockElements(body);
  return blocks[blockIndex] ?? children(body, 'sectPr')[0] ?? null;
}

function isHighSurrogateCodeUnit(code: number): boolean {
  return code >= 0xD800 && code <= 0xDBFF;
}

function isLowSurrogateCodeUnit(code: number): boolean {
  return code >= 0xDC00 && code <= 0xDFFF;
}

function normalizeTextBoundary(text: string, offset: number, side: 'start' | 'end'): number {
  if (offset <= 0 || offset >= text.length) return offset;
  const left = text.charCodeAt(offset - 1);
  const right = text.charCodeAt(offset);
  if (isHighSurrogateCodeUnit(left) && isLowSurrogateCodeUnit(right)) return side === 'start' ? offset - 1 : offset + 1;
  return offset;
}

function fieldInstruction(link: { url?: string; anchor?: string }): string {
  const escapeFieldText = (value: string) => value.replace(/\\/g, '\\\\').replace(/"/g, '\\"');
  const chunks = ['HYPERLINK'];
  if (link.url) chunks.push(`"${escapeFieldText(link.url)}"`);
  if (link.anchor) chunks.push(`\\l "${escapeFieldText(link.anchor)}"`);
  return chunks.join(' ');
}

function preOrderElements(root: Element): Element[] {
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

function cloneCommentInfo(comment: CommentInfo): CommentInfo {
  return JSON.parse(JSON.stringify(comment)) as CommentInfo;
}

function cloneReviewerInfo(reviewer: ReviewerInfo): ReviewerInfo {
  return JSON.parse(JSON.stringify(reviewer)) as ReviewerInfo;
}

function normalizeCommentDate(value: string | undefined): string | undefined {
  if (!value) return undefined;
  return Number.isFinite(Date.parse(value)) ? value : undefined;
}

function updateStyleLength(style: string | null, name: string, points: number): string {
  const normalized = (style ?? '').trim();
  const declaration = `${name}:${Math.max(0, points)}pt`;
  if (!normalized) return declaration;
  const parts = normalized.split(';').map((part) => part.trim()).filter(Boolean);
  let replaced = false;
  const next = parts.map((part) => {
    if (!part.toLowerCase().startsWith(`${name.toLowerCase()}:`)) return part;
    replaced = true;
    return declaration;
  });
  if (!replaced) next.push(declaration);
  return next.join(';');
}

function contentTypesDefaults(types: Element): Element[] {
  return children(types, 'Default', CONTENT_TYPES_NS);
}

function contentTypesOverrides(types: Element): Element[] {
  return children(types, 'Override', CONTENT_TYPES_NS);
}

function createDrawingElement(
  document: Document,
  relationshipId: string,
  size: { widthEmu: number; heightEmu: number },
  options: { alt?: string; title?: string; name?: string; placement: 'inline' | 'floating'; docPrId: number },
): Element {
  const drawing = wordElement(document, 'drawing');
  const container = document.createElementNS(WP_NS, `wp:${options.placement === 'floating' ? 'anchor' : 'inline'}`);
  if (options.placement === 'floating') {
    container.setAttribute('behindDoc', '0');
    container.setAttribute('locked', '0');
    container.setAttribute('layoutInCell', '1');
    container.setAttribute('simplePos', '0');
    container.setAttribute('relativeHeight', '251658240');
    container.setAttribute('allowOverlap', '1');
    const simplePos = document.createElementNS(WP_NS, 'wp:simplePos');
    simplePos.setAttribute('x', '0');
    simplePos.setAttribute('y', '0');
    container.appendChild(simplePos);
    const positionH = document.createElementNS(WP_NS, 'wp:positionH');
    positionH.setAttribute('relativeFrom', 'column');
    const alignH = document.createElementNS(WP_NS, 'wp:align');
    alignH.appendChild(document.createTextNode('left'));
    positionH.appendChild(alignH);
    const positionV = document.createElementNS(WP_NS, 'wp:positionV');
    positionV.setAttribute('relativeFrom', 'paragraph');
    const alignV = document.createElementNS(WP_NS, 'wp:align');
    alignV.appendChild(document.createTextNode('top'));
    positionV.appendChild(alignV);
    container.appendChild(positionH);
    container.appendChild(positionV);
  }
  const extent = document.createElementNS(WP_NS, 'wp:extent');
  extent.setAttribute('cx', String(Math.round(size.widthEmu)));
  extent.setAttribute('cy', String(Math.round(size.heightEmu)));
  const effectExtent = document.createElementNS(WP_NS, 'wp:effectExtent');
  for (const side of ['l', 't', 'r', 'b']) effectExtent.setAttribute(side, '0');
  const docPr = document.createElementNS(WP_NS, 'wp:docPr');
  docPr.setAttribute('id', String(options.docPrId));
  docPr.setAttribute('name', options.name ?? `图片 ${options.docPrId}`);
  setOptionalAttribute(docPr, 'descr', options.alt);
  setOptionalAttribute(docPr, 'title', options.title);
  const framePr = document.createElementNS(WP_NS, 'wp:cNvGraphicFramePr');
  const graphicFrameLocks = document.createElementNS(A_NS, 'a:graphicFrameLocks');
  graphicFrameLocks.setAttribute('noChangeAspect', '1');
  framePr.appendChild(graphicFrameLocks);
  const graphic = document.createElementNS(A_NS, 'a:graphic');
  const graphicData = document.createElementNS(A_NS, 'a:graphicData');
  graphicData.setAttribute('uri', PIC_NS);
  const pic = document.createElementNS(PIC_NS, 'pic:pic');
  const nvPicPr = document.createElementNS(PIC_NS, 'pic:nvPicPr');
  const cNvPr = document.createElementNS(PIC_NS, 'pic:cNvPr');
  cNvPr.setAttribute('id', '0');
  cNvPr.setAttribute('name', options.name ?? `image-${options.docPrId}`);
  setOptionalAttribute(cNvPr, 'descr', options.alt);
  setOptionalAttribute(cNvPr, 'title', options.title);
  const cNvPicPr = document.createElementNS(PIC_NS, 'pic:cNvPicPr');
  nvPicPr.appendChild(cNvPr);
  nvPicPr.appendChild(cNvPicPr);
  const blipFill = document.createElementNS(PIC_NS, 'pic:blipFill');
  const blip = document.createElementNS(A_NS, 'a:blip');
  blip.setAttributeNS(OFFICE_REL_NS, 'r:embed', relationshipId);
  const stretch = document.createElementNS(A_NS, 'a:stretch');
  stretch.appendChild(document.createElementNS(A_NS, 'a:fillRect'));
  blipFill.appendChild(blip);
  blipFill.appendChild(stretch);
  const spPr = document.createElementNS(PIC_NS, 'pic:spPr');
  const xfrm = document.createElementNS(A_NS, 'a:xfrm');
  const off = document.createElementNS(A_NS, 'a:off');
  off.setAttribute('x', '0');
  off.setAttribute('y', '0');
  const ext = document.createElementNS(A_NS, 'a:ext');
  ext.setAttribute('cx', String(Math.round(size.widthEmu)));
  ext.setAttribute('cy', String(Math.round(size.heightEmu)));
  xfrm.appendChild(off);
  xfrm.appendChild(ext);
  const prstGeom = document.createElementNS(A_NS, 'a:prstGeom');
  prstGeom.setAttribute('prst', 'rect');
  prstGeom.appendChild(document.createElementNS(A_NS, 'a:avLst'));
  spPr.appendChild(xfrm);
  spPr.appendChild(prstGeom);
  pic.appendChild(nvPicPr);
  pic.appendChild(blipFill);
  pic.appendChild(spPr);
  graphicData.appendChild(pic);
  graphic.appendChild(graphicData);
  if (options.placement === 'floating') {
    const wrapSquare = document.createElementNS(WP_NS, 'wp:wrapSquare');
    wrapSquare.setAttribute('wrapText', 'bothSides');
    container.appendChild(extent);
    container.appendChild(effectExtent);
    container.appendChild(wrapSquare);
    container.appendChild(docPr);
    container.appendChild(framePr);
    container.appendChild(graphic);
  } else {
    container.appendChild(extent);
    container.appendChild(effectExtent);
    container.appendChild(docPr);
    container.appendChild(framePr);
    container.appendChild(graphic);
  }
  drawing.appendChild(container);
  return drawing;
}

function relativeTargetPath(fromPart: string, toPart: string): string {
  const from = fromPart.split('/').slice(0, -1);
  const to = toPart.split('/');
  while (from.length && to.length && from[0] === to[0]) {
    from.shift();
    to.shift();
  }
  return `${from.map(() => '..').concat(to).join('/')}`;
}

function sourcePartFromRelationshipsPath(relPath: string): string | undefined {
  const match = relPath.match(/^(.*\/)?_rels\/([^/]+)\.rels$/);
  if (!match) return undefined;
  const dir = match[1] ?? '';
  return `${dir}${match[2] ?? ''}`;
}

function documentUsesRelationship(document: Document, relationshipId: string): boolean {
  return Array.from(document.getElementsByTagNameNS(A_NS, 'blip')).some((blip) =>
    [blip.getAttributeNS(OFFICE_REL_NS, 'embed'), blip.getAttributeNS(OFFICE_REL_NS, 'link'),
      blip.getAttribute('r:embed'), blip.getAttribute('r:link')].includes(relationshipId)) ||
    Array.from(document.getElementsByTagNameNS(V_NS, 'imagedata')).some((node) =>
      [node.getAttributeNS(OFFICE_REL_NS, 'id'), node.getAttribute('r:id')].includes(relationshipId));
}

async function readEntry(entry: JSZip.JSZipObject, limit: number): Promise<Uint8Array> {
  return new Promise((resolve, reject) => {
    const chunks: Uint8Array[] = [];
    let size = 0;
    // JSZip exposes internalStream publicly, but omits it from JSZipObject's typings.
    const stream = (entry as JSZip.JSZipObject & {
      internalStream(type: 'uint8array'): JSZip.JSZipStreamHelper<Uint8Array>;
    }).internalStream('uint8array');
    stream.on('data', (chunk: Uint8Array) => {
      size += chunk.length;
      if (size > limit) {
        stream.pause();
        reject(new Error('DOCX exceeds the expanded package/part size limit.'));
        return;
      }
      chunks.push(chunk);
    });
    stream.on('error', reject);
    stream.on('end', () => {
      const result = new Uint8Array(size);
      let offset = 0;
      for (const chunk of chunks) { result.set(chunk, offset); offset += chunk.length; }
      resolve(result);
    });
    stream.resume();
  });
}

function resolveInternalTarget(sourcePart: string, target: string | undefined, mode: string | null): string | undefined {
  if (!target || mode === 'External') return undefined;
  try {
    return resolveTargetPath(sourcePart, decodeURIComponent(target));
  } catch {
    return undefined;
  }
}

export class DocxDocument {
  private parts: Map<string, Uint8Array>;
  private mainPath: string;
  private currentRevision = 0;
  private documents = new Map<string, Document>();
  private dirtyPartXml = new Set<string>();
  private dirtyPartSizes = new Map<string, number>();
  private numberingContextCache?: NumberingContext;
  private stylesCache?: { revision: number; context: StylesContext };
  private outlineCache?: { revision: number; outline: OutlineNode[] };
  private noteStateCache?: { revision: number; state: NoteState };
  private commentStateCache?: { revision: number; comments: CommentInfo[] };
  private contentPartPathsCache?: { revision: number; paths: string[] };
  private commentBindingsCache?: { revision: number; bindings: CommentPartBinding[] };
  private revisionInfoCache?: { revision: number; mainPath: string; stylesRevision: number; revisions: RevisionInfo[] };
  private reviewerInfoCache?: { revision: number; reviewers: ReviewerInfo[] };
  private tableCellLocationCache?: { revision: number; locations: Map<number, TableCellLocation> };
  private imageDataUrls = new Map<string, { revision: number; contentType: string; url: string }>();
  private undoHistory: HistoryState[] = [];
  private redoHistory: HistoryState[] = [];
  private undoHistoryBytes = 0;
  private redoHistoryBytes = 0;
  private historyGroupDepth = 0;
  private historyGroupLabel?: string;
  private historyGroupBase?: { parts: Map<string, Uint8Array>; bytes: number };
  private historyGroupAction: HistoryAction = { kind: 'group' };
  private historyGroupAborted = false;
  private suppressHistory = false;
  private nextHistoryLabel?: string;
  private nextHistoryAction: HistoryAction = { kind: 'other' };
  private pendingMergedHistory?: { label?: string; action: HistoryAction; at: number };
  private revisionAuthor?: string;

  private constructor(parts: Map<string, Uint8Array>) {
    this.parts = parts;
    this.mainPath = this.validatePackage();
  }

  static create(): DocxDocument {
    const files: Record<string, string> = {
      '[Content_Types].xml': `<Types xmlns="${CONTENT_TYPES_NS}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${MAIN_TYPE}"/></Types>`,
      '_rels/.rels': `<Relationships xmlns="${REL_NS}"><Relationship Id="rId1" Type="${OFFICE_DOCUMENT_REL}" Target="word/document.xml"/></Relationships>`,
      'word/document.xml': `<w:document xmlns:w="${WORD_NS}"><w:body><w:p><w:r><w:t xml:space="preserve"></w:t></w:r></w:p><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`,
    };
    return new DocxDocument(new Map(Object.entries(files).map(([path, xml]) => [path, encoder.encode(xml)])));
  }

  static async load(input: Uint8Array | ArrayBuffer | Blob): Promise<DocxDocument> {
    const size = input instanceof Blob ? input.size : input.byteLength;
    if (size > MAX_ARCHIVE) throw new Error('DOCX archive exceeds 50 MiB.');
    const bytes = input instanceof Blob ? await input.arrayBuffer() : input;
    const zip = await JSZip.loadAsync(bytes, { createFolders: false });
    const entries = Object.values(zip.files) as JSZip.JSZipObject[];
    if (entries.length > MAX_PARTS) throw new Error('DOCX contains too many ZIP entries.');
    const parts = new Map<string, Uint8Array>();
    let total = 0;
    for (const entry of entries) {
      if (entry.dir) continue;
      validatePath(entry.name);
      if (entry.unsafeOriginalName !== undefined && entry.unsafeOriginalName !== entry.name) {
        throw new Error('Unsafe ZIP entry path.');
      }
      const data = await readEntry(entry, Math.min(MAX_PART, MAX_TOTAL - total));
      total += data.byteLength;
      parts.set(entry.name, data);
    }
    return new DocxDocument(parts);
  }

  static compare(base: DocxDocument, revised: DocxDocument, options: { author?: string; date?: string } = {}): DocxDocument {
    if (!(base instanceof DocxDocument) || !(revised instanceof DocxDocument)) {
      throw new Error('DocxDocument.compare() expects DocxDocument instances.');
    }
    if (!options || typeof options !== 'object' || Array.isArray(options)) throw new Error('compare options must be an object.');
    if (Object.keys(options).some((key) => !['author', 'date'].includes(key))) throw new Error('Unknown compare option.');
    if (options.author !== undefined) assertText(options.author, 'author');
    if (options.date !== undefined) assertText(options.date, 'date');
    const baseBody = bodyOf(base.getCachedPartDocument(base.mainPath));
    const revisedBody = bodyOf(revised.getCachedPartDocument(revised.mainPath));
    const baseParagraphCount = descendants(baseBody, 'p').length;
    const revisedParagraphCount = descendants(revisedBody, 'p').length;
    if (baseParagraphCount > COMPARE_MAX_PARAGRAPHS || revisedParagraphCount > COMPARE_MAX_PARAGRAPHS) {
      throw new Error(
        `DocxDocument.compare() supports at most ${COMPARE_MAX_PARAGRAPHS} main-document paragraphs per input document ` +
        `(received ${baseParagraphCount} and ${revisedParagraphCount}).`,
      );
    }
    const result = new DocxDocument(new Map(base.parts));
    result.suppressHistory = true;
    result.revisionAuthor = options.author;
    result.setTrackChangesDirect(true);
    const baseBlocks = buildCompareBlocks(baseBody);
    const revisedBlocks = buildCompareBlocks(revisedBody);
    result.applyComparedBlocks(alignCompareBlocks(baseBlocks, revisedBlocks), revisedBlocks, revised.getStylesContext(), options.author, options.date);
    result.materializeAllParts();
    result.currentRevision = 0;
    result.undoHistory = [];
    result.redoHistory = [];
    result.undoHistoryBytes = 0;
    result.redoHistoryBytes = 0;
    result.pendingMergedHistory = undefined;
    result.nextHistoryLabel = undefined;
    result.nextHistoryAction = { kind: 'other' };
    result.numberingContextCache = undefined;
    result.stylesCache = undefined;
    result.noteStateCache = undefined;
    result.commentStateCache = undefined;
    result.contentPartPathsCache = undefined;
    result.commentBindingsCache = undefined;
    result.revisionInfoCache = undefined;
    result.reviewerInfoCache = undefined;
    result.outlineCache = undefined;
    result.tableCellLocationCache = undefined;
    result.imageDataUrls.clear();
    result.suppressHistory = false;
    return result;
  }

  get revision(): number { return this.currentRevision; }
  get mainDocumentPath(): string { return this.mainPath; }

  listParts(): string[] { return [...this.parts.keys()].sort(); }

  canUndo(): boolean {
    return this.undoHistory.length > 0;
  }

  canRedo(): boolean {
    return this.redoHistory.length > 0;
  }

  undo(): DocumentSnapshot {
    if (!this.undoHistory.length) return this.getSnapshot();
    this.materializeAllParts();
    const current = this.snapshotHistoryState(this.revision, this.nextHistoryLabel, this.nextHistoryAction);
    const step = this.undoHistory.pop()!;
    this.undoHistoryBytes -= step.bytes;
    this.restorePartsFromHistory(step.parts);
    this.currentRevision++;
    this.numberingContextCache = undefined;
    this.stylesCache = undefined;
    this.outlineCache = undefined;
    this.noteStateCache = undefined;
    this.commentStateCache = undefined;
    this.revisionInfoCache = undefined;
    this.reviewerInfoCache = undefined;
    this.tableCellLocationCache = undefined;
    this.imageDataUrls.clear();
    this.pushRedoState({ ...current, action: step.action });
    this.nextHistoryLabel = undefined;
    this.nextHistoryAction = { kind: 'other' };
    return this.getSnapshot();
  }

  redo(): DocumentSnapshot {
    if (!this.redoHistory.length) return this.getSnapshot();
    this.materializeAllParts();
    const current = this.snapshotHistoryState(this.revision, this.nextHistoryLabel, this.nextHistoryAction);
    const step = this.redoHistory.pop()!;
    this.redoHistoryBytes -= step.bytes;
    this.restorePartsFromHistory(step.parts);
    this.currentRevision++;
    this.numberingContextCache = undefined;
    this.stylesCache = undefined;
    this.outlineCache = undefined;
    this.noteStateCache = undefined;
    this.commentStateCache = undefined;
    this.revisionInfoCache = undefined;
    this.reviewerInfoCache = undefined;
    this.tableCellLocationCache = undefined;
    this.imageDataUrls.clear();
    this.pushUndoState({ ...current, action: step.action });
    this.enforceHistoryLimits();
    this.nextHistoryLabel = undefined;
    this.nextHistoryAction = { kind: 'other' };
    return this.getSnapshot();
  }

  getHistory(): { undo: HistoryEntry[]; redo: HistoryEntry[] } {
    return {
      undo: this.undoHistory.map((state) => ({ ...state.entry })),
      redo: this.redoHistory.map((state) => ({ ...state.entry })),
    };
  }

  clearHistory(): void {
    this.undoHistory = [];
    this.redoHistory = [];
    this.undoHistoryBytes = 0;
    this.redoHistoryBytes = 0;
    this.historyGroupDepth = 0;
    this.historyGroupLabel = undefined;
    this.historyGroupBase = undefined;
    this.historyGroupAction = { kind: 'group' };
    this.historyGroupAborted = false;
    this.nextHistoryLabel = undefined;
    this.nextHistoryAction = { kind: 'other' };
    this.pendingMergedHistory = undefined;
  }

  beginHistoryGroup(label?: string): void {
    if (this.historyGroupDepth === 0) {
      this.historyGroupLabel = label;
      this.historyGroupBase = undefined;
      this.historyGroupAction = { kind: 'group' };
      this.historyGroupAborted = false;
    }
    this.historyGroupDepth++;
  }

  endHistoryGroup(): void {
    if (this.historyGroupDepth === 0) throw new Error('History group is not active.');
    this.historyGroupDepth--;
    if (this.historyGroupDepth > 0) return;
    if (this.historyGroupBase && !this.historyGroupAborted) {
      this.pushUndoState({
        parts: this.historyGroupBase.parts,
        bytes: this.historyGroupBase.bytes,
        entry: { revision: this.revision, label: this.historyGroupLabel, at: Date.now() },
        action: this.historyGroupAction,
      });
      this.redoHistory = [];
      this.redoHistoryBytes = 0;
      this.enforceHistoryLimits();
    }
    this.historyGroupLabel = undefined;
    this.historyGroupBase = undefined;
    this.historyGroupAction = { kind: 'group' };
    this.historyGroupAborted = false;
  }

  private hasPart(path: string): boolean {
    return this.parts.has(path);
  }

  private getContentType(path: string): string | undefined {
    const types = this.getCachedPartDocument('[Content_Types].xml').documentElement!;
    return children(types, 'Override', CONTENT_TYPES_NS)
      .find(type => type.getAttribute('PartName') === `/${path}`)?.getAttribute('ContentType')
      ?? children(types, 'Default', CONTENT_TYPES_NS)
        .find(type => type.getAttribute('Extension')?.toLowerCase() === path.split('.').pop()?.toLowerCase())?.getAttribute('ContentType')
      ?? undefined;
  }

  private createContentTypeResolver(): (path: string) => string | undefined {
    const types = this.getCachedPartDocument('[Content_Types].xml').documentElement!;
    const overrides = new Map(children(types, 'Override', CONTENT_TYPES_NS)
      .map((type) => [type.getAttribute('PartName') ?? '', type.getAttribute('ContentType') ?? '']));
    const defaults = new Map(children(types, 'Default', CONTENT_TYPES_NS)
      .map((type) => [(type.getAttribute('Extension') ?? '').toLowerCase(), type.getAttribute('ContentType') ?? '']));
    return (path: string): string | undefined =>
      overrides.get(`/${path}`) || defaults.get(path.split('.').pop()?.toLowerCase() ?? '') || undefined;
  }

  private relationshipsFor(partPath: string): Map<string, RelationshipTarget> {
    const relPath = resolveRelationshipsPath(partPath);
    if (!this.hasPart(relPath)) return new Map();
    const rels = this.getCachedPartDocument(relPath).documentElement!;
    const entries: [string, RelationshipTarget][] = children(rels, 'Relationship', REL_NS).flatMap((rel) => {
      const target = rel.getAttribute('Target') ?? undefined;
      const id = rel.getAttribute('Id') ?? '';
      if (!id) return [];
      return [[id, {
        id: rel.getAttribute('Id') ?? '',
        mode: rel.getAttribute('TargetMode') ?? undefined,
        target,
        // 畸形 Target（目录穿越、非法字符、编码错误）降级为无部件路径，读取方法不得因此抛错。
        partPath: resolveInternalTarget(partPath, target, rel.getAttribute('TargetMode')),
      } satisfies RelationshipTarget]];
    });
    return new Map(entries);
  }

  private paragraphsWithRelationships(): ParagraphInfo[] {
    return this.buildParagraphs();
  }

  getPartBytes(path: string): Uint8Array {
    validatePath(path);
    this.materializePart(path);
    const bytes = this.parts.get(path);
    if (!bytes) throw new Error(`Package part not found: ${path}`);
    return Uint8Array.from(bytes);
  }

  getPartXml(path: string): string { return decodeXml(this.getPartBytes(path)); }

  private partDocumentOrUndefined(path: string): Document | undefined {
    try {
      return this.getPartDocument(path);
    } catch {
      return undefined;
    }
  }

  /** Returns a detached DOM; use updatePartXml to persist changes. */
  getPartDocument(path: string): Document {
    return parseXml(serializeXml(this.getCachedPartDocument(path)));
  }

  updatePartXml(path: string, update: (document: Document) => void): void {
    const document = this.getPartDocument(path);
    update(document);
    this.setPartXml(path, serializeXml(document));
  }

  private updatePartXmlInternal(path: string, update: (document: Document) => boolean | void): void {
    validatePath(path);
    const history = this.beginMutationHistory(this.nextHistoryLabel, this.nextHistoryAction);
    this.nextHistoryLabel = undefined;
    this.nextHistoryAction = { kind: 'other' };
    const previous = this.captureState();
    try {
      const document = this.getCachedPartDocument(path);
      const before = this.parts.get(path);
      if (!before) throw new Error(`Package part not found: ${path}`);
      const changed = update(document);
      if (changed === false) {
        this.pendingMergedHistory = undefined;
        return;
      }
      if (changed !== true) {
        const next = encodeXml(serializeXml(document));
        if (equalBytes(before, next)) {
          this.pendingMergedHistory = undefined;
          return;
        }
        this.parts.set(path, next);
        this.dirtyPartXml.delete(path);
        this.dirtyPartSizes.delete(path);
      } else {
        this.dirtyPartXml.add(path);
        this.dirtyPartSizes.delete(path);
      }
      this.finalizeMutation(path);
      if (history) this.recordHistory(history);
      else this.applyPendingMergedHistory();
    } catch (error) {
      this.restoreState(previous);
      this.abortHistoryGroupOnFailure();
      throw error;
    }
  }

  setPartXml(path: string, xml: string): void {
    if (typeof xml !== 'string' || xml.length > MAX_PART) throw new Error('XML part exceeds size limit.');
    const history = this.beginMutationHistory(this.nextHistoryLabel, this.nextHistoryAction);
    this.nextHistoryLabel = undefined;
    this.nextHistoryAction = { kind: 'other' };
    const previous = this.captureState();
    try {
      validatePath(path);
      if (!this.parts.has(path)) throw new Error('Use addPart with a content type to create a new part.');
      const document = parseXml(xml);
      const bytes = encodeXml(xml);
      const current = this.parts.get(path);
      if (!current) throw new Error(`Package part not found: ${path}`);
      if (equalBytes(current, bytes)) {
        this.documents.set(path, document);
        this.pendingMergedHistory = undefined;
        return;
      }
      this.documents.set(path, document);
      this.parts.set(path, bytes);
      this.dirtyPartXml.delete(path);
      this.dirtyPartSizes.delete(path);
      this.finalizeMutation(path);
      if (history) this.recordHistory(history);
      else this.applyPendingMergedHistory();
    } catch (error) {
      this.restoreState(previous);
      this.abortHistoryGroupOnFailure();
      throw error;
    }
  }

  /** Replaces an existing part. Relationships/content types remain under caller control. */
  setPartBytes(path: string, bytes: Uint8Array): void {
    const history = this.beginMutationHistory(this.nextHistoryLabel, this.nextHistoryAction);
    this.nextHistoryLabel = undefined;
    this.nextHistoryAction = { kind: 'other' };
    const previous = this.captureState();
    try {
      const current = this.parts.get(path);
      if (current && equalBytes(current, bytes)) {
        this.pendingMergedHistory = undefined;
        return;
      }
      this.replacePartBytes(path, bytes);
      this.finalizeMutation(path);
      if (history) this.recordHistory(history);
      else this.applyPendingMergedHistory();
    } catch (error) {
      this.restoreState(previous);
      this.abortHistoryGroupOnFailure();
      throw error;
    }
  }

  addPart(path: string, bytes: Uint8Array, contentType: string): void {
    const history = this.beginMutationHistory(this.nextHistoryLabel, this.nextHistoryAction);
    this.nextHistoryLabel = undefined;
    this.nextHistoryAction = { kind: 'other' };
    const previous = this.captureState();
    try {
      validatePath(path);
      assertText(contentType, 'contentType');
      if (!contentType || this.parts.has(path)) throw new Error('New part requires a unique path and content type.');
      const types = this.getCachedPartDocument('[Content_Types].xml');
      const override = types.createElementNS(CONTENT_TYPES_NS, 'Override');
      override.setAttribute('PartName', `/${path}`);
      override.setAttribute('ContentType', contentType);
      types.documentElement!.appendChild(override);
      this.parts.set(path, Uint8Array.from(bytes));
      this.documents.delete(path);
      this.dirtyPartXml.add('[Content_Types].xml');
      this.dirtyPartSizes.delete(path);
      this.finalizeMutation('[Content_Types].xml');
      if (history) this.recordHistory(history);
      else this.applyPendingMergedHistory();
    } catch (error) {
      this.restoreState(previous);
      this.abortHistoryGroupOnFailure();
      throw error;
    }
  }

  private commitParts(parts: Map<string, Uint8Array>): void {
    if (equalPartMap(this.parts, parts)) return;
    const history = this.beginMutationHistory(this.nextHistoryLabel, this.nextHistoryAction);
    this.nextHistoryLabel = undefined;
    this.nextHistoryAction = { kind: 'other' };
    const previous = this.captureState();
    try {
      this.parts = parts;
      this.documents = new Map();
      this.dirtyPartXml = new Set();
      this.dirtyPartSizes = new Map();
      this.assertPackageLimits();
      this.mainPath = this.validatePackage();
      this.currentRevision++;
      this.numberingContextCache = undefined;
      this.stylesCache = undefined;
      this.outlineCache = undefined;
      this.noteStateCache = undefined;
      this.commentStateCache = undefined;
      this.revisionInfoCache = undefined;
      this.reviewerInfoCache = undefined;
      this.tableCellLocationCache = undefined;
      this.imageDataUrls.clear();
      if (history) this.recordHistory(history);
      else this.applyPendingMergedHistory();
    } catch (error) {
      this.restoreState(previous);
      this.abortHistoryGroupOnFailure();
      throw error;
    }
  }


  private validatePackage(): string {
    if (this.parts.size > MAX_PARTS) throw new Error('Too many package parts.');
    let total = 0;
    for (const [path, bytes] of this.parts) {
      validatePath(path);
      total += bytes.byteLength;
      if (bytes.byteLength > MAX_PART || total > MAX_TOTAL) throw new Error('Package exceeds size limits.');
    }
    const rels = this.getCachedPartDocument('_rels/.rels').documentElement;
    if (rels?.namespaceURI !== REL_NS || rels.localName !== 'Relationships') {
      throw new Error('Invalid root package relationships.');
    }
    const main = children(rels, 'Relationship', REL_NS).filter(rel => rel.getAttribute('Type') === OFFICE_DOCUMENT_REL);
    if (main.length !== 1 || main[0]!.getAttribute('TargetMode') === 'External') {
      throw new Error('Package requires one internal officeDocument relationship.');
    }
    const path = decodePackageTarget(main[0]!.getAttribute('Target') ?? '');
    if (!path) throw new Error('Invalid main document relationship target.');
    validatePath(path);
    const types = this.getCachedPartDocument('[Content_Types].xml').documentElement;
    if (types?.namespaceURI !== CONTENT_TYPES_NS || types.localName !== 'Types') {
      throw new Error('Invalid content types part.');
    }
    const overrides = children(types, 'Override', CONTENT_TYPES_NS);
    const type = overrides.find(type => type.getAttribute('PartName') === `/${path}`)?.getAttribute('ContentType')
      ?? children(types, 'Default', CONTENT_TYPES_NS)
        .find(type => type.getAttribute('Extension') === path.split('.').pop())?.getAttribute('ContentType');
    if (type !== MAIN_TYPE) throw new Error('Package is not a supported .docx document (macros/strict OOXML are not supported).');
    bodyOf(this.getCachedPartDocument(path));
    return path;
  }

  private captureState(): {
    parts: Map<string, Uint8Array>;
    documents: Map<string, Document>;
    dirtyPartXml: Set<string>;
    dirtyPartSizes: Map<string, number>;
    mainPath: string;
    revision: number;
    numberingContextCache: NumberingContext | undefined;
    stylesCache: { revision: number; context: StylesContext } | undefined;
    outlineCache: { revision: number; outline: OutlineNode[] } | undefined;
    noteStateCache: { revision: number; state: NoteState } | undefined;
    commentStateCache: { revision: number; comments: CommentInfo[] } | undefined;
    revisionInfoCache: { revision: number; mainPath: string; stylesRevision: number; revisions: RevisionInfo[] } | undefined;
    reviewerInfoCache: { revision: number; reviewers: ReviewerInfo[] } | undefined;
    tableCellLocationCache: { revision: number; locations: Map<number, TableCellLocation> } | undefined;
    imageDataUrls: Map<string, { revision: number; contentType: string; url: string }>;
  } {
    return {
      parts: new Map(this.parts),
      documents: new Map(this.documents),
      dirtyPartXml: new Set(this.dirtyPartXml),
      dirtyPartSizes: new Map(this.dirtyPartSizes),
      mainPath: this.mainPath,
      revision: this.currentRevision,
      numberingContextCache: this.numberingContextCache,
      stylesCache: this.stylesCache,
      outlineCache: this.outlineCache,
      noteStateCache: this.noteStateCache,
      commentStateCache: this.commentStateCache,
      revisionInfoCache: this.revisionInfoCache,
      reviewerInfoCache: this.reviewerInfoCache,
      tableCellLocationCache: this.tableCellLocationCache,
      imageDataUrls: new Map(this.imageDataUrls),
    };
  }

  private restoreState(state: ReturnType<DocxDocument['captureState']>): void {
    this.parts = state.parts;
    this.documents = state.documents;
    this.dirtyPartXml = state.dirtyPartXml;
    this.dirtyPartSizes = state.dirtyPartSizes;
    this.mainPath = state.mainPath;
    this.currentRevision = state.revision;
    this.numberingContextCache = state.numberingContextCache;
    this.stylesCache = state.stylesCache;
    this.outlineCache = state.outlineCache;
    this.noteStateCache = state.noteStateCache;
    this.commentStateCache = state.commentStateCache;
    this.revisionInfoCache = state.revisionInfoCache;
    this.reviewerInfoCache = state.reviewerInfoCache;
    this.tableCellLocationCache = state.tableCellLocationCache;
    this.imageDataUrls = state.imageDataUrls;
  }

  private cloneParts(parts: Map<string, Uint8Array>): Map<string, Uint8Array> {
    return new Map(parts);
  }

  private historyStateBytes(parts: Map<string, Uint8Array>): number {
    let total = 0;
    for (const bytes of parts.values()) total += bytes.byteLength;
    return total;
  }

  private snapshotHistoryState(revision: number, label: string | undefined, action: HistoryAction): HistoryState {
    const parts = this.cloneParts(this.parts);
    return {
      parts,
      bytes: this.historyStateBytes(parts),
      entry: { revision, label, at: Date.now() },
      action,
    };
  }

  private pushUndoState(state: HistoryState): void {
    const previous = this.undoHistory[this.undoHistory.length - 1];
    const canMerge = state.action.kind === 'setParagraphText' &&
      previous?.action.kind === 'setParagraphText' &&
      previous.action.paragraph === state.action.paragraph &&
      state.entry.at - previous.entry.at <= HISTORY_MERGE_WINDOW_MS &&
      !this.historyGroupDepth &&
      this.redoHistory.length === 0;
    if (canMerge) {
      previous.entry.revision = state.entry.revision;
      previous.entry.at = state.entry.at;
      previous.entry.label = state.entry.label ?? previous.entry.label;
      return;
    }
    this.undoHistory.push(state);
    this.undoHistoryBytes += state.bytes;
  }

  private pushRedoState(state: HistoryState): void {
    this.redoHistory.push(state);
    this.redoHistoryBytes += state.bytes;
    const overflow = this.redoHistory.length - HISTORY_MAX_ENTRIES;
    if (overflow > 0) {
      const dropped = this.redoHistory.splice(0, overflow);
      this.redoHistoryBytes -= dropped.reduce((sum, entry) => sum + entry.bytes, 0);
    }
  }

  private enforceHistoryLimits(): void {
    const undoOverflow = this.undoHistory.length - HISTORY_MAX_ENTRIES;
    if (undoOverflow > 0) {
      const dropped = this.undoHistory.splice(0, undoOverflow);
      this.undoHistoryBytes -= dropped.reduce((sum, entry) => sum + entry.bytes, 0);
    }
    let total = this.undoHistoryBytes + this.redoHistoryBytes;
    while (total > HISTORY_MAX_BYTES) {
      if (this.undoHistory.length) {
        let removeCount = 0;
        let removedBytes = 0;
        while (removeCount < this.undoHistory.length && total - removedBytes > HISTORY_MAX_BYTES) {
          removedBytes += this.undoHistory[removeCount]!.bytes;
          removeCount++;
        }
        this.undoHistory.splice(0, removeCount);
        this.undoHistoryBytes -= removedBytes;
        total -= removedBytes;
        continue;
      }
      if (this.redoHistory.length) {
        let removeCount = 0;
        let removedBytes = 0;
        while (removeCount < this.redoHistory.length && total - removedBytes > HISTORY_MAX_BYTES) {
          removedBytes += this.redoHistory[removeCount]!.bytes;
          removeCount++;
        }
        this.redoHistory.splice(0, removeCount);
        this.redoHistoryBytes -= removedBytes;
        total -= removedBytes;
        continue;
      }
      break;
    }
  }

  private restorePartsFromHistory(parts: Map<string, Uint8Array>): void {
    const next = this.cloneParts(parts);
    this.parts = next;
    this.documents = new Map();
    this.dirtyPartXml = new Set();
    this.dirtyPartSizes = new Map();
    this.assertPackageLimits();
    this.mainPath = this.validatePackage();
  }

  private recordHistory(before: HistoryState): void {
    if (this.suppressHistory) return;
    if (this.historyGroupDepth > 0) {
      if (!this.historyGroupBase) this.historyGroupBase = { parts: before.parts, bytes: before.bytes };
      if (before.action.kind !== 'other') this.historyGroupAction = before.action;
      this.redoHistory = [];
      this.redoHistoryBytes = 0;
      return;
    }
    this.pushUndoState({
      parts: before.parts,
      bytes: before.bytes,
      entry: { revision: this.revision, label: before.entry.label, at: Date.now() },
      action: before.action,
    });
    this.redoHistory = [];
    this.redoHistoryBytes = 0;
    this.enforceHistoryLimits();
  }

  private shouldMergeWithPreviousUndo(action: HistoryAction, at: number): boolean {
    const previous = this.undoHistory[this.undoHistory.length - 1];
    return action.kind === 'setParagraphText' &&
      previous?.action.kind === 'setParagraphText' &&
      previous.action.paragraph === action.paragraph &&
      at - previous.entry.at <= HISTORY_MERGE_WINDOW_MS &&
      this.historyGroupDepth === 0 &&
      this.redoHistory.length === 0;
  }

  private applyPendingMergedHistory(): void {
    const pending = this.pendingMergedHistory;
    this.pendingMergedHistory = undefined;
    if (!pending) return;
    const previous = this.undoHistory[this.undoHistory.length - 1];
    if (!previous || !this.shouldMergeWithPreviousUndo(pending.action, pending.at)) return;
    previous.entry.revision = this.revision;
    previous.entry.at = Date.now();
    previous.entry.label = pending.label ?? previous.entry.label;
    this.redoHistory = [];
    this.redoHistoryBytes = 0;
  }

  private beginMutationHistory(label?: string, action: HistoryAction = { kind: 'other' }): HistoryState | undefined {
    if (this.suppressHistory) return undefined;
    const at = Date.now();
    if (this.shouldMergeWithPreviousUndo(action, at)) {
      this.pendingMergedHistory = { label, action, at };
      return undefined;
    }
    this.materializeAllParts();
    return this.snapshotHistoryState(this.revision, label, action);
  }

  private abortHistoryGroupOnFailure(): void {
    if (this.historyGroupDepth > 0) this.historyGroupAborted = true;
    this.pendingMergedHistory = undefined;
  }

  private getCachedPartDocument(path: string): Document {
    validatePath(path);
    let document = this.documents.get(path);
    if (document) return document;
    const bytes = this.parts.get(path);
    if (!bytes) throw new Error(`Package part not found: ${path}`);
    document = parseXml(decodeXml(bytes));
    this.documents.set(path, document);
    return document;
  }

  private materializePart(path: string): void {
    if (!this.dirtyPartXml.has(path)) return;
    const document = this.documents.get(path);
    if (!document) throw new Error(`Package part not found: ${path}`);
    this.parts.set(path, encodeXml(serializeXml(document)));
    this.dirtyPartXml.delete(path);
    this.dirtyPartSizes.delete(path);
  }

  private materializeAllParts(): void {
    for (const path of [...this.dirtyPartXml]) this.materializePart(path);
  }

  private replacePartBytes(path: string, bytes: Uint8Array): void {
    validatePath(path);
    if (!this.parts.has(path)) throw new Error('Use addPart with a content type to create a new part.');
    this.parts.set(path, Uint8Array.from(bytes));
    this.documents.delete(path);
    this.dirtyPartXml.delete(path);
    this.dirtyPartSizes.delete(path);
  }

  private finalizeMutation(path: string): void {
    this.assertPackageLimits();
    if (STRUCTURE_PARTS.has(path)) {
      this.materializeAllParts();
      this.mainPath = this.validatePackage();
    } else if (path === this.mainPath) {
      bodyOf(this.getCachedPartDocument(this.mainPath));
    }
    this.currentRevision++;
    this.numberingContextCache = undefined;
    this.stylesCache = undefined;
    this.outlineCache = undefined;
    this.noteStateCache = undefined;
    this.commentStateCache = undefined;
    this.revisionInfoCache = undefined;
    this.reviewerInfoCache = undefined;
    this.tableCellLocationCache = undefined;
    this.imageDataUrls.clear();
  }

  private assertPackageLimits(): void {
    if (this.parts.size > MAX_PARTS) throw new Error('Too many package parts.');
    let total = 0;
    for (const [path, bytes] of this.parts) {
      validatePath(path);
      const size = this.dirtyPartSizes.get(path) ?? bytes.byteLength;
      total += size;
      if (size > MAX_PART || total > MAX_TOTAL) throw new Error('Package exceeds size limits.');
    }
  }

  private getRelatedPartPath(type: string, fallback?: string, options: { allowMissingPart?: boolean } = {}): string | undefined {
    const relPath = relsPath(this.mainPath);
    if (this.parts.has(relPath)) {
      try {
        const document = this.getCachedPartDocument(relPath).documentElement;
        if (!document) return fallback && this.parts.has(fallback) ? fallback : undefined;
        for (const relation of children(document, 'Relationship', REL_NS)) {
          if (relation.getAttribute('Type') === type && relation.getAttribute('TargetMode') !== 'External') {
            const target = relation.getAttribute('Target');
            if (!target) continue;
            let path: string | undefined;
            try {
              path = resolveTargetPath(this.mainPath, decodeURIComponent(target));
            } catch {
              continue;
            }
            if (path && (options.allowMissingPart || this.parts.has(path))) return path;
          }
        }
      } catch { /* Fall back to conventional paths for malformed optional rels parts. */ }
    }
    return fallback && (options.allowMissingPart || this.parts.has(fallback)) ? fallback : undefined;
  }

  private getSettingsPath(allowMissingPart = false): string | undefined {
    const conventional = `${dirname(this.mainPath) ? `${dirname(this.mainPath)}/` : ''}settings.xml`;
    const fallback = this.parts.has(conventional)
      ? conventional
      : this.parts.has('word/settings.xml') ? 'word/settings.xml' : allowMissingPart ? conventional : undefined;
    return this.getRelatedPartPath(SETTINGS_REL, fallback, { allowMissingPart });
  }

  private getNotePartPath(kind: NoteKind): string | undefined {
    const conventional = `${dirname(this.mainPath) ? `${dirname(this.mainPath)}/` : ''}${kind}s.xml`;
    return this.getRelatedPartPath(noteRelationshipType(kind), this.parts.has(conventional)
      ? conventional
      : this.parts.has(`word/${kind}s.xml`) ? `word/${kind}s.xml` : undefined);
  }

  private ensureMainRelationship(relationType: string, targetPath: string): void {
    const relsPathOfMain = relsPath(this.mainPath);
    const hasRelsPart = this.parts.has(relsPathOfMain);
    const rels = hasRelsPart
      ? this.getPartDocument(relsPathOfMain)
      : parseXml(`<Relationships xmlns="${REL_NS}"/>`);
    const root = rels.documentElement!;
    let resolvedTargetPath: string | undefined;
    try {
      resolvedTargetPath = resolveTargetPath(this.mainPath, targetPath);
    } catch {
      resolvedTargetPath = undefined;
    }
    const existing = children(root, 'Relationship', REL_NS).find((relation) =>
      relation.getAttribute('Type') === relationType &&
      relation.getAttribute('TargetMode') !== 'External' &&
      (() => {
        if (!resolvedTargetPath) return relation.getAttribute('Target') === targetPath;
        const target = relation.getAttribute('Target');
        if (!target) return false;
        try {
          return resolveTargetPath(this.mainPath, decodeURIComponent(target)) === resolvedTargetPath;
        } catch {
          return false;
        }
      })());
    if (!existing) {
      const relationship = rels.createElementNS(REL_NS, 'Relationship');
      relationship.setAttribute('Id', nextRelationshipId(root));
      relationship.setAttribute('Type', relationType);
      relationship.setAttribute('Target', targetPath);
      root.appendChild(relationship);
      if (hasRelsPart) this.setPartXml(relsPathOfMain, serializeXml(rels));
      else this.addPart(relsPathOfMain, encodeXml(serializeXml(rels)), 'application/vnd.openxmlformats-package.relationships+xml');
    }
  }

  private getPackageRelatedPartPath(type: string, fallback?: string, options: { allowMissingPart?: boolean } = {}): string | undefined {
    try {
      const root = this.getCachedPartDocument('_rels/.rels').documentElement;
      if (root?.namespaceURI !== REL_NS || root.localName !== 'Relationships') {
        return fallback && (options.allowMissingPart || this.parts.has(fallback)) ? fallback : undefined;
      }
      for (const relation of children(root, 'Relationship', REL_NS)) {
        if (relation.getAttribute('Type') !== type || relation.getAttribute('TargetMode') === 'External') continue;
        const target = relation.getAttribute('Target');
        if (!target) continue;
        try {
          const path = decodePackageTarget(target);
          if (path) {
            validatePath(path);
            if (options.allowMissingPart || this.parts.has(path)) return path;
          }
        } catch {
          continue;
        }
      }
    } catch { /* Fall back to conventional paths for malformed package rels. */ }
    return fallback && (options.allowMissingPart || this.parts.has(fallback)) ? fallback : undefined;
  }

  private ensurePackageRelationship(relationType: string, targetPath: string): void {
    const hasRelsPart = this.parts.has('_rels/.rels');
    const rels = hasRelsPart
      ? this.getPartDocument('_rels/.rels')
      : parseXml(`<Relationships xmlns="${REL_NS}"/>`);
    const root = rels.documentElement!;
    const normalizedTargetPath = decodePackageTarget(targetPath);
    const existing = children(root, 'Relationship', REL_NS).find((relation) => {
      if (relation.getAttribute('Type') !== relationType || relation.getAttribute('TargetMode') === 'External') return false;
      const target = relation.getAttribute('Target');
      if (!target) return false;
      return decodePackageTarget(target) === normalizedTargetPath;
    });
    if (existing) return;
    const relationship = rels.createElementNS(REL_NS, 'Relationship');
    relationship.setAttribute('Id', nextRelationshipId(root));
    relationship.setAttribute('Type', relationType);
    relationship.setAttribute('Target', targetPath);
    root.appendChild(relationship);
    if (hasRelsPart) this.setPartXml('_rels/.rels', serializeXml(rels));
    else this.addPart('_rels/.rels', encodeXml(serializeXml(rels)), RELS_CONTENT_TYPE);
  }

  private getCorePropertiesPath(allowMissingPart = false): string | undefined {
    const fallback = this.parts.has('docProps/core.xml') ? 'docProps/core.xml' : allowMissingPart ? 'docProps/core.xml' : undefined;
    return this.getPackageRelatedPartPath(CORE_PROPS_REL, fallback, { allowMissingPart });
  }

  private getAppPropertiesPath(allowMissingPart = false): string | undefined {
    const fallback = this.parts.has('docProps/app.xml') ? 'docProps/app.xml' : allowMissingPart ? 'docProps/app.xml' : undefined;
    return this.getPackageRelatedPartPath(APP_PROPS_REL, fallback, { allowMissingPart });
  }

  private ensureNotePart(kind: NoteKind): void {
    let path = this.getNotePartPath(kind);
    if (!path) path = `${dirname(this.mainPath) ? `${dirname(this.mainPath)}/` : ''}${kind}s.xml`;
    if (!this.parts.has(path)) {
      this.addPart(path, encodeXml(defaultNotePartXml(kind)), noteContentType(kind));
    }
    this.ensureMainRelationship(noteRelationshipType(kind), relativeTarget(this.mainPath, path));
  }

  private collectNoteState(body: Element): NoteState {
    const refs = referenceRecords(body);
    const entries: Record<NoteKind, Map<number, ReturnType<typeof parseNoteEntries>[number]>> = {
      footnote: new Map(parseNoteEntries(this.getNotePartPath('footnote') ? this.getPartDocument(this.getNotePartPath('footnote')!) : null, 'footnote').map((entry) => [entry.id, entry])),
      endnote: new Map(parseNoteEntries(this.getNotePartPath('endnote') ? this.getPartDocument(this.getNotePartPath('endnote')!) : null, 'endnote').map((entry) => [entry.id, entry])),
    };
    const settingsPath = this.getSettingsPath();
    const baseSettings = parseDocumentNoteSettings(settingsPath && this.parts.has(settingsPath) ? this.getPartDocument(settingsPath) : null);
    const perSection = sectionSettings(body, baseSettings);
    const byKind: Record<NoteKind, Map<number, { number: number; marker: string }>> = {
      footnote: new Map(),
      endnote: new Map(),
    };
    for (const kind of ['footnote', 'endnote'] as const) {
      const counters = new Map<string, number>();
      for (const reference of refs.filter((item) => item.kind === kind)) {
        if (byKind[kind].has(reference.id)) continue;
        const settings = perSection.get(reference.section)?.[kind] ?? baseSettings[kind];
        const restart = settings.numRestart ?? 'continuous';
        const counterKey = restart === 'eachSect' ? `s${reference.section}` : 'all';
        if (!counters.has(counterKey)) counters.set(counterKey, (settings.numStart ?? 1) - 1);
        const number = (counters.get(counterKey) ?? 0) + 1;
        counters.set(counterKey, number);
        const noteEntry = entries[kind].get(reference.id);
        const customMark = reference.customMarkFollows && noteEntry ? parseCustomMark(noteEntry.element) : undefined;
        byKind[kind].set(reference.id, { number, marker: customMark ?? formatNoteMarker(number, settings.numFmt ?? 'decimal') });
      }
    }
    return { byKind, refs, entries };
  }

  private getNumberingPath(): string | undefined {
    const conventional = `${dirname(this.mainPath) ? `${dirname(this.mainPath)}/` : ''}numbering.xml`;
    return this.getRelatedPartPath(NUMBERING_REL, this.parts.has(conventional)
      ? conventional
      : this.parts.has('word/numbering.xml') ? 'word/numbering.xml' : undefined);
  }

  private getStylesPath(): string | undefined {
    const conventional = `${dirname(this.mainPath) ? `${dirname(this.mainPath)}/` : ''}styles.xml`;
    return this.getRelatedPartPath(STYLES_REL, this.parts.has(conventional)
      ? conventional
      : this.parts.has('word/styles.xml') ? 'word/styles.xml' : undefined);
  }

  private getStylesContext(): StylesContext {
    if (this.stylesCache?.revision === this.revision) return this.stylesCache.context;
    const stylesPath = this.getStylesPath();
    const themePath = this.getRelatedPartPath(THEME_REL, 'word/theme/theme1.xml');
    let stylesRoot: Element | undefined;
    let themeRoot: Element | undefined;
    try { stylesRoot = stylesPath ? this.getCachedPartDocument(stylesPath).documentElement ?? undefined : undefined; } catch { stylesRoot = undefined; }
    try { themeRoot = themePath ? this.getCachedPartDocument(themePath).documentElement ?? undefined : undefined; } catch { themeRoot = undefined; }
    const context = parseStyles(stylesRoot, themeRoot);
    this.stylesCache = { revision: this.revision, context };
    return context;
  }

  private getNoteState(): NoteState {
    if (this.noteStateCache?.revision === this.revision) return this.noteStateCache.state;
    const footnotePath = this.getNotePartPath('footnote');
    const endnotePath = this.getNotePartPath('endnote');
    const state = !footnotePath && !endnotePath
      ? {
          byKind: { footnote: new Map(), endnote: new Map() },
          refs: [],
          entries: { footnote: new Map(), endnote: new Map() },
        }
      : this.collectNoteState(bodyOf(this.getCachedPartDocument(this.mainPath)));
    this.noteStateCache = { revision: this.revision, state };
    return state;
  }

  private contentPartPaths(): string[] {
    if (this.contentPartPathsCache?.revision === this.revision) return [...this.contentPartPathsCache.paths];
    const paths = this.listParts().filter((path) => {
      try {
        const root = this.getCachedPartDocument(path).documentElement;
        return root?.namespaceURI === WORD_NS && ['document', 'hdr', 'ftr', 'footnotes', 'endnotes'].includes(root.localName ?? '');
      } catch {
        return false;
      }
    });
    this.contentPartPathsCache = { revision: this.revision, paths };
    return [...paths];
  }

  private relatedPartPathFor(sourcePartPath: string, relationType: string): string | undefined {
    const relPath = relsPath(sourcePartPath);
    if (!this.parts.has(relPath)) return undefined;
    try {
      const document = this.getCachedPartDocument(relPath).documentElement;
      if (!document) return undefined;
      for (const relation of children(document, 'Relationship', REL_NS)) {
        if (relation.getAttribute('Type') !== relationType || relation.getAttribute('TargetMode') === 'External') continue;
        const target = relation.getAttribute('Target');
        if (!target) continue;
        let partPath: string | undefined;
        try {
          partPath = resolveTargetPath(sourcePartPath, decodeURIComponent(target));
        } catch {
          continue;
        }
        if (partPath && this.parts.has(partPath)) return partPath;
      }
    } catch { /* ignore malformed optional relationships */ }
    return undefined;
  }

  private mayContainComments(path: string): boolean {
    const bytes = this.parts.get(path);
    if (!bytes) return false;
    try {
      const xml = decodeXml(bytes);
      return /commentRange(Start|End)|commentReference|<w:comment\b/i.test(xml);
    } catch {
      return false;
    }
  }

  private commentBindings(): CommentPartBinding[] {
    if (this.commentBindingsCache?.revision === this.revision) return this.commentBindingsCache.bindings.map((binding) => ({ ...binding }));
    const bindings = this.contentPartPaths()
      .map((sourcePartPath) => ({
        sourcePartPath,
        commentsPath: this.relatedPartPathFor(sourcePartPath, COMMENTS_REL),
        commentsExtendedPath: this.relatedPartPathFor(sourcePartPath, COMMENTS_EXTENDED_REL),
      }))
      .filter((binding) => binding.commentsPath || binding.commentsExtendedPath || this.mayContainComments(binding.sourcePartPath));
    this.commentBindingsCache = { revision: this.revision, bindings };
    return bindings.map((binding) => ({ ...binding }));
  }

  private collectCommentLocations(sourcePartPath: string): Map<number, CommentLocation> {
    const document = this.partDocumentOrUndefined(sourcePartPath);
    if (!document) return new Map();
    let container: Element;
    try {
      container = blockContainerOf(document);
    } catch {
      return new Map();
    }
    const paragraphs = descendants(container, 'p');
    const paragraphIndex = new Map(paragraphs.map((paragraph, index) => [paragraph, index]));
    const order = preOrderElements(container);
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
      for (const [runNumber, run] of ownRuns(paragraph).entries()) {
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
        const startParagraph = nearestParagraph(start);
        const endParagraph = nearestParagraph(end);
        const startParagraphNumber = startParagraph ? paragraphIndex.get(startParagraph) : undefined;
        const endParagraphNumber = endParagraph ? paragraphIndex.get(endParagraph) : undefined;
        const startOrder = orderIndex.get(start) ?? Number.MAX_SAFE_INTEGER;
        const endOrder = orderIndex.get(end) ?? Number.MAX_SAFE_INTEGER;
        if (startParagraph && endParagraph && startParagraphNumber !== undefined && endParagraphNumber !== undefined) {
          if (startParagraph === endParagraph) {
            const runs = ownRuns(startParagraph).flatMap((run, runIndex) => {
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

  private readCommentBody(entry: ReturnType<typeof parseCommentEntries>[number], commentsPath: string): { blocks: DocumentBlock[]; text: string } {
    const blocks = this.buildBlocksFromElement(entry.element, this.getStylesContext(), () => null, commentsPath);
    const text = blocks
      .flatMap((block) => block.type === 'paragraph' ? [block.paragraph.text] : [])
      .join('\n')
      .trim();
    return { blocks, text };
  }

  private getAllComments(): CommentInfo[] {
    if (this.commentStateCache?.revision === this.revision) return this.commentStateCache.comments;
    const bindings = this.commentBindings();
    if (!bindings.length) {
      this.commentStateCache = { revision: this.revision, comments: [] };
      return [];
    }
    const locations = new Map<number, CommentLocation>();
    for (const binding of bindings) {
      for (const [id, location] of this.collectCommentLocations(binding.sourcePartPath)) {
        const previous = locations.get(id);
        if (!previous || location.order < previous.order) locations.set(id, location);
      }
    }
    const entries = new Map<number, { entry: ReturnType<typeof parseCommentEntries>[number]; commentsPath: string }>();
    const exEntries = new Map<string, ReturnType<typeof parseCommentExEntries>[number]>();
    const parsedCommentsPaths = new Set<string>();
    const parsedExtendedPaths = new Set<string>();
    for (const binding of bindings) {
      if (binding.commentsPath && !parsedCommentsPaths.has(binding.commentsPath) && this.parts.has(binding.commentsPath)) {
        parsedCommentsPaths.add(binding.commentsPath);
        for (const entry of parseCommentEntries(this.partDocumentOrUndefined(binding.commentsPath) ?? null)) {
          if (!entries.has(entry.id)) entries.set(entry.id, { entry, commentsPath: binding.commentsPath });
        }
      }
      if (binding.commentsExtendedPath && !parsedExtendedPaths.has(binding.commentsExtendedPath) && this.parts.has(binding.commentsExtendedPath)) {
        parsedExtendedPaths.add(binding.commentsExtendedPath);
        for (const entry of parseCommentExEntries(this.partDocumentOrUndefined(binding.commentsExtendedPath) ?? null)) exEntries.set(entry.paraId, entry);
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
      const body = entry && entryInfo ? this.readCommentBody(entry, entryInfo.commentsPath) : undefined;
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
    this.commentStateCache = { revision: this.revision, comments: ordered };
    return ordered;
  }

  private getCommentsPartPath(sourcePartPath = this.mainPath): string | undefined {
    const conventional = `${dirname(sourcePartPath) ? `${dirname(sourcePartPath)}/` : ''}comments.xml`;
    return this.relatedPartPathFor(sourcePartPath, COMMENTS_REL) ?? (this.parts.has(conventional) ? conventional : undefined);
  }

  private getCommentsExtendedPartPath(sourcePartPath = this.mainPath): string | undefined {
    const conventional = `${dirname(sourcePartPath) ? `${dirname(sourcePartPath)}/` : ''}commentsExtended.xml`;
    return this.relatedPartPathFor(sourcePartPath, COMMENTS_EXTENDED_REL) ?? (this.parts.has(conventional) ? conventional : undefined);
  }

  private getNumberingContext(): NumberingContext {
    if (this.numberingContextCache?.revision === this.revision && this.numberingContextCache.mainPath === this.mainPath) {
      return this.numberingContextCache;
    }
    const numberingPath = this.getNumberingPath();
    const stylesPath = this.getStylesPath();
    const numbering = numberingPath && this.parts.has(numberingPath) ? this.getCachedPartDocument(numberingPath) : undefined;
    const styles = stylesPath && this.parts.has(stylesPath) ? this.getCachedPartDocument(stylesPath) : undefined;
    const context = {
      revision: this.revision,
      mainPath: this.mainPath,
      numberingPath,
      stylesPath,
      model: parseNumberingModel(numbering, styles),
    };
    this.numberingContextCache = context;
    return context;
  }

  private buildParagraphs(
    document = this.getCachedPartDocument(this.mainPath),
    styles = this.getStylesContext(),
    numbering = this.getNumberingContext(),
    sourcePartPath = this.mainPath,
    noteState?: NoteState,
  ): ParagraphInfo[] {
    const elements = descendants(blockContainerOf(document), 'p');
    const numberingByParagraph = computeParagraphNumbering(elements, numbering.model);
    const noteNumber = noteState ? (kind: NoteKind, id: number) => noteState.byKind[kind].get(id) ?? null : undefined;
    const imageContext: ImageReadContext = {
      relationships: this.relationshipsFor(sourcePartPath),
      getContentType: this.createContentTypeResolver(),
      sourcePartPath,
    };
    return elements.map((paragraph, index) =>
      readParagraph(paragraph, index, styles, numberingByParagraph.get(paragraph), imageContext, noteNumber));
  }

  private pageBreakMarkers(paragraph: Element): { before: number; after: number } {
    let before = 0;
    let after = 0;
    let seenText = false;
    const walk = (node: Node): void => {
      for (let child = node.firstChild; child; child = child.nextSibling) {
        if (child.nodeType !== 1) continue;
        const element = child as Element;
        if (element.namespaceURI !== WORD_NS) continue;
        if (element.localName === 'p') continue;
        if (element.localName === 'br' && (element.getAttributeNS(WORD_NS, 'type') ?? '') === 'page') {
          if (seenText) after++; else before++;
          continue;
        }
        if (['t', 'tab', 'cr'].includes(element.localName ?? '') ||
            (element.localName === 'br' && (element.getAttributeNS(WORD_NS, 'type') ?? '') !== 'page')) {
          seenText = true;
          continue;
        }
        walk(element);
      }
    };
    for (const run of ownRuns(paragraph)) walk(run);
    return { before, after };
  }

  private sectionBreakByParagraph(document: Document): Map<number, SectionType> {
    const root = document.documentElement;
    if (!root || root.namespaceURI !== WORD_NS || root.localName !== 'document') return new Map();
    const sections = collectSections(document);
    return new Map<number, SectionType>(sections
      .filter(section => section.source === 'paragraph')
      .map(section => [
        section.endParagraph,
        (children(section.sectPr, 'type')[0]?.getAttributeNS(WORD_NS, 'val') ?? 'nextPage') as SectionType,
      ]));
  }

  private buildBlocksFrom(document: Document, paragraphs: ParagraphInfo[]): DocumentBlock[] {
    const body = blockContainerOf(document);
    const paragraphElements = descendants(body, 'p');
    const paragraphByElement = new Map(paragraphElements.map((paragraph, index) => [paragraph, paragraphs[index]!]));
    const paragraphIndexByElement = new Map(paragraphElements.map((paragraph, index) => [paragraph, index]));
    const sectionByParagraph = this.sectionBreakByParagraph(document);
    const walk = (parent: Element): DocumentBlock[] => children(parent).flatMap((child): DocumentBlock[] => {
      if (child.localName === 'p') {
        const paragraph = paragraphByElement.get(child);
        return paragraph ? [{ type: 'paragraph', paragraph }] : [];
      }
      if (child.localName === 'tbl') return [readTable(child, walk)];
      if (isTransparentWordWrapper(child)) return walk(child);
      return [];
    });
    const blocks = walk(body);
    const result: DocumentBlock[] = [];
    let sectionBreaks = 0;
    for (const block of blocks) {
      if (block.type !== 'paragraph') {
        result.push(block);
        continue;
      }
      const paragraphIndex = block.paragraph.index;
      const paragraphElement = paragraphElements[paragraphIndex];
      const markers = paragraphElement ? this.pageBreakMarkers(paragraphElement) : { before: 0, after: 0 };
      for (let i = 0; i < markers.before; i++) result.push({ type: 'pageBreak' });
      result.push(block);
      for (let i = 0; i < markers.after; i++) result.push({ type: 'pageBreak' });
      const mappedIndex = paragraphElement ? (paragraphIndexByElement.get(paragraphElement) ?? paragraphIndex) : paragraphIndex;
      const breakType = sectionByParagraph.get(mappedIndex);
      if (breakType) result.push({ type: 'sectionBreak', section: sectionBreaks++, breakType });
    }
    return result;
  }

  private buildBlocksFromElement(
    root: Element,
    styles: StylesContext,
    noteNumber: (kind: NoteKind, id: number) => { number: number; marker: string } | null,
    sourcePartPath: string,
  ): DocumentBlock[] {
    const imageContext: ImageReadContext = {
      relationships: this.relationshipsFor(sourcePartPath),
      getContentType: this.createContentTypeResolver(),
      sourcePartPath,
    };
    const walk = (parent: Element): DocumentBlock[] => bodyChildren(parent).flatMap((child): DocumentBlock[] => {
      if (child.localName === 'p') {
        return [{ type: 'paragraph', paragraph: readParagraph(child, -1, styles, undefined, imageContext, noteNumber) }];
      }
      if (child.localName === 'tbl') return [readTable(child, walk)];
      return [];
    });
    return walk(root);
  }

  getParagraphs(): ParagraphInfo[] {
    const document = this.getCachedPartDocument(this.mainPath);
    return this.buildParagraphs(document, this.getStylesContext(), this.getNumberingContext(), this.mainPath, this.getNoteState());
  }

  getBlocks(): DocumentBlock[] {
    const document = this.getCachedPartDocument(this.mainPath);
    const styles = this.getStylesContext();
    const paragraphs = this.buildParagraphs(document, styles, this.getNumberingContext(), this.mainPath, this.getNoteState());
    return this.buildBlocksFrom(document, paragraphs);
  }

  private getTableCellLocationIndex(): Map<number, TableCellLocation> {
    if (this.tableCellLocationCache?.revision === this.revision) return this.tableCellLocationCache.locations;
    const locations = new Map<number, TableCellLocation>();
    let tableIndex = 0;
    const setActiveVertical = (
      target: Map<number, { start: number; end: number; location: TableCellLocation }>,
      start: number,
      end: number,
      location: TableCellLocation,
    ): void => {
      for (let index = start; index < end; index++) target.set(index, { start, end, location });
    };
    const walkNestedBlocks = (blocks: DocumentBlock[], location: TableCellLocation, nested: boolean): void => {
      for (const block of blocks) {
        if (block.type === 'paragraph') {
          locations.set(block.paragraph.index, { ...location, nested });
          continue;
        }
        if (block.type !== 'table') continue;
        for (const row of block.rows) {
          for (const cell of row.cells) walkNestedBlocks(cell.blocks, location, true);
        }
      }
    };
    for (const block of this.getBlocks()) {
      if (block.type !== 'table') continue;
      const activeVertical = new Map<number, { start: number; end: number; location: TableCellLocation }>();
      for (const [rowIndex, row] of block.rows.entries()) {
        const nextActive = new Map<number, { start: number; end: number; location: TableCellLocation }>();
        let colIndex = 0;
        let activeHorizontalEnd: number | null = null;
        let activeHorizontalLocation: TableCellLocation | null = null;
        for (const cell of row.cells) {
          const col = colIndex;
          const colSpan = Math.max(1, cell.colSpan);
          colIndex += colSpan;
          let location: TableCellLocation = {
            table: tableIndex,
            row: rowIndex,
            col,
            rowSpan: Math.max(1, cell.rowSpan),
            colSpan,
            nested: false,
          };
          const vertical = activeVertical.get(col);
          if (cell.isMergeContinuation) {
            if (cell.rowSpan === 0 && vertical && vertical.start <= col && vertical.end >= col + colSpan) {
              location = vertical.location;
              setActiveVertical(nextActive, col, col + colSpan, vertical.location);
            } else if (activeHorizontalLocation && activeHorizontalEnd === col) {
              location = activeHorizontalLocation;
              activeHorizontalEnd = col + colSpan;
              activeHorizontalLocation = location;
            } else {
              activeHorizontalEnd = col + colSpan;
              activeHorizontalLocation = location;
              if (cell.rowSpan > 1) setActiveVertical(nextActive, col, col + colSpan, location);
            }
          } else {
            activeHorizontalEnd = col + colSpan;
            activeHorizontalLocation = location;
            if (cell.rowSpan > 1) setActiveVertical(nextActive, col, col + colSpan, location);
          }
          walkNestedBlocks(cell.blocks, location, false);
        }
        activeVertical.clear();
        for (const [start, merge] of nextActive) activeVertical.set(start, merge);
      }
      tableIndex++;
    }
    this.tableCellLocationCache = { revision: this.revision, locations };
    return locations;
  }

  getTableCellAt(paragraphIndex: number): TableCellLocation | null {
    if (!Number.isSafeInteger(paragraphIndex) || paragraphIndex < 0) return null;
    const location = this.getTableCellLocationIndex().get(paragraphIndex);
    return location ? cloneTableCellLocation(location) : null;
  }

  getComments(filter: { authors?: string[]; resolved?: boolean } = {}): CommentInfo[] {
    if (!filter || typeof filter !== 'object' || Array.isArray(filter)) throw new Error('filter must be an object.');
    if ('authors' in filter && filter.authors !== undefined) {
      if (!Array.isArray(filter.authors)) throw new Error('filter.authors must be an array.');
      for (const author of filter.authors) assertText(author, 'filter.authors[]');
    }
    if ('resolved' in filter && filter.resolved !== undefined && typeof filter.resolved !== 'boolean') {
      throw new Error('filter.resolved must be boolean.');
    }
    let comments = this.getAllComments();
    if (filter.authors?.length) {
      const authors = normalizeReviewerFilterAuthors(filter.authors);
      comments = comments.filter((comment) => authors.has(reviewerBucketKey(reviewerBucketOf(comment.author))));
    }
    if (filter.resolved !== undefined) comments = comments.filter((comment) => comment.resolved === filter.resolved);
    return comments.map(cloneCommentInfo);
  }

  getReviewers(): ReviewerInfo[] {
    if (this.reviewerInfoCache?.revision === this.revision) return this.reviewerInfoCache.reviewers.map(cloneReviewerInfo);
    const revisions = this.getRevisions();
    const comments = this.getComments().filter((comment) =>
      comment.anchor?.sourcePartPath === this.mainDocumentPath);
    if (!revisions.length && !comments.length) {
      this.reviewerInfoCache = { revision: this.revision, reviewers: [] };
      return [];
    }
    type Aggregate = ReviewerInfo & {
      firstDateMs?: number;
      lastDateMs?: number;
      initialsCount: Map<string, number>;
      initialsSeenOrder: Map<string, number>;
      initialsOrderSeed: number;
    };
    const byAuthor = new Map<string, Aggregate>();
    const dateUpdate = (entry: Aggregate, value: string | undefined): void => {
      if (!value) return;
      const ms = Date.parse(value);
      if (!Number.isFinite(ms)) return;
      if (entry.firstDateMs === undefined || ms < entry.firstDateMs) {
        entry.firstDateMs = ms;
        entry.firstDate = value;
      }
      if (entry.lastDateMs === undefined || ms > entry.lastDateMs) {
        entry.lastDateMs = ms;
        entry.lastDate = value;
      }
    };
    const aggregateFor = (author: ReviewerInfo['author'], kind: ReviewerInfo['kind']): Aggregate => {
      const existing = byAuthor.get(reviewerBucketKey({ author, kind }));
      if (existing) return existing;
      const created: Aggregate = {
        ...(author !== undefined ? { author } : {}),
        kind,
        revisionCount: 0,
        commentCount: 0,
        unresolvedCommentCount: 0,
        initialsCount: new Map(),
        initialsSeenOrder: new Map(),
        initialsOrderSeed: 0,
      };
      byAuthor.set(reviewerBucketKey(created), created);
      return created;
    };
    const countedMovePairs = new Set<string>();
    for (const revision of revisions) {
      if (revision.kind === 'move' && revision.move?.pairedId !== undefined) {
        if (revision.move.side === 'from') continue;
        const key = `${Math.min(revision.id, revision.move.pairedId)}:${Math.max(revision.id, revision.move.pairedId)}`;
        if (countedMovePairs.has(key)) continue;
        countedMovePairs.add(key);
      }
      const reviewer = reviewerBucketOf(revision.author);
      const entry = aggregateFor(reviewer.author, reviewer.kind);
      entry.revisionCount++;
      dateUpdate(entry, revision.date);
    }
    for (const comment of comments) {
      const reviewer = reviewerBucketOf(comment.author);
      const entry = aggregateFor(reviewer.author, reviewer.kind);
      entry.commentCount++;
      if (comment.resolved !== true) entry.unresolvedCommentCount++;
      dateUpdate(entry, comment.date);
      const initials = comment.initials?.trim();
      if (!initials) continue;
      entry.initialsCount.set(initials, (entry.initialsCount.get(initials) ?? 0) + 1);
      if (!entry.initialsSeenOrder.has(initials)) {
        entry.initialsSeenOrder.set(initials, entry.initialsOrderSeed++);
      }
    }
    const reviewers = [...byAuthor.values()].map((entry): ReviewerInfo => {
      const initials = [...entry.initialsCount.entries()]
        .sort((a, b) =>
          b[1] - a[1] ||
          (entry.initialsSeenOrder.get(a[0]) ?? Number.MAX_SAFE_INTEGER) - (entry.initialsSeenOrder.get(b[0]) ?? Number.MAX_SAFE_INTEGER) ||
          a[0].localeCompare(b[0]))[0]?.[0];
      return {
        ...(entry.author !== undefined ? { author: entry.author } : {}),
        kind: entry.kind,
        ...(initials !== undefined ? { initials } : {}),
        revisionCount: entry.revisionCount,
        commentCount: entry.commentCount,
        unresolvedCommentCount: entry.unresolvedCommentCount,
        ...(entry.firstDate !== undefined ? { firstDate: entry.firstDate } : {}),
        ...(entry.lastDate !== undefined ? { lastDate: entry.lastDate } : {}),
      };
    }).sort((a, b) =>
      b.revisionCount - a.revisionCount ||
      b.commentCount - a.commentCount ||
      b.unresolvedCommentCount - a.unresolvedCommentCount ||
      reviewerBucketKey(a).localeCompare(reviewerBucketKey(b)));
    this.reviewerInfoCache = { revision: this.revision, reviewers };
    return reviewers.map(cloneReviewerInfo);
  }

  getSnapshot(): DocumentSnapshot {
    const document = this.getCachedPartDocument(this.mainPath);
    const noteState = this.getNoteState();
    const stylesContext = this.getStylesContext();
    const numberingContext = this.getNumberingContext();
    const paragraphs = this.buildParagraphs(document, stylesContext, numberingContext, this.mainPath, noteState);
    return {
      revision: this.revision,
      paragraphs,
      blocks: this.buildBlocksFrom(document, paragraphs),
      footnotes: this.getNotesWith('footnote', noteState),
      endnotes: this.getNotesWith('endnote', noteState),
      comments: this.getComments(),
      parts: this.listParts(),
      styles: stylesContext.styles.map(cloneStyleInfo),
      hyperlinks: this.getHyperlinks(),
      bookmarks: this.getBookmarks(),
    };
  }

  getFootnotes(): NoteInfo[] {
    return this.getNotesWith('footnote', this.getNoteState());
  }

  getEndnotes(): NoteInfo[] {
    return this.getNotesWith('endnote', this.getNoteState());
  }

  private getNotesWith(kind: NoteKind, state: NoteState): NoteInfo[] {
    const styles = this.getStylesContext();
    const numberOf = (noteKind: NoteKind, id: number) => state.byKind[noteKind].get(id) ?? null;
    const entries = state.entries[kind];
    const results: NoteInfo[] = [];
    const seen = new Set<number>();
    for (const reference of state.refs.filter((item) => item.kind === kind)) {
      if (seen.has(reference.id)) continue;
      seen.add(reference.id);
      const entry = entries.get(reference.id);
      if (entry && (entry.id < 1 || entry.type !== 'normal')) continue;
      const numbering = state.byKind[kind].get(reference.id) ?? { number: 0, marker: '?' };
      const blocks = entry
        ? this.buildBlocksFromElement(entry.element, styles, numberOf, this.getNotePartPath(kind) ?? this.mainPath)
        : [];
      results.push({
        id: reference.id,
        kind,
        number: numbering.number,
        marker: numbering.marker,
        customMark: reference.customMarkFollows && entry ? parseCustomMark(entry.element) : undefined,
        blocks,
        reference: { paragraph: reference.paragraph, run: reference.run },
      });
    }
    return results;
  }

  getNumberingDefinitions(): NumberingDefinition[] {
    return this.getNumberingContext().model.definitions;
  }

  getStyles(): StyleInfo[] {
    return this.getStylesContext().styles.map(cloneStyleInfo);
  }

  getStyle(id: string): StyleInfo | undefined {
    assertText(id, 'style id');
    const style = this.getStylesContext().byId.get(id);
    return style ? cloneStyleInfo(style) : undefined;
  }

  getStyleGallery(): StyleInfo[] {
    return this.getStylesContext().styles
      .filter((style) => style.quickFormat)
      .slice()
      .sort((a, b) =>
        (a.uiPriority ?? Number.MAX_SAFE_INTEGER) - (b.uiPriority ?? Number.MAX_SAFE_INTEGER) ||
        a.name.localeCompare(b.name) ||
        a.id.localeCompare(b.id))
      .map(cloneStyleInfo);
  }

  private paragraphStyleChain(styleId: string | undefined): StyleInfo[] {
    return resolveStyleChainOrDefault(this.getStylesContext(), styleId, 'paragraph').map(cloneStyleInfo);
  }

  private styleFormats(styleId: string, type: StyleInfo['type']): { paragraph: ParagraphFormat; run: RunFormat; chain: StyleInfo[] } {
    const chain = resolveStyleChainOrDefault(this.getStylesContext(), styleId, type).map(cloneStyleInfo);
    const paragraph: ParagraphFormat = {};
    const run: RunFormat = {};
    for (const style of chain) {
      if (style.paragraph) {
        for (const [key, value] of Object.entries(style.paragraph) as Array<[keyof ParagraphFormat, ParagraphFormat[keyof ParagraphFormat]]>) {
          if (value !== undefined && value !== null) (paragraph as Record<string, unknown>)[key] = cloneRunFormatValue(value);
        }
      }
      if (style.run) {
        for (const [key, value] of Object.entries(style.run) as Array<[keyof RunFormat, RunFormat[keyof RunFormat]]>) {
          if (value !== undefined && value !== null) (run as Record<string, unknown>)[key] = cloneRunFormatValue(value);
        }
      }
    }
    return { paragraph, run, chain };
  }

  private paragraphOutlineLevel(paragraph: ParagraphInfo): number | undefined {
    if (paragraph.outlineLevel !== undefined && paragraph.outlineLevel !== null) {
      return paragraph.outlineLevel >= 0 && paragraph.outlineLevel <= OUTLINE_MAX_LEVEL ? paragraph.outlineLevel : undefined;
    }
    for (const style of this.paragraphStyleChain(paragraph.style ?? undefined)) {
      const level = headingOutlineLevel(style.name);
      if (level !== undefined) return level;
    }
    return undefined;
  }

  private buildOutline(): OutlineNode[] {
    const outline: OutlineNode[] = [];
    const stack: OutlineNode[] = [];
    for (const paragraph of this.getParagraphs()) {
      const level = this.paragraphOutlineLevel(paragraph);
      if (level === undefined) continue;
      const node: OutlineNode = {
        paragraph: paragraph.index,
        level,
        text: paragraph.text,
        ...(paragraph.style ? { styleId: paragraph.style } : {}),
        children: [],
      };
      while (stack.length && stack[stack.length - 1]!.level >= level) stack.pop();
      if (stack.length) stack[stack.length - 1]!.children.push(node);
      else outline.push(node);
      stack.push(node);
    }
    return outline;
  }

  getOutline(): OutlineNode[] {
    if (this.outlineCache?.revision === this.revision) return this.outlineCache.outline.map(cloneOutlineNode);
    const outline = this.buildOutline();
    this.outlineCache = { revision: this.revision, outline };
    return outline.map(cloneOutlineNode);
  }

  private flattenOutline(nodes: OutlineNode[]): OutlineNode[] {
    const flat: OutlineNode[] = [];
    const walk = (items: OutlineNode[]): void => {
      for (const item of items) {
        flat.push(item);
        walk(item.children);
      }
    };
    walk(nodes);
    return flat;
  }

  private bodyBlockEntries(document = this.getCachedPartDocument(this.mainPath)): Array<{ block: Element; parent: Element; paragraphs: number[] }> {
    const body = bodyOf(document);
    const paragraphIndex = new Map<Element, number>();
    let nextIndex = 0;
    const collectParagraphs = (parent: Element): void => {
      for (const child of children(parent)) {
        if (isTransparentWrapper(child)) {
          collectParagraphs(child);
          continue;
        }
        if (child.localName === 'p') {
          paragraphIndex.set(child, nextIndex++);
          continue;
        }
        if (child.localName === 'tbl') {
          for (const row of childrenThroughTransparent(child, 'tr')) {
            for (const cell of childrenThroughTransparent(row, 'tc')) collectParagraphs(cell);
          }
        }
      }
    };
    const blockParagraphs = (block: Element): Element[] => {
      if (block.localName === 'p') return [block];
      if (block.localName !== 'tbl') return [];
      const result: Element[] = [];
      for (const row of childrenThroughTransparent(block, 'tr')) {
        for (const cell of childrenThroughTransparent(row, 'tc')) {
          for (const entry of blockElements(cell)) {
            if (entry.localName === 'p') result.push(entry);
            else result.push(...blockParagraphs(entry));
          }
        }
      }
      return result;
    };
    collectParagraphs(body);
    return blockPositions(body)
      .map(({ block, parent }) => ({
        block,
        parent,
        paragraphs: blockParagraphs(block)
          .map((paragraph) => paragraphIndex.get(paragraph))
          .filter((index): index is number => index !== undefined),
      }))
      .filter((entry) => entry.paragraphs.length);
  }

  getEffectiveParagraphFormat(index: number): ParagraphFormat {
    const paragraph = paragraphAt(this.getCachedPartDocument(this.mainPath), index);
    return computeEffectiveParagraphFormat(this.getStylesContext(), paragraph);
  }

  getEffectiveRunFormat(paragraph: number, run: number): RunFormat {
    assertIndex(run);
    const document = this.getCachedPartDocument(this.mainPath);
    const paragraphElement = paragraphAt(document, paragraph);
    const runElement = ownRuns(paragraphElement)[run];
    if (!runElement) throw new Error(`Run ${run} does not exist.`);
    return computeEffectiveRunFormat(this.getStylesContext(), paragraphElement, runElement);
  }

  applyParagraphStyle(index: number, styleId: string, options: { clearDirectFormat?: boolean } = {}): void {
    assertIndex(index);
    assertText(styleId, 'styleId');
    if (typeof options !== 'object' || !options || Array.isArray(options)) throw new Error('options must be an object.');
    if ('clearDirectFormat' in options && options.clearDirectFormat !== undefined && typeof options.clearDirectFormat !== 'boolean') {
      throw new Error('options.clearDirectFormat must be boolean.');
    }
    const style = this.getStyle(styleId);
    if (!style || style.type !== 'paragraph') throw new Error(`Paragraph style not found: ${styleId}.`);
    const paragraph = this.getParagraphs()[index];
    if (!paragraph) throw new Error(`Paragraph ${index} does not exist.`);
    const clearDirectFormat = options.clearDirectFormat === true;
    if (!clearDirectFormat) {
      if (paragraph.style === styleId) return;
      this.formatParagraph(index, { style: styleId }, { validateStyle: true });
      return;
    }
    const preview = this.getPartDocument(this.mainPath);
    const previewParagraph = paragraphAt(preview, index);
    applyParagraphFormatTo(properties(previewParagraph, 'pPr'), { style: styleId });
    for (const field of PARAGRAPH_DIRECT_FIELDS) {
      if (field === 'style' || paragraph[field] === undefined) continue;
      applyParagraphFormatTo(properties(previewParagraph, 'pPr'), { [field]: null } as ParagraphFormat);
    }
    const previewRuns = ownRuns(previewParagraph);
    for (const [runIndex, run] of previewRuns.entries()) {
      for (const field of RUN_FORMAT_FIELDS) {
        if (field === 'style' || paragraph.runs[runIndex]?.[field] === undefined) continue;
        applyRunFormatTo(properties(run, 'rPr'), { [field]: null } as RunFormat);
      }
    }
    const targetParagraph = computeEffectiveParagraphFormat(this.getStylesContext(), previewParagraph);
    const targetRuns = previewRuns.map((run) => computeEffectiveRunFormat(this.getStylesContext(), previewParagraph, run));
    const paragraphPatch: ParagraphFormat = { style: styleId };
    let paragraphChanged = paragraph.style !== styleId;
    for (const field of PARAGRAPH_DIRECT_FIELDS) {
      if (field === 'style') continue;
      const direct = paragraph[field];
      const target = targetParagraph[field];
      if (direct !== undefined && target !== undefined && !runFormatValueEqual(direct, target)) {
        (paragraphPatch as Record<string, unknown>)[field] = null;
        paragraphChanged = true;
      }
    }
    const runPatches = paragraph.runs.map((run) => {
      const patch: RunFormat = {};
      let changed = false;
      for (const field of RUN_FORMAT_FIELDS) {
        if (field === 'style') continue;
        const direct = run[field];
        const target = targetRuns[run.index]?.[field];
        if (direct !== undefined && target !== undefined && !runFormatValueEqual(direct, target)) {
          (patch as Record<string, unknown>)[field] = null;
          changed = true;
        }
      }
      return changed ? { run: run.index, patch } : undefined;
    }).filter((entry): entry is { run: number; patch: RunFormat } => Boolean(entry));
    if (!paragraphChanged && !runPatches.length) return;
    this.withDraft((draft) => {
      const document = draft.getCachedPartDocument(draft.mainPath);
      const paragraphElement = paragraphAt(document, index);
      if (paragraphChanged) {
        if (draft.trackChangesEnabled()) draft.trackParagraphFormatChange(paragraphElement);
        applyParagraphFormatTo(properties(paragraphElement, 'pPr'), paragraphPatch);
      }
      if (runPatches.length) {
        const runs = ownRuns(paragraphElement);
        for (const { run, patch } of runPatches) {
          const element = runs[run];
          if (!element) continue;
          if (draft.trackChangesEnabled()) draft.trackRunFormatChange(element);
          applyRunFormatTo(properties(element, 'rPr'), patch);
        }
      }
      draft.dirtyPartXml.add(draft.mainPath);
      draft.dirtyPartSizes.delete(draft.mainPath);
    });
  }

  applyCharacterStyle(range: TextRange, styleId: string, options: { clearDirectFormat?: boolean } = {}): void {
    assertText(styleId, 'styleId');
    if (typeof options !== 'object' || !options || Array.isArray(options)) throw new Error('options must be an object.');
    if ('clearDirectFormat' in options && options.clearDirectFormat !== undefined && typeof options.clearDirectFormat !== 'boolean') {
      throw new Error('options.clearDirectFormat must be boolean.');
    }
    const style = this.getStyle(styleId);
    if (!style || style.type !== 'character') throw new Error(`Character style not found: ${styleId}.`);
    const clearDirectFormat = options.clearDirectFormat === true;
    const preview = this.getPartDocument(this.mainPath);
    const normalized = this.normalizeRangeOn(preview, range);
    if (normalized.start === normalized.end) return;
    const previewParagraph = normalized.paragraph;
    this.splitRunAtOffset(previewParagraph, normalized.end);
    this.splitRunAtOffset(previewParagraph, normalized.start);
    const previewRuns = this.runsInRange(previewParagraph, normalized.start, normalized.end);
    const directRuns = previewRuns.map((run) => readRunProperties(children(run, 'rPr')[0], this.getStylesContext().theme));
    for (const run of previewRuns) {
      applyRunFormatTo(properties(run, 'rPr'), { style: styleId });
      for (const field of RUN_FORMAT_FIELDS) {
        if (field === 'style') continue;
        applyRunFormatTo(properties(run, 'rPr'), { [field]: null } as RunFormat);
      }
    }
    const targetRuns = previewRuns.map((run) => computeEffectiveRunFormat(this.getStylesContext(), previewParagraph, run));
    const patches = directRuns.map((direct, index) => {
      const patch: RunFormat = {};
      let changed = direct.style !== styleId;
      if (direct.style !== styleId) patch.style = styleId;
      if (clearDirectFormat) {
        for (const field of RUN_FORMAT_FIELDS) {
          if (field === 'style') continue;
          const current = direct[field];
          const target = targetRuns[index]?.[field];
          if (current !== undefined && target !== undefined && !runFormatValueEqual(current, target)) {
            (patch as Record<string, unknown>)[field] = null;
            changed = true;
          }
        }
      }
      return changed ? { run: index, patch } : undefined;
    }).filter((entry): entry is { run: number; patch: RunFormat } => Boolean(entry));
    if (!patches.length) return;
    this.updatePartXmlInternal(this.mainPath, (document) => {
      const current = this.normalizeRangeOn(document, range);
      this.splitRunAtOffset(current.paragraph, current.end);
      this.splitRunAtOffset(current.paragraph, current.start);
      const runs = this.runsInRange(current.paragraph, current.start, current.end);
      for (const { run, patch } of patches) {
        const element = runs[run];
        if (!element) continue;
        if (this.trackChangesEnabled()) this.trackRunFormatChange(element);
        applyRunFormatTo(properties(element, 'rPr'), patch);
      }
    });
  }

  createStyleFromSelection(range: DocumentRange, style: { id: string; name: string; basedOn?: string }): StyleInfo {
    assertText(style.id, 'style.id');
    assertText(style.name, 'style.name');
    if (!style.id.trim()) throw new Error('style.id must not be empty or whitespace.');
    if (!style.name.trim()) throw new Error('style.name must not be empty or whitespace.');
    if (style.basedOn !== undefined) assertText(style.basedOn, 'style.basedOn');
    if (this.getStyle(style.id)) throw new Error(`Style already exists: ${style.id}.`);
    if (style.basedOn) {
      const base = this.getStyle(style.basedOn);
      if (!base || base.type !== 'paragraph') throw new Error(`Base paragraph style not found: ${style.basedOn}.`);
    }
    const preview = this.getPartDocument(this.mainPath);
    const normalized = this.normalizeDocumentRange(preview, range);
    const paragraphs = this.getParagraphs().slice(normalized.start.paragraph, normalized.end.paragraph + 1);
    if (!paragraphs.length) throw new Error('Range does not contain any paragraphs.');
    const paragraphFormats = paragraphs.map((paragraph) => this.directParagraphFormat(paragraph));
    const paragraphFormat: ParagraphFormat = {};
    for (const field of PARAGRAPH_DIRECT_FIELDS) {
      if (field === 'style') continue;
      const first = paragraphFormats[0]?.[field];
      if (first === undefined) continue;
      if (paragraphFormats.every((format) => runFormatValueEqual(format[field], first))) {
        (paragraphFormat as Record<string, unknown>)[field] = cloneRunFormatValue(first);
      }
    }
    const runFormats: RunFormat[] = [];
    const targetContext = this.getStylesContext();
    for (let paragraphIndex = normalized.start.paragraph; paragraphIndex <= normalized.end.paragraph; paragraphIndex++) {
      const paragraph = paragraphAt(preview, paragraphIndex);
      const length = textOf(paragraph).length;
      const start = paragraphIndex === normalized.start.paragraph ? normalized.start.offset : 0;
      const end = paragraphIndex === normalized.end.paragraph ? normalized.end.offset : length;
      if (start === end) continue;
      this.splitRunAtOffset(paragraph, end);
      this.splitRunAtOffset(paragraph, start);
      for (const run of this.runsInRange(paragraph, start, end)) {
        const effective = computeEffectiveRunFormat(targetContext, paragraph, run);
        if (Object.keys(effective).length) runFormats.push(effective);
      }
    }
    const runFormat = this.mergeRangeFormats(runFormats);
    const baseRunFormat = style.basedOn ? this.styleFormats(style.basedOn, 'paragraph').run : {};
    for (const field of RUN_FORMAT_FIELDS) {
      if (field === 'style') continue;
      if (runFormat[field] !== undefined && runFormatValueEqual(runFormat[field], baseRunFormat[field])) {
        delete (runFormat as Record<string, unknown>)[field];
      }
    }
    const nextStyle: StyleInfo = {
      id: style.id,
      name: style.name,
      type: 'paragraph',
      ...(style.basedOn ? { basedOn: style.basedOn } : {}),
      ...(Object.keys(paragraphFormat).length ? { paragraph: paragraphFormat } : {}),
      ...(Object.keys(runFormat).length ? { run: runFormat } : {}),
    };
    this.defineStyle(nextStyle);
    return cloneStyleInfo(nextStyle);
  }

  setOutlineLevel(index: number, level: number | null): void {
    assertIndex(index);
    if (level !== null && (!Number.isSafeInteger(level) || level < 0 || level > OUTLINE_MAX_LEVEL)) {
      throw new Error(`level must be null or an integer from 0 to ${OUTLINE_MAX_LEVEL}.`);
    }
    const paragraph = this.getParagraphs()[index];
    if (!paragraph) throw new Error(`Paragraph ${index} does not exist.`);
    if ((paragraph.outlineLevel ?? null) === level) return;
    this.formatParagraph(index, { outlineLevel: level });
  }

  moveOutlineSection(from: number, to: number): void {
    assertIndex(from);
    assertIndex(to);
    const paragraphs = this.getParagraphs();
    if (!paragraphs[from]) throw new Error(`Paragraph ${from} does not exist.`);
    if (to > paragraphs.length) throw new Error(`Destination ${to} is out of bounds for ${paragraphs.length} paragraphs.`);
    const tree = this.getOutline();
    const outline = this.flattenOutline(tree);
    const node = outline.find((entry) => entry.paragraph === from);
    if (!node) throw new Error(`Paragraph ${from} is not an outline heading.`);
    const subtree = this.flattenOutline([cloneOutlineNode(node)]);
    const subtreeEnd = subtree.at(-1)?.paragraph ?? from;
    const nextPeer = outline.find((entry) => entry.paragraph > subtreeEnd && entry.level <= node.level);
    const endParagraph = nextPeer ? nextPeer.paragraph - 1 : paragraphs.length - 1;
    const comments = this.getComments().filter((comment) =>
      comment.anchor?.sourcePartPath === this.mainDocumentPath &&
      !('runs' in comment.anchor) &&
      comment.anchor.startParagraph <= endParagraph &&
      comment.anchor.endParagraph >= from &&
      (comment.anchor.startParagraph < from || comment.anchor.endParagraph > endParagraph));
    if (comments.length) throw new Error('Cannot move a section that would split a comment range.');
    const entries = this.bodyBlockEntries(this.getPartDocument(this.mainPath));
    const startBlock = entries.findIndex((entry) => entry.paragraphs.includes(from));
    if (startBlock === -1) throw new Error('Outline section must start on a movable top-level body block.');
    const moved: number[] = [];
    for (let index = startBlock; index < entries.length; index++) {
      const entry = entries[index]!;
      if (!entry.paragraphs.some((paragraph) => paragraph >= from && paragraph <= endParagraph)) break;
      if (entry.paragraphs[0]! < from || entry.paragraphs[entry.paragraphs.length - 1]! > endParagraph) {
        throw new Error('Outline section must align to whole top-level body blocks.');
      }
      moved.push(index);
    }
    if (!moved.length) throw new Error('Outline section must start on a movable top-level body block.');
    const endBlock = moved[moved.length - 1]!;
    const targetBlock = to === paragraphs.length ? entries.length : entries.findIndex((entry) => entry.paragraphs.includes(to));
    if (targetBlock === -1) throw new Error(`Paragraph ${to} does not exist.`);
    if (targetBlock >= startBlock && targetBlock <= endBlock + 1) return;
    if ([...entries.slice(startBlock, endBlock + 1), ...(targetBlock < entries.length ? [entries[targetBlock]!] : [])]
      .some((entry) => entry.parent.localName !== 'body')) {
      throw new Error('moveOutlineSection only supports top-level body sections.');
    }
    this.withDraft((draft) => {
      const document = draft.getCachedPartDocument(draft.mainPath);
      const body = bodyOf(document);
      const currentEntries = draft.bodyBlockEntries(document);
      const blockCount = endBlock - startBlock + 1;
      const fragment = document.createDocumentFragment();
      for (const entry of currentEntries.slice(startBlock, endBlock + 1)) fragment.appendChild(entry.block);
      const remaining = draft.bodyBlockEntries(document).filter((entry) => entry.parent === body);
      const adjustedTarget = targetBlock > endBlock ? targetBlock - blockCount : targetBlock;
      const anchor = adjustedTarget < remaining.length
        ? remaining[adjustedTarget]!.block
        : children(body, 'sectPr')[0] ?? null;
      body.insertBefore(fragment, anchor);
      draft.dirtyPartXml.add(draft.mainPath);
      draft.dirtyPartSizes.delete(draft.mainPath);
    });
  }

  getHyperlinks(): HyperlinkInfo[] {
    const main = this.getPartDocument(this.mainPath);
    const body = bodyOf(main);
    const paragraphs = descendants(body, 'p');
    const relationships = this.relationshipsFor(this.mainPath);
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

  getRevisions(filter: { authors?: string[]; kinds?: RevisionInfo['kind'][] } = {}): RevisionInfo[] {
    const revisions = this.collectRevisions();
    if (!filter.authors && !filter.kinds) return revisions;
    const authors = filter.authors ? normalizeReviewerFilterAuthors(filter.authors) : undefined;
    const kinds = filter.kinds ? new Set(filter.kinds) : undefined;
    return revisions.filter((revision) =>
      (!authors || authors.has(reviewerBucketKey(reviewerBucketOf(revision.author)))) &&
      (!kinds || kinds.has(revision.kind)));
  }

  private collectRevisions(): RevisionInfo[] {
    const styles = this.getStylesContext();
    const stylesRevision = this.stylesCache?.revision ?? this.revision;
    if (this.revisionInfoCache?.revision === this.revision &&
        this.revisionInfoCache.mainPath === this.mainDocumentPath &&
        this.revisionInfoCache.stylesRevision === stylesRevision) {
      return this.revisionInfoCache.revisions;
    }
    const document = this.getCachedPartDocument(this.mainDocumentPath);
    if (!hasRevisionMarkup(document)) {
      this.revisionInfoCache = { revision: this.revision, mainPath: this.mainDocumentPath, stylesRevision, revisions: [] };
      return this.revisionInfoCache.revisions;
    }
    const body = bodyOf(document);
    const theme = styles.theme;
    const paragraphs = descendants(body, 'p');
    const paragraphIndex = new Map(paragraphs.map((paragraph, index) => [paragraph, index]));
    const runIndexByParagraph = new Map<Element, Map<Element, number>>();
    const runIndexOf = (paragraph: Element, run: Element): number | undefined => {
      let entries = runIndexByParagraph.get(paragraph);
      if (!entries) {
        entries = new Map(ownRuns(paragraph).map((item, index) => [item, index]));
        runIndexByParagraph.set(paragraph, entries);
      }
      return entries.get(run);
    };
    const result: RevisionInfo[] = [];
    const moveEntries: { element: Element; info: RevisionInfo }[] = [];
    const push = (element: Element, kind: RevisionInfo['kind']): void => {
      const paragraph = revisionParagraphAnchor(element);
      if (!paragraph) return;
      const paragraphNumber = paragraphIndex.get(paragraph);
      if (paragraphNumber === undefined) return;
      const mark = readRevisionMark(element, kind, theme);
      if (!mark) return;
      const run = revisionRunAnchor(paragraph, element);
      const deletedText = ['del', 'moveFrom'].includes(element.localName ?? '') ? deletedTextOf(element) || undefined : undefined;
      const runNumber = run ? runIndexOf(paragraph, run) : undefined;
      const info: RevisionInfo = {
        id: mark.id,
        kind: mark.kind,
        paragraph: paragraphNumber,
        ...(mark.author !== undefined ? { author: mark.author } : {}),
        ...(mark.date !== undefined ? { date: mark.date } : {}),
        ...(mark.move !== undefined ? { move: { ...mark.move } } : {}),
        ...(runNumber !== undefined ? { run: runNumber } : {}),
        ...(deletedText !== undefined ? { deletedText } : {}),
        ...(mark.previousFormat !== undefined ? { previousFormat: mark.previousFormat } : {}),
      };
      result.push(info);
      if (info.kind === 'move') moveEntries.push({ element, info });
    };
    const walk = (node: Node): void => {
      for (let child = node.firstChild; child; child = child.nextSibling) {
        if (child.nodeType !== 1) continue;
        const element = child as Element;
        if (element.namespaceURI !== WORD_NS) continue;
        switch (element.localName) {
          case 'ins':
          case 'cellIns':
            push(element, 'insertion');
            break;
          case 'del':
          case 'cellDel':
            push(element, 'deletion');
            break;
          case 'moveTo':
          case 'moveFrom':
            push(element, 'move');
            break;
          case 'rPrChange':
            push(element, 'runFormatChange');
            break;
          case 'pPrChange':
            push(element, 'paragraphFormatChange');
            break;
          case 'tblPrChange':
            push(element, 'tableFormatChange');
            break;
          case 'trPrChange':
            push(element, 'rowFormatChange');
            break;
          case 'tcPrChange':
            push(element, 'cellFormatChange');
            break;
        }
        walk(element);
      }
    };
    walk(body);
    const movePairing = pairMoveRevisionMarkers(moveEntries.map((entry) => entry.element));
    const moveInfoByElement = new Map(moveEntries.map((entry) => [entry.element, entry.info] as const));
    for (const { element, info } of moveEntries) {
      const paired = movePairing.pairs.get(element);
      if (!paired) continue;
      const pairedInfo = moveInfoByElement.get(paired);
      if (!pairedInfo || !info.move) continue;
      info.move.pairedId = pairedInfo.id;
    }
    this.revisionInfoCache = { revision: this.revision, mainPath: this.mainDocumentPath, stylesRevision, revisions: result };
    return result;
  }

  acceptRevision(id: number): void {
    assertIndex(id);
    if (!this.getRevisions().some((revision) => revision.id === id)) throw new Error(`Revision ${id} does not exist.`);
    this.withDraft((draft) => draft.applyRevisionById(id, 'accept'));
  }

  rejectRevision(id: number): void {
    assertIndex(id);
    if (!this.getRevisions().some((revision) => revision.id === id)) throw new Error(`Revision ${id} does not exist.`);
    this.withDraft((draft) => draft.applyRevisionById(id, 'reject'));
  }

  acceptAllRevisions(filter: { authors?: string[] } = {}): void {
    const authors = this.normalizeRevisionAuthorFilter(filter);
    const revisions = this.getRevisions();
    if (!revisions.length) return;
    if (authors && !revisions.some((revision) => authors.has(reviewerBucketKey(reviewerBucketOf(revision.author))))) return;
    this.withDraft((draft) => draft.applyRevisionBatch(authors, 'accept'));
  }

  rejectAllRevisions(filter: { authors?: string[] } = {}): void {
    const authors = this.normalizeRevisionAuthorFilter(filter);
    const revisions = this.getRevisions();
    if (!revisions.length) return;
    if (authors && !revisions.some((revision) => authors.has(reviewerBucketKey(reviewerBucketOf(revision.author))))) return;
    this.withDraft((draft) => draft.applyRevisionBatch(authors, 'reject'));
  }

  private normalizeRevisionAuthorFilter(filter: { authors?: string[] }): Set<string> | undefined {
    if (!filter || typeof filter !== 'object' || Array.isArray(filter)) throw new Error('filter must be an object.');
    if (Object.keys(filter).some((key) => key !== 'authors')) throw new Error('Unknown revision filter property.');
    if (!('authors' in filter) || filter.authors === undefined) return undefined;
    if (!Array.isArray(filter.authors) || filter.authors.length > REVISION_FILTER_MAX_AUTHORS) {
      throw new Error(`filter.authors must be an array of at most ${REVISION_FILTER_MAX_AUTHORS} authors.`);
    }
    const authors = new Set<string>();
    for (let index = 0; index < filter.authors.length; index++) {
      const author = filter.authors[index];
      assertText(author, `filter.authors[${index}]`);
      authors.add(reviewerFilterKeyOf(author));
    }
    return authors;
  }

  private applyRevisionById(id: number, action: 'accept' | 'reject'): void {
    this.updatePartXmlInternal(this.mainPath, (document) => {
      const markers = this.collectRevisionMarkers(document, () => true);
      const marker = markers.find((entry) => revisionIdOf(entry) === id);
      if (!marker) throw new Error(`Revision ${id} does not exist.`);
      if (isMoveRevisionMarker(marker)) {
        const paired = pairMoveRevisionMarkers(markers).pairs.get(marker);
        if (paired?.parentNode) {
          this.applyMovePairDecision(marker, paired, action);
          return;
        }
      }
      this.applyRevisionDecision(marker, action);
    });
  }

  private applyRevisionBatch(authors: Set<string> | undefined, action: 'accept' | 'reject'): void {
    this.updatePartXmlInternal(this.mainPath, (document) => {
      const markers = this.collectRevisionMarkers(document, () => true);
      const movePairs = pairMoveRevisionMarkers(markers).pairs;
      if (!authors) {
        const processed = new Set<Element>();
        for (const marker of [...markers].reverse()) {
          if (processed.has(marker) || !marker.parentNode) continue;
          if (isMoveRevisionMarker(marker)) {
            const paired = movePairs.get(marker);
            if (paired && paired.parentNode && !processed.has(paired)) {
              this.applyMovePairDecision(marker, paired, action);
              processed.add(marker);
              processed.add(paired);
              continue;
            }
          }
          if (!marker.parentNode) continue;
          this.applyRevisionDecision(marker, action);
          processed.add(marker);
        }
        return;
      }
      const selected = new Set<Element>();
      const matchesAuthor = (marker: Element): boolean => {
        const author = revisionAuthorOf(marker);
        const bucket = reviewerBucketKey(reviewerBucketOf(author));
        return authors.has(bucket);
      };
      for (const marker of markers) {
        if (!isMoveRevisionMarker(marker)) {
          if (matchesAuthor(marker)) selected.add(marker);
          continue;
        }
        const paired = movePairs.get(marker);
        if (paired) {
          if (matchesAuthor(marker) || matchesAuthor(paired)) {
            selected.add(marker);
            selected.add(paired);
          }
          continue;
        }
        if (matchesAuthor(marker)) selected.add(marker);
      }
      const processed = new Set<Element>();
      for (const marker of [...markers].reverse()) {
        if (!selected.has(marker) || processed.has(marker) || !marker.parentNode) continue;
        if (isMoveRevisionMarker(marker)) {
          const paired = movePairs.get(marker);
          if (paired && selected.has(paired) && paired.parentNode && !processed.has(paired)) {
            this.applyMovePairDecision(marker, paired, action);
            processed.add(marker);
            processed.add(paired);
            continue;
          }
        }
        if (!marker.parentNode) continue;
        this.applyRevisionDecision(marker, action);
        processed.add(marker);
      }
    });
  }

  private processNextRevision(
    document: Document,
    match: (marker: Element, id: number) => boolean,
    action: 'accept' | 'reject',
  ): boolean {
    const marker = this.findFirstRevisionMarker(document, match);
    if (!marker) return false;
    this.applyRevisionDecision(marker, action);
    return true;
  }

  private findFirstRevisionMarker(document: Document, match: (marker: Element, id: number) => boolean): Element | undefined {
    const body = bodyOf(document);
    const stack: Element[] = [body];
    while (stack.length) {
      const node = stack.pop()!;
      const descendantsInOrder = [...children(node)];
      for (let index = descendantsInOrder.length - 1; index >= 0; index--) stack.push(descendantsInOrder[index]!);
      if (node.namespaceURI !== WORD_NS) continue;
      const name = node.localName ?? '';
      if (!REVISION_ELEMENT_NAMES.has(name)) continue;
      const id = revisionIdOf(node);
      if (id === undefined) continue;
      if (match(node, id)) return node;
    }
    return undefined;
  }

  private collectRevisionMarkers(document: Document, match: (marker: Element, id: number) => boolean): Element[] {
    const body = bodyOf(document);
    const result: Element[] = [];
    const walk = (node: Element): void => {
      for (const child of children(node)) {
        if (child.namespaceURI !== WORD_NS) {
          walk(child);
          continue;
        }
        const name = child.localName ?? '';
        if (REVISION_ELEMENT_NAMES.has(name)) {
          const id = revisionIdOf(child);
          if (id !== undefined && match(child, id)) result.push(child);
        }
        walk(child);
      }
    };
    walk(body);
    return result;
  }

  private applyRevisionDecision(marker: Element, action: 'accept' | 'reject'): void {
    const name = marker.localName ?? '';
    if (['rPrChange', 'pPrChange', 'tblPrChange', 'trPrChange', 'tcPrChange'].includes(name)) {
      this.applyFormatChangeRevision(marker, action);
      return;
    }
    if (['ins', 'moveTo', 'del', 'moveFrom'].includes(name)) {
      this.applyInsertionDeletionRevision(marker, action);
      return;
    }
    if (['cellIns', 'cellDel'].includes(name)) {
      const deletesCell = (action === 'accept' && name === 'cellDel') || (action === 'reject' && name === 'cellIns');
      if (deletesCell) {
        this.deleteCellForRevision(marker);
      } else {
        this.removeRevisionMarker(marker);
      }
    }
  }

  private applyMovePairDecision(marker: Element, paired: Element, action: 'accept' | 'reject'): void {
    const from = moveRevisionSideOf(marker) === 'from' ? marker : paired;
    const to = from === marker ? paired : marker;
    const order = action === 'accept' ? [from, to] : [to, from];
    for (const item of order) {
      if (!item.parentNode) continue;
      this.applyInsertionDeletionRevision(item, action);
    }
  }

  private applyFormatChangeRevision(marker: Element, action: 'accept' | 'reject'): void {
    const props = marker.parentNode?.nodeType === 1 ? marker.parentNode as Element : undefined;
    if (!props || props.namespaceURI !== WORD_NS) return;
    if (action === 'reject') {
      const snapshotNameByMarker: Record<string, 'rPr' | 'pPr' | 'tblPr' | 'trPr' | 'tcPr'> = {
        rPrChange: 'rPr',
        pPrChange: 'pPr',
        tblPrChange: 'tblPr',
        trPrChange: 'trPr',
        tcPrChange: 'tcPr',
      };
      const snapshotName = snapshotNameByMarker[marker.localName as keyof typeof snapshotNameByMarker];
      const snapshot = snapshotName ? children(marker, snapshotName)[0] : undefined;
      if (snapshot) {
        for (const child of [...children(props)]) props.removeChild(child);
        for (const child of children(snapshot)) insertPropertyChild(props, child.cloneNode(true) as Element);
      }
    }
    this.removeRevisionMarker(marker);
  }

  private applyInsertionDeletionRevision(marker: Element, action: 'accept' | 'reject'): void {
    const name = marker.localName ?? '';
    const isInsertion = name === 'ins' || name === 'moveTo';
    const parent = marker.parentNode?.nodeType === 1 ? marker.parentNode as Element : undefined;
    const grandparent = parent?.parentNode?.nodeType === 1 ? parent.parentNode as Element : undefined;
    const isParagraphMark =
      parent?.namespaceURI === WORD_NS && parent.localName === 'rPr' &&
      grandparent?.namespaceURI === WORD_NS && grandparent.localName === 'pPr' &&
      nearestWordAncestor(grandparent, 'p');
    const isRowMark = parent?.namespaceURI === WORD_NS && parent.localName === 'trPr' && nearestWordAncestor(parent, 'tr');
    if (isParagraphMark) {
      if ((action === 'accept' && isInsertion) || (action === 'reject' && !isInsertion)) {
        this.removeRevisionMarker(marker);
      } else {
        this.deleteParagraphForRevision(marker);
      }
      return;
    }
    if (isRowMark) {
      if ((action === 'accept' && isInsertion) || (action === 'reject' && !isInsertion)) {
        this.removeRevisionMarker(marker);
      } else {
        this.deleteTableRowForRevision(marker);
      }
      return;
    }
    if ((action === 'accept' && isInsertion) || (action === 'reject' && !isInsertion)) {
      if (!isInsertion) replaceRevisionTextNodes(marker);
      unwrapNode(marker);
      return;
    }
    marker.parentNode?.removeChild(marker);
  }

  private deleteParagraphForRevision(marker: Element): void {
    const paragraph = nearestWordAncestor(marker, 'p');
    if (!paragraph) return;
    const isDeletionMark = marker.localName === 'del' || marker.localName === 'moveFrom';
    if (isDeletionMark && this.paragraphHasUndeletedContent(paragraph)) {
      if (this.tryMergeParagraphAfterMarkDeletion(paragraph, marker)) return;
      this.removeRevisionMarker(marker);
      return;
    }
    this.removeRevisionMarker(marker);
    deleteParagraphElement(paragraph);
  }

  private paragraphHasUndeletedContent(paragraph: Element): boolean {
    const walk = (node: Node, deletedDepth: number): boolean => {
      for (let child = node.firstChild; child; child = child.nextSibling) {
        if (child.nodeType !== 1) continue;
        const element = child as Element;
        if (element.namespaceURI === WORD_NS) {
          const name = element.localName ?? '';
          if (name === 'pPr' || isParagraphAnchorMarker(element)) continue;
          if (name === 'del' || name === 'moveFrom') {
            if (walk(element, deletedDepth + 1)) return true;
            continue;
          }
          if (deletedDepth === 0) return true;
        }
        if (walk(element, deletedDepth)) return true;
      }
      return false;
    };
    return walk(paragraph, 0);
  }

  private tryMergeParagraphAfterMarkDeletion(paragraph: Element, marker: Element): boolean {
    if (children(paragraph, 'pPr').some((props) => children(props, 'sectPr').length)) return false;
    let container: Element | undefined;
    try { container = paragraphContainer(paragraph); } catch { return false; }
    const blocks = blockElements(container);
    const index = blocks.findIndex((block) => block === paragraph);
    const next = index >= 0 ? blocks[index + 1] : undefined;
    if (!next || next.localName !== 'p') return false;
    if (children(next, 'pPr').some((props) => children(props, 'sectPr').length)) return false;
    const insertionPoint = children(next).find((child) => child.localName !== 'pPr') ?? null;
    for (const child of [...children(paragraph)]) {
      if (child.localName === 'pPr') continue;
      next.insertBefore(child, insertionPoint);
    }
    paragraph.parentNode?.removeChild(paragraph);
    removeEmptyParagraphWrappers(next);
    return true;
  }

  private deleteTableRowForRevision(marker: Element): void {
    const row = nearestWordAncestor(marker, 'tr');
    const table = nearestWordAncestor(marker, 'tbl');
    if (!row || !table) return;
    const index = tableRows(table).findIndex((item) => item === row);
    if (index < 0) return;
    if (tableRows(table).length <= 1) {
      table.parentNode?.removeChild(table);
      return;
    }
    deleteTableRowElement(table, index);
  }

  private deleteCellForRevision(marker: Element): void {
    const cell = nearestWordAncestor(marker, 'tc');
    const row = nearestWordAncestor(marker, 'tr');
    if (!cell || !row) return;
    const positions = rowCells(row);
    if (positions.length <= 1) {
      this.removeRevisionMarker(marker);
      clearCellContent(cell);
      return;
    }
    row.removeChild(cell);
  }

  private removeRevisionMarker(marker: Element): void {
    const parent = marker.parentNode?.nodeType === 1 ? marker.parentNode as Element : undefined;
    marker.parentNode?.removeChild(marker);
    removeIfEmpty(parent);
    const grandparent = parent?.parentNode?.nodeType === 1 ? parent.parentNode as Element : undefined;
    if (grandparent && ['pPr', 'rPr', 'trPr', 'tcPr', 'tblPr'].includes(grandparent.localName ?? '')) removeIfEmpty(grandparent);
  }

  getBookmarks(options: { includeInternal?: boolean } = {}): BookmarkInfo[] {
    const body = bodyOf(this.getPartDocument(this.mainPath));
    const paragraphs = descendants(body, 'p');
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

  private hyperlinkNode(hyperlink: HyperlinkInfo, document: Document): Element {
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

  private resolveHyperlink(reference: HyperlinkInfo | number | { paragraph: number; runs: number[]; text: string }): HyperlinkInfo {
    const hyperlinks = this.getHyperlinks();
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

  private splitRunAtOffset(paragraph: Element, offset: number): void {
    const paragraphText = textOf(paragraph);
    const normalizedOffset = normalizeTextBoundary(paragraphText, offset, 'start');
    const runs = ownRuns(paragraph);
    let cursor = 0;
    for (const run of runs) {
      const value = textOf(run);
      const next = cursor + value.length;
      if (normalizedOffset <= cursor || normalizedOffset >= next || !value.length) {
        cursor = next;
        continue;
      }
      const cut = normalizeTextBoundary(value, normalizedOffset - cursor, 'start');
      if (cut <= 0 || cut >= value.length) return;
      const left = value.slice(0, cut);
      const right = value.slice(cut);
      const rightRun = run.cloneNode(true) as Element;
      const leftElements = textElements(run);
      for (const element of leftElements) (element.parentNode as Element).removeChild(element);
      appendText(run, left);
      const rightElements = textElements(rightRun);
      for (const element of rightElements) (element.parentNode as Element).removeChild(element);
      appendText(rightRun, right);
      run.parentNode!.insertBefore(rightRun, run.nextSibling);
      return;
    }
  }

  private validateRangeFields(fields: (keyof RunFormat)[] | undefined): (keyof RunFormat)[] {
    if (fields === undefined) return [...RUN_FORMAT_FIELDS];
    if (!Array.isArray(fields)) throw new Error('fields must be an array.');
    const unique: (keyof RunFormat)[] = [];
    for (const field of fields) {
      if (!RUN_FORMAT_FIELDS.includes(field)) throw new Error(`Unsupported run format field: ${String(field)}`);
      if (!unique.includes(field)) unique.push(field);
    }
    return unique;
  }

  private normalizeTextRange(paragraphElement: Element, range: TextRange): { start: number; end: number } {
    assertIndex(range.paragraph);
    assertIndex(range.start);
    assertIndex(range.end);
    textRangeLength(paragraphElement, range.start, range.end);
    const text = textOf(paragraphElement);
    if (range.start === range.end) {
      const offset = normalizeTextBoundary(text, range.start, 'start');
      return { start: offset, end: offset };
    }
    let start = normalizeTextBoundary(text, range.start, 'start');
    let end = normalizeTextBoundary(text, range.end, 'end');
    if (end < start) end = start;
    return { start, end };
  }

  private normalizeDocumentRange(document: Document, range: DocumentRange): DocumentRange {
    const startParagraph = paragraphAt(document, range.start.paragraph);
    const endParagraph = paragraphAt(document, range.end.paragraph);
    if (range.start.paragraph > range.end.paragraph ||
        (range.start.paragraph === range.end.paragraph && range.start.offset > range.end.offset)) {
      throw new Error('range.start must not be after range.end.');
    }
    if (range.start.paragraph === range.end.paragraph) {
      const normalized = this.normalizeTextRange(startParagraph, {
        paragraph: range.start.paragraph,
        start: range.start.offset,
        end: range.end.offset,
      });
      return {
        start: { paragraph: range.start.paragraph, offset: normalized.start },
        end: { paragraph: range.end.paragraph, offset: normalized.end },
      };
    }
    const startContainer = paragraphContainer(startParagraph);
    const endContainer = paragraphContainer(endParagraph);
    if (startContainer !== endContainer) {
      throw new Error('Cross-container document ranges are not supported.');
    }
    const startNormalized = this.normalizeTextRange(startParagraph, {
      paragraph: range.start.paragraph,
      start: range.start.offset,
      end: range.start.offset,
    });
    const endNormalized = this.normalizeTextRange(endParagraph, {
      paragraph: range.end.paragraph,
      start: range.end.offset,
      end: range.end.offset,
    });
    return {
      start: { paragraph: range.start.paragraph, offset: startNormalized.start },
      end: { paragraph: range.end.paragraph, offset: endNormalized.end },
    };
  }

  private runsInRange(paragraph: Element, start: number, end: number, collapsed = false): Element[] {
    const runs = ownRuns(paragraph);
    let cursor = 0;
    const selected: Element[] = [];
    for (const run of runs) {
      const length = textOf(run).length;
      const next = cursor + length;
      if (length > 0 && start < next && end > cursor) selected.push(run);
      cursor = next;
    }
    if (selected.length || !collapsed) return selected;
    cursor = 0;
    let previous: Element | undefined;
    for (const run of runs) {
      const length = textOf(run).length;
      const next = cursor + length;
      if (!length) continue;
      if (start < next) return [run];
      if (start === next) {
        previous = run;
        cursor = next;
        continue;
      }
      previous = run;
      cursor = next;
    }
    return previous ? [previous] : [];
  }

  private normalizeRangeOn(document: Document, range: TextRange): { paragraph: Element; start: number; end: number } {
    const paragraph = paragraphAt(document, range.paragraph);
    const normalized = this.normalizeTextRange(paragraph, range);
    return { paragraph, start: normalized.start, end: normalized.end };
  }

  private trackedRunFormatSnapshot(run: Element): RunFormat | undefined {
    const format = compactDefined(readRunProperties(children(run, 'rPr')[0], this.getStylesContext().theme));
    return Object.keys(format).length ? format : undefined;
  }

  private trackedParagraphFormatSnapshot(paragraph: Element): ParagraphFormat | undefined {
    const format = compactDefined(readParagraphProperties(children(paragraph, 'pPr')[0]));
    return Object.keys(format).length ? format : undefined;
  }

  private trackRunFormatChange(run: Element, author?: string, date?: string): void {
    const previous = this.trackedRunFormatSnapshot(run) ?? {};
    const props = properties(run, 'rPr');
    markFormatRevision(
      props,
      'rPrChange',
      'rPr',
      this.trackedRevisionAuthor(author),
      date,
    );
    const marker = children(props, 'rPrChange')[0];
    let snapshot = marker ? children(marker, 'rPr')[0] : undefined;
    if (marker && !snapshot) {
      snapshot = wordElement(props.ownerDocument!, 'rPr');
      marker.appendChild(snapshot);
    }
    if (!marker || !snapshot) return;
    if (snapshot.parentNode !== marker) marker.appendChild(snapshot);
    if (!snapshot.firstChild && !snapshot.attributes.length && Object.keys(previous).length) applyRunFormatTo(snapshot, previous);
  }

  private trackParagraphFormatChange(paragraph: Element, author?: string, date?: string): void {
    const previous = this.trackedParagraphFormatSnapshot(paragraph) ?? {};
    const props = properties(paragraph, 'pPr');
    markFormatRevision(
      props,
      'pPrChange',
      'pPr',
      this.trackedRevisionAuthor(author),
      date,
    );
    const marker = children(props, 'pPrChange')[0];
    let snapshot = marker ? children(marker, 'pPr')[0] : undefined;
    if (marker && !snapshot) {
      snapshot = wordElement(props.ownerDocument!, 'pPr');
      marker.appendChild(snapshot);
    }
    if (!marker || !snapshot) return;
    if (snapshot.parentNode !== marker) marker.appendChild(snapshot);
    if (!snapshot.firstChild && !snapshot.attributes.length && Object.keys(previous).length) applyParagraphFormatTo(snapshot, previous);
  }

  private insertTrackedText(paragraph: Element, start: number, text: string, author?: string, date?: string): void {
    if (!text) return;
    this.splitRunAtOffset(paragraph, start);
    const runs = ownRuns(paragraph);
    let cursor = 0;
    let parent = paragraph;
    let reference: Node | null = null;
    let previous: Element | undefined;
    for (const candidate of runs) {
      const length = textOf(candidate).length;
      const next = cursor + length;
      if (length === 0) {
        previous = candidate;
        continue;
      }
      if (start <= cursor || start < next) {
        parent = candidate.parentNode?.nodeType === 1 ? candidate.parentNode as Element : paragraph;
        reference = candidate;
        break;
      }
      cursor = next;
      previous = candidate;
    }
    if (!reference && previous) {
      let node: Node = previous;
      let container: Node | null = node.parentNode;
      while (container?.nodeType === 1 && container !== paragraph && !node.nextSibling) {
        node = container;
        container = node.parentNode;
      }
      if (container?.nodeType === 1) {
        parent = container as Element;
        reference = node.nextSibling;
      }
    }
    const run = wordElement(parent.ownerDocument!, 'r');
    appendText(run, text);
    const wrapper = createRevisionWrapper(parent.ownerDocument!, 'ins', this.trackedRevisionAuthor(author), date);
    wrapper.appendChild(run);
    parent.insertBefore(wrapper, reference);
  }

  private replaceSpanTracked(paragraph: Element, start: number, end: number, replacement: string, author?: string, date?: string): void {
    if (start !== end) {
      this.splitRunAtOffset(paragraph, end);
      this.splitRunAtOffset(paragraph, start);
      const runs = this.runsInRange(paragraph, start, end);
      wrapRunsWithRevision(paragraph, runs, 'del', this.trackedRevisionAuthor(author), date);
    }
    this.insertTrackedText(paragraph, start, replacement, author, date);
  }

  private applyComparedBlocks(
    steps: CompareStep[],
    revisedBlocks: CompareBlockInfo[],
    revisedStyles: StylesContext,
    author?: string,
    date?: string,
  ): void {
    const document = this.getCachedPartDocument(this.mainPath);
    const body = bodyOf(document);
    for (let blockIndex = 0, stepIndex = 0; stepIndex < steps.length; stepIndex++) {
      const step = steps[stepIndex]!;
      const current = blockElements(body)[blockIndex];
      switch (step.kind) {
        case 'equal':
          blockIndex++;
          break;
        case 'modify':
          if (!current || current.localName !== 'p') throw new Error('Compared paragraph does not exist.');
          this.compareParagraphAgainst(current, revisedBlocks[step.revisedIndex]!.element, revisedStyles, author, date);
          blockIndex++;
          break;
        case 'delete':
          if (!current) throw new Error('Compared block does not exist.');
          this.markComparedBlockDeleted(current, author, date);
          blockIndex++;
          break;
        case 'insert':
          this.insertComparedBlock(body, blockIndex, revisedBlocks[step.revisedIndex]!.element, author, date);
          blockIndex++;
          break;
      }
    }
    this.dirtyPartXml.add(this.mainPath);
    this.dirtyPartSizes.delete(this.mainPath);
  }

  private insertComparedBlock(body: Element, blockIndex: number, source: Element, author?: string, date?: string): void {
    const clone = cloneNodeIntoDocument(body.ownerDocument!, source) as Element;
    body.insertBefore(clone, blockInsertionReference(body, blockIndex));
    if (clone.localName === 'p') {
      const runs = ownRuns(clone);
      if (runs.length) wrapRunsWithRevision(clone, runs, 'ins', this.trackedRevisionAuthor(author), date);
      markRevision(property(properties(clone, 'pPr'), 'rPr'), 'ins', this.trackedRevisionAuthor(author), date, 'paraRPr');
      return;
    }
    if (clone.localName === 'tbl') {
      for (const row of tableRows(clone)) markRevision(tableProperty(row, 'trPr'), 'ins', this.trackedRevisionAuthor(author), date);
    }
  }

  private markComparedBlockDeleted(block: Element, author?: string, date?: string): void {
    if (block.localName === 'p') {
      if (children(block, 'pPr').some(props => children(props, 'sectPr').length)) return;
      const runs = ownRuns(block);
      if (runs.length) wrapRunsWithRevision(block, runs, 'del', this.trackedRevisionAuthor(author), date);
      else {
        for (const child of [...children(block)]) {
          if (child.localName === 'pPr' || preservesParagraphMarkers(child)) continue;
          block.removeChild(child);
        }
      }
      markRevision(property(properties(block, 'pPr'), 'rPr'), 'del', this.trackedRevisionAuthor(author), date, 'paraRPr');
      return;
    }
    if (block.localName === 'tbl') {
      for (const row of tableRows(block)) markRevision(tableProperty(row, 'trPr'), 'del', this.trackedRevisionAuthor(author), date);
    }
  }

  private compareParagraphAgainst(
    paragraph: Element,
    revisedParagraph: Element,
    revisedStyles: StylesContext,
    author?: string,
    date?: string,
  ): void {
    const baseTheme = this.getStylesContext().theme;
    const baseParagraphFormat = directParagraphFormatOf(paragraph);
    const baseParagraphClone = paragraph.cloneNode(true) as Element;
    const old = textOf(paragraph);
    const revisedText = textOf(revisedParagraph);
    let start = 0;
    while (start < old.length && start < revisedText.length && old[start] === revisedText[start]) start++;
    let end = old.length;
    let replacementEnd = revisedText.length;
    while (end > start && replacementEnd > start && old[end - 1] === revisedText[replacementEnd - 1]) {
      end--;
      replacementEnd--;
    }
    if (start > 0 && /[\ud800-\udbff]/.test(old[start - 1]!)) start--;
    if (end < old.length && /[\udc00-\udfff]/.test(old[end]!)) { end++; replacementEnd++; }
    if (old !== revisedText) this.replaceSpanTracked(paragraph, start, end, revisedText.slice(start, replacementEnd), author, date);
    const paragraphPatch = buildFormatPatch(
      baseParagraphFormat,
      directParagraphFormatOf(revisedParagraph),
      PARAGRAPH_FORMAT_FIELDS,
    ) as ParagraphFormat;
    if (Object.keys(paragraphPatch).length) {
      this.trackParagraphFormatChange(paragraph, author, date);
      applyParagraphFormatTo(properties(paragraph, 'pPr'), paragraphPatch);
    }
    const revisedVisibleRuns = ownRuns(revisedParagraph)
      .map((run) => ({ run, text: textOf(run) }))
      .filter((entry) => entry.text.length > 0);
    let offset = 0;
    for (const entry of revisedVisibleRuns.slice(0, -1)) {
      offset += entry.text.length;
      this.splitRunAtOffset(paragraph, offset);
    }
    const currentVisibleRuns = ownRuns(paragraph)
      .map((run) => ({ run, text: textOf(run) }))
      .filter((entry) => entry.text.length > 0);
    if (currentVisibleRuns.length !== revisedVisibleRuns.length ||
        currentVisibleRuns.some((entry, index) => entry.text !== revisedVisibleRuns[index]!.text)) {
      return;
    }
    for (let index = 0, runOffset = 0; index < currentVisibleRuns.length; index++) {
      const currentRun = currentVisibleRuns[index]!.run;
      const revisedRun = revisedVisibleRuns[index]!.run;
      const runStart = runOffset;
      const runEnd = runStart + currentVisibleRuns[index]!.text.length;
      runOffset = runEnd;
      const originalFormat = this.originalComparedRunFormat(baseParagraphClone, baseTheme, runStart, runEnd, start, end, replacementEnd);
      const runPatch = buildFormatPatch(
        originalFormat ?? directRunFormatOf(currentRun, baseTheme),
        directRunFormatOf(revisedRun, revisedStyles.theme),
        RUN_FORMAT_FIELDS,
      ) as RunFormat;
      if (!Object.keys(runPatch).length) continue;
      if (!this.runInsideInsertion(currentRun, paragraph) && originalFormat) this.trackRunFormatChange(currentRun, author, date);
      applyRunFormatTo(properties(currentRun, 'rPr'), runPatch);
    }
  }

  private originalComparedRunFormat(
    baseParagraph: Element,
    theme: StylesContext['theme'],
    revisedStart: number,
    revisedEnd: number,
    replaceStart: number,
    baseReplaceEnd: number,
    revisedReplaceEnd: number,
  ): RunFormat | undefined {
    let baseStart = revisedStart;
    let baseEnd = revisedEnd;
    if (revisedEnd <= replaceStart) {
      baseStart = revisedStart;
      baseEnd = revisedEnd;
    } else if (revisedStart >= revisedReplaceEnd) {
      baseStart = revisedStart - revisedReplaceEnd + baseReplaceEnd;
      baseEnd = revisedEnd - revisedReplaceEnd + baseReplaceEnd;
    } else {
      return undefined;
    }
    this.splitRunAtOffset(baseParagraph, baseEnd);
    this.splitRunAtOffset(baseParagraph, baseStart);
    const originalRuns = this.runsInRange(baseParagraph, baseStart, baseEnd);
    if (originalRuns.length !== 1) return undefined;
    return directRunFormatOf(originalRuns[0]!, theme);
  }

  private runInsideInsertion(run: Element, paragraph: Element): boolean {
    for (let parent = run.parentNode; parent && parent !== paragraph; parent = parent.parentNode) {
      if (parent.nodeType !== 1) continue;
      const element = parent as Element;
      if (element.namespaceURI === WORD_NS && ['ins', 'moveTo'].includes(element.localName ?? '')) return true;
    }
    return false;
  }

  private applyRunFormatRangeOnParagraph(paragraph: Element, start: number, end: number, format: RunFormat): void {
    if (start === end) return;
    this.splitRunAtOffset(paragraph, end);
    this.splitRunAtOffset(paragraph, start);
    for (const run of this.runsInRange(paragraph, start, end)) {
      applyRunFormatTo(properties(run, 'rPr'), format);
    }
  }

  private rangeRunFormats(document: Document, range: TextRange, theme: StylesContext['theme']): RunFormat[] {
    const normalized = this.normalizeRangeOn(document, range);
    const runs = this.runsInRange(normalized.paragraph, normalized.start, normalized.end, normalized.start === normalized.end);
    return runs.map((run) => readRunProperties(children(run, 'rPr')[0], theme));
  }

  private mergeRangeFormats(formats: RunFormat[]): RunFormat {
    const merged: RunFormat = {};
    if (!formats.length) return merged;
    for (const field of RUN_FORMAT_FIELDS) {
      const first = formats[0]?.[field];
      let same = true;
      for (let i = 1; i < formats.length; i++) {
        if (!runFormatValueEqual(first, formats[i]?.[field])) {
          same = false;
          break;
        }
      }
      if (same && first !== undefined) (merged as Record<string, unknown>)[field] = cloneRunFormatValue(first);
    }
    return merged;
  }

  private directRunFormat(run: RunInfo): RunFormat {
    const format: RunFormat = {};
    for (const field of RUN_FORMAT_FIELDS) {
      const value = run[field];
      if (value !== undefined) (format as Record<string, unknown>)[field] = cloneRunFormatValue(value);
    }
    return format;
  }

  private directParagraphFormat(paragraph: ParagraphInfo): ParagraphFormat {
    const format: ParagraphFormat = {};
    for (const key of PARAGRAPH_DIRECT_FIELDS) {
      const value = paragraph[key];
      if (value !== undefined) (format as Record<string, unknown>)[key] = cloneRunFormatValue(value);
    }
    return format;
  }

  private normalizeClipboardParagraph(input: ClipboardParagraph, counters: { paragraphs: number; runs: number; images: number }): ClipboardParagraph {
    if (!input || typeof input !== 'object' || !Array.isArray(input.runs)) throw new Error('clipboard paragraph must include runs.');
    counters.paragraphs++;
    if (counters.paragraphs > CLIPBOARD_MAX_PARAGRAPHS) throw new Error(`clipboard paragraph count exceeds ${CLIPBOARD_MAX_PARAGRAPHS}.`);
    if (input.format !== undefined) validateParagraphFormat(input.format);
    if (input.numbering !== undefined) {
      const numbering = input.numbering;
      if (!numbering || typeof numbering !== 'object' || !['bullet', 'decimal'].includes(numbering.kind)) {
        throw new Error('clipboard paragraph numbering.kind must be bullet or decimal.');
      }
      if (numbering.level !== undefined && (!Number.isSafeInteger(numbering.level) || numbering.level < 0 || numbering.level > 8)) {
        throw new Error('clipboard paragraph numbering.level must be an integer between 0 and 8.');
      }
      if (numbering.listId !== undefined && !Number.isSafeInteger(numbering.listId)) {
        throw new Error('clipboard paragraph numbering.listId must be a safe integer.');
      }
    }
    const runs: ClipboardRun[] = [];
    for (const run of input.runs) {
      if (!run || typeof run !== 'object') throw new Error('clipboard run must be an object.');
      counters.runs++;
      if (counters.runs > CLIPBOARD_MAX_RUNS) throw new Error(`clipboard run count exceeds ${CLIPBOARD_MAX_RUNS}.`);
      const normalized: ClipboardRun = {};
      if (run.text !== undefined) {
        const sanitized = sanitizeText(run.text);
        assertText(sanitized, 'clipboard run.text');
        if (sanitized.length > CLIPBOARD_MAX_RUN_TEXT_LENGTH) {
          throw new Error(`clipboard run.text exceeds ${CLIPBOARD_MAX_RUN_TEXT_LENGTH} characters.`);
        }
        if (sanitized) normalized.text = sanitized;
      }
      if (run.format !== undefined) {
        validateRunFormat(run.format);
        normalized.format = cloneRunFormatValue(run.format);
      }
      if (run.hyperlink !== undefined) {
        assertHyperlinkInput(run.hyperlink);
        normalized.hyperlink = { ...run.hyperlink };
      }
      if (run.images !== undefined) {
        if (!Array.isArray(run.images)) throw new Error('clipboard run.images must be an array.');
        normalized.images = [];
        for (const image of run.images) {
          if (!image || typeof image !== 'object') throw new Error('clipboard image must be an object.');
          if (typeof image.bytes !== 'string' || typeof image.contentType !== 'string') {
            throw new Error('clipboard image requires bytes and contentType.');
          }
          assertText(image.contentType, 'clipboard image.contentType');
          const bytes = decodeBase64(image.bytes);
          if (bytes.byteLength > IMAGE_LIMIT) throw new Error(`clipboard image exceeds ${IMAGE_LIMIT} bytes.`);
          counters.images++;
          if (counters.images > CLIPBOARD_MAX_IMAGES) throw new Error(`clipboard image count exceeds ${CLIPBOARD_MAX_IMAGES}.`);
          normalized.images.push({
            bytes: image.bytes,
            contentType: image.contentType,
            ...(image.widthEmu !== undefined ? { widthEmu: image.widthEmu } : {}),
            ...(image.heightEmu !== undefined ? { heightEmu: image.heightEmu } : {}),
            ...(image.alt !== undefined ? { alt: sanitizeText(image.alt) } : {}),
            ...(image.placement !== undefined ? { placement: image.placement } : {}),
          });
        }
      }
      if (normalized.text || normalized.images?.length) runs.push(normalized);
    }
    return {
      runs,
      ...(input.format !== undefined ? { format: cloneRunFormatValue(input.format) } : {}),
      ...(input.numbering !== undefined ? { numbering: { ...input.numbering } } : {}),
    };
  }

  private normalizeClipboardBlocks(fragment: ClipboardFragment): ClipboardBlock[] {
    const counters = { paragraphs: 0, runs: 0, images: 0 };
    const fromParagraphs = (): ClipboardBlock[] => {
      if (!Array.isArray(fragment.paragraphs)) throw new Error('clipboard paragraphs must be an array.');
      return fragment.paragraphs
        .map((paragraph) => this.normalizeClipboardParagraph(paragraph, counters))
        .filter((paragraph) => paragraph.runs.length)
        .map((paragraph) => ({ type: 'paragraph', paragraph } satisfies ClipboardBlock));
    };
    if (!fragment.blocks?.length) return fromParagraphs();
    const blocks: ClipboardBlock[] = [];
    for (const block of fragment.blocks) {
      if (!block || typeof block !== 'object') throw new Error('clipboard block must be an object.');
      if (block.type === 'paragraph') {
        const paragraph = this.normalizeClipboardParagraph(block.paragraph, counters);
        if (paragraph.runs.length) blocks.push({ type: 'paragraph', paragraph });
        continue;
      }
      if (block.type === 'table') {
        const rows = block.table?.rows;
        if (!Array.isArray(rows) || !rows.length) throw new Error('clipboard table.rows must be a non-empty array.');
        const normalizedRows = rows.map((row) => {
          if (!Array.isArray(row) || !row.length) throw new Error('clipboard table row must be a non-empty array.');
          return row.map((cell) => this.normalizeClipboardParagraph(cell, counters));
        });
        blocks.push({ type: 'table', table: { rows: normalizedRows } });
        continue;
      }
      throw new Error(`Unsupported clipboard block type: ${(block as { type?: unknown }).type as string}`);
    }
    return blocks;
  }

  copyClipboardFragment(range: DocumentRange): ClipboardFragment {
    const preview = this.getCachedPartDocument(this.mainPath);
    const normalized = this.normalizeDocumentRange(preview, range);
    const paragraphs = this.getParagraphs();
    const byIndex = new Map(paragraphs.map((paragraph) => [paragraph.index, paragraph]));
    const textChunks: string[] = [];
    const result: ClipboardParagraph[] = [];
    for (let index = normalized.start.paragraph; index <= normalized.end.paragraph; index++) {
      const paragraph = byIndex.get(index);
      if (!paragraph) continue;
      const start = index === normalized.start.paragraph ? normalized.start.offset : 0;
      const end = index === normalized.end.paragraph ? normalized.end.offset : paragraph.text.length;
      if (start >= end) {
        if (index < normalized.end.paragraph) textChunks.push('');
        continue;
      }
      const runs: ClipboardRun[] = [];
      let cursor = 0;
      for (const run of paragraph.runs) {
        const runStart = cursor;
        const runEnd = cursor + run.text.length;
        const overlapStart = Math.max(start, runStart);
        const overlapEnd = Math.min(end, runEnd);
        cursor = runEnd;
        if (overlapStart >= overlapEnd && !(run.images?.length && start <= runStart && end >= runStart)) continue;
        const text = overlapStart < overlapEnd ? run.text.slice(overlapStart - runStart, overlapEnd - runStart) : '';
        const images: ClipboardImage[] = (run.images ?? []).flatMap((image) => {
          if (image.isExternal || !image.partPath) return [];
          try {
            return [{
              bytes: encodeBase64(this.getImageBytes(image)),
              contentType: image.contentType ?? this.getContentType(image.partPath) ?? 'image/png',
              widthEmu: image.widthEmu,
              heightEmu: image.heightEmu,
              alt: image.alt,
              placement: image.placement,
            }];
          } catch {
            return [];
          }
        });
        if (!text && !images.length) continue;
        runs.push({
          ...(text ? { text } : {}),
          format: this.directRunFormat(run),
          ...(run.hyperlink ? { hyperlink: { ...run.hyperlink } } : {}),
          ...(images.length ? { images } : {}),
        });
      }
      const paragraphText = paragraph.text.slice(start, end);
      textChunks.push(paragraphText);
      if (runs.length) {
        result.push({
          runs,
          format: this.directParagraphFormat(paragraph),
          ...(paragraph.numbering ? {
            numbering: {
              kind: paragraph.numbering.isBullet ? 'bullet' : 'decimal',
              level: paragraph.numbering.level,
              listId: paragraph.numbering.numId,
            },
          } : {}),
        });
      }
    }
    return {
      version: 1,
      text: textChunks.join('\n'),
      paragraphs: result,
      blocks: result.map((paragraph) => ({ type: 'paragraph', paragraph })),
    };
  }

  pasteClipboardFragment(range: DocumentRange, fragment: ClipboardFragment): boolean {
    if (!fragment || fragment.version !== 1 || typeof fragment.text !== 'string' || !Array.isArray(fragment.paragraphs)) {
      throw new Error('Invalid clipboard fragment.');
    }
    const blocks = this.normalizeClipboardBlocks(fragment);
    const preview = this.getCachedPartDocument(this.mainPath);
    const normalized = this.normalizeDocumentRange(preview, range);
    if (normalized.start.paragraph !== normalized.end.paragraph) return false;
    if (!blocks.length && !fragment.text && normalized.start.offset === normalized.end.offset) return false;
    this.withDraft((draft) => {
      const listNumbering = new Map<string, number>();
      for (const block of blocks) {
        if (block.type !== 'paragraph' || !block.paragraph.numbering) continue;
        const numbering = block.paragraph.numbering;
        const key = `${numbering.kind}:${numbering.listId ?? 0}`;
        if (!listNumbering.has(key)) listNumbering.set(key, draft.createNumbering(numbering.kind));
      }
      draft.materializeAllParts();
      let next = new Map(draft.parts);
      const main = parseXml(decodeXml(next.get(draft.mainPath)!));
      const paragraph = paragraphAt(main, normalized.start.paragraph);
      const sourcePartPath = draft.mainPath;
      draft.splitRunAtOffset(paragraph, normalized.end.offset);
      draft.splitRunAtOffset(paragraph, normalized.start.offset);
      const selected = draft.runsInRange(paragraph, normalized.start.offset, normalized.end.offset);
      let marker: Node;
      if (selected.length) {
        const firstChild = paragraphDirectChild(paragraph, selected[0]!);
        marker = main.createTextNode('');
        paragraph.insertBefore(marker, firstChild);
        for (const run of selected) {
          const parent = run.parentNode as Element | null;
          if (!parent) continue;
          parent.removeChild(run);
          if (parent !== paragraph && !children(parent, 'r').length && !children(parent, 'proofErr').length) parent.parentNode?.removeChild(parent);
        }
      } else {
        const runs = ownRuns(paragraph);
        let cursor = 0;
        let before: Node | null = null;
        for (const run of runs) {
          const nextCursor = cursor + textOf(run).length;
          if (normalized.start.offset <= cursor || normalized.start.offset < nextCursor) {
            before = paragraphDirectChild(paragraph, run);
            break;
          }
          cursor = nextCursor;
        }
        marker = main.createTextNode('');
        paragraph.insertBefore(marker, before);
      }
      const relPath = resolveRelationshipsPath(sourcePartPath);
      let rels: Document | undefined;
      const hyperlinkRelationByUrl = new Map<string, string>();
      const ensureRels = (): Document => {
        if (rels) return rels;
        const bytes = next.get(relPath);
        rels = bytes ? parseXml(decodeXml(bytes)) : parseXml(`<Relationships xmlns="${REL_NS}"/>`);
        return rels;
      };
      const suffix: Node[] = [];
      for (let current = marker.nextSibling; current; ) {
        const nextNode = current.nextSibling;
        suffix.push(current);
        paragraph.removeChild(current);
        current = nextNode;
      }
      marker.parentNode?.removeChild(marker);
      const trackChanges = draft.trackChangesEnabled();
      const trackedAuthor = trackChanges ? draft.trackedRevisionAuthor() : undefined;
      const buildRunNode = (runFragment: ClipboardRun): { node: Node; trackedRuns: Element[] } | null => {
        const run = wordElement(main, 'r');
        if (runFragment.format) applyRunFormatTo(properties(run, 'rPr'), runFragment.format);
        if (runFragment.text) appendText(run, runFragment.text);
        for (const image of runFragment.images ?? []) {
          const bytes = decodeBase64(image.bytes);
          const size = draft.inferImageSize(bytes, image.contentType, image.widthEmu, image.heightEmu);
          const partPath = draft.nextImagePartPath(image.contentType);
          next = draft.ensureMediaContentType(partPath, image.contentType, next);
          next.set(partPath, Uint8Array.from(bytes));
          const relsDocument = ensureRels();
          const relationshipId = draft.nextRelationshipId(relsDocument);
          const relationship = relsDocument.createElementNS(REL_NS, 'Relationship');
          relationship.setAttribute('Id', relationshipId);
          relationship.setAttribute('Type', IMAGE_REL);
          relationship.setAttribute('Target', relativeTargetPath(sourcePartPath, partPath));
          relsDocument.documentElement!.appendChild(relationship);
          run.appendChild(createDrawingElement(main, relationshipId, size, {
            alt: image.alt,
            title: image.alt,
            placement: image.placement ?? 'inline',
            docPrId: draft.nextDocPrId(main),
          }));
        }
        if (!run.firstChild) return null;
        const link = runFragment.hyperlink;
        if (link && (link.url || link.anchor)) {
          const hyperlink = wordElement(main, 'hyperlink');
          if (link.url) {
            const key = link.url;
            let relationshipId = hyperlinkRelationByUrl.get(key);
            if (!relationshipId) {
              const relsDocument = ensureRels();
              relationshipId = draft.nextRelationshipId(relsDocument);
              const relationship = relsDocument.createElementNS(REL_NS, 'Relationship');
              relationship.setAttribute('Id', relationshipId);
              relationship.setAttribute('Type', HYPERLINK_REL);
              relationship.setAttribute('Target', link.url);
              relationship.setAttribute('TargetMode', 'External');
              relsDocument.documentElement!.appendChild(relationship);
              hyperlinkRelationByUrl.set(key, relationshipId);
            }
            hyperlink.setAttributeNS(OFFICE_REL_NS, 'r:id', relationshipId);
          }
          if (link.anchor) hyperlink.setAttributeNS(WORD_NS, 'w:anchor', link.anchor);
          if (link.tooltip) hyperlink.setAttributeNS(WORD_NS, 'w:tooltip', link.tooltip);
          hyperlink.appendChild(run);
          return { node: hyperlink, trackedRuns: [run] };
        }
        return { node: run, trackedRuns: [run] };
      };
      const createParagraph = (): Element => wordElement(main, 'p');
      const applyParagraphMeta = (targetParagraph: Element, paragraphFragment: ClipboardParagraph): void => {
        if (paragraphFragment.format) applyParagraphFormatTo(properties(targetParagraph, 'pPr'), paragraphFragment.format);
        if (paragraphFragment.numbering) {
          const key = `${paragraphFragment.numbering.kind}:${paragraphFragment.numbering.listId ?? 0}`;
          const numId = listNumbering.get(key);
          if (numId !== undefined) {
            const numPr = property(properties(targetParagraph, 'pPr'), 'numPr');
            setWordValue(numberingProperty(numPr, 'numId'), String(numId));
            setWordValue(numberingProperty(numPr, 'ilvl'), String(paragraphFragment.numbering.level ?? 0));
          }
        }
      };
      const blockParent = paragraph.parentNode as Element;
      const blockAnchor = paragraph.nextSibling;
      const stagedBlocks = main.createDocumentFragment();
      let paragraphTarget: Element | null = null;
      let firstParagraphPlaced = false;
      for (const block of blocks) {
        if (block.type === 'paragraph') {
          const targetParagraph = firstParagraphPlaced ? createParagraph() : paragraph;
          if (firstParagraphPlaced) {
            stagedBlocks.appendChild(targetParagraph);
          } else {
            firstParagraphPlaced = true;
          }
          applyParagraphMeta(targetParagraph, block.paragraph);
          const trackedRuns: Element[] = [];
          for (const runFragment of block.paragraph.runs) {
            const built = buildRunNode(runFragment);
            if (!built) continue;
            targetParagraph.appendChild(built.node);
            trackedRuns.push(...built.trackedRuns);
          }
          if (trackChanges && trackedRuns.length) wrapRunsWithRevision(targetParagraph, trackedRuns, 'ins', trackedAuthor);
          paragraphTarget = targetParagraph;
          continue;
        }
        const rows = block.table.rows.map((row) => row.map((cell) =>
          cell.runs.map((run) => run.text ?? '').join('')));
        const table = buildTable(main, rows.length, Math.max(...rows.map((row) => row.length)), undefined, rows);
        stagedBlocks.appendChild(table);
        paragraphTarget = null;
      }
      if (stagedBlocks.firstChild) blockParent.insertBefore(stagedBlocks, blockAnchor);
      const suffixTarget = paragraphTarget ?? ((blocks.length || fragment.text) ? (() => {
        const tail = createParagraph();
        blockParent.insertBefore(tail, blockAnchor);
        return tail;
      })() : paragraph);
      for (const node of suffix) suffixTarget.appendChild(node);
      if (!blocks.length && fragment.text) {
        const fallback = sanitizeText(fragment.text);
        if (fallback) {
          const run = wordElement(main, 'r');
          appendText(run, fallback);
          suffixTarget.appendChild(run);
        }
      }
      next.set(draft.mainPath, encodeXml(serializeXml(main)));
      if (rels) next.set(relPath, encodeXml(serializeXml(rels)));
      draft.commitParts(next);
    });
    return true;
  }

  insertHyperlink(
    target: { paragraph: number; start: number; end: number },
    link: { url?: string; anchor?: string; tooltip?: string },
  ): HyperlinkInfo {
    assertIndex(target.paragraph);
    assertIndex(target.start);
    assertIndex(target.end);
    assertHyperlinkInput(link);
    return this.withDraft(draft => {
      let createdId: string | undefined;
      let createdMarker: { paragraph: number; run: number; text: string } | undefined;
      let relationshipPath: string | undefined;
      let relationships: Document | undefined;
      if (link.url) {
        relationshipPath = resolveRelationshipsPath(draft.mainPath);
        relationships = draft.hasPart(relationshipPath)
          ? draft.getPartDocument(relationshipPath)
          : parseXml(`<Relationships xmlns="${REL_NS}"/>`);
        createdId = draft.nextRelationshipId(relationships);
        const relationship = relationships.createElementNS(REL_NS, 'Relationship');
        relationship.setAttribute('Id', createdId);
        relationship.setAttribute('Type', HYPERLINK_REL);
        relationship.setAttribute('Target', link.url);
        relationship.setAttribute('TargetMode', 'External');
        relationships.documentElement!.appendChild(relationship);
      }
      draft.updatePartXml(draft.mainPath, document => {
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
          draft.setPartXml(relationshipPath, serializeXml(relationships));
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

  updateHyperlink(
    hyperlink: HyperlinkInfo | number | { paragraph: number; runs: number[]; text: string },
    link: { url?: string; anchor?: string; tooltip?: string },
  ): void {
    assertHyperlinkInput(link);
    const current = this.resolveHyperlink(hyperlink);
    const currentNodeName = this.hyperlinkNode(current, this.getPartDocument(this.mainPath)).localName;
    this.withDraft((draft) => {
      let nextRelationshipId = current.relationshipId;
      const currentRelationshipUsers = current.relationshipId
        ? draft.getHyperlinks().filter(item => item.relationshipId === current.relationshipId).length
        : 0;
      if (currentNodeName !== 'fldSimple' && link.url && link.url !== current.url) {
        const relPath = resolveRelationshipsPath(draft.mainPath);
        const rels = draft.hasPart(relPath) ? draft.getPartDocument(relPath) : parseXml(`<Relationships xmlns="${REL_NS}"/>`);
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
          draft.setPartXml(relPath, serializeXml(rels));
        } else {
          draft.addPart(relPath, encodeXml(serializeXml(rels)), 'application/vnd.openxmlformats-package.relationships+xml');
        }
      }
      draft.updatePartXml(draft.mainPath, document => {
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
            draft.updatePartXml(relPath, rels => {
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

  removeHyperlink(
    hyperlink: HyperlinkInfo | number | { paragraph: number; runs: number[]; text: string },
    options: { keepText?: boolean } = {},
  ): void {
    const link = this.resolveHyperlink(hyperlink);
    this.withDraft((draft) => {
      draft.updatePartXml(draft.mainPath, document => {
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
            draft.updatePartXml(relPath, rels => {
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

  insertBookmark(name: string, range: { startParagraph: number; endParagraph?: number }): BookmarkInfo {
    assertText(name, 'name');
    assertIndex(range.startParagraph);
    if (range.endParagraph !== undefined) assertIndex(range.endParagraph);
    if (range.endParagraph !== undefined && range.endParagraph < range.startParagraph) {
      throw new Error('range.endParagraph must be >= range.startParagraph.');
    }
    if (!name) throw new Error('name must not be empty.');
    if (this.getBookmarks({ includeInternal: true }).some(bookmark => bookmark.name === name)) {
      throw new Error(`Bookmark "${name}" already exists.`);
    }
    let bookmark: BookmarkInfo | undefined;
    this.updatePartXml(this.mainPath, document => {
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

  deleteBookmark(name: string): void {
    assertText(name, 'name');
    this.updatePartXml(this.mainPath, document => {
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

  setParagraphNumbering(index: number, numId: number, level = 0): void {
    assertIndex(numId);
    assertIndex(level);
    if (numId < 1) throw new Error('numId must be at least 1. Use clearParagraphNumbering() to remove numbering.');
    if (level > 8) throw new Error('level must be between 0 and 8.');
    this.updatePartXmlInternal(this.mainPath, document => {
      const props = properties(paragraphAt(document, index), 'pPr');
      const numPr = property(props, 'numPr');
      setWordValue(numberingProperty(numPr, 'ilvl'), String(level));
      setWordValue(numberingProperty(numPr, 'numId'), String(numId));
    });
  }

  clearParagraphNumbering(index: number): void {
    const paragraph = this.getParagraphs()[index];
    if (!paragraph) throw new Error(`Paragraph ${index} does not exist.`);
    const styleHasNumbering = paragraph.style ? this.getNumberingContext().model.paragraphStyles.has(paragraph.style) : false;
    this.updatePartXmlInternal(this.mainPath, document => {
      const props = properties(paragraphAt(document, index), 'pPr');
      const existing = children(props, 'numPr')[0];
      if (existing) props.removeChild(existing);
      if (styleHasNumbering) {
        const numPr = property(props, 'numPr');
        setWordValue(numberingProperty(numPr, 'numId'), '0');
        const ilvl = children(numPr, 'ilvl')[0];
        if (ilvl) numPr.removeChild(ilvl);
      }
    });
  }

  setParagraphLevel(index: number, delta: number): void {
    if (!Number.isSafeInteger(delta)) throw new Error('delta must be a safe integer.');
    const paragraph = this.getParagraphs()[index];
    if (!paragraph) throw new Error(`Paragraph ${index} does not exist.`);
    if (!paragraph.numbering) throw new Error('Paragraph does not have numbering.');
    this.setParagraphNumbering(index, paragraph.numbering.numId, Math.max(0, Math.min(8, paragraph.numbering.level + delta)));
  }

  createNumbering(kind: 'bullet' | 'decimal' | 'multilevel' | NumberingDefinition): number {
    const definition = typeof kind === 'string' ? defaultNumberingDefinition(kind) : kind;
    const levels = [...(definition.levels.length ? definition.levels : defaultNumberingDefinition('decimal').levels)];
    const seenLevels = new Set<number>();
    for (const level of levels) {
      if (!Number.isSafeInteger(level.level) || level.level < 0 || level.level > 8) {
        throw new Error('Numbering levels must be integers between 0 and 8.');
      }
      if (seenLevels.has(level.level)) throw new Error('Numbering definition contains duplicate levels.');
      seenLevels.add(level.level);
    }
    this.materializeAllParts();
    const next = new Map(this.parts);
    const types = this.getCachedPartDocument('[Content_Types].xml');
    const relationsPath = relsPath(this.mainPath);
    const rels = next.has(relationsPath)
      ? this.getCachedPartDocument(relationsPath)
      : parseXml(`<Relationships xmlns="${REL_NS}"/>`);
    let numberingPath = this.getNumberingPath();
    if (!numberingPath) {
      numberingPath = `${dirname(this.mainPath) ? `${dirname(this.mainPath)}/` : ''}numbering.xml`;
      const relationship = rels.createElementNS(REL_NS, 'Relationship');
      relationship.setAttribute('Id', nextRelationshipId(rels.documentElement!));
      relationship.setAttribute('Type', NUMBERING_REL);
      relationship.setAttribute('Target', relativeTarget(this.mainPath, numberingPath));
      rels.documentElement!.appendChild(relationship);
    }
    const numberingDocument = next.has(numberingPath)
      ? this.getCachedPartDocument(numberingPath)
      : parseXml(`<w:numbering xmlns:w="${WORD_NS}"/>`);
    const numberingRoot = numberingDocument.documentElement!;
    const existingAbstractIds = children(numberingRoot, 'abstractNum')
      .map(element => Number(element.getAttributeNS(WORD_NS, 'abstractNumId')))
      .filter(Number.isFinite);
    const existingNumIds = children(numberingRoot, 'num')
      .map(element => Number(element.getAttributeNS(WORD_NS, 'numId')))
      .filter(Number.isFinite);
    const abstractNumId = (existingAbstractIds.length ? Math.max(...existingAbstractIds) : 0) + 1;
    const numId = (existingNumIds.length ? Math.max(...existingNumIds) : 0) + 1;
    const abstract = wordElement(numberingDocument, 'abstractNum');
    setWordAttr(abstract, 'abstractNumId', abstractNumId);
    if (definition.nsid) appendWordValueElement(abstract, 'nsid', definition.nsid);
    if (definition.multiLevelType) appendWordValueElement(abstract, 'multiLevelType', definition.multiLevelType);
    if (definition.tmpl) appendWordValueElement(abstract, 'tmpl', definition.tmpl);
    if (definition.styleLink) appendWordValueElement(abstract, 'styleLink', definition.styleLink);
    if (definition.numStyleLink) appendWordValueElement(abstract, 'numStyleLink', definition.numStyleLink);
    for (const level of levels.sort((a, b) => a.level - b.level)) abstract.appendChild(buildLevelElement(numberingDocument, level));
    insertNumberingNode(numberingRoot, abstract);
    const num = wordElement(numberingDocument, 'num');
    setWordAttr(num, 'numId', numId);
    appendWordValueElement(num, 'abstractNumId', abstractNumId);
    insertNumberingNode(numberingRoot, num);
    const typesRoot = types.documentElement!;
    if (!children(typesRoot, 'Override', CONTENT_TYPES_NS).some(override => override.getAttribute('PartName') === `/${numberingPath}`)) {
      const override = types.createElementNS(CONTENT_TYPES_NS, 'Override');
      override.setAttribute('PartName', `/${numberingPath}`);
      override.setAttribute('ContentType', NUMBERING_TYPE);
      typesRoot.appendChild(override);
    }
    next.set('[Content_Types].xml', encodeXml(serializeXml(types)));
    next.set(relationsPath, encodeXml(serializeXml(rels)));
    next.set(numberingPath, encodeXml(serializeXml(numberingDocument)));
    this.commitParts(next);
    return numId;
  }

  getImages(): ImageInfo[] {
    return this.getParagraphs().flatMap((paragraph) => paragraph.images);
  }

  private resolveImage(image: ImageInfo | string): ImageInfo {
    const matches = this.getImages().filter((item) => typeof image === 'string'
      ? item.id === image || item.relationshipId === image
      : item.relationshipId === image.relationshipId && item.paragraph === image.paragraph &&
        item.run === image.run && item.ordinal === image.ordinal);
    if (!matches.length) throw new Error(`Image ${image} does not exist.`);
    if (typeof image === 'string' && matches.every((item) => item.id !== image) && matches.length > 1) {
      throw new Error(`Image relationshipId ${image} is ambiguous; pass ImageInfo or image.id instead.`);
    }
    if (matches.length > 1) throw new Error(`Image relationshipId ${image} is ambiguous; pass ImageInfo or image.id instead.`);
    return matches[0]!;
  }

  private inferImageSize(bytes: Uint8Array, contentType?: string, widthEmu?: number, heightEmu?: number): { widthEmu: number; heightEmu: number } {
    const detected = detectImageSize(bytes, contentType);
    if (widthEmu && heightEmu) return { widthEmu, heightEmu };
    if (detected) {
      const aspect = detected.width / detected.height;
      if (widthEmu) return { widthEmu, heightEmu: Math.max(1, Math.round(widthEmu / aspect)) };
      if (heightEmu) return { widthEmu: Math.max(1, Math.round(heightEmu * aspect)), heightEmu };
      const width = Math.max(1, Math.round(pxToEmu(detected.width)));
      const height = Math.max(1, Math.round(pxToEmu(detected.height)));
      return { widthEmu: width, heightEmu: height };
    }
    const fallbackWidth = widthEmu ?? 4 * 914400;
    const fallbackHeight = heightEmu ?? fallbackWidth;
    return { widthEmu: fallbackWidth, heightEmu: fallbackHeight };
  }

  private nextImagePartPath(contentType: string): string {
    const extension = extensionForContentType(contentType);
    if (!extension) throw new Error(`Unsupported image content type: ${contentType}`);
    for (let index = 1; index < 10_000; index++) {
      const path = `word/media/image${index}.${extension}`;
      if (!this.hasPart(path)) return path;
    }
    throw new Error('Unable to allocate a unique media part path.');
  }

  private ensureMediaContentType(path: string, contentType: string, base = this.parts): Map<string, Uint8Array> {
    const typesBytes = base.get('[Content_Types].xml');
    if (!typesBytes) throw new Error('Missing [Content_Types].xml.');
    const types = parseXml(decodeXml(typesBytes));
    const defaults = contentTypesDefaults(types.documentElement!);
    const overrides = contentTypesOverrides(types.documentElement!);
    const extension = path.split('.').pop()?.toLowerCase() ?? '';
    const defaultEntry = defaults.find((node) => node.getAttribute('Extension')?.toLowerCase() === extension);
    const overrideEntry = overrides.find((node) => node.getAttribute('PartName') === `/${path}`);
    if (!defaultEntry && !overrideEntry) {
      const addDefault = ['png', 'jpg', 'jpeg', 'gif', 'bmp', 'tiff', 'emf', 'wmf', 'svg'].includes(extension);
      const element = types.createElementNS(CONTENT_TYPES_NS, addDefault ? 'Default' : 'Override');
      if (addDefault) {
        element.setAttribute('Extension', extension);
        element.setAttribute('ContentType', contentType);
      } else {
        element.setAttribute('PartName', `/${path}`);
        element.setAttribute('ContentType', contentType);
      }
      types.documentElement!.appendChild(element);
    } else if (overrideEntry) {
      overrideEntry.setAttribute('ContentType', contentType);
    } else if (defaultEntry?.getAttribute('ContentType') !== contentType) {
      const override = types.createElementNS(CONTENT_TYPES_NS, 'Override');
      override.setAttribute('PartName', `/${path}`);
      override.setAttribute('ContentType', contentType);
      types.documentElement!.appendChild(override);
    }
    const next = new Map(base);
    next.set('[Content_Types].xml', encodeXml(serializeXml(types)));
    return next;
  }

  private ensureRelationshipsDocument(partPath: string, next: Map<string, Uint8Array>): Document {
    const relPath = resolveRelationshipsPath(partPath);
    const existing = next.get(relPath);
    if (existing) return parseXml(decodeXml(existing));
    const document = parseXml(`<Relationships xmlns="${REL_NS}"/>`);
    next.set(relPath, encodeXml(serializeXml(document)));
    return document;
  }

  private nextRelationshipId(rels: Document): string {
    const used = new Set(children(rels.documentElement!, 'Relationship', REL_NS).map((rel) => rel.getAttribute('Id')));
    for (let index = 1; index < 10_000; index++) {
      const id = `rId${index}`;
      if (!used.has(id)) return id;
    }
    throw new Error('Unable to allocate a unique relationship id.');
  }

  private nextDocPrId(document: Document): number {
    const used = new Set(Array.from(document.getElementsByTagNameNS(WP_NS, 'docPr'))
      .map((element) => Number(element.getAttribute('id')))
      .filter((value) => Number.isFinite(value) && value >= 0));
    for (let index = 1; index < 1_000_000; index++) {
      if (!used.has(index)) return index;
    }
    throw new Error('Unable to allocate a unique wp:docPr id.');
  }

  getImageBytes(image: ImageInfo | string): Uint8Array {
    const info = typeof image === 'string' ? this.resolveImage(image) : image;
    if (info.isExternal) throw new Error('External images are not loaded.');
    if (!info.partPath || !this.hasPart(info.partPath)) throw new Error(`Image part not found for ${info.relationshipId}.`);
    return this.getPartBytes(info.partPath);
  }

  getImageDataUrl(image: ImageInfo | string): string {
    const info = typeof image === 'string' ? this.resolveImage(image) : image;
    if (info.isExternal) return placeholderDataUrl('外部图片未加载', info.widthPx || 160, info.heightPx || 90);
    if (!info.partPath || !this.hasPart(info.partPath)) return placeholderDataUrl(info.name ?? '图片缺失', info.widthPx || 160, info.heightPx || 90);
    const contentType = info.contentType ?? this.getContentType(info.partPath);
    if (!isBrowserRenderableContentType(contentType)) {
      return placeholderDataUrl(info.name ?? info.partPath.split('/').pop() ?? '不支持的图片', info.widthPx || 160, info.heightPx || 90);
    }
    const cached = this.imageDataUrls.get(info.partPath);
    if (cached && cached.revision === this.revision && cached.contentType === contentType) return cached.url;
    const url = dataUrlForBytes(this.getPartBytes(info.partPath), contentType!);
    this.imageDataUrls.set(info.partPath, { revision: this.revision, contentType: contentType!, url });
    return url;
  }

  insertImage(options: {
    bytes: Uint8Array;
    contentType: string;
    paragraph?: number;
    run?: number;
    widthEmu?: number;
    heightEmu?: number;
    alt?: string;
    placement?: 'inline' | 'floating';
  }): ImageInfo {
    if (!(options.bytes instanceof Uint8Array) || options.bytes.byteLength > IMAGE_LIMIT) {
      throw new Error('Image bytes must be a Uint8Array no larger than 16 MiB.');
    }
    assertText(options.contentType, 'contentType');
    if (options.paragraph !== undefined) assertIndex(options.paragraph);
    if (options.run !== undefined) assertIndex(options.run);
    if (options.alt !== undefined) assertText(options.alt, 'alt');
    const size = this.inferImageSize(options.bytes, options.contentType, options.widthEmu, options.heightEmu);
    const partPath = this.nextImagePartPath(options.contentType);
    const main = this.getPartDocument(this.mainPath);
    const paragraphs = descendants(bodyOf(main), 'p');
    const paragraph = options.paragraph !== undefined
      ? paragraphAt(main, options.paragraph)
      : paragraphs.at(-1) ?? paragraphAt(main, 0);
    const paragraphIndex = paragraphs.indexOf(paragraph);
    const targetRuns = ownRuns(paragraph);
    const beforeRun = options.run !== undefined ? targetRuns[options.run] : undefined;
    if (options.run !== undefined && !beforeRun) throw new Error(`Run ${options.run} does not exist.`);
    const insertBefore = beforeRun ? paragraphDirectChild(paragraph, beforeRun) : null;
    const relPath = resolveRelationshipsPath(this.mainPath);
    const next = this.ensureMediaContentType(partPath, options.contentType);
    next.set(partPath, Uint8Array.from(options.bytes));
    const rels = this.hasPart(relPath) ? this.getPartDocument(relPath) : parseXml(`<Relationships xmlns="${REL_NS}"/>`);
    const relationshipId = this.nextRelationshipId(rels);
    const relationship = rels.createElementNS(REL_NS, 'Relationship');
    relationship.setAttribute('Id', relationshipId);
    relationship.setAttribute('Type', IMAGE_REL);
    relationship.setAttribute('Target', relativeTargetPath(this.mainPath, partPath));
    rels.documentElement!.appendChild(relationship);
    next.set(relPath, encodeXml(serializeXml(rels)));
    const run = wordElement(main, 'r');
    run.appendChild(createDrawingElement(main, relationshipId, size, {
      alt: options.alt,
      title: options.alt,
      placement: options.placement ?? 'inline',
      docPrId: this.nextDocPrId(main),
    }));
    if (this.trackChangesEnabled()) {
      const wrapper = createRevisionWrapper(paragraph.ownerDocument!, 'ins', this.trackedRevisionAuthor());
      wrapper.appendChild(run);
      paragraph.insertBefore(wrapper, insertBefore);
    } else if (insertBefore) paragraph.insertBefore(run, insertBefore);
    else paragraph.appendChild(run);
    const insertedRunIndex = ownRuns(paragraph).indexOf(run);
    next.set(this.mainPath, encodeXml(serializeXml(main)));
    this.commitParts(next);
    return {
      id: `${this.mainPath}:${paragraphIndex}:${insertedRunIndex}:0:${relationshipId}`,
      paragraph: paragraphIndex,
      run: insertedRunIndex,
      ordinal: 0,
      sourcePartPath: this.mainPath,
      relationshipId,
      partPath,
      contentType: options.contentType,
      widthEmu: size.widthEmu,
      heightEmu: size.heightEmu,
      widthPx: emuToPx(size.widthEmu),
      heightPx: emuToPx(size.heightEmu),
      alt: options.alt,
      title: options.alt,
      placement: options.placement ?? 'inline',
      isExternal: false,
    };
  }

  replaceImageBytes(image: ImageInfo | string, bytes: Uint8Array, contentType?: string): void {
    if (!(bytes instanceof Uint8Array) || bytes.byteLength > IMAGE_LIMIT) {
      throw new Error('Image bytes must be a Uint8Array no larger than 16 MiB.');
    }
    const info = this.resolveImage(image);
    if (info.isExternal || !info.partPath) throw new Error('External images cannot be replaced.');
    const inferredType = contentType ?? detectImageContentType(bytes) ?? info.contentType;
    if (!inferredType) throw new Error('contentType is required when the image format cannot be inferred from bytes.');
    this.materializeAllParts();
    let next = new Map(this.parts);
    let path = info.partPath;
    if (inferredType) {
      assertText(inferredType, 'contentType');
      const currentExtension = path.split('.').pop()?.toLowerCase();
      const nextExtension = extensionForContentType(inferredType);
      if (!nextExtension) throw new Error(`Unsupported image content type: ${inferredType}`);
      if (currentExtension !== nextExtension && !(currentExtension === 'jpg' && nextExtension === 'jpeg') &&
          !(currentExtension === 'jpeg' && nextExtension === 'jpg')) {
        const movedPath = this.nextImagePartPath(inferredType);
        for (const relPath of this.listParts().filter((entry) => entry.endsWith('.rels'))) {
          const sourcePart = sourcePartFromRelationshipsPath(relPath);
          if (!sourcePart) continue;
          const rels = this.getCachedPartDocument(relPath);
          let changed = false;
          for (const rel of children(rels.documentElement!, 'Relationship', REL_NS)) {
            if (rel.getAttribute('Type') !== IMAGE_REL || rel.getAttribute('TargetMode') === 'External') continue;
            const target = rel.getAttribute('Target');
            if (target && resolveTargetPath(sourcePart, decodeURIComponent(target)) === info.partPath) {
              rel.setAttribute('Target', relativeTargetPath(sourcePart, movedPath));
              changed = true;
            }
          }
          if (changed) next.set(relPath, encodeXml(serializeXml(rels)));
        }
        next.delete(info.partPath);
        path = movedPath;
      }
      next = this.ensureMediaContentType(path, inferredType, next);
    }
    next.set(path, Uint8Array.from(bytes));
    this.commitParts(next);
  }

  resizeImage(image: ImageInfo | string, size: { widthEmu?: number; heightEmu?: number; keepAspect?: boolean }): void {
    const info = this.resolveImage(image);
    let nextWidth = size.widthEmu ?? info.widthEmu;
    let nextHeight = size.heightEmu ?? info.heightEmu;
    if (size.keepAspect && info.widthEmu > 0 && info.heightEmu > 0) {
      if (size.widthEmu !== undefined) {
        nextWidth = size.widthEmu;
        nextHeight = Math.round(size.widthEmu * (info.heightEmu / info.widthEmu));
      } else if (size.heightEmu !== undefined) {
        nextHeight = size.heightEmu;
        nextWidth = Math.round(size.heightEmu * (info.widthEmu / info.heightEmu));
      }
    }
    this.updatePartXmlInternal(info.sourcePartPath ?? this.mainPath, (document) => {
      const run = ownRuns(paragraphAt(document, info.paragraph))[info.run];
      if (!run) throw new Error(`Run ${info.run} does not exist.`);
      const imageElement = imageElementForRun(run, info.relationshipId, info.ordinal ?? 0);
      if (!imageElement) throw new Error(`Image ${info.relationshipId} does not exist.`);
      if (imageElement.localName === 'pict') {
        const shape = Array.from(imageElement.getElementsByTagNameNS(V_NS, 'shape'))[0];
        if (!shape) throw new Error('VML image shape is missing.');
        shape.setAttribute('style', updateStyleLength(updateStyleLength(shape.getAttribute('style'), 'width', emuToPx(nextWidth) * 72 / 96), 'height', emuToPx(nextHeight) * 72 / 96));
        return;
      }
      for (const extent of Array.from(imageElement.getElementsByTagNameNS(WP_NS, 'extent'))) {
        extent.setAttribute('cx', String(Math.max(1, Math.round(nextWidth))));
        extent.setAttribute('cy', String(Math.max(1, Math.round(nextHeight))));
      }
      for (const ext of Array.from(imageElement.getElementsByTagNameNS(A_NS, 'ext'))) {
        ext.setAttribute('cx', String(Math.max(1, Math.round(nextWidth))));
        ext.setAttribute('cy', String(Math.max(1, Math.round(nextHeight))));
      }
    });
  }

  setImageAlt(image: ImageInfo | string, alt: string, title?: string): void {
    assertText(alt, 'alt');
    if (title !== undefined) assertText(title, 'title');
    const info = this.resolveImage(image);
    this.updatePartXmlInternal(info.sourcePartPath ?? this.mainPath, (document) => {
      const run = ownRuns(paragraphAt(document, info.paragraph))[info.run];
      if (!run) throw new Error(`Run ${info.run} does not exist.`);
      const imageElement = imageElementForRun(run, info.relationshipId, info.ordinal ?? 0);
      if (!imageElement) throw new Error(`Image ${info.relationshipId} does not exist.`);
      if (imageElement.localName === 'pict') {
        const shape = Array.from(imageElement.getElementsByTagNameNS(V_NS, 'shape'))[0];
        if (!shape) throw new Error('VML image shape is missing.');
        shape.setAttribute('alt', alt);
        setOptionalAttribute(shape, 'title', title);
        return;
      }
      for (const docPr of Array.from(imageElement.getElementsByTagNameNS(WP_NS, 'docPr'))) {
        docPr.setAttribute('descr', alt);
        setOptionalAttribute(docPr, 'title', title);
      }
      for (const cNvPr of Array.from(imageElement.getElementsByTagNameNS(PIC_NS, 'cNvPr'))) {
        cNvPr.setAttribute('descr', alt);
        setOptionalAttribute(cNvPr, 'title', title);
      }
    });
  }

  deleteImage(image: ImageInfo | string): void {
    const info = this.resolveImage(image);
    const sourcePart = info.sourcePartPath ?? this.mainPath;
    const main = this.getCachedPartDocument(sourcePart);
    const run = ownRuns(paragraphAt(main, info.paragraph))[info.run];
    if (!run) throw new Error(`Run ${info.run} does not exist.`);
    const imageElement = imageElementForRun(run, info.relationshipId, info.ordinal ?? 0);
    if (!imageElement) throw new Error(`Image ${info.relationshipId} does not exist.`);
    if (this.trackChangesEnabled()) {
      this.updatePartXmlInternal(sourcePart, (document) => {
        const paragraph = paragraphAt(document, info.paragraph);
        const trackedRun = ownRuns(paragraph)[info.run];
        if (!trackedRun) throw new Error(`Run ${info.run} does not exist.`);
        const trackedImage = imageElementForRun(trackedRun, info.relationshipId, info.ordinal ?? 0);
        if (!trackedImage) throw new Error(`Image ${info.relationshipId} does not exist.`);
        const isolatedRun = isolateRunChild(trackedRun, trackedImage);
        wrapRunsWithRevision(paragraph, [isolatedRun], 'del', this.trackedRevisionAuthor(), undefined, { convertText: false });
      });
      return;
    }
    run.removeChild(imageElement);
    this.materializeAllParts();
    const next = new Map(this.parts);
    if (isEmptyRun(run)) run.parentNode!.removeChild(run);
    next.set(sourcePart, encodeXml(serializeXml(main)));
    const relPath = resolveRelationshipsPath(sourcePart);
    if (this.hasPart(relPath)) {
      const rels = this.getCachedPartDocument(relPath);
      const relationship = children(rels.documentElement!, 'Relationship', REL_NS)
        .find((rel) => rel.getAttribute('Id') === info.relationshipId);
      if (relationship && !documentUsesRelationship(main, relationship.getAttribute('Id') ?? '')) {
        relationship.parentNode?.removeChild(relationship);
        next.set(relPath, encodeXml(serializeXml(rels)));
      }
    }
    if (info.partPath) {
      const stillReferenced = [...next.keys()]
        .filter((path) => path.endsWith('.rels'))
        .some((path) => {
          const sourcePart = sourcePartFromRelationshipsPath(path);
          if (!sourcePart) return false;
          const rels = parseXml(decodeXml(next.get(path)!));
          return children(rels.documentElement!, 'Relationship', REL_NS)
            .filter((rel) => rel.getAttribute('Type') === IMAGE_REL && rel.getAttribute('TargetMode') !== 'External')
            .some((rel) => resolveTargetPath(sourcePart, decodeURIComponent(rel.getAttribute('Target') ?? '')) === info.partPath);
        });
      if (!stillReferenced) next.delete(info.partPath);
    }
    this.commitParts(next);
  }

  private relationshipPart(path: string, create = false): Document {
    const part = relsPath(path);
    if (!this.parts.has(part)) {
      if (!create) throw new Error(`Package part not found: ${part}`);
      const xml = `<Relationships xmlns="${REL_NS}"></Relationships>`;
      this.addPart(part, encodeXml(xml), RELS_CONTENT_TYPE);
    }
    return this.getPartDocument(part);
  }

  private relationshipTargets(path: string): Map<string, { target: string; type: string }> {
    const root = this.parts.has(relsPath(path))
      ? this.relationshipPart(path).documentElement!
      : parseXml(`<Relationships xmlns="${REL_NS}"></Relationships>`).documentElement!;
    const entries: Array<[string, { target: string; type: string }]> = [];
    for (const relationship of children(root, 'Relationship', REL_NS)) {
      if (relationship.getAttribute('TargetMode') === 'External') continue;
      const id = relationship.getAttribute('Id') ?? '';
      if (!id) continue;
      const target = resolvePartPath(path, relationship.getAttribute('Target') ?? '');
      if (!target) continue;
      entries.push([id, {
        target,
        type: relationship.getAttribute('Type') ?? '',
      }]);
    }
    return new Map(entries);
  }

  private nextPartRelationshipId(path: string): string {
    const root = this.relationshipPart(path, true).documentElement!;
    const used = new Set(children(root, 'Relationship', REL_NS).map(node => node.getAttribute('Id') ?? ''));
    let index = 1;
    while (used.has(`rId${index}`)) index++;
    return `rId${index}`;
  }

  private addRelationship(path: string, type: string, target: string): string {
    const relsPathValue = relsPath(path);
    const rels = this.relationshipPart(path, true);
    const root = rels.documentElement!;
    const id = this.nextPartRelationshipId(path);
    const relation = rels.createElementNS(REL_NS, 'Relationship');
    relation.setAttribute('Id', id);
    relation.setAttribute('Type', type);
    relation.setAttribute('Target', target);
    root.appendChild(relation);
    this.setPartXml(relsPathValue, serializeXml(rels));
    return id;
  }

  getSections(): SectionInfo[] {
    const relationships = this.relationshipTargets(this.mainPath);
    const settings = [...relationships.values()].find(relationship => relationship.type === SETTINGS_REL)?.target;
    const evenAndOdd = settings && this.parts.has(settings)
      ? !!children(this.getPartDocument(settings).documentElement!, 'evenAndOddHeaders')[0]
      : false;
    const sections = readSections(this.getPartDocument(this.mainPath), id => relationships.get(id)?.target)
      .map(section => evenAndOdd ? section : { ...section, headers: { ...section.headers, even: undefined }, footers: { ...section.footers, even: undefined } });
    if (sections.length) return sections;
    const paragraphs = this.getParagraphs();
    return [{
      index: 0,
      startParagraph: 0,
      endParagraph: Math.max(paragraphs.length - 1, -1),
      isImplicit: true,
      type: 'nextPage',
      pageWidth: 11906,
      pageHeight: 16838,
      orientation: 'portrait',
      margins: { top: 0, right: 0, bottom: 0, left: 0, header: 0, footer: 0, gutter: 0 },
      columns: { count: 1, space: 720, equalWidth: true },
      titlePage: false,
      headers: {},
      footers: {},
    }];
  }

  getSection(index: number): SectionInfo {
    const section = this.getSections()[index];
    if (!section) throw new Error(`Section ${index} does not exist.`);
    return section;
  }

  setPageSetup(section: number, setup: Partial<PageSetup>): void {
    this.updatePartXml(this.mainPath, document => {
      let descriptors = collectSections(document);
      if (!descriptors.length && section === 0) {
        bodyOf(document).appendChild(wordElement(document, 'sectPr'));
        descriptors = collectSections(document);
      }
      const descriptor = descriptors[section];
      if (!descriptor) throw new Error(`Section ${section} does not exist.`);
      const sectPr = descriptor.sectPr;
      if (setup.type !== undefined) setWordValue(sectionProperty(sectPr, 'type'), setup.type);
      if (setup.pageWidth !== undefined || setup.pageHeight !== undefined || setup.orientation !== undefined) {
        const size = sectionProperty(sectPr, 'pgSz');
        const parsedWidth = Number(size.getAttributeNS(WORD_NS, 'w') ?? size.getAttribute('w:w'));
        const parsedHeight = Number(size.getAttributeNS(WORD_NS, 'h') ?? size.getAttribute('w:h'));
        let width = Number.isFinite(parsedWidth) ? Math.max(0, Math.trunc(parsedWidth)) : 11906;
        let height = Number.isFinite(parsedHeight) ? Math.max(0, Math.trunc(parsedHeight)) : 16838;
        if (setup.pageWidth !== undefined) width = Math.max(0, Math.trunc(setup.pageWidth));
        if (setup.pageHeight !== undefined) height = Math.max(0, Math.trunc(setup.pageHeight));
        if (setup.orientation !== undefined && setup.pageWidth === undefined && setup.pageHeight === undefined) {
          if (setup.orientation === 'landscape' && width < height) [width, height] = [height, width];
          if (setup.orientation === 'portrait' && width > height) [width, height] = [height, width];
        }
        size.setAttributeNS(WORD_NS, 'w:w', String(width));
        size.setAttributeNS(WORD_NS, 'w:h', String(height));
        if (setup.orientation !== undefined) size.setAttributeNS(WORD_NS, 'w:orient', setup.orientation);
      }
      if (setup.margins) {
        const margins = sectionProperty(sectPr, 'pgMar');
        for (const key of ['top', 'right', 'bottom', 'left', 'header', 'footer', 'gutter'] as const) {
          if (setup.margins[key] !== undefined) {
            margins.setAttributeNS(WORD_NS, `w:${key}`, String(Math.max(0, Math.trunc(setup.margins[key]!))));
          }
        }
      }
      if (setup.columns) {
        const columns = sectionProperty(sectPr, 'cols');
        if (setup.columns.count !== undefined) columns.setAttributeNS(WORD_NS, 'w:num', String(Math.max(1, Math.trunc(setup.columns.count))));
        if (setup.columns.space !== undefined) columns.setAttributeNS(WORD_NS, 'w:space', String(Math.max(0, Math.trunc(setup.columns.space))));
        if (setup.columns.equalWidth !== undefined) columns.setAttributeNS(WORD_NS, 'w:equalWidth', setup.columns.equalWidth ? '1' : '0');
        if (setup.columns.widths !== undefined) {
          for (const column of children(columns, 'col')) columns.removeChild(column);
          for (const width of setup.columns.widths) {
            const column = wordElement(document, 'col');
            column.setAttributeNS(WORD_NS, 'w:w', String(Math.max(0, Math.trunc(width))));
            columns.appendChild(column);
          }
        }
      }
      if (setup.pageNumbering !== undefined) {
        const existing = children(sectPr, 'pgNumType')[0];
        if (!setup.pageNumbering) {
          if (existing) sectPr.removeChild(existing);
        } else if (setup.pageNumbering.start === undefined && setup.pageNumbering.format === undefined) {
          if (existing) sectPr.removeChild(existing);
        } else {
          const numbering = existing ?? sectionProperty(sectPr, 'pgNumType');
          if (setup.pageNumbering.start === undefined) numbering.removeAttributeNS(WORD_NS, 'start');
          else numbering.setAttributeNS(WORD_NS, 'w:start', String(Math.max(0, Math.trunc(setup.pageNumbering.start))));
          if (setup.pageNumbering.format === undefined) numbering.removeAttributeNS(WORD_NS, 'fmt');
          else numbering.setAttributeNS(WORD_NS, 'w:fmt', setup.pageNumbering.format);
          if (!numbering.getAttributeNS(WORD_NS, 'start') && !numbering.getAttributeNS(WORD_NS, 'fmt')) sectPr.removeChild(numbering);
        }
      }
      if (setup.titlePage !== undefined) {
        const existing = children(sectPr, 'titlePg')[0];
        if (setup.titlePage) {
          if (!existing) sectionProperty(sectPr, 'titlePg');
        } else if (existing) {
          sectPr.removeChild(existing);
        }
      }
    });
  }

  insertSectionBreak(paragraph: number, type: SectionType): void {
    this.updatePartXml(this.mainPath, document => {
      const target = paragraphAt(document, paragraph);
      let sections = collectSections(document);
      if (!sections.length) {
        bodyOf(document).appendChild(wordElement(document, 'sectPr'));
        sections = collectSections(document);
      }
      const index = sections.findIndex(section => paragraph >= section.startParagraph && paragraph <= section.endParagraph);
      const current = sections[index];
      if (!current) throw new Error(`Section for paragraph ${paragraph} does not exist.`);
      if (children(children(target, 'pPr')[0] ?? target, 'sectPr')[0]) {
        throw new Error('Paragraph already ends with a section break.');
      }
      const source = current.sectPr;
      const props = properties(target, 'pPr');
      const sectionProps = children(props, 'sectPr')[0] ?? property(props, 'sectPr');
      while (sectionProps.firstChild) sectionProps.removeChild(sectionProps.firstChild);
      for (const attribute of [...Array.from(sectionProps.attributes)]) sectionProps.removeAttributeNode(attribute);
      const copy = source.cloneNode(true) as Element;
      for (let child = copy.firstChild; child; child = child.nextSibling) {
        sectionProps.appendChild(child.cloneNode(true));
      }
      for (const attribute of [...Array.from(copy.attributes)]) {
        sectionProps.setAttributeNS(attribute.namespaceURI, attribute.name, attribute.value);
      }
      setWordValue(sectionProperty(sectionProps, 'type'), type);
    });
  }

  deleteSectionBreak(section: number): void {
    this.updatePartXml(this.mainPath, document => {
      const sections = collectSections(document);
      const current = sections[section];
      if (!current || section >= sections.length - 1) throw new Error('Cannot delete the final section break.');
      if (current.source !== 'paragraph') throw new Error('Section break is not paragraph-scoped.');
      const props = children(current.paragraph!, 'pPr')[0];
      const sectPr = props ? children(props, 'sectPr')[0] : undefined;
      if (!props || !sectPr) throw new Error('Section break does not exist.');
      props.removeChild(sectPr);
      if (!props.childNodes.length) current.paragraph!.removeChild(props);
    });
  }

  private getHeaderFooterBlocks(section: number, type: 'header' | 'footer', kind: 'default' | 'first' | 'even'): DocumentBlock[] {
    const sectionInfo = this.getSection(section);
    const map = type === 'header' ? sectionInfo.headers : sectionInfo.footers;
    const entry = map[kind] ?? map.default;
    if (!entry || !this.parts.has(entry)) return [];
    const document = this.getPartDocument(entry);
    const styles = this.getStylesContext();
    const numbering = this.getNumberingContext();
    const paragraphs = this.buildParagraphs(document, styles, numbering, entry);
    return this.buildBlocksFrom(document, paragraphs);
  }

  getHeaderBlocks(section: number, kind: 'default' | 'first' | 'even' = DEFAULT_HEADER_FOOTER_KIND): DocumentBlock[] {
    return this.getHeaderFooterBlocks(section, 'header', kind);
  }

  getFooterBlocks(section: number, kind: 'default' | 'first' | 'even' = DEFAULT_HEADER_FOOTER_KIND): DocumentBlock[] {
    return this.getHeaderFooterBlocks(section, 'footer', kind);
  }

  private createHeaderFooter(section: number, type: 'header' | 'footer', kind: 'default' | 'first' | 'even'): string {
    if (kind === 'even') this.enableEvenAndOddHeaders();
    const rootName = type === 'header' ? 'hdr' : 'ftr';
    const relType = type === 'header' ? HEADER_REL : FOOTER_REL;
    const contentType = type === 'header' ? HEADER_TYPE : FOOTER_TYPE;
    const tag = type === 'header' ? 'headerReference' : 'footerReference';
    let existingPath: string | undefined;
    const targets = this.relationshipTargets(this.mainPath);
    let main = this.getPartDocument(this.mainPath);
    let sections = collectSections(main);
    if (!sections.length && section === 0) {
      this.updatePartXml(this.mainPath, document => { bodyOf(document).appendChild(wordElement(document, 'sectPr')); });
      main = this.getPartDocument(this.mainPath);
      sections = collectSections(main);
    }
    const descriptor = sections[section];
    if (!descriptor) throw new Error(`Section ${section} does not exist.`);
    const existing = children(descriptor.sectPr, tag).find(reference => reference.getAttributeNS(WORD_NS, 'type') === kind);
    if (existing) {
      const id = existing.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id') ?? existing.getAttribute('r:id');
      existingPath = id ? targets.get(id)?.target : undefined;
      if (existingPath) return existingPath;
    }
    const baseDirectory = partDirectory(this.mainPath) || 'word';
    let index = 1;
    let path = `${baseDirectory}/${type}${index}.xml`;
    while (this.parts.has(path)) { index++; path = `${baseDirectory}/${type}${index}.xml`; }
    const xml = `<w:${rootName} xmlns:w="${WORD_NS}"><w:p><w:r><w:t xml:space="preserve"></w:t></w:r></w:p></w:${rootName}>`;
    this.addPart(path, encodeXml(xml), contentType);
    const id = this.addRelationship(this.mainPath, relType, relativeTarget(this.mainPath, path));
    this.updatePartXml(this.mainPath, document => {
      const sectionDescriptor = collectSections(document)[section];
      if (!sectionDescriptor) throw new Error(`Section ${section} does not exist.`);
      const sectPr = sectionDescriptor.sectPr;
      const reference = wordElement(sectPr.ownerDocument!, tag);
      reference.setAttributeNS(WORD_NS, 'w:type', kind);
      reference.setAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'r:id', id);
      const position = SECTION_ORDER.indexOf(tag as typeof SECTION_ORDER[number]);
      const following = children(sectPr).find(child => SECTION_ORDER.indexOf(child.localName as typeof SECTION_ORDER[number]) > position);
      sectPr.insertBefore(reference, following ?? null);
    });
    return path;
  }

  createHeader(section: number, kind: 'default' | 'first' | 'even' = DEFAULT_HEADER_FOOTER_KIND): string {
    return this.createHeaderFooter(section, 'header', kind);
  }

  createFooter(section: number, kind: 'default' | 'first' | 'even' = DEFAULT_HEADER_FOOTER_KIND): string {
    return this.createHeaderFooter(section, 'footer', kind);
  }

  private setHeaderFooterText(section: number, text: string, type: 'header' | 'footer', kind: 'default' | 'first' | 'even'): void {
    assertText(text);
    const map = type === 'header' ? this.getSection(section).headers : this.getSection(section).footers;
    const existingPath = map[kind];
    if (existingPath && this.parts.has(existingPath)) {
      const existingDocument = this.getPartDocument(existingPath);
      const current = descendants(blockContainerOf(existingDocument), 'p').map(textOf).join('\n');
      if (current === text) return;
    }
    const path = type === 'header' ? this.createHeader(section, kind) : this.createFooter(section, kind);
    this.updatePartXml(path, part => {
      const container = blockContainerOf(part);
      while (container.firstChild) container.removeChild(container.firstChild);
      container.appendChild(newParagraph(part, text));
    });
  }

  setHeaderText(section: number, text: string, kind: 'default' | 'first' | 'even' = DEFAULT_HEADER_FOOTER_KIND): void {
    this.setHeaderFooterText(section, text, 'header', kind);
  }

  setFooterText(section: number, text: string, kind: 'default' | 'first' | 'even' = DEFAULT_HEADER_FOOTER_KIND): void {
    this.setHeaderFooterText(section, text, 'footer', kind);
  }

  insertPageNumberField(partPath: string, options: { format?: string; total?: boolean } = {}): void {
    validatePath(partPath);
    const command = options.total ? 'NUMPAGES' : 'PAGE';
    const format = options.format ? ` \\* ${options.format}` : '';
    this.updatePartXml(partPath, document => {
      const container = blockContainerOf(document);
      const paragraph = wordElement(document, 'p');
      const field = wordElement(document, 'fldSimple');
      field.setAttributeNS(WORD_NS, 'w:instr', ` ${command}${format} `);
      const run = wordElement(document, 'r');
      appendText(run, options.total ? '?' : '1');
      field.appendChild(run);
      paragraph.appendChild(field);
      container.appendChild(paragraph);
    });
  }

  private enableEvenAndOddHeaders(): void {
    const relationships = this.relationshipTargets(this.mainPath);
    let settingsPath = [...relationships.values()].find(relationship => relationship.type === SETTINGS_REL)?.target;
    if (!settingsPath) {
      settingsPath = `${partDirectory(this.mainPath) || 'word'}/settings.xml`;
      if (!this.parts.has(settingsPath)) {
        this.addPart(settingsPath, encodeXml(`<w:settings xmlns:w="${WORD_NS}"></w:settings>`), SETTINGS_TYPE);
      }
      this.addRelationship(this.mainPath, SETTINGS_REL, relativeTarget(this.mainPath, settingsPath));
    } else if (!this.parts.has(settingsPath)) {
      this.addPart(settingsPath, encodeXml(`<w:settings xmlns:w="${WORD_NS}"></w:settings>`), SETTINGS_TYPE);
    }
    this.updatePartXml(settingsPath, document => {
      const root = document.documentElement!;
      if (!children(root, 'evenAndOddHeaders')[0]) root.appendChild(wordElement(document, 'evenAndOddHeaders'));
    });
  }

  setParagraphText(index: number, text: string): void {
    assertText(text);
    const normalized = text.replace(/\r\n?/g, '\n');
    this.nextHistoryAction = { kind: 'setParagraphText', paragraph: index };
    this.updatePartXmlInternal(this.mainPath, document => {
      const paragraph = paragraphAt(document, index);
      const old = textOf(paragraph);
      let start = 0;
      while (start < old.length && start < normalized.length && old[start] === normalized[start]) start++;
      let end = old.length;
      let replacementEnd = normalized.length;
      while (end > start && replacementEnd > start && old[end - 1] === normalized[replacementEnd - 1]) {
        end--; replacementEnd--;
      }
      // Never cut a surrogate pair at a shared UTF-16 prefix/suffix.
      if (start > 0 && /[\ud800-\udbff]/.test(old[start - 1]!)) start--;
      if (end < old.length && /[\udc00-\udfff]/.test(old[end]!)) { end++; replacementEnd++; }
      if (old !== normalized) {
        const replacement = normalized.slice(start, replacementEnd);
        if (this.trackChangesEnabled()) this.replaceSpanTracked(paragraph, start, end, replacement);
        else replaceSpan(paragraph, start, end, replacement);
        return true;
      }
      return false;
    });
  }

  insertParagraph(text: string, before?: number): void {
    assertText(text);
    this.updatePartXmlInternal(this.mainPath, document => {
      const paragraph = newParagraph(document, text);
      if (this.trackChangesEnabled()) {
        const runs = children(paragraph, 'r');
        wrapRunsWithRevision(paragraph, runs, 'ins', this.trackedRevisionAuthor());
        markRevision(property(properties(paragraph, 'pPr'), 'rPr'), 'ins', this.trackedRevisionAuthor(), undefined, 'paraRPr');
      }
      if (before !== undefined) {
        const target = paragraphAt(document, before);
        target.parentNode!.insertBefore(paragraph, target);
      } else {
        const body = bodyOf(document);
        body.insertBefore(paragraph, children(body, 'sectPr')[0] ?? null);
      }
      return true;
    });
  }

  deleteParagraph(index: number): void {
    this.updatePartXmlInternal(this.mainPath, document => {
      const paragraph = paragraphAt(document, index);
      if (this.trackChangesEnabled()) {
        if (children(paragraph, 'pPr').some(props => children(props, 'sectPr').length)) {
          throw new Error('Cannot delete a section-break paragraph while track changes is enabled; edit its XML explicitly.');
        }
        const runs = ownRuns(paragraph);
        if (runs.length) wrapRunsWithRevision(paragraph, runs, 'del', this.trackedRevisionAuthor());
        else {
          for (const child of [...children(paragraph)]) {
            if (child.localName === 'pPr' || isParagraphAnchorMarker(child)) continue;
            paragraph.removeChild(child);
          }
        }
        markRevision(property(properties(paragraph, 'pPr'), 'rPr'), 'del', this.trackedRevisionAuthor(), undefined, 'paraRPr');
        return;
      }
      deleteParagraphElement(paragraph);
    });
  }

  formatParagraph(index: number, format: ParagraphFormat, options: { validateStyle?: boolean } = {}): void {
    validateParagraphFormat(format);
    if (options.validateStyle && typeof format.style === 'string' && !this.getStyle(format.style)) {
      throw new Error(`Paragraph style not found: ${format.style} (styles.xml is missing or does not define it).`);
    }
    this.updatePartXmlInternal(this.mainPath, document => {
      const paragraph = paragraphAt(document, index);
      if (this.trackChangesEnabled()) this.trackParagraphFormatChange(paragraph);
      const props = properties(paragraph, 'pPr');
      applyParagraphFormatTo(props, format);
    });
  }

  formatRun(paragraph: number, run: number, format: RunFormat): void {
    assertIndex(run);
    validateRunFormat(format);
    this.updatePartXmlInternal(this.mainPath, document => {
      const element = ownRuns(paragraphAt(document, paragraph))[run];
      if (!element) throw new Error(`Run ${run} does not exist.`);
      if (this.trackChangesEnabled()) this.trackRunFormatChange(element);
      const props = properties(element, 'rPr');
      applyRunFormatTo(props, format);
    });
  }

  formatRange(range: TextRange, format: RunFormat): void {
    validateRunFormat(format);
    const preflightDocument = this.getCachedPartDocument(this.mainPath);
    const preflight = this.normalizeRangeOn(preflightDocument, range);
    if (preflight.start === preflight.end) return;
    if (!this.runsInRange(preflight.paragraph, preflight.start, preflight.end).length) return;
    this.updatePartXmlInternal(this.mainPath, document => {
      const normalized = this.normalizeRangeOn(document, range);
      if (this.trackChangesEnabled()) {
        this.splitRunAtOffset(normalized.paragraph, normalized.end);
        this.splitRunAtOffset(normalized.paragraph, normalized.start);
        for (const run of this.runsInRange(normalized.paragraph, normalized.start, normalized.end)) {
          this.trackRunFormatChange(run);
          applyRunFormatTo(properties(run, 'rPr'), format);
        }
        return;
      }
      this.applyRunFormatRangeOnParagraph(normalized.paragraph, normalized.start, normalized.end, format);
    });
  }

  getRangeFormat(range: TextRange): RunFormat {
    const document = this.getCachedPartDocument(this.mainPath);
    const theme = this.getStylesContext().theme;
    return this.mergeRangeFormats(this.rangeRunFormats(document, range, theme));
  }

  copyFormat(range: TextRange): RunFormat {
    return this.getRangeFormat(range);
  }

  clearRangeFormat(range: TextRange, fields?: (keyof RunFormat)[]): void {
    const targets = this.validateRangeFields(fields);
    if (!targets.length) return;
    const format: RunFormat = {};
    for (const field of targets) (format as Record<string, unknown>)[field] = null;
    this.formatRange(range, format);
  }

  formatDocumentRange(range: DocumentRange, format: RunFormat): void {
    validateRunFormat(format);
    const preflightDocument = this.getCachedPartDocument(this.mainPath);
    const normalized = this.normalizeDocumentRange(preflightDocument, range);
    const hasTarget = (() => {
      for (let index = normalized.start.paragraph; index <= normalized.end.paragraph; index++) {
        const paragraph = paragraphAt(preflightDocument, index);
        const length = textOf(paragraph).length;
        const start = index === normalized.start.paragraph ? normalized.start.offset : 0;
        const end = index === normalized.end.paragraph ? normalized.end.offset : length;
        if (start === end) continue;
        if (this.runsInRange(paragraph, start, end).length) return true;
      }
      return false;
    })();
    if (!hasTarget) return;
    this.withDraft((draft) => {
      draft.updatePartXmlInternal(draft.mainPath, (document) => {
        const normalized = draft.normalizeDocumentRange(document, range);
        for (let index = normalized.start.paragraph; index <= normalized.end.paragraph; index++) {
          const paragraph = paragraphAt(document, index);
          const length = textOf(paragraph).length;
          const start = index === normalized.start.paragraph ? normalized.start.offset : 0;
          const end = index === normalized.end.paragraph ? normalized.end.offset : length;
          if (draft.trackChangesEnabled()) {
            draft.splitRunAtOffset(paragraph, end);
            draft.splitRunAtOffset(paragraph, start);
            for (const run of draft.runsInRange(paragraph, start, end)) {
              draft.trackRunFormatChange(run);
              applyRunFormatTo(properties(run, 'rPr'), format);
            }
          } else {
            draft.applyRunFormatRangeOnParagraph(paragraph, start, end, format);
          }
        }
      });
    });
  }

  getDocumentRangeFormat(range: DocumentRange): RunFormat {
    const document = this.getCachedPartDocument(this.mainPath);
    const theme = this.getStylesContext().theme;
    const normalized = this.normalizeDocumentRange(document, range);
    const formats: RunFormat[] = [];
    for (let index = normalized.start.paragraph; index <= normalized.end.paragraph; index++) {
      const paragraph = paragraphAt(document, index);
      const length = textOf(paragraph).length;
      const start = index === normalized.start.paragraph ? normalized.start.offset : 0;
      const end = index === normalized.end.paragraph ? normalized.end.offset : length;
      const rangeFormats = this.rangeRunFormats(document, { paragraph: index, start, end }, theme);
      formats.push(...rangeFormats);
    }
    return this.mergeRangeFormats(formats);
  }

  applyFormat(range: DocumentRange, format: RunFormat): void {
    this.formatDocumentRange(range, format);
  }

  setParagraphTabs(index: number, tabs: TabStop[]): void {
    validateTabs(tabs);
    this.formatParagraph(index, { tabs: tabs.length ? tabs : null });
  }

  setParagraphBorders(index: number, borders: ParagraphFormat['borders']): void {
    validateParagraphBorders(borders);
    this.formatParagraph(index, { borders });
  }

  setParagraphShading(index: number, shading: Shading): void {
    validateDocShading(shading);
    this.formatParagraph(index, { shading });
  }

  insertBreak(paragraph: number, run: number, type: 'textWrapping' | 'page' | 'column'): void {
    assertIndex(run);
    if (!['textWrapping', 'page', 'column'].includes(type)) throw new Error('Invalid break type.');
    this.updatePartXml(this.mainPath, document => {
      const target = ownRuns(paragraphAt(document, paragraph))[run];
      if (!target) throw new Error(`Run ${run} does not exist.`);
      const br = wordElement(document, 'br');
      if (type !== 'textWrapping') br.setAttributeNS(WORD_NS, 'w:type', type);
      target.appendChild(br);
    });
  }

  insertSymbol(paragraph: number, run: number, font: string, charCode: number): void {
    assertIndex(run);
    assertText(font, 'font');
    if (!Number.isInteger(charCode) || charCode < 0 || charCode > 0xFFFF || !isValidXmlCharCode(charCode)) {
      throw new Error('charCode must be an XML-valid BMP code point.');
    }
    this.updatePartXml(this.mainPath, document => {
      const target = ownRuns(paragraphAt(document, paragraph))[run];
      if (!target) throw new Error(`Run ${run} does not exist.`);
      const symbol = wordElement(document, 'sym');
      symbol.setAttributeNS(WORD_NS, 'w:font', font);
      symbol.setAttributeNS(WORD_NS, 'w:char', charCode.toString(16).toUpperCase().padStart(4, '0'));
      target.appendChild(symbol);
    });
  }

  getSettings(): { defaultTabStop: number; evenAndOddHeaders: boolean; trackChanges: boolean; [key: string]: unknown } {
    const base = { defaultTabStop: 720, evenAndOddHeaders: false, trackChanges: false };
    const path = this.getSettingsPath();
    if (!path) return base;
    try {
      const root = this.getPartDocument(path).documentElement;
      if (!root || root.namespaceURI !== WORD_NS || root.localName !== 'settings') return base;
      const tab = children(root, 'defaultTabStop')[0];
      const raw = tab?.getAttributeNS(WORD_NS, 'val') ?? tab?.getAttribute('w:val');
      const value = raw !== null && raw !== undefined && /^-?\d+$/.test(raw) ? Number(raw) : 720;
      const odd = children(root, 'evenAndOddHeaders')[0];
      const enabled = odd ? !['0', 'false', 'off'].includes((wordValue(odd) ?? '1').toLowerCase()) : false;
      const track = children(root, 'trackChanges')[0];
      const trackChanges = track ? !['0', 'false', 'off'].includes((wordValue(track) ?? '1').toLowerCase()) : false;
      return { ...base, defaultTabStop: Number.isFinite(value) && value > 0 ? value : 720, evenAndOddHeaders: enabled, trackChanges };
    } catch {
      return base;
    }
  }

  setRevisionAuthor(author: string): void {
    assertText(author, 'author');
    this.revisionAuthor = author;
  }

  private trackedRevisionAuthor(author?: string): string | undefined {
    return author ?? this.revisionAuthor;
  }

  private trackChangesEnabled(): boolean {
    return this.getSettings().trackChanges;
  }

  setTrackChanges(enabled: boolean): void {
    if (this.hasExplicitTrackChangesSetting(enabled)) return;
    this.withDraft((draft) => draft.setTrackChangesDirect(enabled));
  }

  private setTrackChangesDirect(enabled: boolean): void {
    let path = this.getSettingsPath(true);
    if (!path) path = `${dirname(this.mainPath) ? `${dirname(this.mainPath)}/` : ''}settings.xml`;
    if (!this.parts.has(path)) this.addPart(path, encodeXml(`<w:settings xmlns:w="${WORD_NS}"/>`), SETTINGS_TYPE);
    this.ensureMainRelationship(SETTINGS_REL, relativeTarget(this.mainPath, path));
    this.updatePartXml(path, (document) => setTrackChangesOn(document, enabled));
  }

  private hasExplicitTrackChangesSetting(enabled: boolean): boolean {
    const path = this.getSettingsPath();
    if (!path || !this.parts.has(path)) return false;
    try {
      const root = this.getPartDocument(path).documentElement;
      if (!root || root.namespaceURI !== WORD_NS || root.localName !== 'settings') return false;
      const trackChanges = children(root, 'trackChanges')[0];
      if (!trackChanges) return false;
      const current = !['0', 'false', 'off'].includes((wordValue(trackChanges) ?? '1').toLowerCase());
      return current === enabled;
    } catch {
      return false;
    }
  }

  defineStyle(style: StyleInfo): void {
    assertText(style.id, 'style.id');
    const styleName = style.name || style.id;
    assertText(styleName, 'style.name');
    const type = style.type;
    if (!['paragraph', 'character', 'table', 'numbering'].includes(type)) throw new Error(`Unsupported style type: ${String(type)}`);
    if (style.uiPriority !== undefined && (!Number.isSafeInteger(style.uiPriority) || style.uiPriority < 0)) {
      throw new Error('style.uiPriority must be a non-negative safe integer.');
    }
    if (style.paragraph !== undefined) {
      validateParagraphFormat(style.paragraph);
      rejectNullFormatValues(style.paragraph, 'style.paragraph');
    }
    if (style.run !== undefined) {
      validateRunFormat(style.run);
      rejectNullFormatValues(style.run, 'style.run');
    }
    const draft = new DocxDocument(new Map([...this.parts].map(([path, bytes]) => [path, Uint8Array.from(bytes)])));
    const stylesPath = draft.getStylesPath() ?? (() => {
      const path = `${dirname(draft.mainPath) ? `${dirname(draft.mainPath)}/` : ''}styles.xml`;
      const relsDocument = draft.parts.has(relsPath(draft.mainPath))
        ? draft.getPartDocument(relsPath(draft.mainPath))
        : parseXml(`<Relationships xmlns="${REL_NS}"/>`);
      const relsRoot = relsDocument.documentElement!;
      const exists = children(relsRoot, 'Relationship', REL_NS)
        .some(relation => relation.getAttribute('Type') === STYLES_REL);
      if (!exists) {
        const relation = relsDocument.createElementNS(REL_NS, 'Relationship');
        relation.setAttribute('Id', nextRelationshipId(relsRoot));
        relation.setAttribute('Type', STYLES_REL);
        relation.setAttribute('Target', basename(path));
        relsRoot.appendChild(relation);
        draft.parts.set(relsPath(draft.mainPath), encodeXml(serializeXml(relsDocument)));
        draft.documents.delete(relsPath(draft.mainPath));
        draft.dirtyPartXml.delete(relsPath(draft.mainPath));
        draft.dirtyPartSizes.delete(relsPath(draft.mainPath));
      }
      if (!draft.parts.has(path)) {
        const types = draft.getPartDocument('[Content_Types].xml');
        const typesRoot = types.documentElement!;
        if (!children(typesRoot, 'Override', CONTENT_TYPES_NS)
          .some(override => override.getAttribute('PartName') === `/${path}`)) {
          const override = types.createElementNS(CONTENT_TYPES_NS, 'Override');
          override.setAttribute('PartName', `/${path}`);
          override.setAttribute('ContentType', STYLES_TYPE);
          typesRoot.appendChild(override);
          draft.parts.set('[Content_Types].xml', encodeXml(serializeXml(types)));
          draft.documents.delete('[Content_Types].xml');
          draft.dirtyPartXml.delete('[Content_Types].xml');
          draft.dirtyPartSizes.delete('[Content_Types].xml');
        }
        draft.parts.set(path, encodeXml(`<w:styles xmlns:w="${WORD_NS}"/>`));
        draft.documents.delete(path);
        draft.dirtyPartXml.delete(path);
        draft.dirtyPartSizes.delete(path);
      }
      return path;
    })();
    draft.updatePartXml(stylesPath, document => {
      const root = document.documentElement!;
      let styleElement = children(root, 'style').find(element => element.getAttributeNS(WORD_NS, 'styleId') === style.id);
      if (!styleElement) {
        styleElement = wordElement(document, 'style');
        root.appendChild(styleElement);
      } else {
        removeProperty(styleElement, 'name');
        removeProperty(styleElement, 'basedOn');
        removeProperty(styleElement, 'next');
        removeProperty(styleElement, 'link');
        removeProperty(styleElement, 'aliases');
        removeProperty(styleElement, 'qFormat');
        removeProperty(styleElement, 'pPr');
        removeProperty(styleElement, 'rPr');
      }
      styleElement.setAttributeNS(WORD_NS, 'w:type', style.type);
      styleElement.setAttributeNS(WORD_NS, 'w:styleId', style.id);
      removeWordAttribute(styleElement, 'default');
      if (style.isDefault) styleElement.setAttributeNS(WORD_NS, 'w:default', '1');
      const name = property(styleElement, 'name');
      setWordValue(name, styleName);
      if (style.aliases?.length) setWordValue(property(styleElement, 'aliases'), style.aliases.join(', '));
      if (style.basedOn) setWordValue(property(styleElement, 'basedOn'), style.basedOn);
      if (style.next) setWordValue(property(styleElement, 'next'), style.next);
      if (style.link) setWordValue(property(styleElement, 'link'), style.link);
      if (style.uiPriority !== undefined) setWordValue(property(styleElement, 'uiPriority'), String(style.uiPriority));
      if (style.quickFormat) property(styleElement, 'qFormat');
      if (style.paragraph) applyParagraphFormatTo(property(styleElement, 'pPr'), style.paragraph);
      if (style.run) applyRunFormatTo(property(styleElement, 'rPr'), style.run);
    });
    if (equalPartMap(this.parts, draft.parts)) return;
    this.parts = draft.parts;
    this.mainPath = draft.mainPath;
    this.currentRevision++;
  }

  replaceText(search: string, replacement: string): void {
    assertText(search, 'search'); assertText(replacement, 'replacement');
    if (!search) throw new Error('search must not be empty.');
    this.updatePartXmlInternal(this.mainPath, document => {
      for (const paragraph of descendants(bodyOf(document), 'p')) {
        const text = textOf(paragraph);
        const matches: number[] = [];
        for (let i = text.indexOf(search); i !== -1; i = text.indexOf(search, i + search.length)) matches.push(i);
        for (const start of matches.reverse()) replaceSpan(paragraph, start, start + search.length, replacement);
      }
    });
  }

  insertTable(rows: string[][]): void {
    validateRows(rows);
    this.updatePartXmlInternal(this.mainPath, document => {
      const body = bodyOf(document);
      const section = children(body, 'sectPr')[0] ?? null;
      const table = buildTable(document, rows.length, Math.max(...rows.map(row => row.length)), undefined, rows);
      body.insertBefore(table, section);
      body.insertBefore(newParagraph(document, ''), section);
    });
  }

  insertTableAt(rows: number, cols: number, before?: number, format?: TableFormat): void {
    assertIndex(rows); assertIndex(cols);
    if (rows < 1 || cols < 1) throw new Error('Table must contain at least one row and one column.');
    this.updatePartXmlInternal(this.mainPath, document => {
      const table = buildTable(document, rows, cols, format);
      const body = bodyOf(document);
      const section = children(body, 'sectPr')[0] ?? null;
      if (before !== undefined) {
        const target = bodyBlockAt(document, before);
        const previous = target.previousSibling?.nodeType === 1 ? target.previousSibling as Element : null;
        if (target.localName === 'tbl' && previous?.namespaceURI === WORD_NS && previous.localName === 'p' && textOf(previous) === '') {
          body.insertBefore(table, previous);
        } else {
          body.insertBefore(table, target);
          if (target.localName !== 'p') body.insertBefore(newParagraph(document, ''), target);
        }
      } else {
        body.insertBefore(table, section);
        body.insertBefore(newParagraph(document, ''), section);
      }
    });
  }

  getTable(index: number): TableInfo {
    const document = this.getCachedPartDocument(this.mainPath);
    const paragraphs = this.buildParagraphs(document);
    const body = bodyOf(document);
    const indices = new Map(descendants(body, 'p').map((paragraph, i) => [paragraph, paragraphs[i]!]));
    const walk = (parent: Element): DocumentBlock[] => children(parent).flatMap((child): DocumentBlock[] => {
      if (child.localName === 'p') return [{ type: 'paragraph', paragraph: indices.get(child)! }];
      if (child.localName === 'tbl') return [readTable(child, walk)];
      if (isTransparentWordWrapper(child)) return walk(child);
      return [];
    });
    const table = readTable(tableAt(document, index), walk);
    return { index, rows: table.rows, format: table.format, grid: table.grid };
  }

  insertTableRow(table: number, at: number): void {
    assertIndex(table);
    assertIndex(at);
    this.updatePartXmlInternal(this.mainPath, document => {
      const element = tableAt(document, table);
      const rows = tableRows(element);
      if (at > rows.length) throw new Error(`Row ${at} does not exist.`);
      const tr = wordElement(document, 'tr');
      for (let i = 0; i < tableGrid(element).length; i++) tr.appendChild(blankCell(document));
      const anchor = rows[at] ?? null;
      const parent = (anchor?.parentNode?.nodeType === 1 ? anchor.parentNode as Element : undefined)
        ?? (rows.at(-1)?.parentNode?.nodeType === 1 ? rows.at(-1)!.parentNode as Element : undefined)
        ?? element;
      parent.insertBefore(tr, anchor);
      if (this.trackChangesEnabled()) markRevision(tableProperty(tr, 'trPr'), 'ins', this.trackedRevisionAuthor());
    });
  }

  deleteTableRow(table: number, at: number): void {
    this.updatePartXmlInternal(this.mainPath, document => {
      const element = tableAt(document, table);
      if (this.trackChangesEnabled()) {
        const rows = tableRows(element);
        const row = rows[at];
        if (!row) throw new Error(`Row ${at} does not exist.`);
        if (rows.length <= 1) throw new Error('Cannot delete the only table row.');
        markRevision(tableProperty(row, 'trPr'), 'del', this.trackedRevisionAuthor());
        return;
      }
      deleteTableRowElement(element, at);
    });
  }

  insertTableColumn(table: number, at: number): void {
    assertIndex(table);
    assertIndex(at);
    this.updatePartXmlInternal(this.mainPath, document => {
      const element = tableAt(document, table);
      const grid = ensureTableGrid(element);
      const widths = tableGrid(element);
      if (at > widths.length) throw new Error(`Column ${at} does not exist.`);
      grid.insertBefore(gridCol(document, widths[Math.max(0, Math.min(at, widths.length - 1))] ?? 2250), children(grid, 'gridCol')[at] ?? null);
      const model = tableModel(element);
      const handled = new Set<Element>();
      for (const [rowIndex, row] of tableRows(element).entries()) {
        const covering = at < widths.length ? model.matrix[rowIndex]?.[at] : undefined;
        if (covering && covering.rowSpan > 1) {
          if (handled.has(covering.cell)) continue;
          handled.add(covering.cell);
          for (let index = covering.row; index < covering.row + covering.rowSpan; index++) {
            const chain = rowCells(model.rows[index]!).find(position => position.start === covering.start);
            if (!chain) continue;
            const props = tableProperty(chain.cell, 'tcPr');
            mergeElement(props, 'gridSpan', chain.span + 1 > 1 ? chain.span + 1 : undefined);
          }
          continue;
        }
        const positions = rowCells(row);
        let inserted = false;
        for (const position of positions) {
          if (at > position.start && at < position.start + position.span) {
            const props = tableProperty(position.cell, 'tcPr');
            mergeElement(props, 'gridSpan', position.span + 1 > 1 ? position.span + 1 : undefined);
            inserted = true;
            break;
          }
          if (at === position.start) {
            row.insertBefore(blankCell(document), position.cell);
            inserted = true;
            break;
          }
        }
        if (!inserted) row.appendChild(blankCell(document));
      }
    });
  }

  deleteTableColumn(table: number, at: number): void {
    this.updatePartXmlInternal(this.mainPath, document => {
      const element = tableAt(document, table);
      deleteTableColumnElement(element, at);
    });
  }

  mergeCells(table: number, range: { row: number; col: number; rowSpan: number; colSpan: number }): void {
    assertIndex(range.row); assertIndex(range.col); assertIndex(range.rowSpan); assertIndex(range.colSpan);
    if (range.rowSpan < 1 || range.colSpan < 1) throw new Error('merge range must be at least 1 × 1.');
    this.updatePartXmlInternal(this.mainPath, document => {
      const element = tableAt(document, table);
      const model = tableModel(element);
      const master = model.matrix[range.row]?.[range.col];
      if (!master || master.row !== range.row || master.start !== range.col) throw new Error('mergeCells must start at a visible top-left cell.');
      if (range.row + range.rowSpan > model.rows.length || range.col + range.colSpan > model.grid.length) {
        throw new Error('merge range exceeds table bounds.');
      }
      for (let rowIndex = range.row; rowIndex < range.row + range.rowSpan; rowIndex++) {
        const positions = rowCells(model.rows[rowIndex]!);
        const covered = positions.filter(position => position.start >= range.col && position.start + position.span <= range.col + range.colSpan);
        if (!covered.length || covered[0]!.start !== range.col) throw new Error('mergeCells requires a rectangular, grid-aligned selection.');
        for (const position of positions) {
          const intersects = position.start < range.col + range.colSpan && position.start + position.span > range.col;
          const contained = position.start >= range.col && position.start + position.span <= range.col + range.colSpan;
          if (intersects && !contained) throw new Error('mergeCells cannot partially cover existing merged cells.');
        }
        if (rowIndex === range.row) {
          for (const position of covered) {
            if (position.cell === master.cell) continue;
            appendCellContent(master.cell, position.cell);
            model.rows[rowIndex]!.removeChild(position.cell);
          }
        } else {
          const continuation = covered[0]!;
          appendCellContent(master.cell, continuation.cell);
          for (const position of covered.slice(1)) {
            appendCellContent(master.cell, position.cell);
            model.rows[rowIndex]!.removeChild(position.cell);
          }
          clearCellContent(continuation.cell);
          const props = tableProperty(continuation.cell, 'tcPr');
          mergeElement(props, 'gridSpan', range.colSpan > 1 ? range.colSpan : undefined);
          mergeElement(props, 'vMerge', 'continue');
        }
      }
      const masterProps = tableProperty(master.cell, 'tcPr');
      mergeElement(masterProps, 'gridSpan', range.colSpan > 1 ? range.colSpan : undefined);
      mergeElement(masterProps, 'vMerge', range.rowSpan > 1 ? 'restart' : undefined);
      ensureCellParagraph(master.cell);
    });
  }

  splitCell(table: number, row: number, col: number, rows: number, cols: number): void {
    assertIndex(row); assertIndex(col); assertIndex(rows); assertIndex(cols);
    this.updatePartXmlInternal(this.mainPath, document => {
      const element = tableAt(document, table);
      const model = tableModel(element);
      const master = model.matrix[row]?.[col];
      if (!master || master.row !== row || master.start !== col) throw new Error('splitCell must target a visible top-left cell.');
      if (rows !== master.rowSpan || cols !== master.colSpan) {
        throw new Error('splitCell currently supports restoring a merged cell to its original grid span.');
      }
      for (let rowIndex = row; rowIndex < row + rows; rowIndex++) {
        const rowElement = model.rows[rowIndex]!;
        const position = rowCells(rowElement).find(item => item.start === col);
        const cell = rowIndex === row ? master.cell : position?.cell;
        if (!cell) throw new Error('splitCell found an invalid merged-cell structure.');
        const anchor = rowCells(rowElement).find(item => item.start >= col + cols)?.cell ?? null;
        const props = tableProperty(cell, 'tcPr');
        mergeElement(props, 'gridSpan', undefined);
        mergeElement(props, 'vMerge', undefined);
        if (rowIndex !== row) clearCellContent(cell);
        const existingCells = rowIndex === row ? 1 : 1;
        for (let i = existingCells; i < cols; i++) {
          const extra = blankCell(document);
          rowElement.insertBefore(extra, anchor);
        }
      }
    });
  }

  formatTable(table: number, format: TableFormat): void {
    this.updatePartXmlInternal(this.mainPath, document => setTableFormat(tableAt(document, table), format));
  }

  formatTableRow(table: number, row: number, format: RowFormat): void {
    this.updatePartXmlInternal(this.mainPath, document => {
      const element = tableRows(tableAt(document, table))[row];
      if (!element) throw new Error(`Row ${row} does not exist.`);
      setRowFormat(element, format);
    });
  }

  formatCell(table: number, row: number, col: number, format: CellFormat): void {
    this.updatePartXmlInternal(this.mainPath, document => setCellFormat(cellAt(tableAt(document, table), row, col).cell, format));
  }

  setCellText(table: number, row: number, col: number, text: string): void {
    assertText(text);
    this.updatePartXmlInternal(this.mainPath, document => {
      const cell = cellAt(tableAt(document, table), row, col).cell;
      ensureCellParagraph(cell);
      const paragraph = childrenThroughTransparent(cell, 'p')[0];
      if (!paragraph) throw new Error('Cell paragraph does not exist.');
      const indices = new Map(descendants(bodyOf(document), 'p').map((item, index) => [item, index]));
      const index = indices.get(paragraph);
      if (index === undefined) throw new Error('Cell paragraph index does not exist.');
      const old = textOf(paragraph);
      let start = 0;
      while (start < old.length && start < text.length && old[start] === text[start]) start++;
      let end = old.length;
      let replacementEnd = text.length;
      while (end > start && replacementEnd > start && old[end - 1] === text[replacementEnd - 1]) { end--; replacementEnd--; }
      if (start > 0 && /[\ud800-\udbff]/.test(old[start - 1]!)) start--;
      if (start > 0 && /[\ud800-\udbff]/.test(text[start - 1]!)) start--;
      if (end < old.length && /[\udc00-\udfff]/.test(old[end]!)) { end++; replacementEnd++; }
      if (replacementEnd < text.length && /[\udc00-\udfff]/.test(text[replacementEnd]!)) replacementEnd++;
      if (old !== text) {
        replaceSpan(paragraph, start, end, text.slice(start, replacementEnd));
        return true;
      }
      readParagraph(paragraph, index, this.getStylesContext());
      return false;
    });
  }

  private withDraft<T>(action: (draft: DocxDocument) => T): T {
    this.materializeAllParts();
    const draft = new DocxDocument(new Map(this.parts));
    draft.currentRevision = this.currentRevision;
    draft.revisionAuthor = this.revisionAuthor;
    draft.undoHistory = this.undoHistory.map((state) => ({
      parts: draft.cloneParts(state.parts),
      bytes: state.bytes,
      entry: { ...state.entry },
      action: state.action,
    }));
    draft.redoHistory = this.redoHistory.map((state) => ({
      parts: draft.cloneParts(state.parts),
      bytes: state.bytes,
      entry: { ...state.entry },
      action: state.action,
    }));
    draft.undoHistoryBytes = this.undoHistoryBytes;
    draft.redoHistoryBytes = this.redoHistoryBytes;
    draft.suppressHistory = true;
    const history = this.beginMutationHistory(this.nextHistoryLabel, this.nextHistoryAction);
    this.nextHistoryLabel = undefined;
    this.nextHistoryAction = { kind: 'other' };
    try {
      const result = action(draft);
      draft.materializeAllParts();
      if (equalPartMap(this.parts, draft.parts)) {
        this.pendingMergedHistory = undefined;
        return result;
      }
      this.parts = draft.parts;
      this.documents = draft.documents;
      this.dirtyPartXml = draft.dirtyPartXml;
      this.dirtyPartSizes = draft.dirtyPartSizes;
      this.mainPath = draft.mainPath;
      this.revisionAuthor = draft.revisionAuthor;
      this.currentRevision++;
      this.numberingContextCache = undefined;
      this.stylesCache = undefined;
      this.outlineCache = undefined;
      this.noteStateCache = undefined;
      this.commentStateCache = undefined;
      this.revisionInfoCache = undefined;
      this.reviewerInfoCache = undefined;
      this.tableCellLocationCache = undefined;
      this.imageDataUrls.clear();
      if (history) this.recordHistory(history);
      return result;
    } catch (error) {
      this.abortHistoryGroupOnFailure();
      throw error;
    }
  }

  private nextNoteId(kind: NoteKind): number {
    const used = new Set<number>();
    const path = this.getNotePartPath(kind);
    if (path && this.parts.has(path)) {
      for (const entry of parseNoteEntries(this.getPartDocument(path), kind)) {
        if (entry.id >= 1) used.add(entry.id);
      }
    }
    const body = bodyOf(this.getPartDocument(this.mainPath));
    for (const reference of referenceRecords(body)) {
      if (reference.kind === kind && reference.id >= 1) used.add(reference.id);
    }
    let id = 1;
    while (used.has(id)) id++;
    return id;
  }

  getNoteSettings(): NoteSettings {
    const settingsPath = this.getSettingsPath();
    return parseDocumentNoteSettings(settingsPath && this.parts.has(settingsPath) ? this.getPartDocument(settingsPath) : null);
  }

  setNoteSettings(settings: Partial<NoteSettings>): void {
    const safe = normalizedNoteSettingsPatch(settings ?? {});
    if (!safe.footnote && !safe.endnote) return;
    this.withDraft((draft) => draft.setNoteSettingsDirect(safe));
  }

  private setNoteSettingsDirect(settings: Partial<NoteSettings>): void {
    let path = this.getSettingsPath(true);
    if (!path) path = `${dirname(this.mainPath) ? `${dirname(this.mainPath)}/` : ''}settings.xml`;
    if (!this.parts.has(path)) this.addPart(path, encodeXml(`<w:settings xmlns:w="${WORD_NS}"/>`), SETTINGS_TYPE);
    this.ensureMainRelationship(SETTINGS_REL, relativeTarget(this.mainPath, path));
    this.updatePartXml(path, (document) => setNoteSettingsOn(document, settings));
  }

  getDocumentProperties(): DocumentProperties {
    const corePath = this.getCorePropertiesPath();
    const appPath = this.getAppPropertiesPath();
    return parseDocumentProperties(
      corePath && this.parts.has(corePath) ? this.getPartDocument(corePath) : null,
      appPath && this.parts.has(appPath) ? this.getPartDocument(appPath) : null,
    );
  }

  setDocumentProperties(patch: Partial<DocumentProperties>): void {
    assertDocumentPropertiesPatch(patch);
    if (!Object.keys(patch).length) return;
    this.withDraft((draft) => draft.setDocumentPropertiesDirect(patch));
  }

  private setDocumentPropertiesDirect(patch: Partial<DocumentProperties>): void {
    const coreKeys = CORE_PROPERTY_KEYS;
    const appKeys = APP_PROPERTY_KEYS;
    const corePatch = Object.fromEntries(coreKeys.filter((key) => key in patch).map((key) => [key, patch[key]])) as Partial<DocumentProperties>;
    const appPatch = Object.fromEntries(appKeys.filter((key) => key in patch).map((key) => [key, patch[key]])) as Partial<DocumentProperties>;
    if (Object.keys(corePatch).length) {
      const hasDefined = coreKeys.some((key) => key in corePatch && corePatch[key] !== undefined);
      let path = this.getCorePropertiesPath(hasDefined);
      if (path || hasDefined) {
        path ??= 'docProps/core.xml';
        this.ensurePackageRelationship(CORE_PROPS_REL, path);
        if (!this.parts.has(path)) this.addPart(path, encodeXml(defaultCorePropertiesXml()), CORE_PROPS_TYPE);
        this.updatePartXml(path, (document) => setCoreDocumentPropertiesOn(document, corePatch));
      }
    }
    if (Object.keys(appPatch).length) {
      const hasDefined = appKeys.some((key) => key in appPatch && appPatch[key] !== undefined);
      let path = this.getAppPropertiesPath(hasDefined);
      if (path || hasDefined) {
        path ??= 'docProps/app.xml';
        this.ensurePackageRelationship(APP_PROPS_REL, path);
        if (!this.parts.has(path)) this.addPart(path, encodeXml(defaultAppPropertiesXml()), APP_PROPS_TYPE);
        this.updatePartXml(path, (document) => setAppDocumentPropertiesOn(document, appPatch));
      }
    }
  }

  getDocumentProtection(): DocumentProtection {
    const path = this.getSettingsPath();
    return parseDocumentProtection(path && this.parts.has(path) ? this.getPartDocument(path) : null);
  }

  setDocumentProtection(value: DocumentProtection): void {
    assertDocumentProtection(value);
    if (!value.enabled && value.edit === undefined && value.enforced === undefined && !this.getDocumentProtection().enabled && !this.hasDocumentProtectionElement()) return;
    this.withDraft((draft) => draft.setDocumentProtectionDirect(value));
  }

  private setDocumentProtectionDirect(value: DocumentProtection): void {
    let path = this.getSettingsPath(true);
    if (!path) path = `${dirname(this.mainPath) ? `${dirname(this.mainPath)}/` : ''}settings.xml`;
    if (!this.parts.has(path)) this.addPart(path, encodeXml(`<w:settings xmlns:w="${WORD_NS}"/>`), SETTINGS_TYPE);
    this.ensureMainRelationship(SETTINGS_REL, relativeTarget(this.mainPath, path));
    this.updatePartXml(path, (document) => setDocumentProtectionOn(document, value));
  }

  private hasDocumentProtectionElement(): boolean {
    const path = this.getSettingsPath();
    if (!path || !this.parts.has(path)) return false;
    try {
      const root = this.getPartDocument(path).documentElement;
      return Boolean(root && root.namespaceURI === WORD_NS && root.localName === 'settings' && children(root, 'documentProtection')[0]);
    } catch {
      return false;
    }
  }

  insertFootnote(paragraph: number, run: number, text: string, options: { customMark?: string } = {}): NoteInfo {
    return this.withDraft((draft) => draft.insertNoteDirect('footnote', paragraph, run, text, options));
  }

  insertEndnote(paragraph: number, run: number, text: string, options: { customMark?: string } = {}): NoteInfo {
    return this.withDraft((draft) => draft.insertNoteDirect('endnote', paragraph, run, text, options));
  }

  private insertNoteDirect(kind: NoteKind, paragraph: number, run: number, text: string, options: { customMark?: string } = {}): NoteInfo {
    assertIndex(paragraph);
    assertIndex(run);
    assertText(text);
    if (options.customMark !== undefined) assertText(options.customMark, 'customMark');
    this.updatePartXml(this.mainPath, (document) => {
      const paragraphElement = paragraphAt(document, paragraph);
      const runs = ownRuns(paragraphElement);
      if (run > runs.length) throw new Error(`Run ${run} does not exist.`);
    });
    this.ensureNotePart(kind);
    const notePath = this.getNotePartPath(kind);
    if (!notePath) throw new Error(`Unable to resolve ${kind} part path.`);
    const id = this.nextNoteId(kind);
    this.updatePartXml(this.mainPath, (document) => {
      const paragraphElement = paragraphAt(document, paragraph);
      const runs = ownRuns(paragraphElement);
      const referenceRun = wordElement(document, 'r');
      const props = wordElement(document, 'rPr');
      const style = wordElement(document, 'rStyle');
      setWordValue(style, noteReferenceStyle(kind));
      props.appendChild(style);
      const vertAlign = wordElement(document, 'vertAlign');
      setWordValue(vertAlign, 'superscript');
      props.appendChild(vertAlign);
      referenceRun.appendChild(props);
      const reference = wordElement(document, noteReferenceName(kind));
      reference.setAttributeNS(WORD_NS, 'w:id', String(id));
      if (options.customMark) reference.setAttributeNS(WORD_NS, 'w:customMarkFollows', '1');
      referenceRun.appendChild(reference);
      const anchor = runs[run];
      if (anchor?.parentNode) anchor.parentNode.insertBefore(referenceRun, anchor);
      else paragraphElement.insertBefore(referenceRun, null);
    });
    this.updatePartXml(notePath, (document) => {
      const root = document.documentElement!;
      const note = wordElement(document, kind);
      note.setAttributeNS(WORD_NS, 'w:id', String(id));
      const paragraphElement = wordElement(document, 'p');
      if (options.customMark) {
        const markRun = wordElement(document, 'r');
        appendText(markRun, options.customMark);
        paragraphElement.appendChild(markRun);
      }
      const markerRun = wordElement(document, 'r');
      const markerProps = wordElement(document, 'rPr');
      const markerStyle = wordElement(document, 'rStyle');
      setWordValue(markerStyle, noteReferenceStyle(kind));
      markerProps.appendChild(markerStyle);
      markerRun.appendChild(markerProps);
      markerRun.appendChild(wordElement(document, noteRefName(kind)));
      paragraphElement.appendChild(markerRun);
      if (text) {
        const textRun = wordElement(document, 'r');
        appendText(textRun, ` ${text}`);
        paragraphElement.appendChild(textRun);
      }
      note.appendChild(paragraphElement);
      root.appendChild(note);
    });
    const body = bodyOf(this.getPartDocument(this.mainPath));
    return this.getNotesWith(kind, this.collectNoteState(body)).find((item) => item.id === id)!;
  }

  setNoteText(kind: NoteKind, id: number, text: string): void {
    this.withDraft((draft) => draft.setNoteTextDirect(kind, id, text));
  }

  private setNoteTextDirect(kind: NoteKind, id: number, text: string): void {
    assertIndex(id);
    assertText(text);
    const path = this.getNotePartPath(kind);
    if (!path) throw new Error(`${kind} ${id} does not exist.`);
    this.updatePartXml(path, (document) => {
      const note = parseNoteEntries(document, kind).find((item) => item.id === id && item.type === 'normal');
      if (!note) throw new Error(`${kind} ${id} does not exist.`);
      const customMark = parseCustomMark(note.element);
      while (note.element.firstChild) note.element.removeChild(note.element.firstChild);
      const paragraph = wordElement(document, 'p');
      if (customMark) {
        const custom = wordElement(document, 'r');
        appendText(custom, customMark);
        paragraph.appendChild(custom);
      }
      const marker = wordElement(document, 'r');
      const markerProps = wordElement(document, 'rPr');
      const style = wordElement(document, 'rStyle');
      setWordValue(style, noteReferenceStyle(kind));
      markerProps.appendChild(style);
      marker.appendChild(markerProps);
      marker.appendChild(wordElement(document, noteRefName(kind)));
      paragraph.appendChild(marker);
      if (text) {
        const textRun = wordElement(document, 'r');
        appendText(textRun, ` ${text}`);
        paragraph.appendChild(textRun);
      }
      note.element.appendChild(paragraph);
    });
  }

  deleteNote(kind: NoteKind, id: number): void {
    this.withDraft((draft) => draft.deleteNoteDirect(kind, id));
  }

  private deleteNoteDirect(kind: NoteKind, id: number): void {
    assertIndex(id);
    const path = this.getNotePartPath(kind);
    if (path && this.parts.has(path)) {
      this.updatePartXml(path, (document) => {
        const entry = parseNoteEntries(document, kind).find((item) => item.id === id);
        if (entry) entry.element.parentNode?.removeChild(entry.element);
      });
    }
    this.updatePartXml(this.mainPath, (document) => {
      for (const paragraph of descendants(bodyOf(document), 'p')) {
        for (const run of ownRuns(paragraph)) {
          const reference = children(run, noteReferenceName(kind))[0];
          if (reference && Number(reference.getAttributeNS(WORD_NS, 'id')) === id) {
            run.removeChild(reference);
            if (children(run).every((child) => child.localName === 'rPr')) run.parentNode?.removeChild(run);
          }
        }
      }
    });
  }

  convertNote(kind: NoteKind, id: number): void {
    this.withDraft((draft) => draft.convertNoteDirect(kind, id));
  }

  private convertNoteDirect(kind: NoteKind, id: number): void {
    assertIndex(id);
    const targetKind: NoteKind = kind === 'footnote' ? 'endnote' : 'footnote';
    const sourcePath = this.getNotePartPath(kind);
    if (!sourcePath) throw new Error(`${kind} ${id} does not exist.`);
    const sourceDocument = this.getPartDocument(sourcePath);
    const source = parseNoteEntries(sourceDocument, kind).find((item) => item.id === id && item.type === 'normal');
    if (!source) throw new Error(`${kind} ${id} does not exist.`);
    this.ensureNotePart(targetKind);
    const targetPath = this.getNotePartPath(targetKind);
    if (!targetPath) throw new Error(`Unable to resolve ${targetKind} part path.`);
    const targetId = this.nextNoteId(targetKind);
    const targetDocument = this.getPartDocument(targetPath);
    const clone = wordElement(targetDocument, targetKind);
    for (let i = 0; i < source.element.attributes.length; i++) {
      const attribute = source.element.attributes.item(i);
      if (!attribute) continue;
      clone.setAttributeNS(attribute.namespaceURI, attribute.name, attribute.value);
    }
    for (let child = source.element.firstChild; child; child = child.nextSibling) clone.appendChild(targetDocument.importNode(child, true));
    for (const marker of descendants(clone, noteRefName(kind))) {
      const replacement = wordElement(targetDocument, noteRefName(targetKind));
      marker.parentNode?.replaceChild(replacement, marker);
      const markerRun = replacement.parentNode as Element | null;
      const markerProps = markerRun ? children(markerRun, 'rPr')[0] : null;
      if (markerProps) setWordValue(property(markerProps, 'rStyle'), noteReferenceStyle(targetKind));
    }
    clone.setAttributeNS(WORD_NS, 'w:id', String(targetId));
    source.element.parentNode?.removeChild(source.element);
    targetDocument.documentElement!.appendChild(clone);
    this.setPartXml(sourcePath, serializeXml(sourceDocument));
    this.setPartXml(targetPath, serializeXml(targetDocument));
    this.updatePartXml(this.mainPath, (document) => {
      for (const paragraph of descendants(bodyOf(document), 'p')) {
        for (const runElement of ownRuns(paragraph)) {
          const from = children(runElement, noteReferenceName(kind))[0];
          if (!from || Number(from.getAttributeNS(WORD_NS, 'id')) !== id) continue;
          const to = wordElement(document, noteReferenceName(targetKind));
          to.setAttributeNS(WORD_NS, 'w:id', String(targetId));
          if (from.getAttributeNS(WORD_NS, 'customMarkFollows')) to.setAttributeNS(WORD_NS, 'w:customMarkFollows', from.getAttributeNS(WORD_NS, 'customMarkFollows')!);
          runElement.replaceChild(to, from);
          const runProps = children(runElement, 'rPr')[0];
          if (runProps) setWordValue(property(runProps, 'rStyle'), noteReferenceStyle(targetKind));
        }
      }
    });
  }

  private ensurePartRelationship(sourcePartPath: string, relationType: string, targetPath: string): void {
    const relationshipPath = relsPath(sourcePartPath);
    const hasRelationshipPart = this.parts.has(relationshipPath);
    const rels = hasRelationshipPart ? this.getPartDocument(relationshipPath) : parseXml(`<Relationships xmlns="${REL_NS}"/>`);
    const root = rels.documentElement!;
    const target = relativeTarget(sourcePartPath, targetPath);
    const existing = children(root, 'Relationship', REL_NS)
      .find((relation) => relation.getAttribute('Type') === relationType && relation.getAttribute('TargetMode') !== 'External');
    if (existing) {
      if (existing.getAttribute('Target') === target) return;
      existing.setAttribute('Target', target);
    } else {
      const relationship = rels.createElementNS(REL_NS, 'Relationship');
      relationship.setAttribute('Id', nextRelationshipId(root));
      relationship.setAttribute('Type', relationType);
      relationship.setAttribute('Target', target);
      root.appendChild(relationship);
    }
    if (hasRelationshipPart) this.setPartXml(relationshipPath, serializeXml(rels));
    else this.addPart(relationshipPath, encodeXml(serializeXml(rels)), RELS_CONTENT_TYPE);
  }

  private ensureCommentsParts(sourcePartPath: string, includeExtended = false): { commentsPath: string; commentsExtendedPath?: string } {
    const sourceBase = basename(sourcePartPath).replace(/\.xml$/i, '');
    let commentsPath = this.getCommentsPartPath(sourcePartPath);
    if (!commentsPath) {
      // Keep new comment parts next to the source part; non-main fallback names stay source-specific
      // so future cleanup/interoperability logic can reconstruct the relationship target deterministically.
      commentsPath = sourcePartPath === this.mainPath
        ? `${dirname(sourcePartPath) ? `${dirname(sourcePartPath)}/` : ''}comments.xml`
        : `${dirname(sourcePartPath) ? `${dirname(sourcePartPath)}/` : ''}${sourceBase}-comments.xml`;
    }
    if (!this.parts.has(commentsPath)) this.addPart(commentsPath, encodeXml(defaultCommentsXml()), COMMENTS_TYPE);
    this.ensurePartRelationship(sourcePartPath, COMMENTS_REL, commentsPath);
    let commentsExtendedPath: string | undefined;
    if (includeExtended) {
      commentsExtendedPath = this.getCommentsExtendedPartPath(sourcePartPath);
      if (!commentsExtendedPath) {
        commentsExtendedPath = sourcePartPath === this.mainPath
          ? `${dirname(sourcePartPath) ? `${dirname(sourcePartPath)}/` : ''}commentsExtended.xml`
          : `${dirname(sourcePartPath) ? `${dirname(sourcePartPath)}/` : ''}${sourceBase}-commentsExtended.xml`;
      }
      if (!this.parts.has(commentsExtendedPath)) {
        this.addPart(commentsExtendedPath, encodeXml(defaultCommentsExtendedXml()), COMMENTS_EXTENDED_TYPE);
      }
      this.ensurePartRelationship(sourcePartPath, COMMENTS_EXTENDED_REL, commentsExtendedPath);
    }
    return { commentsPath, commentsExtendedPath };
  }

  private nextCommentId(): number {
    const max = this.getAllComments().reduce((current, comment) => Math.max(current, comment.id), 0);
    return max + 1;
  }

  private nextCommentParaId(): string {
    const used = new Set<number>();
    for (const binding of this.commentBindings()) {
      if (binding.commentsPath && this.parts.has(binding.commentsPath)) {
        for (const entry of parseCommentEntries(this.partDocumentOrUndefined(binding.commentsPath) ?? null)) {
          if (entry.paraId && /^[0-9a-f]{8}$/i.test(entry.paraId)) used.add(Number.parseInt(entry.paraId, 16));
        }
      }
      if (binding.commentsExtendedPath && this.parts.has(binding.commentsExtendedPath)) {
        for (const entry of parseCommentExEntries(this.partDocumentOrUndefined(binding.commentsExtendedPath) ?? null)) {
          used.add(Number.parseInt(entry.paraId, 16));
          if (entry.paraIdParent) used.add(Number.parseInt(entry.paraIdParent, 16));
        }
      }
    }
    let next = 1;
    while (used.has(next)) next++;
    return next.toString(16).toUpperCase().padStart(8, '0');
  }

  private locateComment(id: number): { sourcePartPath: string; commentsPath: string; commentsExtendedPath?: string; entry: ReturnType<typeof parseCommentEntries>[number] } {
    assertIndex(id);
    const parsed = new Set<string>();
    for (const binding of this.commentBindings()) {
      if (!binding.commentsPath || parsed.has(binding.commentsPath) || !this.parts.has(binding.commentsPath)) continue;
      parsed.add(binding.commentsPath);
      const entry = parseCommentEntries(this.partDocumentOrUndefined(binding.commentsPath) ?? null).find((item) => item.id === id);
      if (entry) return { sourcePartPath: binding.sourcePartPath, commentsPath: binding.commentsPath, commentsExtendedPath: binding.commentsExtendedPath, entry };
    }
    throw new Error(`Comment ${id} does not exist.`);
  }

  private ensureCommentParaId(commentsPath: string, id: number): string {
    let resolved = '';
    this.updatePartXml(commentsPath, (document) => {
      const entry = parseCommentEntries(document).find((item) => item.id === id);
      if (!entry) throw new Error(`Comment ${id} does not exist.`);
      if (entry.paraId) {
        resolved = entry.paraId.toUpperCase();
        return;
      }
      const paragraphs = children(entry.element, 'p');
      const paragraph = paragraphs.at(-1) ?? wordElement(document, 'p');
      if (!paragraph.parentNode) entry.element.appendChild(paragraph);
      resolved = this.nextCommentParaId();
      ensureCommentParagraphParaId(paragraph, resolved);
    });
    return resolved;
  }

  private writeCommentBody(commentElement: Element, paraId: string, text: string): void {
    const document = commentElement.ownerDocument!;
    while (commentElement.firstChild) commentElement.removeChild(commentElement.firstChild);
    const paragraph = wordElement(document, 'p');
    ensureCommentParagraphParaId(paragraph, paraId);
    const paragraphProps = properties(paragraph, 'pPr');
    setWordValue(property(paragraphProps, 'pStyle'), commentParagraphStyle());
    paragraph.appendChild(makeCommentAnnotationRun(document));
    if (text) {
      const run = wordElement(document, 'r');
      const props = properties(run, 'rPr');
      setWordValue(property(props, 'rStyle'), commentParagraphStyle());
      appendText(run, text);
      paragraph.appendChild(run);
    }
    commentElement.appendChild(paragraph);
  }

  private upsertCommentEx(commentsExtendedPath: string, paraId: string, patch: { parentParaId?: string | null; done?: boolean }): void {
    this.updatePartXml(commentsExtendedPath, (document) => {
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

  private boundaryRun(paragraph: Element, offset: number): Element | null {
    const runs = ownRuns(paragraph);
    let cursor = 0;
    for (const run of runs) {
      const text = textOf(run);
      const next = cursor + text.length;
      if (offset <= cursor) return run;
      cursor = next;
    }
    return null;
  }

  private removeRelationshipTarget(sourcePartPath: string, relationType: string, targetPath: string): void {
    const relationshipPath = relsPath(sourcePartPath);
    if (!this.parts.has(relationshipPath)) return;
    this.updatePartXml(relationshipPath, (document) => {
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

  private relationshipCount(targetPath: string, relationType: string): number {
    let count = 0;
    for (const sourcePartPath of this.contentPartPaths()) {
      const relationshipPath = relsPath(sourcePartPath);
      if (!this.parts.has(relationshipPath)) continue;
      const document = this.partDocumentOrUndefined(relationshipPath)?.documentElement;
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

  private removePartOverride(path: string): void {
    this.updatePartXml('[Content_Types].xml', (document) => {
      for (const override of children(document.documentElement!, 'Override', CONTENT_TYPES_NS)) {
        if (override.getAttribute('PartName') === `/${path}`) override.parentNode?.removeChild(override);
      }
    });
  }

  private removePartAndOverride(path: string): void {
    if (!this.parts.has(path)) return;
    this.parts.delete(path);
    this.documents.delete(path);
    this.dirtyPartXml.delete(path);
    this.dirtyPartSizes.delete(path);
    this.removePartOverride(path);
  }

  private cleanupCommentParts(commentsPath: string, commentsExtendedPath?: string): void {
    const hasComments = this.parts.has(commentsPath) && parseCommentEntries(this.partDocumentOrUndefined(commentsPath) ?? null).length > 0;
    if (!hasComments) {
      for (const binding of this.commentBindings()) {
        this.removeRelationshipTarget(binding.sourcePartPath, COMMENTS_REL, commentsPath);
      }
      if (this.relationshipCount(commentsPath, COMMENTS_REL) === 0) this.removePartAndOverride(commentsPath);
    }
    if (commentsExtendedPath && this.parts.has(commentsExtendedPath) && parseCommentExEntries(this.partDocumentOrUndefined(commentsExtendedPath) ?? null).length === 0) {
      for (const binding of this.commentBindings()) {
        this.removeRelationshipTarget(binding.sourcePartPath, COMMENTS_EXTENDED_REL, commentsExtendedPath);
      }
      if (this.relationshipCount(commentsExtendedPath, COMMENTS_EXTENDED_REL) === 0) this.removePartAndOverride(commentsExtendedPath);
    }
  }

  addComment(range: TextRange | DocumentRange, comment: { author?: string; initials?: string; text: string }): number {
    // TextRange/DocumentRange use the main-document paragraph namespace in this API surface.
    // Header/footer/footnote/endnote comment anchors are readable via getComments(), but write APIs
    // currently insert anchors only into the main document part.
    return this.withDraft((draft) => draft.addCommentDirect(range, comment));
  }

  private addCommentDirect(range: TextRange | DocumentRange, comment: { author?: string; initials?: string; text: string }): number {
    if (comment.author !== undefined) assertText(comment.author, 'comment.author');
    if (comment.initials !== undefined) assertText(comment.initials, 'comment.initials');
    assertText(comment.text, 'comment.text');
    const preview = this.getPartDocument(this.mainPath);
    const normalized = 'paragraph' in range
      ? (() => {
          const current = this.normalizeRangeOn(preview, range);
          return {
            startParagraph: range.paragraph,
            startOffset: current.start,
            endParagraph: range.paragraph,
            endOffset: current.end,
          };
        })()
      : (() => {
          const current = this.normalizeDocumentRange(preview, range);
          return {
            startParagraph: current.start.paragraph,
            startOffset: current.start.offset,
            endParagraph: current.end.paragraph,
            endOffset: current.end.offset,
          };
        })();
    const id = this.nextCommentId();
    const paraId = this.nextCommentParaId();
    const date = new Date().toISOString();
    this.updatePartXml(this.mainPath, (document) => {
      const startParagraph = paragraphAt(document, normalized.startParagraph);
      const endParagraph = paragraphAt(document, normalized.endParagraph);
      const collapsed = normalized.startParagraph === normalized.endParagraph && normalized.startOffset === normalized.endOffset;
      if (startParagraph === endParagraph) {
        this.splitRunAtOffset(startParagraph, normalized.endOffset);
        this.splitRunAtOffset(startParagraph, normalized.startOffset);
      } else {
        this.splitRunAtOffset(endParagraph, normalized.endOffset);
        this.splitRunAtOffset(startParagraph, normalized.startOffset);
      }
      const startMarker = wordElement(document, 'commentRangeStart');
      startMarker.setAttributeNS(WORD_NS, 'w:id', String(id));
      const startBoundary = this.boundaryRun(startParagraph, normalized.startOffset);
      if (startBoundary?.parentNode) startBoundary.parentNode.insertBefore(startMarker, startBoundary);
      else startParagraph.insertBefore(startMarker, startParagraph.firstChild);

      const endMarker = wordElement(document, 'commentRangeEnd');
      endMarker.setAttributeNS(WORD_NS, 'w:id', String(id));
      const referenceRun = wordElement(document, 'r');
      const props = properties(referenceRun, 'rPr');
      setWordValue(property(props, 'rStyle'), commentReferenceStyle());
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
        const endBoundary = this.boundaryRun(endParagraph, normalized.endOffset);
        if (endBoundary?.parentNode) endBoundary.parentNode.insertBefore(endMarker, endBoundary);
        else endParagraph.appendChild(endMarker);
        if (endBoundary?.parentNode) endBoundary.parentNode.insertBefore(referenceRun, endBoundary);
        else endParagraph.appendChild(referenceRun);
      }
    });
    const { commentsPath, commentsExtendedPath } = this.ensureCommentsParts(this.mainPath, true);
    this.updatePartXml(commentsPath, (document) => {
      const root = document.documentElement!;
      const commentElement = wordElement(document, 'comment');
      commentElement.setAttributeNS(WORD_NS, 'w:id', String(id));
      if (comment.author) commentElement.setAttributeNS(WORD_NS, 'w:author', comment.author);
      if (comment.initials) commentElement.setAttributeNS(WORD_NS, 'w:initials', comment.initials);
      commentElement.setAttributeNS(WORD_NS, 'w:date', date);
      this.writeCommentBody(commentElement, paraId, comment.text);
      root.appendChild(commentElement);
    });
    this.upsertCommentEx(commentsExtendedPath!, paraId, { done: false, parentParaId: null });
    return id;
  }

  replyComment(parentId: number, comment: { author?: string; initials?: string; text: string }): number {
    return this.withDraft((draft) => draft.replyCommentDirect(parentId, comment));
  }

  private replyCommentDirect(parentId: number, comment: { author?: string; initials?: string; text: string }): number {
    if (comment.author !== undefined) assertText(comment.author, 'comment.author');
    if (comment.initials !== undefined) assertText(comment.initials, 'comment.initials');
    assertText(comment.text, 'comment.text');
    const parent = this.locateComment(parentId);
    const parentParaId = this.ensureCommentParaId(parent.commentsPath, parentId);
    const id = this.nextCommentId();
    const paraId = this.nextCommentParaId();
    const { commentsPath, commentsExtendedPath } = this.ensureCommentsParts(parent.sourcePartPath, true);
    this.updatePartXml(commentsPath, (document) => {
      const root = document.documentElement!;
      const commentElement = wordElement(document, 'comment');
      commentElement.setAttributeNS(WORD_NS, 'w:id', String(id));
      if (comment.author) commentElement.setAttributeNS(WORD_NS, 'w:author', comment.author);
      if (comment.initials) commentElement.setAttributeNS(WORD_NS, 'w:initials', comment.initials);
      commentElement.setAttributeNS(WORD_NS, 'w:date', new Date().toISOString());
      this.writeCommentBody(commentElement, paraId, comment.text);
      root.appendChild(commentElement);
    });
    this.upsertCommentEx(commentsExtendedPath!, paraId, { parentParaId, done: false });
    return id;
  }

  setCommentResolved(id: number, resolved: boolean): void {
    this.withDraft((draft) => draft.setCommentResolvedDirect(id, resolved));
  }

  private setCommentResolvedDirect(id: number, resolved: boolean): void {
    assertIndex(id);
    if (typeof resolved !== 'boolean') throw new Error('resolved must be boolean.');
    const target = this.locateComment(id);
    const paraId = this.ensureCommentParaId(target.commentsPath, id);
    const { commentsExtendedPath } = this.ensureCommentsParts(target.sourcePartPath, true);
    this.upsertCommentEx(commentsExtendedPath!, paraId, { done: resolved });
  }

  setCommentText(id: number, text: string): void {
    this.withDraft((draft) => draft.setCommentTextDirect(id, text));
  }

  private setCommentTextDirect(id: number, text: string): void {
    assertIndex(id);
    assertText(text, 'text');
    const target = this.locateComment(id);
    const paraId = this.ensureCommentParaId(target.commentsPath, id);
    this.updatePartXml(target.commentsPath, (document) => {
      const entry = parseCommentEntries(document).find((item) => item.id === id);
      if (!entry) throw new Error(`Comment ${id} does not exist.`);
      this.writeCommentBody(entry.element, paraId, text);
    });
  }

  deleteComment(id: number, options: { withReplies?: boolean } = {}): void {
    this.withDraft((draft) => draft.deleteCommentDirect(id, options));
  }

  private deleteCommentDirect(id: number, options: { withReplies?: boolean } = {}): void {
    assertIndex(id);
    const withReplies = options.withReplies ?? true;
    const allComments = this.getAllComments();
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
      const target = this.locateComment(commentId);
      const paraId = this.ensureCommentParaId(target.commentsPath, commentId);
      const list = byCommentsPath.get(target.commentsPath) ?? [];
      list.push(commentId);
      byCommentsPath.set(target.commentsPath, list);
      paraIds.set(commentId, { commentsPath: target.commentsPath, commentsExtendedPath: target.commentsExtendedPath, paraId, sourcePartPath: target.sourcePartPath });
    }
    for (const [commentsPath, ids] of byCommentsPath) {
      this.updatePartXml(commentsPath, (document) => {
        for (const entry of parseCommentEntries(document)) {
          if (ids.includes(entry.id)) entry.element.parentNode?.removeChild(entry.element);
        }
      });
    }
    const extendedPaths = new Set<string>();
    for (const info of paraIds.values()) {
      if (!info.commentsExtendedPath || extendedPaths.has(info.commentsExtendedPath) || !this.parts.has(info.commentsExtendedPath)) continue;
      extendedPaths.add(info.commentsExtendedPath);
      const deletedParaIds = new Set([...paraIds.values()]
        .filter((entry) => entry.commentsExtendedPath === info.commentsExtendedPath)
        .map((entry) => entry.paraId));
      this.updatePartXml(info.commentsExtendedPath, (document) => {
        for (const entry of parseCommentExEntries(document)) {
          if (deletedParaIds.has(entry.paraId)) entry.element.parentNode?.removeChild(entry.element);
        }
      });
    }
    const affectedSourceParts = new Set([...paraIds.values()].map((info) => info.sourcePartPath));
    for (const sourcePartPath of affectedSourceParts) {
      this.updatePartXml(sourcePartPath, (document) => {
        const container = blockContainerOf(document);
        for (const nodeName of ['commentRangeStart', 'commentRangeEnd'] as const) {
          for (const node of descendants(container, nodeName)) {
            const commentId = Number(node.getAttributeNS(WORD_NS, 'id') ?? node.getAttribute('w:id'));
            if (deleteIds.has(commentId)) node.parentNode?.removeChild(node);
          }
        }
        for (const paragraph of descendants(container, 'p')) {
          for (const run of ownRuns(paragraph)) {
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
    for (const info of paraIds.values()) this.cleanupCommentParts(info.commentsPath, info.commentsExtendedPath);
  }

  /** All operations succeed together, or the original package/revision is unchanged. */
  applyOperations(request: AgentRequest): DocumentSnapshot {
    validateRequest(request);
    if (request.expectedRevision !== undefined && request.expectedRevision !== this.revision) {
      throw new Error(`Revision conflict: expected ${request.expectedRevision}, current ${this.revision}.`);
    }
    if (!request.operations.length) return this.getSnapshot();
    const historyOnlyBatch = request.operations.every((operation) => operation.type === 'undo' || operation.type === 'redo');
    if (historyOnlyBatch) {
      for (const operation of request.operations) {
        if (operation.type === 'undo') this.undo();
        else this.redo();
      }
      return this.getSnapshot();
    }
    this.materializeAllParts();
    const draft = new DocxDocument(new Map(this.parts));
    draft.currentRevision = this.currentRevision;
    draft.revisionAuthor = this.revisionAuthor;
    draft.undoHistory = this.undoHistory.map((state) => ({
      parts: draft.cloneParts(state.parts),
      bytes: state.bytes,
      entry: { ...state.entry },
      action: state.action,
    }));
    draft.redoHistory = this.redoHistory.map((state) => ({
      parts: draft.cloneParts(state.parts),
      bytes: state.bytes,
      entry: { ...state.entry },
      action: state.action,
    }));
    draft.undoHistoryBytes = this.undoHistoryBytes;
    draft.redoHistoryBytes = this.redoHistoryBytes;
    draft.suppressHistory = true;
    const history = historyOnlyBatch ? undefined : this.beginMutationHistory(this.nextHistoryLabel, { kind: 'transaction' });
    this.nextHistoryLabel = undefined;
    this.nextHistoryAction = { kind: 'other' };
    try {
      for (const operation of request.operations) {
        switch (operation.type) {
        case 'setTrackChanges': draft.setTrackChanges(operation.enabled); break;
        case 'setRevisionAuthor': draft.setRevisionAuthor(operation.author); break;
        case 'acceptRevision': draft.acceptRevision(operation.id); break;
        case 'rejectRevision': draft.rejectRevision(operation.id); break;
        case 'acceptAllRevisions': draft.acceptAllRevisions(operation.filter); break;
        case 'rejectAllRevisions': draft.rejectAllRevisions(operation.filter); break;
        case 'setParagraphText': draft.setParagraphText(operation.index, operation.text); break;
        case 'insertParagraph': draft.insertParagraph(operation.text, operation.before); break;
        case 'deleteParagraph': draft.deleteParagraph(operation.index); break;
        case 'formatParagraph': draft.formatParagraph(operation.index, operation.format); break;
        case 'applyParagraphStyle': draft.applyParagraphStyle(operation.index, operation.styleId, operation.options); break;
        case 'setParagraphNumbering': draft.setParagraphNumbering(operation.index, operation.numId, operation.level); break;
        case 'clearParagraphNumbering': draft.clearParagraphNumbering(operation.index); break;
        case 'setParagraphLevel': draft.setParagraphLevel(operation.index, operation.delta); break;
        case 'formatRun': draft.formatRun(operation.paragraph, operation.run, operation.format); break;
        case 'formatRange': draft.formatRange(operation.range, operation.format); break;
        case 'applyCharacterStyle': draft.applyCharacterStyle(operation.range, operation.styleId, operation.options); break;
        case 'clearRangeFormat': draft.clearRangeFormat(operation.range, operation.fields); break;
        case 'formatDocumentRange': draft.formatDocumentRange(operation.range, operation.format); break;
        case 'setOutlineLevel': draft.setOutlineLevel(operation.index, operation.level); break;
        case 'moveOutlineSection': draft.moveOutlineSection(operation.from, operation.to); break;
        case 'setParagraphTabs': draft.setParagraphTabs(operation.index, operation.tabs); break;
        case 'setParagraphBorders': draft.setParagraphBorders(operation.index, operation.borders); break;
        case 'setParagraphShading': draft.setParagraphShading(operation.index, operation.shading); break;
        case 'insertBreak': draft.insertBreak(operation.paragraph, operation.run, operation.breakType); break;
        case 'insertSymbol': draft.insertSymbol(operation.paragraph, operation.run, operation.font, operation.charCode); break;
        case 'replaceText': draft.replaceText(operation.search, operation.replacement); break;
        case 'insertTable': draft.insertTable(operation.rows); break;
        case 'insertTableAt': draft.insertTableAt(operation.rows, operation.cols, operation.before, operation.format); break;
        case 'insertTableRow': draft.insertTableRow(operation.table, operation.at); break;
        case 'deleteTableRow': draft.deleteTableRow(operation.table, operation.at); break;
        case 'insertTableColumn': draft.insertTableColumn(operation.table, operation.at); break;
        case 'deleteTableColumn': draft.deleteTableColumn(operation.table, operation.at); break;
        case 'mergeCells': draft.mergeCells(operation.table, operation.range); break;
        case 'splitCell': draft.splitCell(operation.table, operation.row, operation.col, operation.rows, operation.cols); break;
        case 'formatTable': draft.formatTable(operation.table, operation.format); break;
        case 'formatTableRow': draft.formatTableRow(operation.table, operation.row, operation.format); break;
        case 'formatCell': draft.formatCell(operation.table, operation.row, operation.col, operation.format); break;
        case 'setCellText': draft.setCellText(operation.table, operation.row, operation.col, operation.text); break;
        case 'insertHyperlink': draft.insertHyperlink(operation.target, operation.link); break;
        case 'updateHyperlink': draft.updateHyperlink(operation.hyperlink, operation.link); break;
        case 'removeHyperlink': draft.removeHyperlink(operation.hyperlink, operation.options); break;
        case 'insertBookmark': draft.insertBookmark(operation.name, operation.range); break;
        case 'deleteBookmark': draft.deleteBookmark(operation.name); break;
        case 'insertImage': draft.insertImage({
          bytes: decodeBase64(operation.bytes),
          contentType: operation.contentType,
          paragraph: operation.paragraph,
          run: operation.run,
          widthEmu: operation.widthEmu,
          heightEmu: operation.heightEmu,
          alt: operation.alt,
          placement: operation.placement,
        }); break;
        case 'replaceImageBytes': draft.replaceImageBytes(operation.image, decodeBase64(operation.bytes), operation.contentType); break;
        case 'resizeImage': draft.resizeImage(operation.image, operation.size); break;
        case 'setImageAlt': draft.setImageAlt(operation.image, operation.alt, operation.title); break;
        case 'deleteImage': draft.deleteImage(operation.image); break;
        case 'setPartXml': draft.setPartXml(operation.path, operation.xml); break;
        case 'insertFootnote': draft.insertFootnote(operation.paragraph, operation.run, operation.text, { customMark: operation.customMark }); break;
        case 'insertEndnote': draft.insertEndnote(operation.paragraph, operation.run, operation.text, { customMark: operation.customMark }); break;
        case 'setNoteText': draft.setNoteText(operation.kind, operation.id, operation.text); break;
        case 'deleteNote': draft.deleteNote(operation.kind, operation.id); break;
        case 'convertNote': draft.convertNote(operation.kind, operation.id); break;
        case 'addComment': draft.addComment(operation.range, operation.comment); break;
        case 'replyComment': draft.replyComment(operation.parentId, operation.comment); break;
        case 'setCommentResolved': draft.setCommentResolved(operation.id, operation.resolved); break;
        case 'setCommentText': draft.setCommentText(operation.id, operation.text); break;
        case 'deleteComment': draft.deleteComment(operation.id, operation.options); break;
        case 'undo': draft.undo(); break;
        case 'redo': draft.redo(); break;
        }
      }
      draft.materializeAllParts();
      if (equalPartMap(this.parts, draft.parts)) {
        this.pendingMergedHistory = undefined;
        return this.getSnapshot();
      }
      const snapshot = draft.getSnapshot();
      this.parts = draft.parts;
      this.documents = draft.documents;
      this.dirtyPartXml = draft.dirtyPartXml;
      this.dirtyPartSizes = draft.dirtyPartSizes;
      this.mainPath = draft.mainPath;
      this.revisionAuthor = draft.revisionAuthor;
      this.currentRevision++;
      this.undoHistory = draft.undoHistory;
      this.redoHistory = draft.redoHistory;
      this.undoHistoryBytes = draft.undoHistoryBytes;
      this.redoHistoryBytes = draft.redoHistoryBytes;
      this.numberingContextCache = undefined;
      this.stylesCache = undefined;
      this.outlineCache = undefined;
      this.noteStateCache = undefined;
      this.commentStateCache = undefined;
      this.revisionInfoCache = undefined;
      this.reviewerInfoCache = undefined;
      this.tableCellLocationCache = undefined;
      this.imageDataUrls.clear();
      if (history) this.recordHistory(history);
      return { ...snapshot, revision: this.revision };
    } catch (error) {
      this.abortHistoryGroupOnFailure();
      throw error;
    }
  }

  async toUint8Array(): Promise<Uint8Array> {
    this.materializeAllParts();
    const zip = new JSZip();
    for (const [path, bytes] of this.parts) zip.file(path, bytes, { createFolders: false });
    return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  }

  async toBlob(): Promise<Blob> {
    const bytes = await this.toUint8Array();
    return new Blob([bytes as Uint8Array<ArrayBuffer>], { type: DOCX_TYPE });
  }
}
