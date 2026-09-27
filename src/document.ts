import JSZip from 'jszip';
import type { Document, Element, Node } from '@xmldom/xmldom';
import type {
  AgentRequest, DocumentBlock, DocumentSnapshot, PageSetup, ParagraphFormat, ParagraphInfo, RunFormat, RunInfo,
  SectionInfo, SectionType,
} from './types.js';
import {
  assertText, children, CONTENT_TYPES_NS, descendants, OFFICE_DOCUMENT_REL, parseXml, REL_NS,
  serializeXml, setWordValue, validatePath, WORD_NS, wordElement, wordValue,
} from './xml.js';
import { assertIndex, validateParagraphFormat, validateRequest, validateRows, validateRunFormat } from './operations.js';
import { collectSections, readSections, SECTION_ORDER } from './section.js';

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
const encoder = new TextEncoder();

function decodeXml(bytes: Uint8Array): string {
  const utf16le = (bytes[0] === 0xff && bytes[1] === 0xfe) || (bytes[0] === 0x3c && bytes[1] === 0);
  const utf16be = (bytes[0] === 0xfe && bytes[1] === 0xff) || (bytes[0] === 0 && bytes[1] === 0x3c);
  return new TextDecoder(utf16le ? 'utf-16le' : utf16be ? 'utf-16be' : 'utf-8', { fatal: true }).decode(bytes);
}

function encodeXml(xml: string): Uint8Array {
  return encoder.encode(xml.replace(/^(<\?xml\b[^?]*\bencoding\s*=\s*)(["'])[^"']*\2/i, '$1"UTF-8"'));
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
  if (['hdr', 'ftr'].includes(root.localName ?? '')) return root;
  throw new Error('Part does not contain block-level WordprocessingML content.');
}

function textElements(element: Element): Element[] {
  const result: Element[] = [];
  function walk(node: Node): void {
    for (let child = node.firstChild; child; child = child.nextSibling) {
      if (child.nodeType !== 1) continue;
      const element = child as Element;
      if (element.namespaceURI === WORD_NS) {
        if (element.localName === 'p') continue;
        if (['t', 'tab', 'br', 'cr'].includes(element.localName ?? '')) {
          result.push(element);
          continue;
        }
      }
      walk(element);
    }
  }
  walk(element);
  return result;
}

function elementText(element: Element): string {
  if (element.localName === 't') return element.textContent ?? '';
  if (element.localName === 'tab') return '\t';
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
  const text = textElements(element).map(elementText).join('');
  if (text) return text;
  return fieldPlaceholder(element) ?? '';
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
};

function property(parent: Element, name: string): Element {
  let result = children(parent, name)[0];
  if (!result) {
    result = wordElement(parent.ownerDocument!, name);
    const order = PROPERTY_ORDER[parent.localName as keyof typeof PROPERTY_ORDER];
    const position = order.indexOf(name);
    const following = children(parent).find(child => order.indexOf(child.localName!) > position);
    parent.insertBefore(result, following ?? null);
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

function readRun(run: Element, index: number): RunInfo {
  const props = children(run, 'rPr')[0];
  const get = (name: string) => props ? children(props, name)[0] : undefined;
  const toggle = (name: string) => get(name) ? !['0', 'false', 'off'].includes(wordValue(get(name)) ?? '') : undefined;
  const size = wordValue(get('sz'));
  const underline = get('u');
  const color = wordValue(get('color'));
  return {
    index, text: textOf(run), bold: toggle('b'), italic: toggle('i'),
    underline: underline ? !['none', '0', 'false'].includes(wordValue(underline) ?? '') : undefined,
    fontSize: size && Number.isFinite(Number(size)) ? Number(size) / 2 : undefined,
    fontFamily: get('rFonts')?.getAttributeNS(WORD_NS, 'ascii') ?? undefined,
    color: color && /^[a-f\d]{6}$/i.test(color) ? color : undefined,
  };
}

function readParagraph(paragraph: Element, index: number): ParagraphInfo {
  const props = children(paragraph, 'pPr')[0];
  const alignment = props ? wordValue(children(props, 'jc')[0]) : undefined;
  return {
    index, text: textOf(paragraph), runs: ownRuns(paragraph).map(readRun),
    style: props ? wordValue(children(props, 'pStyle')[0]) : undefined,
    alignment: ['left', 'center', 'right', 'both'].includes(alignment ?? '')
      ? alignment as ParagraphFormat['alignment'] : undefined,
  };
}

async function readEntry(entry: JSZip.JSZipObject, limit: number): Promise<Uint8Array> {
  return new Promise((resolve, reject) => {
    const chunks: Uint8Array[] = [];
    let size = 0;
    // JSZip exposes internalStream publicly, but omits it from JSZipObject's typings.
    const stream = (entry as JSZip.JSZipObject & {
      internalStream(type: 'uint8array'): JSZip.JSZipStreamHelper<Uint8Array>;
    }).internalStream('uint8array');
    stream.on('data', chunk => {
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

export class DocxDocument {
  private parts: Map<string, Uint8Array>;
  private mainPath: string;
  private currentRevision = 0;

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
    const entries = Object.values(zip.files);
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

  get revision(): number { return this.currentRevision; }
  get mainDocumentPath(): string { return this.mainPath; }

  listParts(): string[] { return [...this.parts.keys()].sort(); }

  getPartBytes(path: string): Uint8Array {
    validatePath(path);
    const bytes = this.parts.get(path);
    if (!bytes) throw new Error(`Package part not found: ${path}`);
    return Uint8Array.from(bytes);
  }

  getPartXml(path: string): string { return decodeXml(this.getPartBytes(path)); }

  /** Returns a detached DOM; use updatePartXml to persist changes. */
  getPartDocument(path: string): Document { return parseXml(this.getPartXml(path)); }

  updatePartXml(path: string, update: (document: Document) => void): void {
    const document = this.getPartDocument(path);
    update(document);
    this.setPartXml(path, serializeXml(document));
  }

  setPartXml(path: string, xml: string): void {
    if (typeof xml !== 'string' || xml.length > MAX_PART) throw new Error('XML part exceeds size limit.');
    parseXml(xml);
    this.setPartBytes(path, encodeXml(xml));
  }

  /** Replaces an existing part. Relationships/content types remain under caller control. */
  setPartBytes(path: string, bytes: Uint8Array): void {
    validatePath(path);
    if (!this.parts.has(path)) throw new Error('Use addPart with a content type to create a new part.');
    const next = new Map(this.parts);
    next.set(path, Uint8Array.from(bytes));
    this.commitParts(next);
  }

  addPart(path: string, bytes: Uint8Array, contentType: string): void {
    validatePath(path);
    assertText(contentType, 'contentType');
    if (!contentType || this.parts.has(path)) throw new Error('New part requires a unique path and content type.');
    const types = this.getPartDocument('[Content_Types].xml');
    const override = types.createElementNS(CONTENT_TYPES_NS, 'Override');
    override.setAttribute('PartName', `/${path}`);
    override.setAttribute('ContentType', contentType);
    types.documentElement!.appendChild(override);
    const next = new Map(this.parts);
    next.set('[Content_Types].xml', encodeXml(serializeXml(types)));
    next.set(path, Uint8Array.from(bytes));
    this.commitParts(next);
  }

  private commitParts(parts: Map<string, Uint8Array>): void {
    const draft = new DocxDocument(parts);
    this.parts = draft.parts;
    this.mainPath = draft.mainPath;
    this.currentRevision++;
  }

  private validatePackage(): string {
    if (this.parts.size > MAX_PARTS) throw new Error('Too many package parts.');
    let total = 0;
    for (const [path, bytes] of this.parts) {
      validatePath(path);
      total += bytes.byteLength;
      if (bytes.byteLength > MAX_PART || total > MAX_TOTAL) throw new Error('Package exceeds size limits.');
    }
    const rels = this.getPartDocument('_rels/.rels').documentElement;
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
    const types = this.getPartDocument('[Content_Types].xml').documentElement;
    if (types?.namespaceURI !== CONTENT_TYPES_NS || types.localName !== 'Types') {
      throw new Error('Invalid content types part.');
    }
    const overrides = children(types, 'Override', CONTENT_TYPES_NS);
    const type = overrides.find(type => type.getAttribute('PartName') === `/${path}`)?.getAttribute('ContentType')
      ?? children(types, 'Default', CONTENT_TYPES_NS)
        .find(type => type.getAttribute('Extension') === path.split('.').pop())?.getAttribute('ContentType');
    if (type !== MAIN_TYPE) throw new Error('Package is not a supported .docx document (macros/strict OOXML are not supported).');
    bodyOf(this.getPartDocument(path));
    return path;
  }

  getParagraphs(): ParagraphInfo[] {
    return descendants(bodyOf(this.getPartDocument(this.mainPath)), 'p').map(readParagraph);
  }

  getBlocks(): DocumentBlock[] {
    const main = this.getPartDocument(this.mainPath);
    const body = bodyOf(main);
    const sections = collectSections(main);
    const sectionByParagraph = new Map<number, SectionType>(sections
      .filter(section => section.source === 'paragraph')
      .map(section => [section.endParagraph, (children(section.sectPr, 'type')[0]?.getAttributeNS(WORD_NS, 'val') ?? 'nextPage') as SectionType]));
    const paragraphs = descendants(body, 'p');
    const indices = new Map(paragraphs.map((p, i) => [p, i]));
    const pageBreakMarkers = (paragraph: Element): { before: number; after: number } => {
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
    };
    const walk = (parent: Element): DocumentBlock[] => children(parent).flatMap((child): DocumentBlock[] => {
      if (child.localName === 'p') return [{ type: 'paragraph', paragraph: readParagraph(child, indices.get(child)!) }];
      if (child.localName === 'tbl') return [{
        type: 'table',
        rows: children(child, 'tr').map(row => ({
          cells: children(row, 'tc').map(cell => ({ blocks: walk(cell) })),
        })),
      }];
      if (['sdt', 'sdtContent', 'customXml'].includes(child.localName ?? '')) return walk(child);
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
      const paragraph = paragraphs[block.paragraph.index];
      const markers = paragraph ? pageBreakMarkers(paragraph) : { before: 0, after: 0 };
      for (let i = 0; i < markers.before; i++) result.push({ type: 'pageBreak' });
      result.push(block);
      for (let i = 0; i < markers.after; i++) result.push({ type: 'pageBreak' });
      const breakType = sectionByParagraph.get(block.paragraph.index);
      if (breakType) result.push({ type: 'sectionBreak', section: sectionBreaks++, breakType });
    }
    return result;
  }

  getSnapshot(): DocumentSnapshot {
    return { revision: this.revision, paragraphs: this.getParagraphs(), blocks: this.getBlocks(), parts: this.listParts() };
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

  private nextRelationshipId(path: string): string {
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
    const id = this.nextRelationshipId(path);
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
        if (setup.pageWidth !== undefined) size.setAttributeNS(WORD_NS, 'w:w', String(Math.max(0, Math.trunc(setup.pageWidth))));
        if (setup.pageHeight !== undefined) size.setAttributeNS(WORD_NS, 'w:h', String(Math.max(0, Math.trunc(setup.pageHeight))));
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
    const container = blockContainerOf(this.getPartDocument(entry));
    const indices = new Map(descendants(container, 'p').map((p, i) => [p, i]));
    const walk = (parent: Element): DocumentBlock[] => children(parent).flatMap((child): DocumentBlock[] => {
      if (child.localName === 'p') return [{ type: 'paragraph', paragraph: readParagraph(child, indices.get(child)!) }];
      if (child.localName === 'tbl') return [{
        type: 'table',
        rows: children(child, 'tr').map(row => ({
          cells: children(row, 'tc').map(cell => ({ blocks: walk(cell) })),
        })),
      }];
      if (['sdt', 'sdtContent', 'customXml'].includes(child.localName ?? '')) return walk(child);
      return [];
    });
    return walk(container);
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
    this.updatePartXml(this.mainPath, document => {
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
      if (old !== normalized) replaceSpan(paragraph, start, end, normalized.slice(start, replacementEnd));
    });
  }

  insertParagraph(text: string, before?: number): void {
    assertText(text);
    this.updatePartXml(this.mainPath, document => {
      const paragraph = newParagraph(document, text);
      if (before !== undefined) {
        const target = paragraphAt(document, before);
        target.parentNode!.insertBefore(paragraph, target);
      } else {
        const body = bodyOf(document);
        body.insertBefore(paragraph, children(body, 'sectPr')[0] ?? null);
      }
    });
  }

  deleteParagraph(index: number): void {
    this.updatePartXml(this.mainPath, document => {
      const paragraph = paragraphAt(document, index);
      const parent = paragraph.parentNode as Element;
      // A cell must end with a paragraph, and section properties must not be silently lost.
      if (children(paragraph, 'pPr').some(props => children(props, 'sectPr').length)) {
        throw new Error('Cannot delete a section-break paragraph; edit its XML explicitly.');
      }
      parent.removeChild(paragraph);
      const last = children(parent).filter(child => child.localName !== 'sectPr').at(-1);
      if ((parent.localName === 'tc' && last?.localName !== 'p') ||
          (parent.localName === 'body' && !children(parent, 'p').length)) {
        parent.insertBefore(newParagraph(document, ''), children(parent, 'sectPr')[0] ?? null);
      }
    });
  }

  formatParagraph(index: number, format: ParagraphFormat): void {
    validateParagraphFormat(format);
    this.updatePartXml(this.mainPath, document => {
      const props = properties(paragraphAt(document, index), 'pPr');
      if (format.style !== undefined) setWordValue(property(props, 'pStyle'), format.style);
      if (format.alignment !== undefined) setWordValue(property(props, 'jc'), format.alignment);
    });
  }

  formatRun(paragraph: number, run: number, format: RunFormat): void {
    assertIndex(run);
    validateRunFormat(format);
    this.updatePartXml(this.mainPath, document => {
      const element = ownRuns(paragraphAt(document, paragraph))[run];
      if (!element) throw new Error(`Run ${run} does not exist.`);
      const props = properties(element, 'rPr');
      for (const [key, tag] of [['bold', 'b'], ['italic', 'i'], ['underline', 'u']] as const) {
        if (format[key] !== undefined) {
          setWordValue(property(props, tag), key === 'underline' ? (format[key] ? 'single' : 'none') : (format[key] ? '1' : '0'));
        }
      }
      if (format.fontSize !== undefined) {
        setWordValue(property(props, 'sz'), String(format.fontSize * 2));
        setWordValue(property(props, 'szCs'), String(format.fontSize * 2));
      }
      if (format.color !== undefined) setWordValue(property(props, 'color'), format.color);
      if (format.fontFamily !== undefined) {
        const fonts = property(props, 'rFonts');
        for (const name of ['ascii', 'hAnsi', 'eastAsia', 'cs']) fonts.setAttributeNS(WORD_NS, `w:${name}`, format.fontFamily);
      }
    });
  }

  replaceText(search: string, replacement: string): void {
    assertText(search, 'search'); assertText(replacement, 'replacement');
    if (!search) throw new Error('search must not be empty.');
    this.updatePartXml(this.mainPath, document => {
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
    this.updatePartXml(this.mainPath, document => {
      const body = bodyOf(document);
      const table = wordElement(document, 'tbl');
      const grid = wordElement(document, 'tblGrid');
      const columns = Math.max(...rows.map(row => row.length));
      for (let i = 0; i < columns; i++) {
        const column = wordElement(document, 'gridCol');
        column.setAttributeNS(WORD_NS, 'w:w', String(Math.floor(9000 / columns)));
        grid.appendChild(column);
      }
      table.appendChild(grid);
      for (const row of rows) {
        const tr = wordElement(document, 'tr');
        for (let i = 0; i < columns; i++) {
          const cell = wordElement(document, 'tc');
          cell.appendChild(newParagraph(document, row[i] ?? ''));
          tr.appendChild(cell);
        }
        table.appendChild(tr);
      }
      const section = children(body, 'sectPr')[0] ?? null;
      body.insertBefore(table, section);
      body.insertBefore(newParagraph(document, ''), section);
    });
  }

  /** All operations succeed together, or the original package/revision is unchanged. */
  applyOperations(request: AgentRequest): DocumentSnapshot {
    validateRequest(request);
    if (request.expectedRevision !== undefined && request.expectedRevision !== this.revision) {
      throw new Error(`Revision conflict: expected ${request.expectedRevision}, current ${this.revision}.`);
    }
    if (!request.operations.length) return this.getSnapshot();
    const draft = new DocxDocument(new Map(this.parts));
    for (const operation of request.operations) {
      switch (operation.type) {
        case 'setParagraphText': draft.setParagraphText(operation.index, operation.text); break;
        case 'insertParagraph': draft.insertParagraph(operation.text, operation.before); break;
        case 'deleteParagraph': draft.deleteParagraph(operation.index); break;
        case 'formatParagraph': draft.formatParagraph(operation.index, operation.format); break;
        case 'formatRun': draft.formatRun(operation.paragraph, operation.run, operation.format); break;
        case 'replaceText': draft.replaceText(operation.search, operation.replacement); break;
        case 'insertTable': draft.insertTable(operation.rows); break;
        case 'setPartXml': draft.setPartXml(operation.path, operation.xml); break;
      }
    }
    const snapshot = draft.getSnapshot();
    this.parts = draft.parts;
    this.mainPath = draft.mainPath;
    this.currentRevision++;
    return { ...snapshot, revision: this.revision };
  }

  async toUint8Array(): Promise<Uint8Array> {
    const zip = new JSZip();
    for (const [path, bytes] of this.parts) zip.file(path, bytes, { createFolders: false });
    return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  }

  async toBlob(): Promise<Blob> {
    const bytes = await this.toUint8Array();
    return new Blob([bytes as Uint8Array<ArrayBuffer>], { type: DOCX_TYPE });
  }
}
