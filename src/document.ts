import JSZip from 'jszip';
import type { Document, Element, Node } from '@xmldom/xmldom';
import type {
  AgentRequest, BookmarkInfo, DocumentBlock, DocumentSnapshot, HyperlinkInfo,
  ParagraphFormat, ParagraphInfo, RunFormat, RunInfo,
} from './types.js';
import {
  assertText, children, CONTENT_TYPES_NS, descendants, OFFICE_DOCUMENT_REL, parseXml, REL_NS,
  serializeXml, setWordValue, validatePath, WORD_NS, wordElement, wordValue,
} from './xml.js';
import { assertIndex, validateParagraphFormat, validateRequest, validateRows, validateRunFormat } from './operations.js';
import {
  assertHyperlinkInput, isInternalBookmark, isUnsafeHyperlink, parseFldSimpleHyperlink,
  readExternalRelationship, relationshipPath,
} from './hyperlink.js';

const MAX_ARCHIVE = 50 * 1024 * 1024;
const MAX_PART = 16 * 1024 * 1024;
const MAX_TOTAL = 64 * 1024 * 1024;
const MAX_PARTS = 2048;
const DOCX_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
const MAIN_TYPE = `${DOCX_TYPE}.main+xml`;
const DOC_REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
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

function paragraphAt(document: Document, index: number): Element {
  assertIndex(index);
  const paragraph = descendants(bodyOf(document), 'p')[index];
  if (!paragraph) throw new Error(`Paragraph ${index} does not exist.`);
  return paragraph;
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
  return element.localName === 't' ? element.textContent ?? '' : element.localName === 'tab' ? '\t' : '\n';
}

function textOf(element: Element): string {
  return textElements(element).map(elementText).join('');
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
  return hyperlink.getAttributeNS(DOC_REL_NS, 'id') ?? hyperlink.getAttribute('r:id') ?? undefined;
}

function readRelationships(document: Document | undefined): Element | undefined {
  const root = document?.documentElement;
  return root?.namespaceURI === REL_NS && root.localName === 'Relationships' ? root : undefined;
}

function runHyperlinkInfo(run: Element, paragraph: Element, relationships: Element | undefined): RunInfo['hyperlink'] {
  let parent = run.parentNode;
  while (parent && parent !== paragraph) {
    if (parent.nodeType === 1 && (parent as Element).namespaceURI === WORD_NS &&
        (parent as Element).localName === 'hyperlink') {
      const hyperlink = parent as Element;
      const relationshipId = relationshipIdOf(hyperlink);
      const url = relationshipId ? readExternalRelationship(relationships, relationshipId) : undefined;
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

function readRun(run: Element, index: number, paragraph: Element, relationships: Element | undefined): RunInfo {
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
    hyperlink: runHyperlinkInfo(run, paragraph, relationships),
  };
}

function readParagraph(paragraph: Element, index: number, relationships: Element | undefined): ParagraphInfo {
  const props = children(paragraph, 'pPr')[0];
  const alignment = props ? wordValue(children(props, 'jc')[0]) : undefined;
  return {
    index, text: textOf(paragraph), runs: ownRuns(paragraph).map((run, runIndex) => readRun(run, runIndex, paragraph, relationships)),
    style: props ? wordValue(children(props, 'pStyle')[0]) : undefined,
    alignment: ['left', 'center', 'right', 'both'].includes(alignment ?? '')
      ? alignment as ParagraphFormat['alignment'] : undefined,
  };
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

function fieldInstruction(link: { url?: string; anchor?: string }): string {
  const chunks = ['HYPERLINK'];
  if (link.url) chunks.push(`"${link.url.replace(/"/g, '""')}"`);
  if (link.anchor) chunks.push(`\\l "${link.anchor.replace(/"/g, '""')}"`);
  return chunks.join(' ');
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

  private partDocumentOrUndefined(path: string): Document | undefined {
    try {
      return this.getPartDocument(path);
    } catch {
      return undefined;
    }
  }

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
    const path = decodeURIComponent(main[0]!.getAttribute('Target') ?? '').replace(/^\//, '');
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
    const main = this.getPartDocument(this.mainPath);
    const relationships = readRelationships(this.partDocumentOrUndefined(relationshipPath(this.mainPath)));
    return descendants(bodyOf(main), 'p').map((paragraph, index) => readParagraph(paragraph, index, relationships));
  }

  getBlocks(): DocumentBlock[] {
    const body = bodyOf(this.getPartDocument(this.mainPath));
    const relationships = readRelationships(this.partDocumentOrUndefined(relationshipPath(this.mainPath)));
    const indices = new Map(descendants(body, 'p').map((p, i) => [p, i]));
    const walk = (parent: Element): DocumentBlock[] => children(parent).flatMap((child): DocumentBlock[] => {
      if (child.localName === 'p') return [{ type: 'paragraph', paragraph: readParagraph(child, indices.get(child)!, relationships) }];
      if (child.localName === 'tbl') return [{
        type: 'table',
        rows: children(child, 'tr').map(row => ({
          cells: children(row, 'tc').map(cell => ({ blocks: walk(cell) })),
        })),
      }];
      if (['sdt', 'sdtContent', 'customXml'].includes(child.localName ?? '')) return walk(child);
      return [];
    });
    return walk(body);
  }

  getSnapshot(): DocumentSnapshot {
    return {
      revision: this.revision,
      paragraphs: this.getParagraphs(),
      blocks: this.getBlocks(),
      parts: this.listParts(),
      hyperlinks: this.getHyperlinks(),
      bookmarks: this.getBookmarks(),
    };
  }

  private mainRelationshipsPath(): string {
    return relationshipPath(this.mainPath);
  }

  private ensureMainRelationshipsPart(): void {
    const relsPath = this.mainRelationshipsPath();
    if (this.parts.has(relsPath)) return;
    this.addPart(
      relsPath,
      encoder.encode(`<Relationships xmlns="${REL_NS}"/>`),
      'application/vnd.openxmlformats-package.relationships+xml',
    );
  }

  private nextRelationshipId(document: Document): string {
    const ids = children(document.documentElement!, 'Relationship', REL_NS)
      .map(rel => rel.getAttribute('Id') ?? '')
      .map(id => Number(id.match(/^rId(\d+)$/)?.[1] ?? NaN))
      .filter(value => Number.isFinite(value));
    let next = 1;
    while (ids.includes(next)) next++;
    return `rId${next}`;
  }

  getHyperlinks(): HyperlinkInfo[] {
    const main = this.getPartDocument(this.mainPath);
    const body = bodyOf(main);
    const paragraphs = descendants(body, 'p');
    const relationships = readRelationships(this.partDocumentOrUndefined(this.mainRelationshipsPath()));
    const result: HyperlinkInfo[] = [];
    const paragraphIndex = new Map(paragraphs.map((paragraph, index) => [paragraph, index]));
    for (const paragraph of paragraphs) {
      const runs = ownRuns(paragraph);
      const runIndices = new Map(runs.map((run, index) => [run, index]));
      for (const hyperlink of descendants(paragraph, 'hyperlink')) {
        const linkedRuns = descendants(hyperlink, 'r')
          .map(run => runIndices.get(run))
          .filter((index): index is number => index !== undefined);
        if (!linkedRuns.length) continue;
        const relationshipId = relationshipIdOf(hyperlink);
        const url = relationshipId ? readExternalRelationship(relationships, relationshipId) : undefined;
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
        const linkedRuns = descendants(field, 'r')
          .map(run => runIndices.get(run))
          .filter((index): index is number => index !== undefined);
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

  getBookmarks(options: { includeInternal?: boolean } = {}): BookmarkInfo[] {
    const body = bodyOf(this.getPartDocument(this.mainPath));
    const paragraphs = descendants(body, 'p');
    const paragraphIndex = new Map(paragraphs.map((paragraph, index) => [paragraph, index]));
    const starts = new Map<number, { name: string; paragraph: number }>();
    const ends = new Map<number, number>();
    for (const start of descendants(body, 'bookmarkStart')) {
      const id = Number(start.getAttributeNS(WORD_NS, 'id') ?? start.getAttribute('w:id'));
      const name = start.getAttributeNS(WORD_NS, 'name') ?? start.getAttribute('w:name') ?? '';
      const paragraph = nearestParagraph(start);
      if (!Number.isSafeInteger(id) || !paragraph || !name) continue;
      starts.set(id, { name, paragraph: paragraphIndex.get(paragraph)! });
    }
    for (const end of descendants(body, 'bookmarkEnd')) {
      const id = Number(end.getAttributeNS(WORD_NS, 'id') ?? end.getAttribute('w:id'));
      const paragraph = nearestParagraph(end);
      if (!Number.isSafeInteger(id) || !paragraph) continue;
      ends.set(id, paragraphIndex.get(paragraph)!);
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

  private resolveHyperlink(reference: HyperlinkInfo | number): HyperlinkInfo {
    const hyperlinks = this.getHyperlinks();
    if (typeof reference === 'number') {
      const hyperlink = hyperlinks[reference];
      if (!hyperlink) throw new Error('Hyperlink does not exist.');
      return hyperlink;
    }
    const hyperlink = hyperlinks.find(item => item.paragraph === reference.paragraph &&
      item.runs[0] === reference.runs[0] && item.text === reference.text);
    if (!hyperlink) throw new Error('Hyperlink does not exist.');
    return hyperlink;
  }

  private splitRunAtOffset(paragraph: Element, offset: number): void {
    const runs = ownRuns(paragraph);
    let cursor = 0;
    for (const run of runs) {
      const value = textOf(run);
      const next = cursor + value.length;
      if (offset <= cursor || offset >= next || !value.length) {
        cursor = next;
        continue;
      }
      const left = value.slice(0, offset - cursor);
      const right = value.slice(offset - cursor);
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

  insertHyperlink(
    target: { paragraph: number; start: number; end: number },
    link: { url?: string; anchor?: string; tooltip?: string },
  ): HyperlinkInfo {
    assertIndex(target.paragraph);
    assertIndex(target.start);
    assertIndex(target.end);
    assertHyperlinkInput(link);
    let createdId: string | undefined;
    this.ensureMainRelationshipsPart();
    if (link.url) {
      this.updatePartXml(this.mainRelationshipsPath(), rels => {
        const relationship = rels.createElementNS(REL_NS, 'Relationship');
        createdId = this.nextRelationshipId(rels);
        relationship.setAttribute('Id', createdId);
        relationship.setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink');
        relationship.setAttribute('Target', link.url!);
        relationship.setAttribute('TargetMode', 'External');
        rels.documentElement!.appendChild(relationship);
      });
    }
    this.updatePartXml(this.mainPath, document => {
      const paragraph = paragraphAt(document, target.paragraph);
      textRangeLength(paragraph, target.start, target.end);
      this.splitRunAtOffset(paragraph, target.end);
      this.splitRunAtOffset(paragraph, target.start);
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
      const hyperlink = wordElement(document, 'hyperlink');
      if (createdId) hyperlink.setAttributeNS(DOC_REL_NS, 'r:id', createdId);
      if (link.anchor) hyperlink.setAttributeNS(WORD_NS, 'w:anchor', link.anchor);
      if (link.tooltip) hyperlink.setAttributeNS(WORD_NS, 'w:tooltip', link.tooltip);
      paragraph.insertBefore(hyperlink, selected[0]!);
      for (const run of selected) hyperlink.appendChild(run);
      const props = properties(selected[0]!, 'rPr');
      if (!children(props, 'rStyle').length) setWordValue(property(props, 'rStyle'), 'Hyperlink');
      if (!children(props, 'color').length) setWordValue(property(props, 'color'), '0563C1');
      if (!children(props, 'u').length) setWordValue(property(props, 'u'), 'single');
    });
    return this.getHyperlinks().at(-1)!;
  }

  updateHyperlink(
    hyperlink: HyperlinkInfo | number,
    link: { url?: string; anchor?: string; tooltip?: string },
  ): void {
    assertHyperlinkInput(link);
    const current = this.resolveHyperlink(hyperlink);
    const currentNodeName = this.hyperlinkNode(current, this.getPartDocument(this.mainPath)).localName;
    let nextRelationshipId = current.relationshipId;
    const currentRelationshipUsers = current.relationshipId
      ? this.getHyperlinks().filter(item => item.relationshipId === current.relationshipId).length
      : 0;
    this.ensureMainRelationshipsPart();
    if (currentNodeName !== 'fldSimple' && link.url && link.url !== current.url) {
      this.updatePartXml(this.mainRelationshipsPath(), rels => {
        const canReuse = Boolean(current.relationshipId && currentRelationshipUsers <= 1);
        nextRelationshipId = canReuse ? current.relationshipId : this.nextRelationshipId(rels);
        const existing = canReuse
          ? children(rels.documentElement!, 'Relationship', REL_NS)
            .find(item => item.getAttribute('Id') === nextRelationshipId)
          : undefined;
        const relationship = existing ?? rels.createElementNS(REL_NS, 'Relationship');
        relationship.setAttribute('Id', nextRelationshipId!);
        relationship.setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink');
        relationship.setAttribute('Target', link.url!);
        relationship.setAttribute('TargetMode', 'External');
        if (!existing) rels.documentElement!.appendChild(relationship);
      });
    }
    this.updatePartXml(this.mainPath, document => {
      const node = this.hyperlinkNode(current, document);
      if (node.localName === 'fldSimple') {
        node.setAttributeNS(WORD_NS, 'w:instr', fieldInstruction(link));
        if (link.tooltip) node.setAttributeNS(WORD_NS, 'w:tooltip', link.tooltip);
        else node.removeAttributeNS(WORD_NS, 'tooltip');
      } else {
        if (link.url) node.setAttributeNS(DOC_REL_NS, 'r:id', nextRelationshipId!);
        else { node.removeAttributeNS(DOC_REL_NS, 'id'); node.removeAttribute('r:id'); }
        if (link.anchor) node.setAttributeNS(WORD_NS, 'w:anchor', link.anchor);
        else node.removeAttributeNS(WORD_NS, 'anchor');
        if (link.tooltip) node.setAttributeNS(WORD_NS, 'w:tooltip', link.tooltip);
        else node.removeAttributeNS(WORD_NS, 'tooltip');
      }
    });
    if (current.relationshipId && (!link.url || current.relationshipId !== nextRelationshipId)) {
      const stillUsed = this.getHyperlinks().some(item => item.relationshipId === current.relationshipId);
      if (!stillUsed && this.parts.has(this.mainRelationshipsPath())) {
        this.updatePartXml(this.mainRelationshipsPath(), rels => {
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

  removeHyperlink(hyperlink: HyperlinkInfo | number, options: { keepText?: boolean } = {}): void {
    const link = this.resolveHyperlink(hyperlink);
    this.updatePartXml(this.mainPath, document => {
      const node = this.hyperlinkNode(link, document);
      const parent = node.parentNode as Element;
      if (options.keepText === false) {
        parent.removeChild(node);
      } else {
        while (node.firstChild) parent.insertBefore(node.firstChild, node);
        parent.removeChild(node);
      }
    });
    if (link.relationshipId) {
      const stillUsed = this.getHyperlinks().some(item => item.relationshipId === link.relationshipId);
      if (!stillUsed && this.parts.has(this.mainRelationshipsPath())) {
        this.updatePartXml(this.mainRelationshipsPath(), rels => {
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
        case 'insertHyperlink': draft.insertHyperlink(operation.target, operation.link); break;
        case 'updateHyperlink': draft.updateHyperlink(operation.hyperlink, operation.link); break;
        case 'removeHyperlink': draft.removeHyperlink(operation.hyperlink, operation.options); break;
        case 'insertBookmark': draft.insertBookmark(operation.name, operation.range); break;
        case 'deleteBookmark': draft.deleteBookmark(operation.name); break;
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
