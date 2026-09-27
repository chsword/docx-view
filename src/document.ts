import JSZip from 'jszip';
import type { Document, Element, Node } from '@xmldom/xmldom';
import type {
  AgentRequest, DocumentBlock, DocumentSnapshot, NoteInfo, NoteSettings, ParagraphFormat, ParagraphInfo, RunFormat, RunInfo,
} from './types.js';
import {
  assertText, children, CONTENT_TYPES_NS, descendants, OFFICE_DOCUMENT_REL, parseXml, REL_NS,
  serializeXml, setWordValue, validatePath, WORD_NS, wordElement, wordValue,
} from './xml.js';
import { assertIndex, validateParagraphFormat, validateRequest, validateRows, validateRunFormat } from './operations.js';
import {
  defaultNotePartXml, formatNoteMarker, noteContentType, notePartPath as conventionalNotePartPath, noteRefName, noteReferenceName,
  noteReferenceStyle, noteRelationshipType, parseCustomMark, parseDocumentNoteSettings, parseNoteEntries,
  parseSectionNoteSettings, setNoteSettingsOn,
} from './notes.js';

const MAX_ARCHIVE = 50 * 1024 * 1024;
const MAX_PART = 16 * 1024 * 1024;
const MAX_TOTAL = 64 * 1024 * 1024;
const MAX_PARTS = 2048;
const DOCX_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
const MAIN_TYPE = `${DOCX_TYPE}.main+xml`;
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

function noteReferenceInRun(run: Element): { kind: 'footnote' | 'endnote'; id: number; customMarkFollows: boolean } | null {
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

function readRun(
  run: Element,
  index: number,
  noteNumber: ((kind: 'footnote' | 'endnote', id: number) => { number: number; marker: string } | null) | null = null,
): RunInfo {
  const props = children(run, 'rPr')[0];
  const get = (name: string) => props ? children(props, name)[0] : undefined;
  const toggle = (name: string) => get(name) ? !['0', 'false', 'off'].includes(wordValue(get(name)) ?? '') : undefined;
  const size = wordValue(get('sz'));
  const underline = get('u');
  const color = wordValue(get('color'));
  const note = noteReferenceInRun(run);
  const resolved = note && noteNumber ? noteNumber(note.kind, note.id) : null;
  return {
    index, text: textOf(run), bold: toggle('b'), italic: toggle('i'),
    underline: underline ? !['none', '0', 'false'].includes(wordValue(underline) ?? '') : undefined,
    fontSize: size && Number.isFinite(Number(size)) ? Number(size) / 2 : undefined,
    fontFamily: get('rFonts')?.getAttributeNS(WORD_NS, 'ascii') ?? undefined,
    color: color && /^[a-f\d]{6}$/i.test(color) ? color : undefined,
    noteReference: note && resolved ? { kind: note.kind, id: note.id, number: resolved.number, marker: resolved.marker } : undefined,
  };
}

function readParagraph(
  paragraph: Element,
  index: number,
  noteNumber: ((kind: 'footnote' | 'endnote', id: number) => { number: number; marker: string } | null) | null = null,
): ParagraphInfo {
  const props = children(paragraph, 'pPr')[0];
  const alignment = props ? wordValue(children(props, 'jc')[0]) : undefined;
  return {
    index, text: textOf(paragraph), runs: ownRuns(paragraph).map((run, runIndex) => readRun(run, runIndex, noteNumber)),
    style: props ? wordValue(children(props, 'pStyle')[0]) : undefined,
    alignment: ['left', 'center', 'right', 'both'].includes(alignment ?? '')
      ? alignment as ParagraphFormat['alignment'] : undefined,
  };
}

interface NoteReferenceRecord {
  kind: 'footnote' | 'endnote';
  id: number;
  customMarkFollows: boolean;
  paragraph: number;
  run: number;
  section: number;
}

interface NoteState {
  byKind: Record<'footnote' | 'endnote', Map<number, { number: number; marker: string }>>;
  refs: NoteReferenceRecord[];
  entries: Record<'footnote' | 'endnote', Map<number, ReturnType<typeof parseNoteEntries>[number]>>;
  notePaths: Record<'footnote' | 'endnote', string | null>;
}

function bodyChildren(parent: Element): Element[] {
  return children(parent).flatMap(child => ['sdt', 'sdtContent', 'customXml'].includes(child.localName ?? '') ? bodyChildren(child) : [child]);
}

function sectionOfParagraphs(body: Element): Map<Element, number> {
  const mapping = new Map<Element, number>();
  let current = 0;
  const walk = (element: Element): void => {
    for (const child of bodyChildren(element)) {
      if (child.localName === 'p') {
        mapping.set(child, current);
        if (children(children(child, 'pPr')[0] ?? child, 'sectPr').length) current++;
        continue;
      }
      if (child.localName === 'tbl') {
        for (const p of descendants(child, 'p')) mapping.set(p, current);
      }
    }
  };
  walk(body);
  return mapping;
}

function relativePath(base: string, target: string): string {
  const from = base.split('/').slice(0, -1);
  const to = target.split('/');
  while (from.length && to.length && from[0] === to[0]) { from.shift(); to.shift(); }
  return `${'../'.repeat(from.length)}${to.join('/')}`;
}

function resolvePartPath(base: string, target: string): string {
  const resolved = decodeURIComponent(target).replace(/^\//, '');
  const stack = base.split('/').slice(0, -1);
  for (const segment of resolved.split('/')) {
    if (!segment || segment === '.') continue;
    if (segment === '..') {
      if (!stack.length) throw new Error(`Invalid relationship target: ${target}`);
      stack.pop();
      continue;
    }
    stack.push(segment);
  }
  const result = stack.join('/');
  validatePath(result);
  return result;
}

function relsPath(partPath: string): string {
  const segments = partPath.split('/');
  const file = segments.pop()!;
  const dir = segments.join('/');
  return `${dir ? `${dir}/` : ''}_rels/${file}.rels`;
}

function nextRelationshipId(rels: Element): string {
  const used = new Set(children(rels, 'Relationship', REL_NS).map(item => item.getAttribute('Id') ?? '').filter(Boolean));
  for (let i = 1; i < Number.MAX_SAFE_INTEGER; i++) {
    const candidate = `rId${i}`;
    if (!used.has(candidate)) return candidate;
  }
  throw new Error('No available relationship id.');
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
  const references: NoteReferenceRecord[] = [];
  for (const [paragraphIndex, paragraph] of paragraphs.entries()) {
    for (const [runIndex, run] of ownRuns(paragraph).entries()) {
      const reference = noteReferenceInRun(run);
      if (!reference) continue;
      references.push({
        kind: reference.kind,
        id: reference.id,
        customMarkFollows: reference.customMarkFollows,
        paragraph: paragraphIndex,
        run: runIndex,
        section: sectionMap.get(paragraph) ?? 0,
      });
    }
  }
  return references;
}

function noteBodyBlocks(
  note: Element,
  paragraphIndices = new Map<Element, number>(),
  noteNumber: ((kind: 'footnote' | 'endnote', id: number) => { number: number; marker: string } | null) | null = null,
): DocumentBlock[] {
  const walk = (parent: Element): DocumentBlock[] => bodyChildren(parent).flatMap((child): DocumentBlock[] => {
    if (child.localName === 'p') return [{ type: 'paragraph', paragraph: readParagraph(child, paragraphIndices.get(child) ?? 0, noteNumber) }];
    if (child.localName === 'tbl') return [{
      type: 'table',
      rows: children(child, 'tr').map(row => ({
        cells: children(row, 'tc').map(cell => ({ blocks: walk(cell) })),
      })),
    }];
    return [];
  });
  return walk(note);
}

function noteParagraphOrder(note: Element): Element[] {
  const result: Element[] = [];
  const walk = (parent: Element): void => {
    for (const child of bodyChildren(parent)) {
      if (child.localName === 'p') { result.push(child); continue; }
      if (child.localName === 'tbl') {
        for (const row of children(child, 'tr')) for (const cell of children(row, 'tc')) walk(cell);
      }
    }
  };
  walk(note);
  return result;
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

  private withDraft<T>(action: (draft: DocxDocument) => T): T {
    const draft = new DocxDocument(new Map(this.parts));
    const result = action(draft);
    this.parts = draft.parts;
    this.mainPath = draft.mainPath;
    this.currentRevision++;
    return result;
  }

  private mainRelationshipsPath(): string {
    return relsPath(this.mainPath);
  }

  private mainRelationshipsDocument(): Document | null {
    const path = this.mainRelationshipsPath();
    return this.parts.has(path) ? this.getPartDocument(path) : null;
  }

  private relationshipTargets(type: string): string[] {
    const rels = this.mainRelationshipsDocument()?.documentElement;
    if (!rels || rels.namespaceURI !== REL_NS || rels.localName !== 'Relationships') return [];
    return children(rels, 'Relationship', REL_NS)
      .filter(item => item.getAttribute('Type') === type && item.getAttribute('TargetMode') !== 'External')
      .map(item => resolvePartPath(this.mainPath, item.getAttribute('Target') ?? ''))
      .filter((path, index, list) => list.indexOf(path) === index);
  }

  private firstRelationshipTarget(type: string): string | null {
    return this.relationshipTargets(type)[0] ?? null;
  }

  private defaultSiblingPath(filename: string): string {
    const segments = this.mainPath.split('/');
    segments.pop();
    return `${segments.length ? `${segments.join('/')}/` : ''}${filename}`;
  }

  private ensureMainRelationship(name: string, type: string, target: string): void {
    const path = this.mainRelationshipsPath();
    if (!this.parts.has(path)) {
      this.addPart(path, encodeXml(`<Relationships xmlns="${REL_NS}"/>`), 'application/vnd.openxmlformats-package.relationships+xml');
    }
    this.updatePartXml(path, document => {
      const root = document.documentElement;
      if (!root || root.namespaceURI !== REL_NS || root.localName !== 'Relationships') throw new Error(`Invalid ${name} relationships part.`);
      const exists = children(root, 'Relationship', REL_NS)
        .some(item => item.getAttribute('Type') === type && decodeURIComponent(item.getAttribute('Target') ?? '') === target);
      if (exists) return;
      const rel = document.createElementNS(REL_NS, 'Relationship');
      rel.setAttribute('Id', nextRelationshipId(root));
      rel.setAttribute('Type', type);
      rel.setAttribute('Target', target);
      root.appendChild(rel);
    });
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

  private notePartPath(kind: 'footnote' | 'endnote'): string | null {
    const fromRelationship = this.firstRelationshipTarget(noteRelationshipType(kind));
    if (fromRelationship) return fromRelationship;
    const conventional = conventionalNotePartPath(kind);
    return this.parts.has(conventional) ? conventional : null;
  }

  private noteDocument(kind: 'footnote' | 'endnote'): Document | null {
    const path = this.notePartPath(kind);
    return path ? this.getPartDocument(path) : null;
  }

  private settingsPartPath(): string | null {
    const fromRelationship = this.firstRelationshipTarget('http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings');
    if (fromRelationship) return fromRelationship;
    const conventional = this.defaultSiblingPath('settings.xml');
    return this.parts.has(conventional) ? conventional : null;
  }

  private collectNoteState(body: Element): NoteState {
    const refs = referenceRecords(body);
    const notePaths = {
      footnote: this.notePartPath('footnote'),
      endnote: this.notePartPath('endnote'),
    };
    const entries = {
      footnote: new Map(parseNoteEntries(notePaths.footnote && this.parts.has(notePaths.footnote) ? this.getPartDocument(notePaths.footnote) : null, 'footnote').map(entry => [entry.id, entry])),
      endnote: new Map(parseNoteEntries(notePaths.endnote && this.parts.has(notePaths.endnote) ? this.getPartDocument(notePaths.endnote) : null, 'endnote').map(entry => [entry.id, entry])),
    };
    const settingsPath = this.settingsPartPath();
    const settingsPart = settingsPath && this.parts.has(settingsPath) ? this.getPartDocument(settingsPath) : null;
    const baseSettings = parseDocumentNoteSettings(settingsPart);
    const perSection = sectionSettings(body, baseSettings);
    const byKind = {
      footnote: new Map<number, { number: number; marker: string }>(),
      endnote: new Map<number, { number: number; marker: string }>(),
    };
    for (const kind of ['footnote', 'endnote'] as const) {
      const counters = new Map<string, number>();
      for (const reference of refs.filter(item => item.kind === kind)) {
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
    return { byKind, refs, entries, notePaths };
  }

  private getParagraphsWith(body: Element, state: NoteState): ParagraphInfo[] {
    const numberOf = (kind: 'footnote' | 'endnote', id: number) => state.byKind[kind].get(id) ?? null;
    return descendants(body, 'p').map((paragraph, index) => readParagraph(paragraph, index, numberOf));
  }

  getParagraphs(): ParagraphInfo[] {
    const body = bodyOf(this.getPartDocument(this.mainPath));
    return this.getParagraphsWith(body, this.collectNoteState(body));
  }

  private getBlocksWith(body: Element, state: NoteState): DocumentBlock[] {
    const numberOf = (kind: 'footnote' | 'endnote', id: number) => state.byKind[kind].get(id) ?? null;
    const indices = new Map(descendants(body, 'p').map((p, i) => [p, i]));
    const walk = (parent: Element): DocumentBlock[] => children(parent).flatMap((child): DocumentBlock[] => {
      if (child.localName === 'p') return [{ type: 'paragraph', paragraph: readParagraph(child, indices.get(child)!, numberOf) }];
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

  getBlocks(): DocumentBlock[] {
    const body = bodyOf(this.getPartDocument(this.mainPath));
    return this.getBlocksWith(body, this.collectNoteState(body));
  }

  getFootnotes(): NoteInfo[] {
    const body = bodyOf(this.getPartDocument(this.mainPath));
    return this.getNotesWith('footnote', this.collectNoteState(body));
  }

  getEndnotes(): NoteInfo[] {
    const body = bodyOf(this.getPartDocument(this.mainPath));
    return this.getNotesWith('endnote', this.collectNoteState(body));
  }

  private getNotesWith(kind: 'footnote' | 'endnote', state: NoteState): NoteInfo[] {
    const numberOf = (noteKind: 'footnote' | 'endnote', id: number) => state.byKind[noteKind].get(id) ?? null;
    const entries = state.entries[kind];
    const results: NoteInfo[] = [];
    const seen = new Set<number>();
    for (const reference of state.refs.filter(item => item.kind === kind)) {
      if (seen.has(reference.id)) continue;
      seen.add(reference.id);
      const entry = entries.get(reference.id);
      if (entry && (entry.id < 1 || entry.type !== 'normal')) continue;
      const numbering = state.byKind[kind].get(reference.id) ?? { number: 0, marker: '?' };
      results.push({
        id: reference.id,
        kind,
        number: numbering.number,
        marker: numbering.marker,
        customMark: reference.customMarkFollows && entry ? parseCustomMark(entry.element) : undefined,
        blocks: entry ? noteBodyBlocks(entry.element, new Map(noteParagraphOrder(entry.element).map((paragraph, index) => [paragraph, index])), numberOf) : [],
        reference: { paragraph: reference.paragraph, run: reference.run },
      });
    }
    return results;
  }

  getNoteSettings(): NoteSettings {
    const settingsPath = this.settingsPartPath();
    return parseDocumentNoteSettings(settingsPath && this.parts.has(settingsPath) ? this.getPartDocument(settingsPath) : null);
  }

  setNoteSettings(settings: Partial<NoteSettings>): void {
    const safe = settings ?? {};
    if (!safe.footnote && !safe.endnote) return;
    this.withDraft(draft => {
      draft.setNoteSettingsDirect(safe);
    });
  }

  private setNoteSettingsDirect(settings: Partial<NoteSettings>): void {
    let path = this.settingsPartPath();
    if (!path) path = this.defaultSiblingPath('settings.xml');
    if (!this.parts.has(path)) {
      this.addPart(path, encodeXml(`<w:settings xmlns:w="${WORD_NS}"/>`), 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml');
    }
    this.ensureMainRelationship('settings', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings', relativePath(this.mainPath, path));
    this.updatePartXml(path, document => setNoteSettingsOn(document, settings));
  }

  getSnapshot(): DocumentSnapshot {
    const mainDocument = this.getPartDocument(this.mainPath);
    const body = bodyOf(mainDocument);
    const state = this.collectNoteState(body);
    return {
      revision: this.revision,
      paragraphs: this.getParagraphsWith(body, state),
      blocks: this.getBlocksWith(body, state),
      footnotes: this.getNotesWith('footnote', state),
      endnotes: this.getNotesWith('endnote', state),
      parts: this.listParts(),
    };
  }

  private ensureNotePart(kind: 'footnote' | 'endnote'): void {
    const path = this.notePartPath(kind) ?? this.defaultSiblingPath(`${kind}s.xml`);
    if (!this.parts.has(path)) {
      this.addPart(path, encodeXml(defaultNotePartXml(kind)), noteContentType(kind));
      this.ensureMainRelationship(kind, noteRelationshipType(kind), relativePath(this.mainPath, path));
    }
  }

  private nextNoteId(kind: 'footnote' | 'endnote'): number {
    const used = new Set(parseNoteEntries(this.noteDocument(kind), kind).map(entry => entry.id).filter(id => id >= 1));
    const body = bodyOf(this.getPartDocument(this.mainPath));
    for (const reference of referenceRecords(body)) {
      if (reference.kind === kind && reference.id >= 1) used.add(reference.id);
    }
    let id = 1;
    while (used.has(id)) id++;
    return id;
  }

  insertFootnote(paragraph: number, run: number, text: string, options: { customMark?: string } = {}): NoteInfo {
    return this.withDraft(draft => draft.insertNoteDirect('footnote', paragraph, run, text, options));
  }

  insertEndnote(paragraph: number, run: number, text: string, options: { customMark?: string } = {}): NoteInfo {
    return this.withDraft(draft => draft.insertNoteDirect('endnote', paragraph, run, text, options));
  }

  private insertNoteDirect(
    kind: 'footnote' | 'endnote',
    paragraph: number,
    run: number,
    text: string,
    options: { customMark?: string } = {},
  ): NoteInfo {
    assertIndex(paragraph);
    assertIndex(run);
    assertText(text);
    if (options.customMark !== undefined) assertText(options.customMark, 'customMark');
    this.updatePartXml(this.mainPath, document => {
      const paragraphElement = paragraphAt(document, paragraph);
      const runs = ownRuns(paragraphElement);
      if (run > runs.length) throw new Error(`Run ${run} does not exist.`);
    });
    this.ensureNotePart(kind);
    const notePath = this.notePartPath(kind)!;
    const id = this.nextNoteId(kind);
    this.updatePartXml(this.mainPath, document => {
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
      if (anchor?.parentNode) {
        anchor.parentNode.insertBefore(referenceRun, anchor);
      } else {
        paragraphElement.insertBefore(referenceRun, null);
      }
    });
    this.updatePartXml(notePath, document => {
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
    return this.getNotesWith(kind, this.collectNoteState(body)).find(item => item.id === id)!;
  }

  setNoteText(kind: 'footnote' | 'endnote', id: number, text: string): void {
    this.withDraft(draft => draft.setNoteTextDirect(kind, id, text));
  }

  private setNoteTextDirect(kind: 'footnote' | 'endnote', id: number, text: string): void {
    assertIndex(id);
    assertText(text);
    const path = this.notePartPath(kind);
    if (!path) throw new Error(`${kind} ${id} does not exist.`);
    this.updatePartXml(path, document => {
      const note = parseNoteEntries(document, kind).find(item => item.id === id && item.type === 'normal');
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

  deleteNote(kind: 'footnote' | 'endnote', id: number): void {
    this.withDraft(draft => draft.deleteNoteDirect(kind, id));
  }

  private deleteNoteDirect(kind: 'footnote' | 'endnote', id: number): void {
    assertIndex(id);
    const path = this.notePartPath(kind);
    if (path && this.parts.has(path)) {
      this.updatePartXml(path, document => {
        const entry = parseNoteEntries(document, kind).find(item => item.id === id);
        if (entry) entry.element.parentNode?.removeChild(entry.element);
      });
    }
    this.updatePartXml(this.mainPath, document => {
      const body = bodyOf(document);
      for (const paragraph of descendants(body, 'p')) {
        for (const run of ownRuns(paragraph)) {
          const reference = children(run, noteReferenceName(kind))[0];
          if (reference && Number(reference.getAttributeNS(WORD_NS, 'id')) === id) {
            run.removeChild(reference);
            if (children(run).every(child => child.localName === 'rPr')) run.parentNode?.removeChild(run);
          }
        }
      }
    });
  }

  convertNote(kind: 'footnote' | 'endnote', id: number): void {
    this.withDraft(draft => draft.convertNoteDirect(kind, id));
  }

  private convertNoteDirect(kind: 'footnote' | 'endnote', id: number): void {
    assertIndex(id);
    const targetKind = kind === 'footnote' ? 'endnote' : 'footnote';
    const sourceDocument = this.noteDocument(kind);
    const source = parseNoteEntries(sourceDocument, kind).find(item => item.id === id && item.type === 'normal');
    if (!source) throw new Error(`${kind} ${id} does not exist.`);
    this.ensureNotePart(targetKind);
    const targetId = this.nextNoteId(targetKind);
    const targetPath = this.notePartPath(targetKind)!;
    const targetDocument = this.getPartDocument(targetPath);
    const clone = wordElement(targetDocument, targetKind);
    for (let i = 0; i < source.element.attributes.length; i++) {
      const attribute = source.element.attributes.item(i);
      if (!attribute) continue;
      clone.setAttributeNS(attribute.namespaceURI, attribute.name, attribute.value);
    }
    for (let child = source.element.firstChild; child; child = child.nextSibling) {
      clone.appendChild(targetDocument.importNode(child, true));
    }
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
    this.setPartXml(this.notePartPath(kind)!, serializeXml(sourceDocument!));
    this.setPartXml(targetPath, serializeXml(targetDocument));
    this.updatePartXml(this.mainPath, document => {
      for (const paragraph of descendants(bodyOf(document), 'p')) {
        for (const runElement of ownRuns(paragraph)) {
          const from = children(runElement, noteReferenceName(kind))[0];
          if (!from || Number(from.getAttributeNS(WORD_NS, 'id')) !== id) continue;
          const to = wordElement(document, noteReferenceName(targetKind));
          to.setAttributeNS(WORD_NS, 'w:id', String(targetId));
          if (from.getAttributeNS(WORD_NS, 'customMarkFollows')) {
            to.setAttributeNS(WORD_NS, 'w:customMarkFollows', from.getAttributeNS(WORD_NS, 'customMarkFollows')!);
          }
          runElement.replaceChild(to, from);
          const runProps = children(runElement, 'rPr')[0];
          if (runProps) setWordValue(property(runProps, 'rStyle'), noteReferenceStyle(targetKind));
        }
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
        case 'setPartXml': draft.setPartXml(operation.path, operation.xml); break;
        case 'insertFootnote': draft.insertNoteDirect('footnote', operation.paragraph, operation.run, operation.text, { customMark: operation.customMark }); break;
        case 'insertEndnote': draft.insertNoteDirect('endnote', operation.paragraph, operation.run, operation.text, { customMark: operation.customMark }); break;
        case 'setNoteText': draft.setNoteTextDirect(operation.kind, operation.id, operation.text); break;
        case 'deleteNote': draft.deleteNoteDirect(operation.kind, operation.id); break;
        case 'convertNote': draft.convertNoteDirect(operation.kind, operation.id); break;
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
