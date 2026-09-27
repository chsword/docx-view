import JSZip from 'jszip';
import type { Document, Element, Node } from '@xmldom/xmldom';
import type {
  AgentRequest, DocumentBlock, DocumentSnapshot, ParagraphFormat, ParagraphInfo, RunFormat, RunInfo, StyleInfo,
} from './types.js';
import {
  assertText, children, CONTENT_TYPES_NS, descendants, OFFICE_DOCUMENT_REL, parseXml, REL_NS,
  serializeXml, setWordValue, validatePath, WORD_NS, wordElement, wordValue,
} from './xml.js';
import { assertIndex, validateParagraphFormat, validateRequest, validateRows, validateRunFormat } from './operations.js';
import {
  cloneStyleInfo,
  computeEffectiveParagraphFormat,
  computeEffectiveRunFormat,
  parseStyles,
  readParagraphProperties,
  readRunProperties,
  type StylesContext,
} from './styles.js';

const MAX_ARCHIVE = 50 * 1024 * 1024;
const MAX_PART = 16 * 1024 * 1024;
const MAX_TOTAL = 64 * 1024 * 1024;
const MAX_PARTS = 2048;
const DOCX_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
const MAIN_TYPE = `${DOCX_TYPE}.main+xml`;
const STYLES_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles';
const STYLES_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml';
const THEME_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme';
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

function styleChild(parent: Element, name: string): Element {
  let result = children(parent, name)[0];
  if (!result) {
    result = wordElement(parent.ownerDocument!, name);
    parent.appendChild(result);
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

function dirname(path: string): string {
  const slash = path.lastIndexOf('/');
  return slash === -1 ? '' : path.slice(0, slash);
}

function basename(path: string): string {
  return path.slice(path.lastIndexOf('/') + 1);
}

function relsPath(path: string): string {
  const dir = dirname(path);
  return `${dir ? `${dir}/` : ''}_rels/${basename(path)}.rels`;
}

function resolveTarget(sourcePart: string, target: string): string {
  const decoded = decodeURIComponent(target).replace(/^\//, '');
  if (!decoded) throw new Error('Invalid relationship target.');
  if (!target.startsWith('/')) {
    const base = dirname(sourcePart).split('/').filter(Boolean);
    for (const segment of decoded.split('/')) {
      if (!segment || segment === '.') continue;
      if (segment === '..') base.pop();
      else base.push(segment);
    }
    const path = base.join('/');
    validatePath(path);
    return path;
  }
  validatePath(decoded);
  return decoded;
}

function readRun(run: Element, index: number, styles: StylesContext, paragraph: Element): RunInfo {
  const direct = readRunProperties(children(run, 'rPr')[0], styles.theme);
  return {
    index,
    text: textOf(run),
    ...direct,
    effective: computeEffectiveRunFormat(styles, paragraph, run),
  };
}

function readParagraph(paragraph: Element, index: number, styles: StylesContext): ParagraphInfo {
  const direct = readParagraphProperties(children(paragraph, 'pPr')[0]);
  return {
    index,
    text: textOf(paragraph),
    ...direct,
    runs: ownRuns(paragraph).map((run, runIndex) => readRun(run, runIndex, styles, paragraph)),
    effective: computeEffectiveParagraphFormat(styles, paragraph),
  };
}

function removeProperty(parent: Element, name: string): void {
  for (const child of children(parent, name)) parent.removeChild(child);
}

function setOnOff(parent: Element, name: string, value: boolean, onValue = '1', offValue = '0'): void {
  setWordValue(property(parent, name), value ? onValue : offValue);
}

function applyParagraphFormatTo(props: Element, format: ParagraphFormat): void {
  if (format.style !== undefined) setWordValue(property(props, 'pStyle'), format.style);
  if (format.alignment !== undefined) setWordValue(property(props, 'jc'), format.alignment);
  if ([format.keepNext, format.keepLines, format.pageBreakBefore, format.widowControl].some(value => value !== undefined)) {
    if (format.keepNext !== undefined) setOnOff(props, 'keepNext', format.keepNext);
    if (format.keepLines !== undefined) setOnOff(props, 'keepLines', format.keepLines);
    if (format.pageBreakBefore !== undefined) setOnOff(props, 'pageBreakBefore', format.pageBreakBefore);
    if (format.widowControl !== undefined) setOnOff(props, 'widowControl', format.widowControl);
  }
  if ([format.indentLeft, format.indentRight, format.indentFirstLine, format.indentHanging].some(value => value !== undefined)) {
    const indent = property(props, 'ind');
    if (format.indentLeft !== undefined) indent.setAttributeNS(WORD_NS, 'w:left', String(format.indentLeft));
    if (format.indentRight !== undefined) indent.setAttributeNS(WORD_NS, 'w:right', String(format.indentRight));
    if (format.indentFirstLine !== undefined) indent.setAttributeNS(WORD_NS, 'w:firstLine', String(format.indentFirstLine));
    if (format.indentHanging !== undefined) indent.setAttributeNS(WORD_NS, 'w:hanging', String(format.indentHanging));
  }
  if ([format.spacingBefore, format.spacingAfter, format.lineSpacing, format.lineSpacingRule].some(value => value !== undefined)) {
    const spacing = property(props, 'spacing');
    if (format.spacingBefore !== undefined) spacing.setAttributeNS(WORD_NS, 'w:before', String(format.spacingBefore));
    if (format.spacingAfter !== undefined) spacing.setAttributeNS(WORD_NS, 'w:after', String(format.spacingAfter));
    if (format.lineSpacing !== undefined) spacing.setAttributeNS(WORD_NS, 'w:line', String(format.lineSpacing));
    if (format.lineSpacingRule !== undefined) spacing.setAttributeNS(WORD_NS, 'w:lineRule', format.lineSpacingRule);
  }
  if (format.outlineLevel !== undefined) setWordValue(property(props, 'outlineLvl'), String(format.outlineLevel));
}

function applyRunFormatTo(props: Element, format: RunFormat): void {
  if (format.style !== undefined) setWordValue(property(props, 'rStyle'), format.style);
  for (const [key, tag] of [
    ['bold', 'b'],
    ['italic', 'i'],
    ['strike', 'strike'],
    ['doubleStrike', 'dstrike'],
    ['smallCaps', 'smallCaps'],
    ['allCaps', 'caps'],
  ] as const) {
    if (format[key] !== undefined) setOnOff(props, tag, format[key]!);
  }
  if (format.underline !== undefined || format.underlineStyle !== undefined || format.underlineColor !== undefined) {
    const underline = property(props, 'u');
    if (format.underlineColor !== undefined) underline.setAttributeNS(WORD_NS, 'w:color', format.underlineColor);
    if (format.underline !== undefined || format.underlineStyle !== undefined) {
      const value = format.underline === false ? 'none' : (format.underlineStyle ?? (format.underline ? 'single' : 'none'));
      setWordValue(underline, value);
    }
  }
  if (format.fontSize !== undefined) {
    setWordValue(property(props, 'sz'), String(format.fontSize * 2));
    setWordValue(property(props, 'szCs'), String(format.fontSize * 2));
  }
  if (format.color !== undefined) setWordValue(property(props, 'color'), format.color);
  if (format.fontFamily !== undefined || format.fontFamilyEastAsia !== undefined) {
    const fonts = property(props, 'rFonts');
    if (format.fontFamily !== undefined) {
      for (const name of ['ascii', 'hAnsi', 'cs']) fonts.setAttributeNS(WORD_NS, `w:${name}`, format.fontFamily);
      if (format.fontFamilyEastAsia === undefined) fonts.setAttributeNS(WORD_NS, 'w:eastAsia', format.fontFamily);
    }
    if (format.fontFamilyEastAsia !== undefined) fonts.setAttributeNS(WORD_NS, 'w:eastAsia', format.fontFamilyEastAsia);
  }
  if (format.verticalAlign !== undefined) {
    if (format.verticalAlign === 'baseline') removeProperty(props, 'vertAlign');
    else setWordValue(property(props, 'vertAlign'), format.verticalAlign);
  }
  if (format.highlight !== undefined) setWordValue(property(props, 'highlight'), format.highlight);
  if (format.characterSpacing !== undefined) setWordValue(property(props, 'spacing'), String(format.characterSpacing));
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

  private getRelatedPartPath(type: string, fallback?: string): string | undefined {
    const rels = this.parts.get(relsPath(this.mainPath));
    if (rels) {
      try {
        const document = parseXml(decodeXml(rels)).documentElement;
        if (!document) return fallback && this.parts.has(fallback) ? fallback : undefined;
        for (const relation of children(document, 'Relationship', REL_NS)) {
          if (relation.getAttribute('Type') === type && relation.getAttribute('TargetMode') !== 'External') {
            const target = relation.getAttribute('Target');
            if (!target) continue;
            const path = resolveTarget(this.mainPath, target);
            if (this.parts.has(path)) return path;
          }
        }
      } catch { /* Fall back to conventional paths for malformed optional rels parts. */ }
    }
    return fallback && this.parts.has(fallback) ? fallback : undefined;
  }

  private getStylesContext(): StylesContext {
    const stylesPath = this.getRelatedPartPath(STYLES_REL, 'word/styles.xml');
    const themePath = this.getRelatedPartPath(THEME_REL, 'word/theme/theme1.xml');
    let stylesRoot: Element | undefined;
    let themeRoot: Element | undefined;
    try { stylesRoot = stylesPath ? this.getPartDocument(stylesPath).documentElement ?? undefined : undefined; } catch { stylesRoot = undefined; }
    try { themeRoot = themePath ? this.getPartDocument(themePath).documentElement ?? undefined : undefined; } catch { themeRoot = undefined; }
    return parseStyles(stylesRoot, themeRoot);
  }

  private buildParagraphs(document = this.getPartDocument(this.mainPath), styles = this.getStylesContext()): ParagraphInfo[] {
    return descendants(bodyOf(document), 'p').map((paragraph, index) => readParagraph(paragraph, index, styles));
  }

  getParagraphs(): ParagraphInfo[] {
    return this.buildParagraphs();
  }

  getBlocks(): DocumentBlock[] {
    const document = this.getPartDocument(this.mainPath);
    const styles = this.getStylesContext();
    const body = bodyOf(document);
    const paragraphs = this.buildParagraphs(document, styles);
    const indices = new Map(descendants(body, 'p').map((paragraph, index) => [paragraph, paragraphs[index]!]));
    const walk = (parent: Element): DocumentBlock[] => children(parent).flatMap((child): DocumentBlock[] => {
      if (child.localName === 'p') return [{ type: 'paragraph', paragraph: indices.get(child)! }];
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
    const styles = this.getStyles();
    return { revision: this.revision, paragraphs: this.getParagraphs(), blocks: this.getBlocks(), parts: this.listParts(), styles };
  }

  getStyles(): StyleInfo[] {
    return this.getStylesContext().styles.map(cloneStyleInfo);
  }

  getStyle(id: string): StyleInfo | undefined {
    assertText(id, 'style id');
    const style = this.getStylesContext().byId.get(id);
    return style ? cloneStyleInfo(style) : undefined;
  }

  getEffectiveParagraphFormat(index: number): ParagraphFormat {
    const paragraph = paragraphAt(this.getPartDocument(this.mainPath), index);
    return computeEffectiveParagraphFormat(this.getStylesContext(), paragraph);
  }

  getEffectiveRunFormat(paragraph: number, run: number): RunFormat {
    const document = this.getPartDocument(this.mainPath);
    const paragraphElement = paragraphAt(document, paragraph);
    const runElement = ownRuns(paragraphElement)[run];
    if (!runElement) throw new Error(`Run ${run} does not exist.`);
    return computeEffectiveRunFormat(this.getStylesContext(), paragraphElement, runElement);
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

  formatParagraph(index: number, format: ParagraphFormat, options: { validateStyle?: boolean } = {}): void {
    validateParagraphFormat(format);
    if (options.validateStyle && format.style !== undefined && !this.getStyle(format.style)) {
      throw new Error(`Paragraph style not found: ${format.style} (styles.xml is missing or does not define it).`);
    }
    this.updatePartXml(this.mainPath, document => {
      const props = properties(paragraphAt(document, index), 'pPr');
      applyParagraphFormatTo(props, format);
    });
  }

  formatRun(paragraph: number, run: number, format: RunFormat): void {
    assertIndex(run);
    validateRunFormat(format);
    this.updatePartXml(this.mainPath, document => {
      const element = ownRuns(paragraphAt(document, paragraph))[run];
      if (!element) throw new Error(`Run ${run} does not exist.`);
      const props = properties(element, 'rPr');
      applyRunFormatTo(props, format);
    });
  }

  defineStyle(style: StyleInfo): void {
    assertText(style.id, 'style.id');
    assertText(style.name, 'style.name');
    const type = style.type;
    if (!['paragraph', 'character', 'table', 'numbering'].includes(type)) throw new Error(`Unsupported style type: ${String(type)}`);
    if (style.paragraph !== undefined) validateParagraphFormat(style.paragraph);
    if (style.run !== undefined) validateRunFormat(style.run);
    const draft = new DocxDocument(new Map(this.parts));
    const stylesPath = draft.getRelatedPartPath(STYLES_REL, 'word/styles.xml') ?? (() => {
      const path = `${dirname(draft.mainPath) ? `${dirname(draft.mainPath)}/` : ''}styles.xml`;
      const relsDocument = draft.parts.has(relsPath(draft.mainPath))
        ? draft.getPartDocument(relsPath(draft.mainPath))
        : parseXml(`<Relationships xmlns="${REL_NS}"/>`);
      const relsRoot = relsDocument.documentElement!;
      const exists = children(relsRoot, 'Relationship', REL_NS)
        .some(relation => relation.getAttribute('Type') === STYLES_REL);
      if (!exists) {
        const relation = relsDocument.createElementNS(REL_NS, 'Relationship');
        relation.setAttribute('Id', `rId${children(relsRoot, 'Relationship', REL_NS).length + 1}`);
        relation.setAttribute('Type', STYLES_REL);
        relation.setAttribute('Target', basename(path));
        relsRoot.appendChild(relation);
        draft.parts.set(relsPath(draft.mainPath), encodeXml(serializeXml(relsDocument)));
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
        }
        draft.parts.set(path, encodeXml(`<w:styles xmlns:w="${WORD_NS}"/>`));
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
      if (style.isDefault !== undefined) styleElement.setAttributeNS(WORD_NS, 'w:default', style.isDefault ? '1' : '0');
      const name = styleChild(styleElement, 'name');
      setWordValue(name, style.name);
      if (style.basedOn) setWordValue(styleChild(styleElement, 'basedOn'), style.basedOn);
      if (style.next) setWordValue(styleChild(styleElement, 'next'), style.next);
      if (style.link) setWordValue(styleChild(styleElement, 'link'), style.link);
      if (style.aliases?.length) setWordValue(styleChild(styleElement, 'aliases'), style.aliases.join(', '));
      if (style.quickFormat) styleChild(styleElement, 'qFormat');
      if (style.paragraph) applyParagraphFormatTo(styleChild(styleElement, 'pPr'), style.paragraph);
      if (style.run) applyRunFormatTo(styleChild(styleElement, 'rPr'), style.run);
    });
    this.parts = draft.parts;
    this.mainPath = draft.mainPath;
    this.currentRevision++;
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
