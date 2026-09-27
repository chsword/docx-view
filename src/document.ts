import JSZip from 'jszip';
import type { Document, Element, Node } from '@xmldom/xmldom';
import type {
  AgentRequest, DocumentBlock, DocumentSnapshot, ImageInfo, ParagraphFormat, ParagraphInfo, RunFormat, RunInfo,
} from './types.js';
import {
  A_NS, dataUrlForBytes, decodeBase64, detectImageSize, emuToPx, extensionForContentType, IMAGE_REL,
  isBrowserRenderableContentType, OFFICE_REL_NS, PIC_NS, placeholderDataUrl, pxToEmu, readRunImages,
  resolveRelationshipsPath, resolveTargetPath, V_NS, WP_NS,
} from './drawing.js';
import type { RelationshipTarget } from './drawing.js';
import {
  assertText, children, CONTENT_TYPES_NS, descendants, OFFICE_DOCUMENT_REL, parseXml, REL_NS,
  serializeXml, setWordValue, validatePath, WORD_NS, wordElement, wordValue,
} from './xml.js';
import { assertIndex, validateParagraphFormat, validateRequest, validateRows, validateRunFormat } from './operations.js';

const MAX_ARCHIVE = 50 * 1024 * 1024;
const MAX_PART = 16 * 1024 * 1024;
const MAX_TOTAL = 64 * 1024 * 1024;
const MAX_PARTS = 2048;
const DOCX_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
const MAIN_TYPE = `${DOCX_TYPE}.main+xml`;
const encoder = new TextEncoder();
const IMAGE_LIMIT = 16 * 1024 * 1024;

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

function readRun(
  run: Element,
  index: number,
  paragraph: number,
  relationships: Map<string, RelationshipTarget>,
  getContentType: (path: string) => string | undefined,
  sourcePartPath: string,
): RunInfo {
  const props = children(run, 'rPr')[0];
  const get = (name: string) => props ? children(props, name)[0] : undefined;
  const toggle = (name: string) => get(name) ? !['0', 'false', 'off'].includes(wordValue(get(name)) ?? '') : undefined;
  const size = wordValue(get('sz'));
  const underline = get('u');
  const color = wordValue(get('color'));
  const images = readRunImages(run, paragraph, index, relationships, getContentType, sourcePartPath);
  return {
    index, text: textOf(run), bold: toggle('b'), italic: toggle('i'),
    underline: underline ? !['none', '0', 'false'].includes(wordValue(underline) ?? '') : undefined,
    fontSize: size && Number.isFinite(Number(size)) ? Number(size) / 2 : undefined,
    fontFamily: get('rFonts')?.getAttributeNS(WORD_NS, 'ascii') ?? undefined,
    color: color && /^[a-f\d]{6}$/i.test(color) ? color : undefined,
    image: images[0],
  };
}

function readParagraph(
  paragraph: Element,
  index: number,
  relationships: Map<string, RelationshipTarget>,
  getContentType: (path: string) => string | undefined,
  sourcePartPath: string,
): ParagraphInfo {
  const props = children(paragraph, 'pPr')[0];
  const alignment = props ? wordValue(children(props, 'jc')[0]) : undefined;
  const runElements = ownRuns(paragraph);
  const runs = runElements.map((run, runIndex) => readRun(run, runIndex, index, relationships, getContentType, sourcePartPath));
  return {
    index, text: textOf(paragraph), runs,
    style: props ? wordValue(children(props, 'pStyle')[0]) : undefined,
    alignment: ['left', 'center', 'right', 'both'].includes(alignment ?? '')
      ? alignment as ParagraphFormat['alignment'] : undefined,
    images: runElements.flatMap((run, runIndex) => readRunImages(run, index, runIndex, relationships, getContentType, sourcePartPath)),
  };
}

function setOptionalAttribute(element: Element, name: string, value: string | undefined): void {
  if (value === undefined || value === '') element.removeAttribute(name);
  else element.setAttribute(name, value);
}

function imageElementForRun(run: Element, relationshipId: string, ordinal = 0): Element | undefined {
  let index = 0;
  for (const drawing of elementChildren(run, WORD_NS, 'drawing')) {
    const blips = Array.from(drawing.getElementsByTagNameNS(A_NS, 'blip'));
    if (blips.some((blip) => [blip.getAttributeNS(OFFICE_REL_NS, 'embed'), blip.getAttributeNS(OFFICE_REL_NS, 'link'),
      blip.getAttribute('r:embed'), blip.getAttribute('r:link')].includes(relationshipId))) {
      if (index === ordinal) return drawing;
      index++;
    }
  }
  for (const pict of elementChildren(run, WORD_NS, 'pict')) {
    const imageData = Array.from(pict.getElementsByTagNameNS(V_NS, 'imagedata'));
    if (imageData.some((node) => [node.getAttributeNS(OFFICE_REL_NS, 'id'), node.getAttribute('r:id')].includes(relationshipId))) {
      if (index === ordinal) return pict;
      index++;
    }
  }
  return undefined;
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
    container.setAttribute('simplePos', '0');
    container.setAttribute('relativeHeight', '251658240');
    container.setAttribute('allowOverlap', '1');
    container.appendChild(document.createElementNS(WP_NS, 'wp:simplePos'));
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

  private hasPart(path: string): boolean {
    return this.parts.has(path);
  }

  private getContentType(path: string): string | undefined {
    const types = this.getPartDocument('[Content_Types].xml').documentElement!;
    return children(types, 'Override', CONTENT_TYPES_NS)
      .find(type => type.getAttribute('PartName') === `/${path}`)?.getAttribute('ContentType')
      ?? children(types, 'Default', CONTENT_TYPES_NS)
        .find(type => type.getAttribute('Extension')?.toLowerCase() === path.split('.').pop()?.toLowerCase())?.getAttribute('ContentType')
      ?? undefined;
  }

  private relationshipsFor(partPath: string): Map<string, RelationshipTarget> {
    const relPath = resolveRelationshipsPath(partPath);
    if (!this.hasPart(relPath)) return new Map();
    const rels = this.getPartDocument(relPath).documentElement!;
    const entries: [string, RelationshipTarget][] = children(rels, 'Relationship', REL_NS).flatMap((rel) => {
      const target = rel.getAttribute('Target') ?? undefined;
      const id = rel.getAttribute('Id') ?? '';
      if (!id) return [];
      return [[id, {
        id: rel.getAttribute('Id') ?? '',
        mode: rel.getAttribute('TargetMode') ?? undefined,
        target,
        partPath: target && rel.getAttribute('TargetMode') !== 'External' ? resolveTargetPath(partPath, decodeURIComponent(target)) : undefined,
      } satisfies RelationshipTarget]];
    });
    return new Map(entries);
  }

  private paragraphsWithRelationships(): ParagraphInfo[] {
    const relationships = this.relationshipsFor(this.mainPath);
    return descendants(bodyOf(this.getPartDocument(this.mainPath)), 'p')
      .map((paragraph, index) => readParagraph(paragraph, index, relationships, (path) => this.getContentType(path), this.mainPath));
  }

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

  getParagraphs(): ParagraphInfo[] {
    return this.paragraphsWithRelationships();
  }

  getBlocks(): DocumentBlock[] {
    const body = bodyOf(this.getPartDocument(this.mainPath));
    const indices = new Map(descendants(body, 'p').map((p, i) => [p, i]));
    const relationships = this.relationshipsFor(this.mainPath);
    const walk = (parent: Element): DocumentBlock[] => children(parent).flatMap((child): DocumentBlock[] => {
      if (child.localName === 'p') return [{
        type: 'paragraph',
        paragraph: readParagraph(child, indices.get(child)!, relationships, (path) => this.getContentType(path), this.mainPath),
      }];
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
    return { revision: this.revision, paragraphs: this.getParagraphs(), blocks: this.getBlocks(), parts: this.listParts() };
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
    const info = this.resolveImage(image);
    if (info.isExternal) throw new Error('External images are not loaded.');
    if (!info.partPath || !this.hasPart(info.partPath)) throw new Error(`Image part not found for ${info.relationshipId}.`);
    return this.getPartBytes(info.partPath);
  }

  getImageDataUrl(image: ImageInfo | string): string {
    const info = this.resolveImage(image);
    if (info.isExternal) return placeholderDataUrl('外部图片未加载', info.widthPx || 160, info.heightPx || 90);
    if (!info.partPath || !this.hasPart(info.partPath)) return placeholderDataUrl(info.name ?? '图片缺失', info.widthPx || 160, info.heightPx || 90);
    const contentType = info.contentType ?? this.getContentType(info.partPath);
    if (!isBrowserRenderableContentType(contentType)) {
      return placeholderDataUrl(info.name ?? info.partPath.split('/').pop() ?? '不支持的图片', info.widthPx || 160, info.heightPx || 90);
    }
    return dataUrlForBytes(this.getPartBytes(info.partPath), contentType!);
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
    const insertedRunIndex = options.run ?? targetRuns.length;
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
    if (beforeRun) paragraph.insertBefore(run, beforeRun);
    else paragraph.appendChild(run);
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
    let next = new Map(this.parts);
    let path = info.partPath;
    if (contentType) {
      assertText(contentType, 'contentType');
      const currentExtension = path.split('.').pop()?.toLowerCase();
      const nextExtension = extensionForContentType(contentType);
      if (!nextExtension) throw new Error(`Unsupported image content type: ${contentType}`);
      if (currentExtension !== nextExtension && !(currentExtension === 'jpg' && nextExtension === 'jpeg') &&
          !(currentExtension === 'jpeg' && nextExtension === 'jpg')) {
        const movedPath = this.nextImagePartPath(contentType);
        for (const relPath of this.listParts().filter((entry) => entry.endsWith('.rels'))) {
          const sourcePart = sourcePartFromRelationshipsPath(relPath);
          if (!sourcePart) continue;
          const rels = this.getPartDocument(relPath);
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
      next = this.ensureMediaContentType(path, contentType, next);
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
    this.updatePartXml(info.sourcePartPath ?? this.mainPath, (document) => {
      const run = ownRuns(paragraphAt(document, info.paragraph))[info.run];
      if (!run) throw new Error(`Run ${info.run} does not exist.`);
      const imageElement = imageElementForRun(run, info.relationshipId, info.ordinal ?? 0);
      if (!imageElement) throw new Error(`Image ${info.relationshipId} does not exist.`);
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
    this.updatePartXml(info.sourcePartPath ?? this.mainPath, (document) => {
      const run = ownRuns(paragraphAt(document, info.paragraph))[info.run];
      if (!run) throw new Error(`Run ${info.run} does not exist.`);
      const imageElement = imageElementForRun(run, info.relationshipId, info.ordinal ?? 0);
      if (!imageElement) throw new Error(`Image ${info.relationshipId} does not exist.`);
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
    const main = this.getPartDocument(sourcePart);
    const run = ownRuns(paragraphAt(main, info.paragraph))[info.run];
    if (!run) throw new Error(`Run ${info.run} does not exist.`);
    const imageElement = imageElementForRun(run, info.relationshipId, info.ordinal ?? 0);
    if (!imageElement) throw new Error(`Image ${info.relationshipId} does not exist.`);
    run.removeChild(imageElement);
    const next = new Map(this.parts);
    if (isEmptyRun(run)) run.parentNode!.removeChild(run);
    next.set(sourcePart, encodeXml(serializeXml(main)));
    const relPath = resolveRelationshipsPath(sourcePart);
    if (this.hasPart(relPath)) {
      const rels = this.getPartDocument(relPath);
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
