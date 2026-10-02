import type { Document, Element } from '@xmldom/xmldom';
import type { DocumentProperties, DocumentStatistics } from './types.js';
import { assertText } from './xml.js';

export const CORE_PROPS_REL = 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties';
export const APP_PROPS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties';
export const CORE_PROPS_TYPE = 'application/vnd.openxmlformats-package.core-properties+xml';
export const APP_PROPS_TYPE = 'application/vnd.openxmlformats-officedocument.extended-properties+xml';

const CP_NS = 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties';
const DC_NS = 'http://purl.org/dc/elements/1.1/';
const DCTERMS_NS = 'http://purl.org/dc/terms/';
const DCMITYPE_NS = 'http://purl.org/dc/dcmitype/';
const XSI_NS = 'http://www.w3.org/2001/XMLSchema-instance';
const APP_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties';
const VT_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes';

export const CORE_PROPERTY_KEYS = ['title', 'subject', 'creator', 'lastModifiedBy', 'keywords', 'description', 'category', 'created', 'modified', 'revisionNumber'] as const;
export const APP_PROPERTY_KEYS = ['company', 'manager'] as const;
const DATE_KEYS = new Set<DocumentPropertyKey>(['created', 'modified']);
const TEXT_KEYS = new Set<DocumentPropertyKey>(['title', 'subject', 'creator', 'lastModifiedBy', 'keywords', 'description', 'category', 'company', 'manager']);
const VALID_KEYS = new Set<DocumentPropertyKey>([...CORE_PROPERTY_KEYS, ...APP_PROPERTY_KEYS]);
const ISO_8601_RE = /^\d{4}-\d{2}-\d{2}(?:T\d{2}:\d{2}(?::\d{2}(?:\.\d{1,9})?)?(?:Z|[+-]\d{2}:\d{2})?)?$/;

type DocumentPropertyKey = keyof DocumentProperties;

const CORE_FIELD_INFO: Record<typeof CORE_PROPERTY_KEYS[number], { namespace: string; qualifiedName: string; localName: string; date?: boolean; numeric?: boolean }> = {
  title: { namespace: DC_NS, qualifiedName: 'dc:title', localName: 'title' },
  subject: { namespace: DC_NS, qualifiedName: 'dc:subject', localName: 'subject' },
  creator: { namespace: DC_NS, qualifiedName: 'dc:creator', localName: 'creator' },
  lastModifiedBy: { namespace: CP_NS, qualifiedName: 'cp:lastModifiedBy', localName: 'lastModifiedBy' },
  keywords: { namespace: CP_NS, qualifiedName: 'cp:keywords', localName: 'keywords' },
  description: { namespace: DC_NS, qualifiedName: 'dc:description', localName: 'description' },
  category: { namespace: CP_NS, qualifiedName: 'cp:category', localName: 'category' },
  created: { namespace: DCTERMS_NS, qualifiedName: 'dcterms:created', localName: 'created', date: true },
  modified: { namespace: DCTERMS_NS, qualifiedName: 'dcterms:modified', localName: 'modified', date: true },
  revisionNumber: { namespace: CP_NS, qualifiedName: 'cp:revision', localName: 'revision', numeric: true },
};

const APP_FIELD_INFO: Record<typeof APP_PROPERTY_KEYS[number], { qualifiedName: string; localName: string }> = {
  company: { qualifiedName: 'Company', localName: 'Company' },
  manager: { qualifiedName: 'Manager', localName: 'Manager' },
};
const CORE_PROPERTY_ORDER = ['category', 'contentStatus', 'contentType', 'created', 'creator', 'description', 'identifier', 'keywords', 'language', 'lastModifiedBy', 'lastPrinted', 'modified', 'revision', 'subject', 'title', 'version'];
const APP_PROPERTY_ORDER = ['Template', 'Manager', 'Company', 'Pages', 'Words', 'Characters', 'PresentationFormat', 'Lines', 'Paragraphs', 'Slides', 'Notes', 'TotalTime', 'HiddenSlides', 'MMClips', 'ScaleCrop', 'HeadingPairs', 'TitlesOfParts', 'LinksUpToDate', 'CharactersWithSpaces', 'SharedDoc', 'HyperlinkBase', 'HyperlinksChanged', 'AppVersion', 'DocSecurity'];

function childInNamespace(parent: Element, namespace: string, localName: string): Element | undefined {
  for (let child = parent.firstChild; child; child = child.nextSibling) {
    if (child.nodeType !== 1) continue;
    const element = child as Element;
    if (element.namespaceURI === namespace && element.localName === localName) return element;
  }
  return undefined;
}

function removeChild(parent: Element, namespace: string, localName: string): void {
  const element = childInNamespace(parent, namespace, localName);
  element?.parentNode?.removeChild(element);
}

function textContentOf(element: Element | undefined): string | undefined {
  return element ? element.textContent ?? '' : undefined;
}

function setElementText(element: Element, value: string): void {
  while (element.firstChild) element.removeChild(element.firstChild);
  element.appendChild(element.ownerDocument!.createTextNode(value));
}

function ensureChild(parent: Element, namespace: string, qualifiedName: string, localName: string, order: string[]): Element {
  const existing = childInNamespace(parent, namespace, localName);
  if (existing) return existing;
  const element = parent.ownerDocument!.createElementNS(namespace, qualifiedName);
  const position = order.indexOf(localName);
  const following = Array.from({ length: parent.childNodes.length }, (_, index) => parent.childNodes.item(index))
    .find((child) => child?.nodeType === 1 && order.indexOf((child as Element).localName ?? '') > position);
  parent.insertBefore(element, following ?? null);
  return element;
}

function rootOrNull(document: Document | null, namespace: string, localName: string): Element | null {
  const root = document?.documentElement;
  if (!root || root.namespaceURI !== namespace || root.localName !== localName) return null;
  return root;
}

function validateIsoDate(value: string, name: string): void {
  assertText(value, name);
  if (!ISO_8601_RE.test(value) || !Number.isFinite(Date.parse(value))) {
    throw new Error(`${name} must be a valid ISO 8601 date/time string.`);
  }
}

export function assertDocumentPropertiesPatch(patch: unknown): asserts patch is Partial<DocumentProperties> {
  if (!patch || typeof patch !== 'object' || Array.isArray(patch)) {
    throw new Error('document properties patch must be an object.');
  }
  for (const key of Object.keys(patch)) {
    if (!VALID_KEYS.has(key as DocumentPropertyKey)) throw new Error(`Unknown document property: ${key}.`);
  }
  const record = patch as Partial<DocumentProperties>;
  for (const key of TEXT_KEYS) {
    const value = record[key];
    if (key in record && value !== undefined) assertText(value, key);
  }
  for (const key of DATE_KEYS) {
    const value = record[key];
    if (key in record && value !== undefined) validateIsoDate(value as string, key);
  }
  if ('revisionNumber' in record && record.revisionNumber !== undefined &&
      (!Number.isSafeInteger(record.revisionNumber) || record.revisionNumber < 0)) {
    throw new Error('revisionNumber must be a non-negative integer.');
  }
}

export function defaultCorePropertiesXml(): string {
  return `<cp:coreProperties xmlns:cp="${CP_NS}" xmlns:dc="${DC_NS}" xmlns:dcterms="${DCTERMS_NS}" xmlns:dcmitype="${DCMITYPE_NS}" xmlns:xsi="${XSI_NS}"/>`;
}

export function defaultAppPropertiesXml(): string {
  return `<Properties xmlns="${APP_NS}" xmlns:vt="${VT_NS}"/>`;
}

export function parseDocumentProperties(coreDocument: Document | null, appDocument: Document | null): DocumentProperties {
  const result: DocumentProperties = {};
  const coreRoot = rootOrNull(coreDocument, CP_NS, 'coreProperties');
  if (coreRoot) {
    for (const key of CORE_PROPERTY_KEYS) {
      const info = CORE_FIELD_INFO[key];
      const value = textContentOf(childInNamespace(coreRoot, info.namespace, info.localName));
      if (value === undefined) continue;
      if (info.numeric) {
        if (/^\d+$/.test(value)) (result as Record<DocumentPropertyKey, unknown>)[key] = Number(value);
        continue;
      }
      (result as Record<DocumentPropertyKey, unknown>)[key] = value;
    }
  }
  const appRoot = rootOrNull(appDocument, APP_NS, 'Properties');
  if (appRoot) {
    for (const key of APP_PROPERTY_KEYS) {
      const info = APP_FIELD_INFO[key];
      const value = textContentOf(childInNamespace(appRoot, APP_NS, info.localName));
      if (value !== undefined) result[key] = value;
    }
  }
  return result;
}

export function setCoreDocumentPropertiesOn(document: Document, patch: Partial<DocumentProperties>): void {
  const root = rootOrNull(document, CP_NS, 'coreProperties');
  if (!root) throw new Error('Invalid core properties root.');
  for (const key of CORE_PROPERTY_KEYS) {
    if (!(key in patch)) continue;
    const info = CORE_FIELD_INFO[key];
    const value = patch[key];
    if (value === undefined) {
      removeChild(root, info.namespace, info.localName);
      continue;
    }
    const element = ensureChild(root, info.namespace, info.qualifiedName, info.localName, CORE_PROPERTY_ORDER);
    if (info.numeric) {
      setElementText(element, String(value));
      continue;
    }
    setElementText(element, value as string);
    if (info.date) element.setAttributeNS(XSI_NS, 'xsi:type', 'dcterms:W3CDTF');
  }
}

export function setAppDocumentPropertiesOn(document: Document, patch: Partial<DocumentProperties>): void {
  const root = rootOrNull(document, APP_NS, 'Properties');
  if (!root) throw new Error('Invalid app properties root.');
  for (const key of APP_PROPERTY_KEYS) {
    if (!(key in patch)) continue;
    const info = APP_FIELD_INFO[key];
    const value = patch[key];
    if (value === undefined) {
      removeChild(root, APP_NS, info.localName);
      continue;
    }
    setElementText(ensureChild(root, APP_NS, info.qualifiedName, info.localName, APP_PROPERTY_ORDER), value);
  }
}

/**
 * 字数统计按 Word 的口径：东亚文字（汉字、假名、谚文）**每个字算一个词**，其余按空白切分的连续
 * 非空白串算一个词；东亚标点不算词。字符数按 Unicode 码点计（一个 emoji 是一个字符，不是两个
 * UTF-16 单元），「不含空格」去掉所有空白。
 */
const EAST_ASIAN_LETTER = /[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Hangul}]/u;
const EAST_ASIAN_PUNCTUATION = /[　-〿＀-／：-＠［-｀｛-･‘-‟…]/u;

export function countTextStatistics(text: string): { words: number; characters: number; charactersWithSpaces: number } {
  let words = 0;
  let inWord = false;
  let characters = 0;
  let charactersWithSpaces = 0;
  for (const character of text) {
    if (character === '\n' || character === '\r') { inWord = false; continue; }
    charactersWithSpaces++;
    if (/\s/u.test(character)) { inWord = false; continue; }
    characters++;
    if (EAST_ASIAN_LETTER.test(character)) { words++; inWord = false; continue; }
    if (EAST_ASIAN_PUNCTUATION.test(character)) { inWord = false; continue; }
    if (!inWord) { words++; inWord = true; }
  }
  return { words, characters, charactersWithSpaces };
}

const STATISTIC_ELEMENTS = [
  ['pages', 'Pages'], ['words', 'Words'], ['characters', 'Characters'], ['lines', 'Lines'],
  ['paragraphs', 'Paragraphs'], ['charactersWithSpaces', 'CharactersWithSpaces'],
] as const;

export function setAppStatisticsOn(document: Document, statistics: DocumentStatistics): void {
  const root = rootOrNull(document, APP_NS, 'Properties');
  if (!root) throw new Error('Invalid app properties root.');
  for (const [key, localName] of STATISTIC_ELEMENTS) {
    const value = statistics[key];
    if (value === undefined) continue;
    setElementText(ensureChild(root, APP_NS, localName, localName, APP_PROPERTY_ORDER), String(value));
  }
}
