import type { Document, Element } from '@xmldom/xmldom';
import type { DocumentProperties } from './types.js';
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

const CORE_KEYS = ['title', 'subject', 'creator', 'lastModifiedBy', 'keywords', 'description', 'category', 'created', 'modified', 'revisionNumber'] as const;
const APP_KEYS = ['company', 'manager'] as const;
const DATE_KEYS = new Set<DocumentPropertyKey>(['created', 'modified']);
const TEXT_KEYS = new Set<DocumentPropertyKey>(['title', 'subject', 'creator', 'lastModifiedBy', 'keywords', 'description', 'category', 'company', 'manager']);
const VALID_KEYS = new Set<DocumentPropertyKey>([...CORE_KEYS, ...APP_KEYS]);
const ISO_8601_RE = /^\d{4}-\d{2}-\d{2}(?:T\d{2}:\d{2}(?::\d{2}(?:\.\d{1,9})?)?(?:Z|[+-]\d{2}:\d{2})?)?$/;

type DocumentPropertyKey = keyof DocumentProperties;

const CORE_FIELD_INFO: Record<typeof CORE_KEYS[number], { namespace: string; qualifiedName: string; localName: string; date?: boolean; numeric?: boolean }> = {
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

const APP_FIELD_INFO: Record<typeof APP_KEYS[number], { qualifiedName: string; localName: string }> = {
  company: { qualifiedName: 'Company', localName: 'Company' },
  manager: { qualifiedName: 'Manager', localName: 'Manager' },
};

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

function ensureChild(parent: Element, namespace: string, qualifiedName: string, localName: string): Element {
  const existing = childInNamespace(parent, namespace, localName);
  if (existing) return existing;
  const element = parent.ownerDocument!.createElementNS(namespace, qualifiedName);
  parent.appendChild(element);
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
    for (const key of CORE_KEYS) {
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
    for (const key of APP_KEYS) {
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
  for (const key of CORE_KEYS) {
    if (!(key in patch)) continue;
    const info = CORE_FIELD_INFO[key];
    const value = patch[key];
    if (value === undefined) {
      removeChild(root, info.namespace, info.localName);
      continue;
    }
    const element = ensureChild(root, info.namespace, info.qualifiedName, info.localName);
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
  for (const key of APP_KEYS) {
    if (!(key in patch)) continue;
    const info = APP_FIELD_INFO[key];
    const value = patch[key];
    if (value === undefined) {
      removeChild(root, APP_NS, info.localName);
      continue;
    }
    setElementText(ensureChild(root, APP_NS, info.qualifiedName, info.localName), value);
  }
}
