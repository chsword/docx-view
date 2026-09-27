import type { Element } from '@xmldom/xmldom';
import { REL_NS } from './xml.js';
import { assertText } from './xml.js';

const MAX_LINK_LENGTH = 8_192;
const MAX_ANCHOR_LENGTH = 256;
const MAX_TOOLTIP_LENGTH = 2_048;
const ALLOWED_URL_SCHEMES = new Set(['http', 'https', 'mailto']);
const BLOCKED_URL_SCHEMES = new Set(['javascript', 'data', 'vbscript', 'file']);

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

export function isInternalBookmark(name: string): boolean {
  return /^_GoBack$/i.test(name) || /^_Toc/i.test(name);
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

export function readExternalRelationship(relationships: Element | undefined, id: string): string | undefined {
  if (!relationships) return undefined;
  for (const rel of Array.from(relationships.getElementsByTagNameNS(REL_NS, 'Relationship'))) {
    if (rel.getAttribute('Id') !== id) continue;
    if (rel.getAttribute('TargetMode') !== 'External') return undefined;
    return rel.getAttribute('Target') ?? undefined;
  }
  return undefined;
}
