import type { Element } from '@xmldom/xmldom';
import { MC_NS, selectAlternateContentBranch, validatePath } from './xml.js';
import type { ImageInfo } from './types.js';

export const OFFICE_REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
export const WP_NS = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
export const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
export const PIC_NS = 'http://schemas.openxmlformats.org/drawingml/2006/picture';
export const V_NS = 'urn:schemas-microsoft-com:vml';

export const IMAGE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
export const EMU_PER_INCH = 914400;
export const PX_PER_INCH = 96;
export const PT_PER_INCH = 72;

export interface RelationshipTarget {
  id: string;
  mode?: string;
  target?: string;
  partPath?: string;
}

const CONTENT_TYPE_EXTENSIONS: Record<string, string> = {
  'image/png': 'png',
  'image/jpeg': 'jpg',
  'image/gif': 'gif',
  'image/bmp': 'bmp',
  'image/tiff': 'tiff',
  'image/x-emf': 'emf',
  'image/x-wmf': 'wmf',
  'image/svg+xml': 'svg',
};

const EXTENSION_CONTENT_TYPES: Record<string, string> = Object.fromEntries(
  Object.entries(CONTENT_TYPE_EXTENSIONS).flatMap(([type, extension]) =>
    type === 'image/jpeg' ? [[extension, type], ['jpeg', type]] : [[extension, type]]),
);
const BASE64_RE = /^(?:[A-Za-z0-9+/]{4})*(?:[A-Za-z0-9+/]{2}==|[A-Za-z0-9+/]{3}=)?$/;

export function emuToPx(emu: number): number {
  return emu * PX_PER_INCH / EMU_PER_INCH;
}

export function pxToEmu(px: number): number {
  return px * EMU_PER_INCH / PX_PER_INCH;
}

export function emuToPt(emu: number): number {
  return emu * PT_PER_INCH / EMU_PER_INCH;
}

export function ptToEmu(pt: number): number {
  return pt * EMU_PER_INCH / PT_PER_INCH;
}

export function extensionForContentType(contentType: string): string | undefined {
  return CONTENT_TYPE_EXTENSIONS[contentType.toLowerCase()];
}

export function contentTypeForExtension(extension: string): string | undefined {
  return EXTENSION_CONTENT_TYPES[extension.replace(/^\./, '').toLowerCase()];
}

export function isBrowserRenderableContentType(contentType?: string): boolean {
  return ['image/png', 'image/jpeg', 'image/gif', 'image/bmp', 'image/svg+xml'].includes(contentType ?? '');
}

export function assertBase64(text: string): void {
  if (typeof text !== 'string' || !text || text.length > 22_500_000 || !BASE64_RE.test(text)) {
    throw new Error('bytes must be a valid base64 string.');
  }
}

export function resolveRelationshipsPath(partPath: string): string {
  validatePath(partPath);
  const slash = partPath.lastIndexOf('/');
  const dir = slash === -1 ? '' : `${partPath.slice(0, slash + 1)}_rels/`;
  const name = slash === -1 ? partPath : partPath.slice(slash + 1);
  return `${dir}${name}.rels`;
}

export function resolveTargetPath(partPath: string, target: string): string | undefined {
  if (!target || /^[a-z][a-z\d+.-]*:/i.test(target)) return undefined;
  const slash = partPath.lastIndexOf('/');
  const base = slash === -1 ? [] : partPath.slice(0, slash).split('/');
  const segments = target.replace(/\\/g, '/').split('/');
  for (const segment of segments) {
    if (!segment || segment === '.') continue;
    if (segment === '..') base.pop();
    else base.push(segment);
  }
  const resolved = base.join('/');
  validatePath(resolved);
  return resolved;
}

export function encodeBase64(bytes: Uint8Array): string {
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

export function decodeBase64(text: string): Uint8Array {
  assertBase64(text);
  const binary = atob(text);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index++) bytes[index] = binary.charCodeAt(index);
  return bytes;
}

export function dataUrlForBytes(bytes: Uint8Array, contentType: string): string {
  return `data:${contentType};base64,${encodeBase64(bytes)}`;
}

export function detectImageContentType(bytes: Uint8Array): string | undefined {
  if (bytes.length >= 8 && bytes.slice(0, 8).every((value, index) => value === [137, 80, 78, 71, 13, 10, 26, 10][index])) {
    return 'image/png';
  }
  if (bytes.length >= 6 && /^GIF8[79]a$/.test(new TextDecoder('ascii').decode(bytes.slice(0, 6)))) {
    return 'image/gif';
  }
  if (bytes.length >= 2 && bytes[0] === 0xff && bytes[1] === 0xd8) {
    return 'image/jpeg';
  }
  if (bytes.length >= 2 && bytes[0] === 0x42 && bytes[1] === 0x4d) {
    return 'image/bmp';
  }
  return undefined;
}

export function placeholderDataUrl(label: string, widthPx = 160, heightPx = 90): string {
  const safe = label.replace(/[&<>"']/g, char => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', '\'': '&apos;',
  }[char]!));
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${Math.max(48, Math.round(widthPx))}" height="${Math.max(48, Math.round(heightPx))}" viewBox="0 0 ${Math.max(48, Math.round(widthPx))} ${Math.max(48, Math.round(heightPx))}"><rect width="100%" height="100%" fill="#f7f9fd" stroke="#c7d3e5" stroke-dasharray="6 4"/><text x="50%" y="50%" fill="#7285a5" font-family="Arial, sans-serif" font-size="12" text-anchor="middle" dominant-baseline="middle">${safe}</text></svg>`;
  return dataUrlForBytes(new TextEncoder().encode(svg), 'image/svg+xml');
}

function numberAttribute(element: Element | undefined, name: string, fallback = 0): number {
  const value = element?.getAttribute(name);
  return value && Number.isFinite(Number(value)) ? Number(value) : fallback;
}

function firstChild(parent: Element | undefined, ns: string, localName: string): Element | undefined {
  if (!parent) return undefined;
  for (let child = parent.firstChild; child; child = child.nextSibling) {
    if (child.nodeType === 1) {
      const element = child as Element;
      if (element.namespaceURI === ns && element.localName === localName) return element;
    }
  }
  return undefined;
}

function descendantsByTag(parent: Element, ns: string, localName: string): Element[] {
  return Array.from(parent.getElementsByTagNameNS(ns, localName));
}

function readCrop(srcRect: Element | undefined): ImageInfo['crop'] | undefined {
  if (!srcRect) return undefined;
  const left = numberAttribute(srcRect, 'l');
  const top = numberAttribute(srcRect, 't');
  const right = numberAttribute(srcRect, 'r');
  const bottom = numberAttribute(srcRect, 'b');
  if (!left && !top && !right && !bottom) return undefined;
  return { left: left / 100000, top: top / 100000, right: right / 100000, bottom: bottom / 100000 };
}

function clampEmu(value: number): number {
  return Number.isFinite(value) && value > 0 ? value : 0;
}

function readDocPr(container: Element | undefined): Pick<ImageInfo, 'name' | 'alt' | 'title'> {
  return {
    name: container?.getAttribute('name') ?? undefined,
    alt: container?.getAttribute('descr') ?? undefined,
    title: container?.getAttribute('title') ?? undefined,
  };
}

function readDrawingImage(
  node: Element,
  paragraph: number,
  run: number,
  relationships: Map<string, RelationshipTarget>,
  getContentType: (path: string) => string | undefined,
): ImageInfo[] {
  const result: ImageInfo[] = [];
  for (const container of [...descendantsByTag(node, WP_NS, 'inline'), ...descendantsByTag(node, WP_NS, 'anchor')]) {
    const placement = container.localName === 'anchor' ? 'floating' : 'inline';
    const extent = firstChild(container, WP_NS, 'extent');
    const docPr = firstChild(container, WP_NS, 'docPr');
    const wrap = placement === 'floating'
      ? (firstChild(container, WP_NS, 'wrapNone') ? 'none'
        : firstChild(container, WP_NS, 'wrapSquare') ? 'square'
        : firstChild(container, WP_NS, 'wrapTight') ? 'tight'
        : firstChild(container, WP_NS, 'wrapThrough') ? 'through'
        : firstChild(container, WP_NS, 'wrapTopAndBottom') ? 'topAndBottom'
        : undefined)
      : undefined;
    const graphic = firstChild(container, A_NS, 'graphic') ?? descendantsByTag(container, A_NS, 'graphic')[0];
    const graphicData = firstChild(graphic, A_NS, 'graphicData');
    const pic = firstChild(graphicData, PIC_NS, 'pic');
    const blip = firstChild(firstChild(pic, PIC_NS, 'blipFill'), A_NS, 'blip');
    const embed = blip?.getAttributeNS(OFFICE_REL_NS, 'embed') ?? blip?.getAttribute('r:embed') ?? undefined;
    const link = blip?.getAttributeNS(OFFICE_REL_NS, 'link') ?? blip?.getAttribute('r:link') ?? undefined;
    const relationshipId = embed ?? link;
    if (!relationshipId) continue;
    const relation = relationships.get(relationshipId);
    const isExternal = relation?.mode === 'External' || (!embed && Boolean(link));
    const xfrm = firstChild(firstChild(pic, PIC_NS, 'spPr'), A_NS, 'xfrm');
    const xfrmExtent = firstChild(xfrm, A_NS, 'ext');
    const widthEmu = clampEmu(numberAttribute(extent, 'cx') || numberAttribute(xfrmExtent, 'cx'));
    const heightEmu = clampEmu(numberAttribute(extent, 'cy') || numberAttribute(xfrmExtent, 'cy'));
    const rotation = numberAttribute(xfrm, 'rot') / 60000 || undefined;
    const contentType = relation?.partPath ? getContentType(relation.partPath) : undefined;
    const crop = readCrop(firstChild(firstChild(pic, PIC_NS, 'blipFill'), A_NS, 'srcRect'));
    result.push({
      id: '',
      paragraph,
      run,
      relationshipId,
      sourcePartPath: '',
      partPath: relation?.partPath,
      contentType,
      widthEmu,
      heightEmu,
      widthPx: emuToPx(widthEmu),
      heightPx: emuToPx(heightEmu),
      placement,
      wrap,
      rotation,
      flipH: xfrm?.getAttribute('flipH') === '1' || xfrm?.getAttribute('flipH') === 'true',
      flipV: xfrm?.getAttribute('flipV') === '1' || xfrm?.getAttribute('flipV') === 'true',
      crop,
      isExternal,
      ...readDocPr(docPr),
      behindDoc: container.localName === 'anchor' && ['1', 'true'].includes(container.getAttribute('behindDoc') ?? ''),
    });
  }
  return result;
}

function readVmlImage(
  node: Element,
  paragraph: number,
  run: number,
  relationships: Map<string, RelationshipTarget>,
  getContentType: (path: string) => string | undefined,
): ImageInfo[] {
  return descendantsByTag(node, V_NS, 'imagedata').flatMap((imageData) => {
    const relationshipId = imageData.getAttributeNS(OFFICE_REL_NS, 'id') ?? imageData.getAttribute('r:id') ?? undefined;
    if (!relationshipId) return [];
    const relation = relationships.get(relationshipId);
    const shape = imageData.parentNode as Element | null;
    const style = shape?.getAttribute('style');
    const widthMatch = style?.match(/width:([\d.]+)pt/i);
    const heightMatch = style?.match(/height:([\d.]+)pt/i);
    const widthEmu = clampEmu(widthMatch ? ptToEmu(Number(widthMatch[1])) : 0);
    const heightEmu = clampEmu(heightMatch ? ptToEmu(Number(heightMatch[1])) : 0);
    return [{
      id: '',
      paragraph,
      run,
      sourcePartPath: '',
      relationshipId,
      partPath: relation?.partPath,
      contentType: relation?.partPath ? getContentType(relation.partPath) : undefined,
      widthEmu,
      heightEmu,
      widthPx: emuToPx(widthEmu),
      heightPx: emuToPx(heightEmu),
      placement: 'inline',
      isExternal: relation?.mode === 'External',
      name: shape?.getAttribute('id') ?? undefined,
      alt: shape?.getAttribute('alt') ?? undefined,
      title: shape?.getAttribute('title') ?? undefined,
    } satisfies ImageInfo];
  });
}

export function readRunImages(
  runElement: Element,
  paragraph: number,
  run: number,
  relationships: Map<string, RelationshipTarget>,
  getContentType: (path: string) => string | undefined,
  sourcePartPath = '',
): ImageInfo[] {
  const images: ImageInfo[] = [];
  const elements: Element[] = [];
  for (let child = runElement.firstChild; child; child = child.nextSibling) {
    if (child.nodeType !== 1) continue;
    const element = child as Element;
    if (element.namespaceURI === MC_NS && element.localName === 'AlternateContent') {
      const branch = selectAlternateContentBranch(element);
      if (branch) {
        for (let branchChild = branch.firstChild; branchChild; branchChild = branchChild.nextSibling) {
          if (branchChild.nodeType === 1) elements.push(branchChild as Element);
        }
      }
      continue;
    }
    elements.push(element);
  }
  for (const element of elements) {
    if (element.namespaceURI === 'http://schemas.openxmlformats.org/wordprocessingml/2006/main' && element.localName === 'drawing') {
      images.push(...readDrawingImage(element, paragraph, run, relationships, getContentType));
    }
    if (element.namespaceURI === 'http://schemas.openxmlformats.org/wordprocessingml/2006/main' && element.localName === 'pict') {
      images.push(...readVmlImage(element, paragraph, run, relationships, getContentType));
    }
  }
  return images.map((image, ordinal) => ({
    ...image,
    ordinal,
    sourcePartPath,
    id: `${sourcePartPath}:${paragraph}:${run}:${ordinal}:${image.relationshipId}`,
  }));
}

export function detectImageSize(bytes: Uint8Array, contentType?: string): { width: number; height: number } | null {
  const type = contentType?.toLowerCase();
  if (type === 'image/png' || bytes.slice(0, 8).every((value, index) => value === [137, 80, 78, 71, 13, 10, 26, 10][index])) {
    if (bytes.length < 24) return null;
    const width = (bytes[16]! << 24) | (bytes[17]! << 16) | (bytes[18]! << 8) | bytes[19]!;
    const height = (bytes[20]! << 24) | (bytes[21]! << 16) | (bytes[22]! << 8) | bytes[23]!;
    return width > 0 && height > 0 ? { width, height } : null;
  }
  if (type === 'image/gif' || /^GIF8[79]a$/.test(new TextDecoder('ascii').decode(bytes.slice(0, 6)))) {
    if (bytes.length < 10) return null;
    const width = bytes[6]! | (bytes[7]! << 8);
    const height = bytes[8]! | (bytes[9]! << 8);
    return width > 0 && height > 0 ? { width, height } : null;
  }
  if (type === 'image/jpeg' || (bytes[0] === 0xff && bytes[1] === 0xd8)) {
    let offset = 2;
    while (offset + 9 < bytes.length) {
      if (bytes[offset] !== 0xff) { offset++; continue; }
      const marker = bytes[offset + 1]!;
      const length = (bytes[offset + 2]! << 8) | bytes[offset + 3]!;
      if ([0xc0, 0xc1, 0xc2, 0xc3, 0xc5, 0xc6, 0xc7, 0xc9, 0xca, 0xcb, 0xcd, 0xce, 0xcf].includes(marker)) {
        const height = (bytes[offset + 5]! << 8) | bytes[offset + 6]!;
        const width = (bytes[offset + 7]! << 8) | bytes[offset + 8]!;
        return width > 0 && height > 0 ? { width, height } : null;
      }
      offset += 2 + length;
    }
    return null;
  }
  if (type === 'image/bmp' || (bytes[0] === 0x42 && bytes[1] === 0x4d)) {
    if (bytes.length < 26) return null;
    const width = bytes[18]! | (bytes[19]! << 8) | (bytes[20]! << 16) | (bytes[21]! << 24);
    const height = Math.abs(bytes[22]! | (bytes[23]! << 8) | (bytes[24]! << 16) | (bytes[25]! << 24));
    return width > 0 && height > 0 ? { width, height } : null;
  }
  return null;
}
