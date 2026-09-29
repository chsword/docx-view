import type { Element } from '@xmldom/xmldom';
import { emuToPx, V_NS, WP_NS, A_NS } from './drawing.js';
import type { ShapeInfo, ShapeKind } from './types.js';

const WORD_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const MC_NS = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const WPS_NS = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';
const DIAGRAM_NS = 'http://schemas.openxmlformats.org/drawingml/2006/diagram';
const CHART_NS = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const OLE_NS = 'urn:schemas-microsoft-com:office:office';

function descendants(parent: Element, namespace: string, localName: string): Element[] {
  return Array.from(parent.getElementsByTagNameNS(namespace, localName));
}

function first(parent: Element | undefined, namespace: string, localName: string): Element | undefined {
  return parent ? descendants(parent, namespace, localName)[0] : undefined;
}

function numberAttribute(element: Element | undefined, name: string): number {
  const value = element?.getAttribute(name);
  return value && Number.isFinite(Number(value)) && Number(value) > 0 ? Number(value) : 0;
}

function shapeKind(uri: string | undefined, hasTextContent: boolean): ShapeKind {
  if (uri === WPS_NS || hasTextContent) return 'textbox';
  if (uri === DIAGRAM_NS || uri?.includes('/diagram')) return 'smartArt';
  if (uri === CHART_NS || uri?.includes('/chart')) return 'chart';
  if (uri === OLE_NS || uri?.includes('/ole')) return 'ole';
  return 'shape';
}

function textContent(node: Element): boolean {
  return descendants(node, WORD_NS, 'txbxContent').length > 0;
}

function readDrawingShape(node: Element, paragraph: number, run: number, ordinal: number): ShapeInfo[] {
  const result: ShapeInfo[] = [];
  const containers = [
    ...descendants(node, WP_NS, 'inline'),
    ...descendants(node, WP_NS, 'anchor'),
  ];
  for (const container of containers) {
    const placement = container.localName === 'anchor' ? 'floating' : 'inline';
    const extent = first(container, WP_NS, 'extent');
    const graphicData = first(first(container, A_NS, 'graphic'), A_NS, 'graphicData');
    const chosen = first(graphicData, MC_NS, 'Choice');
    const source = chosen ?? graphicData ?? container;
    const hasTextContent = textContent(source);
    const uri = graphicData?.getAttribute('uri') ?? undefined;
    const docPr = first(container, WP_NS, 'docPr');
    const wps = first(source, WPS_NS, 'wsp');
    const geometry = first(first(wps, WPS_NS, 'spPr'), A_NS, 'prstGeom')?.getAttribute('prst') ?? undefined;
    const wrap = placement === 'floating'
      ? (first(container, WP_NS, 'wrapNone') ? 'none'
        : first(container, WP_NS, 'wrapSquare') ? 'square'
        : first(container, WP_NS, 'wrapTight') ? 'tight'
        : first(container, WP_NS, 'wrapThrough') ? 'through'
        : first(container, WP_NS, 'wrapTopAndBottom') ? 'topAndBottom' : undefined)
      : undefined;
    const id = docPr?.getAttribute('id') ?? `${paragraph}:${run}:${ordinal + result.length}`;
    result.push({
      id,
      paragraph,
      run,
      kind: shapeKind(uri, hasTextContent),
      form: 'drawingml',
      name: docPr?.getAttribute('name') ?? undefined,
      alt: docPr?.getAttribute('descr') ?? undefined,
      title: docPr?.getAttribute('title') ?? undefined,
      widthPx: emuToPx(numberAttribute(extent, 'cx')),
      heightPx: emuToPx(numberAttribute(extent, 'cy')),
      placement,
      wrap,
      hasTextContent,
      geometry,
    });
  }
  return result;
}

function readVmlShape(node: Element, paragraph: number, run: number, ordinal: number): ShapeInfo[] {
  return descendants(node, V_NS, 'shape').map((shape, index) => {
    const style = shape.getAttribute('style') ?? '';
    const width = style.match(/(?:^|;)width:\s*([\d.]+)pt/i);
    const height = style.match(/(?:^|;)height:\s*([\d.]+)pt/i);
    const hasTextContent = textContent(shape);
    return {
      id: shape.getAttribute('id') ?? `${paragraph}:${run}:${ordinal + index}`,
      paragraph,
      run,
      kind: hasTextContent ? 'textbox' : 'shape',
      form: 'vml',
      name: shape.getAttribute('id') ?? undefined,
      alt: shape.getAttribute('alt') ?? undefined,
      title: shape.getAttribute('title') ?? undefined,
      widthPx: width ? Number(width[1]) * 96 / 72 : 0,
      heightPx: height ? Number(height[1]) * 96 / 72 : 0,
      placement: 'inline',
      hasTextContent,
      geometry: shape.getAttribute('type') ?? undefined,
    };
  });
}

export function readRunShapes(runElement: Element, paragraph: number, run: number, sourcePartPath = ''): ShapeInfo[] {
  const result: ShapeInfo[] = [];
  const elements: Element[] = [];
  for (let child = runElement.firstChild; child; child = child.nextSibling) {
    if (child.nodeType !== 1) continue;
    const element = child as Element;
    if (element.namespaceURI === MC_NS && element.localName === 'AlternateContent') {
      const choices = descendants(element, MC_NS, 'Choice');
      const choice = choices.find((candidate) => (candidate.getAttribute('Requires') ?? '').split(/\s+/).includes('wps'))
        ?? choices[0];
      const branch = choice ?? first(element, MC_NS, 'Fallback');
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
    if (element.namespaceURI === WORD_NS && element.localName === 'drawing') {
      result.push(...readDrawingShape(element, paragraph, run, result.length));
    } else if (element.namespaceURI === WORD_NS && element.localName === 'pict') {
      result.push(...readVmlShape(element, paragraph, run, result.length));
    }
  }
  return result.map((shape, index) => ({
    ...shape,
    id: `${sourcePartPath}:${paragraph}:${run}:${index}:${shape.id}`,
  }));
}

export function shapeTextElements(shape: Element): Element[] {
  return descendants(shape, WORD_NS, 'txbxContent')
    .filter((content) => {
      let ancestor = content.parentNode as Element | null;
      let inChoice = false;
      while (ancestor && ancestor !== shape) {
        if (ancestor.namespaceURI === MC_NS && ancestor.localName === 'Choice') inChoice = true;
        if (ancestor.namespaceURI === MC_NS && ancestor.localName === 'Fallback') return false;
        ancestor = ancestor.parentNode as Element | null;
      }
      return inChoice || !descendants(shape, MC_NS, 'Choice').length;
    })
    .flatMap((content) => descendants(content, WORD_NS, 'p'));
}
