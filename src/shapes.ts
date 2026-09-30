import type { Document, Element } from '@xmldom/xmldom';
import { emuToPx, V_NS, WP_NS, A_NS, OFFICE_REL_NS } from './drawing.js';
import type { RelationshipTarget } from './drawing.js';
import type { CustomGeometry, CustomGeometryCommand, ShapeChildInfo, ShapeInfo, ShapeKind } from './types.js';
import { MC_NS, selectAlternateContentBranch } from './xml.js';
import { resolveDrawingColor, type ThemeInfo } from './styles.js';

const WORD_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const WPS_NS = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';
const DIAGRAM_NS = 'http://schemas.openxmlformats.org/drawingml/2006/diagram';
const DSP_NS = 'http://schemas.microsoft.com/office/drawing/2008/diagram';
const CHART_NS = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const OLE_NS = 'urn:schemas-microsoft-com:office:office';

function descendants(parent: Element, namespace: string, localName: string): Element[] {
  return Array.from(parent.getElementsByTagNameNS(namespace, localName));
}

function descendantsInNamespace(parent: Document, namespace: string, localName: string): Element[] {
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

function cssColor(value: string | undefined): string | undefined {
  if (!value) return undefined;
  if (/^#[\da-f]{3}(?:[\da-f]{3})?$/i.test(value)) return value;
  return /^(?:black|white|red|green|blue|yellow|gray|grey|silver|navy|teal|olive|maroon|purple|fuchsia|lime|aqua)$/i.test(value)
    ? value.toLowerCase() : undefined;
}

function colorWithAlpha(theme: ThemeInfo, color: Element | undefined): string | undefined {
  const hex = resolveDrawingColor(theme, color);
  if (!hex) return undefined;
  const alpha = color ? descendants(color, A_NS, 'alpha')[0]?.getAttribute('val') : undefined;
  if (alpha === undefined || !Number.isFinite(Number(alpha))) return `#${hex}`;
  const opacity = Math.max(0, Math.min(1, Number(alpha) / 100000));
  const [r, g, b] = hex.match(/../g)!.map((channel: string) => parseInt(channel, 16));
  return `rgba(${r}, ${g}, ${b}, ${opacity})`;
}

function readShapeAppearance(spPr: Element | undefined, theme: ThemeInfo, relationships: Map<string, RelationshipTarget>, style?: Element): Pick<ShapeInfo, 'fill' | 'line' | 'geometry' | 'customGeometry'> {
  const presetGeom = first(spPr, A_NS, 'prstGeom');
  const geometry = presetGeom?.getAttribute('prst') ?? undefined;
  const fill = readDrawingFill(spPr, theme, relationships, style);
  const line = readLine(spPr, theme, style);
  const customGeometry = parseCustomGeometry(spPr);
  return { ...(geometry ? { geometry } : {}), ...(fill ? { fill } : {}), ...(line ? { line } : {}), ...(customGeometry ? { customGeometry } : {}) };
}

function readDrawingFill(spPr: Element | undefined, theme: ThemeInfo, relationships: Map<string, RelationshipTarget>, style?: Element): NonNullable<ShapeInfo['fill']> | undefined {
  const fillNode = spPr && Array.from(spPr.childNodes).find((child) =>
    child.nodeType === 1 && (child as Element).namespaceURI === A_NS &&
    ['noFill', 'solidFill', 'gradFill', 'blipFill'].includes((child as Element).localName ?? '')) as Element | undefined;
  if (!fillNode) {
    const fillRef = style && descendants(style, A_NS, 'fillRef')[0];
    const color = fillRef && (descendants(fillRef, A_NS, 'schemeClr')[0] ?? descendants(fillRef, A_NS, 'srgbClr')[0]);
    const resolved = colorWithAlpha(theme, color);
    return resolved ? { type: 'solid', color: resolved } : undefined;
  }
  if (fillNode.localName === 'noFill') return { type: 'none' };
  if (fillNode.localName === 'solidFill') return { type: 'solid', color: colorWithAlpha(theme, descendants(fillNode, A_NS, 'srgbClr')[0] ?? descendants(fillNode, A_NS, 'schemeClr')[0]) };
  if (fillNode.localName === 'gradFill') {
    const stops = descendants(fillNode, A_NS, 'gs').flatMap((stop) => {
      const color = descendants(stop, A_NS, 'srgbClr')[0] ?? descendants(stop, A_NS, 'schemeClr')[0];
      const position = Number(stop.getAttribute('pos'));
      const resolved = colorWithAlpha(theme, color);
      return Number.isFinite(position) && resolved ? [{ position: Math.max(0, Math.min(1, position / 100000)), color: resolved }] : [];
    });
    const angle = Number(descendants(fillNode, A_NS, 'lin')[0]?.getAttribute('ang'));
    return { type: 'gradient', stops, ...(Number.isFinite(angle) ? { angle: angle / 60000 } : {}) };
  }
  const blip = descendants(fillNode, A_NS, 'blip')[0];
  const relationshipId = blip?.getAttributeNS(OFFICE_REL_NS, 'embed');
  const relationship = relationshipId ? relationships.get(relationshipId) : undefined;
  return {
    type: 'picture',
    ...(relationship?.mode !== 'External' && relationship?.partPath ? { imagePartPath: relationship.partPath } : {}),
  };
}

function readLine(spPr: Element | undefined, theme: ThemeInfo, style?: Element): ShapeInfo['line'] {
  const line = spPr && Array.from(spPr.childNodes).find((child) =>
    child.nodeType === 1 && (child as Element).namespaceURI === A_NS && (child as Element).localName === 'ln') as Element | undefined;
  const lineRef = !line && style ? descendants(style, A_NS, 'lnRef')[0] : undefined;
  if (!line && !lineRef) return undefined;
  const color = line
    ? descendants(line, A_NS, 'srgbClr')[0] ?? descendants(line, A_NS, 'schemeClr')[0]
    : lineRef && (descendants(lineRef, A_NS, 'schemeClr')[0] ?? descendants(lineRef, A_NS, 'srgbClr')[0]);
  const width = line ? Number(line.getAttribute('w')) : Number.NaN;
  const dashName = line ? descendants(line, A_NS, 'prstDash')[0]?.getAttribute('val') : undefined;
  const dashes: Record<string, string> = {
    dash: '6 3', dashDot: '6 3 1 3', dot: '1 3', lgDash: '10 3', lgDashDot: '10 3 1 3',
    lgDashDotDot: '10 3 1 3 1 3', sysDash: '4 2', sysDashDot: '4 2 1 2', sysDot: '1 2',
  };
  return {
    ...(colorWithAlpha(theme, color) ? { color: colorWithAlpha(theme, color) } : {}),
    ...(line && Number.isFinite(width) && width >= 0 ? { widthPx: emuToPx(width) } : {}),
    ...(dashName && dashes[dashName] ? { dash: dashes[dashName] } : {}),
  };
}

function parseCustomGeometry(spPr: Element | undefined): CustomGeometry | undefined {
  const custom = spPr && descendants(spPr, A_NS, 'custGeom')[0];
  const path = custom && descendants(custom, A_NS, 'path')[0];
  if (!path) return undefined;
  const width = Number(path.getAttribute('w'));
  const height = Number(path.getAttribute('h'));
  if (!Number.isFinite(width) || !Number.isFinite(height) || width <= 0 || height <= 0) return undefined;
  const commands: CustomGeometryCommand[] = [];
  for (let child = path.firstChild; child; child = child.nextSibling) {
    if (child.nodeType !== 1) continue;
    const command = child as Element;
    if (command.namespaceURI !== A_NS) continue;
    if (command.localName === 'close') {
      commands.push({ type: 'close' });
      continue;
    }
    if (!['moveTo', 'lnTo', 'cubicBezTo'].includes(command.localName ?? '')) return undefined;
    const points = descendants(command, A_NS, 'pt').map((point) => [Number(point.getAttribute('x')), Number(point.getAttribute('y'))]);
    if (points.some(([x, y]) => !Number.isFinite(x) || !Number.isFinite(y))) return undefined;
    if (command.localName === 'cubicBezTo' && points.length === 3) {
      commands.push({ type: 'cubicBezTo', x1: points[0]![0]!, y1: points[0]![1]!, x2: points[1]![0]!, y2: points[1]![1]!, x: points[2]![0]!, y: points[2]![1]! });
    } else if (['moveTo', 'lnTo'].includes(command.localName ?? '') && points.length === 1) {
      commands.push({ type: command.localName as 'moveTo' | 'lnTo', x: points[0]![0]!, y: points[0]![1]! });
    } else return undefined;
  }
  return commands.length ? { width, height, commands } : undefined;
}

function readSmartArtChildren(
  graphicData: Element | undefined,
  relationships: Map<string, RelationshipTarget>,
  theme: ThemeInfo,
  getPartDocument?: (path: string) => Document | undefined,
  getPartRelationships?: (path: string) => Map<string, RelationshipTarget>,
): ShapeChildInfo[] | undefined {
  if (!graphicData || !getPartDocument) return undefined;
  const relIds = first(graphicData, DIAGRAM_NS, 'relIds');
  const dataId = relIds?.getAttributeNS(OFFICE_REL_NS, 'dm') ?? relIds?.getAttribute('r:dm');
  const dataPath = dataId ? relationships.get(dataId)?.partPath : undefined;
  const dataDocument = dataPath ? getPartDocument(dataPath) : undefined;
  const dataModelExt = dataDocument
    ? descendantsInNamespace(dataDocument, DSP_NS, 'dataModelExt')[0]
    : undefined;
  const drawingId = dataModelExt?.getAttribute('relId')
    ?? dataModelExt?.getAttributeNS(OFFICE_REL_NS, 'id')
    ?? dataModelExt?.getAttribute('r:id');
  const relation = (drawingId ? relationships.get(drawingId) : undefined)
    ?? (dataPath ? [...(getPartRelationships?.(dataPath) ?? new Map()).values()].find((entry) =>
      entry.type?.endsWith('/diagramDrawing')) : undefined)
    ?? (() => {
      const candidates = [...relationships.values()].filter((entry) =>
        entry.type?.endsWith('/diagramDrawing') ||
        entry.partPath && /diagramDrawing/i.test(entry.target ?? ''));
      return candidates.length === 1 ? candidates[0] : undefined;
    })();
  const drawingPath = relation?.partPath;
  if (!drawingPath) return undefined;
  const drawing = getPartDocument(drawingPath);
  if (!drawing) return undefined;
  const drawingRelationships = getPartRelationships?.(drawingPath) ?? new Map();
  const shapes = descendants(drawing.documentElement!, DSP_NS, 'sp');
  if (!shapes.length) return undefined;
  return shapes.map((shape) => {
    const spPr = first(shape, DSP_NS, 'spPr');
    const style = first(shape, DSP_NS, 'style');
    const appearance = readShapeAppearance(spPr, theme, drawingRelationships, style);
    const xfrm = first(spPr, A_NS, 'xfrm');
    const off = first(xfrm, A_NS, 'off');
    const ext = first(xfrm, A_NS, 'ext');
    const text = descendants(shape, A_NS, 't').map((node) => node.textContent ?? '').join('');
    const rotation = Number(xfrm?.getAttribute('rot'));
    return {
      offsetXPx: emuToPx(numberAttribute(off, 'x')),
      offsetYPx: emuToPx(numberAttribute(off, 'y')),
      widthPx: emuToPx(numberAttribute(ext, 'cx')),
      heightPx: emuToPx(numberAttribute(ext, 'cy')),
      ...appearance,
      ...(Number.isFinite(rotation) && rotation ? { rotation: rotation / 60000 } : {}),
      ...(xfrm?.getAttribute('flipH') === '1' || xfrm?.getAttribute('flipH') === 'true' ? { flipH: true } : {}),
      ...(xfrm?.getAttribute('flipV') === '1' || xfrm?.getAttribute('flipV') === 'true' ? { flipV: true } : {}),
      ...(text ? { text } : {}),
    };
  });
}

function readDrawingShape(
  node: Element,
  paragraph: number,
  run: number,
  ordinal: number,
  relationships: Map<string, RelationshipTarget>,
  theme: ThemeInfo,
  getPartDocument?: (path: string) => Document | undefined,
  getPartRelationships?: (path: string) => Map<string, RelationshipTarget>,
): ShapeInfo[] {
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
    const spPr = first(wps, WPS_NS, 'spPr');
    const shapeStyle = first(wps, WPS_NS, 'style');
    const xfrm = first(spPr, A_NS, 'xfrm');
    const rotation = Number(xfrm?.getAttribute('rot'));
    const appearance = readShapeAppearance(spPr, theme, relationships, shapeStyle);
    const presetGeom = first(spPr, A_NS, 'prstGeom');
    const adjustments = first(presetGeom, A_NS, 'avLst');
    const parsedAdjustments = adjustments ? descendants(adjustments, A_NS, 'gd').flatMap((gd) => {
      const value = Number(gd.getAttribute('fmla')?.replace(/^val\s+/, ''));
      return gd.getAttribute('name') && Number.isFinite(value) ? [{ name: gd.getAttribute('name')!, value: value / 100000 }] : [];
    }) : [];
    const wrap = placement === 'floating'
      ? (first(container, WP_NS, 'wrapNone') ? 'none'
        : first(container, WP_NS, 'wrapSquare') ? 'square'
        : first(container, WP_NS, 'wrapTight') ? 'tight'
        : first(container, WP_NS, 'wrapThrough') ? 'through'
        : first(container, WP_NS, 'wrapTopAndBottom') ? 'topAndBottom' : undefined)
      : undefined;
    const id = docPr?.getAttribute('id') ?? `${paragraph}:${run}:${ordinal + result.length}`;
    const fill = appearance.fill;
    const line = appearance.line;
    const hasAppearance = fill !== undefined || line !== undefined;
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
      ...appearance,
      ...(fill ? { fill } : hasAppearance ? {} : { fill: { type: 'solid' as const, color: '#f7f9fd' } }),
      ...(line ? { line } : hasAppearance ? {} : { line: { color: '#c7d3e5', widthPx: 1 } }),
      ...(Number.isFinite(rotation) && rotation ? { rotation: rotation / 60000 } : {}),
      ...(xfrm?.getAttribute('flipH') === '1' || xfrm?.getAttribute('flipH') === 'true' ? { flipH: true } : {}),
      ...(xfrm?.getAttribute('flipV') === '1' || xfrm?.getAttribute('flipV') === 'true' ? { flipV: true } : {}),
      ...(parsedAdjustments.length ? { adjustments: parsedAdjustments } : {}),
      ...(appearance.customGeometry ? { customGeometry: appearance.customGeometry } : {}),
      ...(shapeKind(uri, hasTextContent) === 'smartArt'
        ? { children: readSmartArtChildren(graphicData, relationships, theme, getPartDocument, getPartRelationships) }
        : {}),
    });
  }
  return result;
}

function vmlPreset(type: string | undefined): string | undefined {
  if (!type) return undefined;
  const simple: Record<string, string> = { rect: 'rect', roundRect: 'roundRect', ellipse: 'ellipse', oval: 'ellipse', line: 'line', rightArrow: 'rightArrow' };
  if (simple[type]) return simple[type];
  const code = Number(type.match(/_x0000_t(\d+)/i)?.[1]);
  return ({ 1: 'rect', 2: 'roundRect', 3: 'ellipse', 4: 'diamond', 5: 'line', 202: 'wedgeRectCallout', 203: 'cloudCallout' } as Record<number, string>)[code];
}

function readVmlShape(node: Element, paragraph: number, run: number, ordinal: number): ShapeInfo[] {
  const shapes = Array.from(node.getElementsByTagNameNS(V_NS, '*'))
    .filter((shape) => ['shape', 'rect', 'oval', 'line'].includes(shape.localName ?? ''));
  return shapes.map((shape, index) => {
    const style = shape.getAttribute('style') ?? '';
    const width = style.match(/(?:^|;)width:\s*([\d.]+)pt/i);
    const height = style.match(/(?:^|;)height:\s*([\d.]+)pt/i);
    const hasTextContent = textContent(shape);
    const type = shape.getAttribute('type') || shape.localName || undefined;
    const fillColor = cssColor(shape.getAttribute('fillcolor') ?? descendants(shape, V_NS, 'fill')[0]?.getAttribute('color') ?? undefined);
    const strokeColor = cssColor(shape.getAttribute('strokecolor') ?? descendants(shape, V_NS, 'stroke')[0]?.getAttribute('color') ?? undefined);
    const strokeWeight = (shape.getAttribute('strokeweight') ?? descendants(shape, V_NS, 'stroke')[0]?.getAttribute('weight') ?? '').match(/^([\d.]+)(pt|px)?$/i);
    const customPath = descendants(shape, V_NS, 'path').length > 0;
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
      geometry: customPath ? undefined : vmlPreset(type),
      fill: { type: fillColor ? 'solid' : 'none', ...(fillColor ? { color: fillColor } : {}) },
      line: {
        ...(strokeColor ? { color: strokeColor } : {}),
        ...(strokeWeight ? { widthPx: Number(strokeWeight[1]) * (strokeWeight[2]?.toLowerCase() === 'pt' ? 96 / 72 : 1) } : {}),
      },
    };
  });
}

export function readRunShapes(
  runElement: Element,
  paragraph: number,
  run: number,
  sourcePartPath = '',
  relationships: Map<string, RelationshipTarget> = new Map(),
  theme: ThemeInfo = { colors: {}, fonts: {} },
  getPartDocument?: (path: string) => Document | undefined,
  getPartRelationships?: (path: string) => Map<string, RelationshipTarget>,
): ShapeInfo[] {
  const result: ShapeInfo[] = [];
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
    if (element.namespaceURI === WORD_NS && element.localName === 'drawing') {
      result.push(...readDrawingShape(element, paragraph, run, result.length, relationships, theme, getPartDocument, getPartRelationships));
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
