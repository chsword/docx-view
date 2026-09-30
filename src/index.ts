export { DocxDocument } from './document.js';
export { DocxEditor } from './editor.js';
export type { DocxEditorOptions, EditorReviewFilter } from './editor.js';
export { AGENT_OPERATION_SCHEMA } from './operations.js';
export type { FieldInfo, FieldKind, FieldSwitch, MathInfo, MathMlNode, PaginationInfo } from './types.js';
export { columnWidthsPx, effectiveKinsoku, lineNumbersFor, paginate, paragraphSpacingPx, snapLineHeightPx } from './layout.js';
export type {
  FlowItem, LayoutMeasurer, LayoutTable, LineBox, MeasureContext, PageBox, ParagraphMeasureArea, WrapExclusion,
} from './layout.js';
export {
  A_NS, EMU_PER_INCH, IMAGE_REL, PIC_NS, PT_PER_INCH, PX_PER_INCH, V_NS, WP_NS,
  contentTypeForExtension, dataUrlForBytes, decodeBase64, detectImageSize, emuToPt, emuToPx,
  extensionForContentType, isBrowserRenderableContentType, placeholderDataUrl, ptToEmu, pxToEmu,
  resolveRelationshipsPath, resolveTargetPath,
} from './drawing.js';
export { WORD_NS, REL_NS, CONTENT_TYPES_NS, OFFICE_DOCUMENT_REL, OFFICE_REL_NS } from './xml.js';
export { readRunShapes } from './shapes.js';
export { customGeometryPath, presetGeometryPath } from './geometry.js';
export { axisTicks, barRects, pieSlicePath, valueToPx } from './chart.js';
export { MATH_NS, ommlToLinearText, ommlToLinearTextWithInfo, ommlToMathMl, ommlToMathMlWithInfo } from './math.js';
export type { LinearConversion, MathConversion } from './math.js';
export type {
  AgentOperation, AgentRequest, BookmarkInfo, BorderFormat, BordersFormat, CellFormat, CommentAnchor, CommentInfo, ContentControlInfo, ContentControlKind,
  EditableRegionEditorGroup, EditableRegionInfo,
  CompatibilitySettings, DocumentBlock, DocumentProperties, DocumentProtection, DocumentSnapshot, HistoryEntry, HyperlinkInfo, ImageInfo, NoteInfo, NoteSettings, NoteSettingsValue,
  RevisionInfo, RevisionMark, ReviewerAuthorKind, ReviewerFilterAuthor, ReviewerInfo,
  MarginFormat, NumberingDefinition, NumberingInfo, NumberingLevelDefinition, ParagraphFormat, ParagraphInfo,
  RowFormat, RunFormat, RunInfo, Shading, ShadingFormat, StyleInfo, OutlineNode, TabStop, BorderSide, TableCellInfo, TableCellLocation, TableFormat, TableInfo, TableRowInfo,
  TextRange, DocumentRange, ClipboardFragment,
  WidthFormat,
  ChartInfo, CustomGeometry, CustomGeometryCommand, PageSetup, SectionInfo, SectionType, ShapeChildInfo, ShapeInfo, ShapeKind,
} from './types.js';
