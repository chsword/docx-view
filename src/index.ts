export { DocxDocument } from './document.js';
export { DocxEditor } from './editor.js';
export type { DocxEditorOptions } from './editor.js';
export { AGENT_OPERATION_SCHEMA } from './operations.js';
export {
  A_NS, EMU_PER_INCH, IMAGE_REL, PIC_NS, PT_PER_INCH, PX_PER_INCH, V_NS, WP_NS,
  contentTypeForExtension, dataUrlForBytes, decodeBase64, detectImageSize, emuToPt, emuToPx,
  extensionForContentType, isBrowserRenderableContentType, placeholderDataUrl, ptToEmu, pxToEmu,
  resolveRelationshipsPath, resolveTargetPath,
} from './drawing.js';
export { WORD_NS, REL_NS, CONTENT_TYPES_NS, OFFICE_DOCUMENT_REL, OFFICE_REL_NS } from './xml.js';
export type {
  AgentOperation, AgentRequest, BookmarkInfo, BorderFormat, BordersFormat, CellFormat, DocumentBlock,
  DocumentSnapshot, HyperlinkInfo, ImageInfo, NoteInfo, NoteSettings, NoteSettingsValue, RevisionInfo, RevisionMark,
  MarginFormat, NumberingDefinition, NumberingInfo, NumberingLevelDefinition, ParagraphFormat, ParagraphInfo,
  RowFormat, RunFormat, RunInfo, Shading, ShadingFormat, StyleInfo, TabStop, BorderSide, TableCellInfo, TableFormat, TableInfo, TableRowInfo,
  TextRange, DocumentRange,
  WidthFormat,
  PageSetup, SectionInfo, SectionType,
} from './types.js';
