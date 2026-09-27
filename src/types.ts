export interface RunFormat {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  fontSize?: number;
  fontFamily?: string;
  color?: string;
}

export interface ParagraphFormat {
  alignment?: 'left' | 'center' | 'right' | 'both';
  style?: string;
}

export interface RunInfo extends RunFormat {
  index: number;
  text: string;
  image?: ImageInfo;
}

export interface ParagraphInfo extends ParagraphFormat {
  index: number;
  text: string;
  runs: RunInfo[];
  images: ImageInfo[];
}

export interface ImageInfo {
  id: string;
  paragraph: number;
  run: number;
  ordinal?: number;
  sourcePartPath?: string;
  relationshipId: string;
  partPath?: string;
  contentType?: string;
  widthEmu: number;
  heightEmu: number;
  widthPx: number;
  heightPx: number;
  name?: string;
  alt?: string;
  title?: string;
  placement: 'inline' | 'floating';
  wrap?: 'none' | 'square' | 'tight' | 'through' | 'topAndBottom';
  rotation?: number;
  flipH?: boolean;
  flipV?: boolean;
  crop?: { left: number; top: number; right: number; bottom: number };
  isExternal: boolean;
  behindDoc?: boolean;
}

export type DocumentBlock =
  | { type: 'paragraph'; paragraph: ParagraphInfo }
  | { type: 'table'; rows: { cells: { blocks: DocumentBlock[] }[] }[] };

export interface DocumentSnapshot {
  revision: number;
  paragraphs: ParagraphInfo[];
  blocks: DocumentBlock[];
  parts: string[];
}

export type AgentOperation =
  | { type: 'setParagraphText'; index: number; text: string }
  | { type: 'insertParagraph'; text: string; before?: number }
  | { type: 'deleteParagraph'; index: number }
  | { type: 'formatParagraph'; index: number; format: ParagraphFormat }
  | { type: 'formatRun'; paragraph: number; run: number; format: RunFormat }
  | { type: 'replaceText'; search: string; replacement: string }
  | { type: 'insertTable'; rows: string[][] }
  | { type: 'insertImage'; bytes: string; contentType: string; paragraph?: number; run?: number; widthEmu?: number; heightEmu?: number; alt?: string; placement?: 'inline' | 'floating' }
  | { type: 'replaceImageBytes'; image: string; bytes: string; contentType?: string }
  | { type: 'resizeImage'; image: string; size: { widthEmu?: number; heightEmu?: number; keepAspect?: boolean } }
  | { type: 'setImageAlt'; image: string; alt: string; title?: string }
  | { type: 'deleteImage'; image: string }
  | { type: 'setPartXml'; path: string; xml: string };

export interface AgentRequest {
  expectedRevision?: number;
  operations: AgentOperation[];
}
