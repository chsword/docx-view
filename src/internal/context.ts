import type { Document } from '@xmldom/xmldom';

/** Access to package parts without exposing the backing parts map. */
export interface PartAccess {
  readonly mainPath: string;
  readonly revision: number;
  hasPart(path: string): boolean;
  readPart(path: string): Uint8Array | undefined;
  deletePart(path: string): void;
  forgetPartDocument(path: string): void;
  forgetDirtyPartXml(path: string): void;
  forgetDirtyPartSize(path: string): void;
  getPartDocument(path: string): Document;
  partDocumentOrUndefined(path: string): Document | undefined;
  updatePartXml(path: string, update: (document: Document) => boolean | void): void;
  addPart(path: string, bytes: Uint8Array, contentType: string): void;
  relatedPartPathFor(sourcePartPath: string, relationType: string): string | undefined;
  contentPartPaths(): string[];
  ensurePartRelationship(sourcePartPath: string, relationType: string, targetPath: string): void;
}

/** The ten revision-scoped document caches; their invalidation remains with the owner. */
export interface CacheBundle {
  numberingContextCache?: unknown;
  stylesCache?: unknown;
  outlineCache?: unknown;
  noteStateCache?: unknown;
  commentStateCache?: unknown;
  contentPartPathsCache?: unknown;
  commentBindingsCache?: unknown;
  revisionInfoCache?: unknown;
  reviewerInfoCache?: unknown;
  tableCellLocationCache?: unknown;
}

/** Records a single draft transaction, including its existing history behavior. */
export interface HistoryRecorder<Context> {
  withDraft<T>(action: (draft: Context) => T): T;
}
