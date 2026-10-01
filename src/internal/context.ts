import type { Document } from '@xmldom/xmldom';

/**
 * Access to package parts without exposing the backing parts map.
 *
 * 文档访问分两套语义，名字已经写明是哪一套，不要混用：
 *
 * - `partDocumentCopy` / `updatePartXmlFromCopy` / `writePartXml` —— **分离副本**。
 *   副本上的修改不会进包，必须整体写回；每次取副本都要 serialize + parse 一遍。
 * - `livePartDocument` / `mutateLivePartXml` —— **活的缓存文档**。就地修改由
 *   `mutateLivePartXml` 提交；回调返回 `true` 表示已改（走脏标记，省掉一次序列化比对）。
 *
 * 选错不会报错，只会静默出问题：该用活文档而用了副本，回调外的修改丢失；
 * 该用缓存而用了副本读取，每次读都重新解析整个部件（违反约定第 8 条）。
 */
export interface PartAccess {
  readonly mainPath: string;
  readonly revision: number;
  hasPart(path: string): boolean;
  readPart(path: string): Uint8Array | undefined;
  deletePart(path: string): void;
  forgetPartDocument(path: string): void;
  forgetDirtyPartXml(path: string): void;
  forgetDirtyPartSize(path: string): void;
  /** 分离副本；就地修改不会进包。 */
  partDocumentCopy(path: string): Document;
  /** 分离副本；部件不存在时返回 undefined 而不抛错。 */
  partDocumentCopyOrUndefined(path: string): Document | undefined;
  /** 取一份副本交给回调修改，回调返回后整体提交。 */
  updatePartXmlFromCopy(path: string, update: (document: Document) => boolean | void): void;
  /** 整体写回一份在外部构造或修改好的 XML（用于已持有副本的场景）。 */
  writePartXml(path: string, xml: string): void;
  /** 活的缓存文档；就地修改由 mutateLivePartXml 提交。 */
  livePartDocument(path: string): Document;
  /** 在活文档上修改；回调返回 true 表示已改（脏标记），false 表示无改动。 */
  mutateLivePartXml(path: string, update: (document: Document) => boolean | void): void;
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
