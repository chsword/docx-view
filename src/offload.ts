import { DocxDocument } from './document.js';
import type { DocumentSnapshot, TextMatch } from './types.js';
import type { ZipParts } from './zip.js';

/**
 * 可以整个搬进 Worker 的批处理：字节进，纯数据或字节出，两头都过得了结构化克隆。和
 * `zipParts()` 一样，**库不替宿主创建 Worker**——Worker 的创建方式、打包与 CSP 都是宿主决定的；
 * 宿主在 Worker 里 import 这几个函数、用 postMessage 收发即可。
 *
 * 搬不了的：可编辑的文档对象。编辑模型是活的 XML 树，传不过线程边界；分页测量要
 * `getClientRects()`，只有主线程有。
 */

/** 比较两份 .docx，返回带修订的结果文档的字节；主线程拿到后 `DocxDocument.load()`。 */
export async function compareDocxBytes(base: Uint8Array, revised: Uint8Array,
  options: { author?: string; date?: string; zip?: ZipParts } = {}): Promise<Uint8Array> {
  const { zip, ...compareOptions } = options;
  const result = DocxDocument.compare(await DocxDocument.load(base), await DocxDocument.load(revised), compareOptions);
  return result.toUint8Array(zip ? { zip } : {});
}

/** 只读打开：解析整份文档，交回快照（段落、块结构、批注、样式……），不交回可编辑对象。 */
export async function readDocxSnapshot(bytes: Uint8Array): Promise<DocumentSnapshot> {
  return (await DocxDocument.load(bytes)).getSnapshot();
}

/** 全文搜索，见 `DocxDocument.findText()`。 */
export async function searchDocxText(bytes: Uint8Array, query: string,
  options: { caseSensitive?: boolean; maxResults?: number } = {}): Promise<TextMatch[]> {
  return (await DocxDocument.load(bytes)).findText(query, options);
}
