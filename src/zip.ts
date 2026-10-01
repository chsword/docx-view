import JSZip from 'jszip';

/**
 * 把部件字节打成 .docx(ZIP)。
 *
 * 单独成一个模块、只依赖字节,是为了让它能在**任何线程**上跑:DEFLATE 占存盘耗时的
 * 37%~67%(1500 段文档实测 92 / 138 ms),把它搬到 Web Worker 能把这部分从主线程挪走。
 * 而序列化必须留在主线程——它要读那棵活的 XML 树,而 xmldom 节点不是结构化可克隆的。
 *
 * 这里**不**替宿主创建 Worker:那要替它决定打包器与 CSP 策略。宿主自己写 worker,
 * 在里面调这个函数,再把 `zip` 注入 `toUint8Array()`。README 里有五行的样例。
 */
export type ZipParts = (parts: ReadonlyMap<string, Uint8Array>) => Promise<Uint8Array>;

export const zipParts: ZipParts = async (parts) => {
  const zip = new JSZip();
  for (const [path, bytes] of parts) zip.file(path, bytes, { createFolders: false });
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
};
