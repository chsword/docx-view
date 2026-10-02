/**
 * Strict OOXML（ISO/IEC 29500 Strict）→ Transitional。两者的结构相同，Strict 换了一套命名空间
 * URI（`http://purl.oclc.org/ooxml/…`），关系类型也跟着换，另有几处取值写法不同。打开时整包换成
 * Transitional，之后所有读写都照常走；**存盘是 Transitional**——大多数工具都这么处理，Word 打开
 * 两种都认。
 *
 * 只换精确的 URI 字符串：它们只出现在命名空间声明和关系的 `Type` 里，不会误伤正文。只处理
 * UTF-8 的 XML 部件：UTF-16 部件重新编码会和它的 XML 声明对不上，宁可不动（Word 存 Strict 用的
 * 都是 UTF-8）。
 */
const T = 'http://schemas.openxmlformats.org';
const S = 'http://purl.oclc.org/ooxml';

/** 先放名字本身不同的关系类型，再放按前缀对应的命名空间；替换时按长度从长到短匹配。 */
const STRICT_TO_TRANSITIONAL: Record<string, string> = {
  [`${S}/officeDocument/relationships/extendedProperties`]: `${T}/officeDocument/2006/relationships/extended-properties`,
  [`${S}/officeDocument/relationships/customProperties`]: `${T}/officeDocument/2006/relationships/custom-properties`,
  [`${S}/officeDocument/relationships`]: `${T}/officeDocument/2006/relationships`,
  [`${S}/wordprocessingml/main`]: `${T}/wordprocessingml/2006/main`,
  [`${S}/officeDocument/math`]: `${T}/officeDocument/2006/math`,
  [`${S}/officeDocument/extendedProperties`]: `${T}/officeDocument/2006/extended-properties`,
  [`${S}/officeDocument/customProperties`]: `${T}/officeDocument/2006/custom-properties`,
  [`${S}/officeDocument/docPropsVTypes`]: `${T}/officeDocument/2006/docPropsVTypes`,
  [`${S}/officeDocument/sharedTypes`]: `${T}/officeDocument/2006/sharedTypes`,
  [`${S}/officeDocument/bibliography`]: `${T}/officeDocument/2006/bibliography`,
  [`${S}/officeDocument/customXml`]: `${T}/officeDocument/2006/customXml`,
  [`${S}/drawingml/main`]: `${T}/drawingml/2006/main`,
  [`${S}/drawingml/wordprocessingDrawing`]: `${T}/drawingml/2006/wordprocessingDrawing`,
  [`${S}/drawingml/picture`]: `${T}/drawingml/2006/picture`,
  [`${S}/drawingml/chartDrawing`]: `${T}/drawingml/2006/chartDrawing`,
  [`${S}/drawingml/chart`]: `${T}/drawingml/2006/chart`,
  [`${S}/drawingml/diagram`]: `${T}/drawingml/2006/diagram`,
  [`${S}/drawingml/lockedCanvas`]: `${T}/drawingml/2006/lockedCanvas`,
  [`${S}/drawingml/compatibility`]: `${T}/drawingml/2006/compatibility`,
  [`${S}/schemaLibrary/main`]: `${T}/schemaLibrary/2006/main`,
};

const STRICT_URI = new RegExp(
  Object.keys(STRICT_TO_TRANSITIONAL).sort((a, b) => b.length - a.length).map((uri) => uri.replace(/[.*+?^${}()|[\]\\/]/g, '\\$&')).join('|'),
  'g',
);

/**
 * 原地把包里的 Strict 部件换成 Transitional，返回有没有换过。取值写法上处理最常见的两处：
 * 百分比宽度 `w:w="50%"`（Transitional 是五十分之一个百分点：2500）与字符缩放
 * `<w:w w:val="150%"/>`（Transitional 是 150）。其余 Strict 专有写法按畸形值降级——读不出来
 * 就当没写，不会报错。
 */
export function convertStrictParts(parts: Map<string, Uint8Array>): boolean {
  // 只有包本身是 Strict（根关系里的 officeDocument 用的是 Strict 的关系类型）才换。否则一份
  // 讲 OOXML 的 Transitional 文档，正文里引用了 Strict 的命名空间 URI，就会被改掉正文。
  const decoder = new TextDecoder('utf-8', { fatal: true });
  const rootRels = parts.get('_rels/.rels');
  try {
    if (!rootRels || !decoder.decode(rootRels).includes(`${S}/officeDocument/relationships/officeDocument`)) return false;
  } catch {
    return false;
  }
  let converted = false;
  const encoder = new TextEncoder();
  for (const [path, bytes] of parts) {
    if (!/\.(xml|rels)$/i.test(path)) continue;
    if (bytes[0] === 0xff || bytes[0] === 0xfe) continue;
    let text: string;
    try {
      text = decoder.decode(bytes);
    } catch {
      continue;
    }
    if (!text.includes(`${S}/`)) continue;
    let next = text.replace(STRICT_URI, (uri) => STRICT_TO_TRANSITIONAL[uri]!);
    next = next
      .replace(/\s[\w.-]+:conformance="strict"/g, '')
      .replace(/(<[\w.-]+:(?:tblW|tcW|wBefore|wAfter|tblInd|tblCellSpacing)\b[^>]*?\s[\w.-]+:w=")(\d+(?:\.\d+)?)%"/g,
        (_match, head: string, value: string) => `${head}${Math.round(Number(value) * 50)}"`)
      .replace(/(<[\w.-]+:w\s[^>]*?[\w.-]+:val=")(\d+(?:\.\d+)?)%"/g, (_match, head: string, value: string) => `${head}${Math.round(Number(value))}"`);
    if (next !== text) {
      parts.set(path, encoder.encode(next));
      converted = true;
    }
  }
  return converted;
}
