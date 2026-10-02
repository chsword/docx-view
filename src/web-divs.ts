import type { Element } from '@xmldom/xmldom';
import type { BorderFormat, WebDivInfo } from './types.js';
import { children, WORD_NS, wordValue } from './xml.js';

/**
 * `webSettings.xml` 里的 `w:divs`：文档经 HTML 往返（另存为网页、邮件回复）时，Word 把 HTML
 * `<div>` / `<blockquote>` 的嵌套结构记在这里，段落与表格行用 `w:divId` 指向其中一个。
 *
 * 树形结构（`w:divsChild` 里套 `w:div`）摊平成数组，用 `parentId` 还原层级；畸形输入（非数字
 * `w:id`、重复 id、成环的嵌套）不抛错，重复 id 只取第一个。
 */
export function parseWebDivs(webSettings: Element | undefined): WebDivInfo[] {
  if (!webSettings) return [];
  const result: WebDivInfo[] = [];
  const seen = new Set<number>();
  const walk = (container: Element | undefined, parentId: number | undefined, depth: number) => {
    // 嵌套深度与 OMML 一样设上限：真实文档里 div 很少超过十几层，深度只会来自构造的输入。
    if (!container || depth > 64) return;
    for (const div of children(container, 'div')) {
      const id = Number(div.getAttributeNS(WORD_NS, 'id') ?? div.getAttribute('w:id'));
      if (!Number.isSafeInteger(id) || seen.has(id)) continue;
      seen.add(id);
      const twips = (name: string) => {
        const value = Number(wordValue(children(div, name)[0]));
        return Number.isFinite(value) ? value : 0;
      };
      const info: WebDivInfo = {
        id,
        ...(parentId !== undefined ? { parentId } : {}),
        blockQuote: onOff(children(div, 'blockQuote')[0]),
        bodyDiv: onOff(children(div, 'bodyDiv')[0]),
        marginLeft: twips('marLeft'),
        marginRight: twips('marRight'),
        marginTop: twips('marTop'),
        marginBottom: twips('marBottom'),
      };
      const borders = readDivBorders(children(div, 'divBdr')[0]);
      if (borders) info.borders = borders;
      result.push(info);
      walk(children(div, 'divsChild')[0], id, depth + 1);
    }
  };
  walk(children(webSettings, 'divs')[0], undefined, 0);
  return result;
}

function onOff(element: Element | undefined): boolean {
  if (!element) return false;
  const value = wordValue(element);
  return value === undefined || !['0', 'false', 'off'].includes(value.toLowerCase());
}

function readDivBorders(element: Element | undefined): WebDivInfo['borders'] {
  if (!element) return undefined;
  const borders: NonNullable<WebDivInfo['borders']> = {};
  for (const side of ['top', 'left', 'bottom', 'right'] as const) {
    const border = children(element, side)[0];
    if (!border) continue;
    const style = wordValue(border) ?? undefined;
    const size = Number(border.getAttributeNS(WORD_NS, 'sz'));
    const space = Number(border.getAttributeNS(WORD_NS, 'space'));
    const color = border.getAttributeNS(WORD_NS, 'color') ?? undefined;
    const entry: BorderFormat = { style, none: style === 'nil' || style === 'none' };
    if (Number.isFinite(size) && border.hasAttributeNS(WORD_NS, 'sz')) entry.size = size;
    if (Number.isFinite(space) && border.hasAttributeNS(WORD_NS, 'space')) entry.space = space;
    if (color && /^(auto|[0-9a-f]{6})$/i.test(color)) entry.color = color;
    borders[side] = entry;
  }
  return Object.keys(borders).length ? borders : undefined;
}

/**
 * 每个 div 的**累计**左右缩进（缇）：自己的 `marLeft` / `marRight` 加上所有祖先的。Word 把 div 的
 * 边距加在段落自己的缩进之外——邮件里层层引用的回复就是这么一级级往里缩的。
 */
export function webDivIndents(divs: WebDivInfo[]): Map<number, { left: number; right: number }> {
  const byId = new Map(divs.map((div) => [div.id, div]));
  const result = new Map<number, { left: number; right: number }>();
  for (const div of divs) {
    let left = 0;
    let right = 0;
    const seen = new Set<number>();
    for (let cursor: WebDivInfo | undefined = div; cursor && !seen.has(cursor.id);
      seen.add(cursor.id), cursor = cursor.parentId === undefined ? undefined : byId.get(cursor.parentId)) {
      left += cursor.marginLeft;
      right += cursor.marginRight;
    }
    result.set(div.id, { left, right });
  }
  return result;
}
