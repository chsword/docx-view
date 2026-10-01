import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';

import { DocxDocument } from '../dist/document.js';
const WORD_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const source = (path) => readFileSync(new URL(path, import.meta.url), 'utf8');
const imports = (text) => [...text.matchAll(/from ['"]([^'"]+)['"]/g)].map((match) => match[1]);

test('the bookmark module stays a leaf with only its declared dependencies', () => {
  const text = source('../src/bookmark.ts');
  assert.deepEqual(imports(text), [
    '@xmldom/xmldom',
    './types.js',
    './internal/context.js',
    './operations.js',
    './xml.js',
    './internal/elements.js',
  ]);
  // 书签只用副本语义：回调外只读，所有修改都在 updatePartXmlFromCopy 的回调里。
  assert.match(text, /interface BookmarkContext extends Pick<PartAccess, 'mainPath' \| 'partDocumentCopy' \| 'updatePartXmlFromCopy'>/);
  assert.match(text, /getBookmarks: Bound<typeof getBookmarks>/);
});

test('the content-control module stays a leaf with only its declared dependencies', () => {
  const text = source('../src/content-control.ts');
  assert.deepEqual(imports(text), [
    '@xmldom/xmldom',
    './types.js',
    './internal/context.js',
    './operations.js',
    './xml.js',
    './internal/elements.js',
    './revisions.js',
  ]);
  // 内容控件要活文档：读取不能每次重新解析（约定第 8 条），写入回调靠 return true 走脏标记。
  assert.match(text, /'mainPath' \| 'livePartDocument' \| 'mutateLivePartXml'/);
  assert.doesNotMatch(text, /livePartDocument\(path: string\): Document;/);
});

test('the hyperlink module declares only its actual dependencies', () => {
  const text = source('../src/hyperlink.ts');
  assert.deepEqual([...new Set(imports(text))], [
    '@xmldom/xmldom',
    './types.js',
    './drawing.js',
    './internal/context.js',
    './operations.js',
    './xml.js',
    './content-control.js',
    './internal/elements.js',
  ]);
  // content-control 是单向依赖：它不 import hyperlink，所以不成环。
  assert.doesNotMatch(source('../src/content-control.ts'), /from '\.\/hyperlink\.js'/);
  assert.doesNotMatch(text, /from '\.\/document\.js'/);
  // setPartXml 与 nextRelationshipId 刻意借用公开 API：relationships 是 getPartDocument
  // 返回的分离副本，只能整体写回；PartAccess 的 updatePartXml 会在副本上重取、丢掉修改。
  // 超链接持有的 relationships 是 partDocumentCopy 返回的副本，只能整体写回。
  assert.match(text, /'writePartXml'/);
  assert.doesNotMatch(text, /writePartXml\(path: string, xml: string\): void;/);
});

test('document.ts supplies contexts explicitly instead of casting through unknown', () => {
  // `as unknown as XContext` 会关掉结构检查，于是成员名绑到「类上恰好有的那个」，
  // 而「副本 vs 活文档」选错是静默失败。四个 context 都由显式 getter 提供，TypeScript 会检查。
  const text = source('../src/document.ts');
  assert.doesNotMatch(text, /as unknown as \w*Context/);
  for (const getter of ['commentContext', 'bookmarkContext', 'contentControlContext', 'hyperlinkContext']) {
    assert.match(text, new RegExp(`private get ${getter}\\(\\): \\w+ \\{`));
  }
});

test('the copy and live document semantics are observably different', () => {
  // 两套语义的区别只能由「是不是同一个对象」观察到，而选错是静默的，所以在这里把它钉住：
  // 副本每次都是新对象（改它不进包），活文档每次都是同一个（改它由 mutateLivePartXml 提交）。
  const doc = DocxDocument.create();
  const first = doc.getPartDocument(doc.mainDocumentPath);
  const second = doc.getPartDocument(doc.mainDocumentPath);
  assert.notStrictEqual(first, second, 'partDocumentCopy 必须每次返回新的分离副本');

  // 改副本不影响包：再取一次读不到这次的修改。
  const body = first.getElementsByTagName('w:body')[0];
  body.appendChild(first.createElementNS(WORD_NS, 'w:p'));
  const reread = doc.getPartDocument(doc.mainDocumentPath);
  assert.equal(
    reread.getElementsByTagName('w:p').length,
    second.getElementsByTagName('w:p').length,
    '对副本的修改不得进入包',
  );
});
