import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';

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
  assert.match(text, /interface BookmarkContext extends Pick<PartAccess, 'mainPath' \| 'getPartDocument' \| 'updatePartXml'>/);
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
  assert.match(text, /interface ContentControlContext extends Pick<PartAccess, 'mainPath'>/);
  assert.match(text, /getCachedPartDocument\(path: string\): Document/);
  assert.match(text, /updatePartXmlInternal\(path: string/);
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
  assert.match(text, /setPartXml\(path: string, xml: string\): void/);
});
