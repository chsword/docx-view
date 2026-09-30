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
