import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';

// src/internal/context.ts 是后续每个功能簇搬迁都要依赖的叶子模块。
// 一旦有人往里加一条指向 src/*.ts 的 import 就会成环，而成环没有任何东西会报警。
test('the internal context module stays a leaf with no DOM dependency', () => {
  const source = readFileSync(new URL('../src/internal/context.ts', import.meta.url), 'utf8');
  assert.doesNotMatch(source, /\b(?:document|window)\./);
  assert.doesNotMatch(source, /HTMLElement|globalThis/);
  assert.deepEqual(
    [...source.matchAll(/from ['"]([^'"]+)['"]/g)].map((match) => match[1]),
    ['@xmldom/xmldom'],
  );
});
