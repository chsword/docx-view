import { test } from 'node:test';
import assert from 'node:assert/strict';
import { Worker } from 'node:worker_threads';
import { DocxDocument } from '../dist/document.js';
import { WORD_NS } from '../dist/xml.js';

async function bytesOf(texts) {
  const doc = DocxDocument.create();
  doc.setPartXml(doc.mainDocumentPath, `<w:document xmlns:w="${WORD_NS}"><w:body>${texts.map((text) => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`).join('')}<w:sectPr/></w:body></w:document>`);
  return doc.toUint8Array();
}

test('findText returns ranges that address the matched text', () => {
  const doc = DocxDocument.create();
  doc.setPartXml(doc.mainDocumentPath, `<w:document xmlns:w="${WORD_NS}"><w:body>
    <w:p><w:r><w:t xml:space="preserve">Apple pie and apple </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>APPLE</w:t></w:r></w:p>
    <w:p><w:r><w:t>没有</w:t></w:r></w:p><w:p><w:r><w:t>苹果 apple</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`);
  assert.deepEqual(doc.findText('apple'), [
    { paragraph: 0, start: 0, end: 5 }, { paragraph: 0, start: 14, end: 19 }, { paragraph: 0, start: 20, end: 25 },
    { paragraph: 2, start: 3, end: 8 },
  ], '默认不区分大小写，跨 run 也找得到');
  assert.deepEqual(doc.findText('apple', { caseSensitive: true }).map((match) => match.start), [14, 3]);
  assert.equal(doc.findText('apple', { maxResults: 2 }).length, 2);
  // 偏移能直接拿去改正文。
  const [hit] = doc.findText('pie');
  doc.formatDocumentRange({ start: { paragraph: hit.paragraph, offset: hit.start }, end: { paragraph: hit.paragraph, offset: hit.end } }, { italic: true });
  assert.equal(doc.getParagraphs()[0].runs.find((run) => run.text === 'pie').italic, true);
  assert.throws(() => doc.findText(''), /empty/);
  assert.throws(() => doc.findText('x', { maxResults: 0 }), /maxResults/);
});

test('compare, read-only open and search run across a real worker boundary', async () => {
  const base = await bytesOf(['first line', 'second line', 'third line']);
  const revised = await bytesOf(['first line', 'second line changed', 'third line']);
  const worker = new Worker(new URL('./offload-worker-fixture.mjs', import.meta.url));
  const call = (method, ...args) => new Promise((resolve, reject) => {
    worker.once('message', (message) => (message.ok ? resolve(message.value) : reject(new Error(message.message))));
    worker.postMessage({ method, args });
  });
  try {
    const comparedBytes = await call('compareDocxBytes', base, revised, { author: 'Worker' });
    const compared = await DocxDocument.load(comparedBytes);
    assert.ok(compared.getRevisions().some((revision) => revision.author === 'Worker'));
    compared.acceptAllRevisions();
    assert.deepEqual(compared.getParagraphs().map((paragraph) => paragraph.text), ['first line', 'second line changed', 'third line']);

    const snapshot = await call('readDocxSnapshot', revised);
    assert.deepEqual(snapshot.paragraphs.map((paragraph) => paragraph.text), ['first line', 'second line changed', 'third line']);
    assert.ok(snapshot.parts.includes('word/document.xml'));

    assert.deepEqual(await call('searchDocxText', revised, 'LINE'), [
      { paragraph: 0, start: 6, end: 10 }, { paragraph: 1, start: 7, end: 11 }, { paragraph: 2, start: 6, end: 10 },
    ]);
    await assert.rejects(call('searchDocxText', revised, ''), /empty/);
  } finally {
    await worker.terminate();
  }
});
