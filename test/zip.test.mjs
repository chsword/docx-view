import { test } from 'node:test';
import assert from 'node:assert/strict';
import { Worker } from 'node:worker_threads';
import { DocxDocument } from '../dist/document.js';
import { zipParts } from '../dist/zip.js';

const seeded = async (paragraphs = 30) => {
  const doc = DocxDocument.create();
  doc.applyOperations({
    operations: Array.from({ length: paragraphs }, (_unused, index) => ({
      type: 'insertParagraph', text: `第 ${index} 段 ${'正文内容'.repeat(10)}`,
    })),
  });
  return DocxDocument.load(await doc.toUint8Array());
};

test('toUint8Array lets the host move the ZIP step off the calling thread', async () => {
  const doc = await seeded();
  doc.setParagraphText(0, '改过了');

  // 注入的函数拿到的是已物化的部件字节，够它自己打包。
  let received;
  const bytes = await doc.toUint8Array({
    zip: async (parts) => {
      received = parts;
      return zipParts(parts);
    },
  });
  assert.ok(received.has('word/document.xml'));
  assert.ok(received.get('word/document.xml') instanceof Uint8Array);
  assert.match(new TextDecoder().decode(received.get('word/document.xml')), /改过了/);

  // 注入与默认两条路产出的字节必须一致。
  assert.deepEqual(bytes, await doc.toUint8Array());
  // 交出去之后本文档的部件不能受影响（所以不要在宿主那边转移 ArrayBuffer）。
  assert.match(doc.getPartXml('word/document.xml'), /改过了/);

  // 返回值原样用，但形状不对要拦住——否则错误会推迟到宿主保存文件时才暴露。
  await assert.rejects(() => doc.toUint8Array({ zip: 'nope' }), /zip must be a function/);
  await assert.rejects(() => doc.toUint8Array({ zip: async () => 'not bytes' }), /must resolve to a Uint8Array/);

  // toBlob 也走同一条路。注入一个产出可辨认内容的 zip —— 用「真的压一遍」做断言是分不出
  // toBlob 有没有透传 options 的，因为两条路的字节本来就一样。
  const marker = new Uint8Array([1, 2, 3, 4, 5]);
  const blob = await doc.toBlob({ zip: async () => marker });
  assert.equal(blob.type, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
  assert.equal(blob.size, marker.length);
  assert.deepEqual(new Uint8Array(await blob.arrayBuffer()), marker);
});

test('zipParts actually deflates instead of storing', async () => {
  // 没有这条断言的话，把 compression 改成 STORE 一样全绿：STORE 产出的 ZIP 也合法可读。
  const doc = await seeded(200);
  doc.setParagraphText(0, '改过了');
  const bytes = await doc.toUint8Array();
  let raw = 0;
  await doc.toUint8Array({
    zip: async (parts) => {
      for (const part of parts.values()) raw += part.length;
      return zipParts(parts);
    },
  });
  assert.ok(raw > 20_000, `未压缩字节应当足够大才测得出压缩，实际 ${raw}`);
  assert.ok(bytes.length < raw / 2, `应当明显压缩过：${bytes.length} 对 ${raw}`);
});

test('the parts map survives a real worker boundary and the package still loads', async () => {
  // 只有真的跨一次线程边界才能证明 Map<string, Uint8Array> 过得了结构化克隆；
  // 用假函数替代是测不出来的。
  const doc = await seeded();
  doc.setParagraphText(0, '在 worker 里压缩');
  const worker = new Worker(new URL('./zip-worker-fixture.mjs', import.meta.url));
  try {
    const bytes = await doc.toUint8Array({
      zip: (parts) => new Promise((resolve, reject) => {
        worker.once('message', (message) => (message.ok ? resolve(message.bytes) : reject(new Error(message.message))));
        worker.postMessage(parts);
      }),
    });
    assert.deepEqual(bytes, await doc.toUint8Array(), 'worker 与主线程产出的字节一致');
    const reloaded = await DocxDocument.load(bytes);
    assert.equal(reloaded.getParagraph(0).text, '在 worker 里压缩');
    assert.equal(reloaded.getParagraphs().length, doc.getParagraphs().length);
  } finally {
    await worker.terminate();
  }
});
