// 跨文档规模的测量必须在独立进程里做。同一个进程里，小规模那侧会因为 JIT 越跑越快、
// 又会被大规模那侧留下的垃圾拖慢，配对交替抵消不了这种工作集大小的不对称 ——
// 实测同进程的配对比值会在 3.8 到 11.0 之间跳。
// 用法: node test/perf-worker.mjs <文档段数> <操作次数> <轮数>
import { DocxDocument } from '../dist/document.js';

const MAX_OPERATIONS_PER_REQUEST = 1000;
const [documentParagraphs, operationCount, rounds] = process.argv.slice(2).map(Number);

function seededDocument() {
  const doc = DocxDocument.create();
  let remaining = documentParagraphs - 1;
  while (remaining > 0) {
    const take = Math.min(MAX_OPERATIONS_PER_REQUEST, remaining);
    doc.applyOperations({
      operations: Array.from({ length: take }, (_, i) => ({ type: 'insertParagraph', text: `seed-${i}` })),
    });
    remaining -= take;
  }
  return doc;
}

const step = Math.max(1, Math.floor(documentParagraphs / operationCount));
const request = {
  operations: Array.from({ length: operationCount }, (_, i) => ({
    type: 'setParagraphText', index: (i * step) % documentParagraphs, text: `更新-${i}`,
  })),
};

const samples = [];
for (let round = 0; round < rounds; round++) {
  const doc = seededDocument();
  const start = process.hrtime.bigint();
  doc.applyOperations(request);
  samples.push(Number(process.hrtime.bigint() - start) / 1e6);
}
samples.sort((a, b) => a - b);
const median = samples.length % 2
  ? samples[(samples.length - 1) / 2]
  : (samples[samples.length / 2 - 1] + samples[samples.length / 2]) / 2;
process.stdout.write(JSON.stringify({ median, samples, perOperation: median / operationCount }));
