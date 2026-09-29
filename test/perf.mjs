import { test } from 'node:test';
import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { DocxDocument } from '../dist/document.js';

const ROUNDS = 5;
const WARMUPS = 2;
const INITIAL_PARAGRAPH_COUNT = 1;

function elapsedMs(run) {
  const start = process.hrtime.bigint();
  run();
  return Number(process.hrtime.bigint() - start) / 1e6;
}

function median(values) {
  const sorted = [...values].sort((a, b) => a - b);
  const middle = Math.floor(sorted.length / 2);
  if (sorted.length % 2) return sorted[middle];
  return (sorted[middle - 1] + sorted[middle]) / 2;
}

// 只给 run 计时，setup 不算。播种文档、构造操作数组都属于 setup：
// 把它们算进耗时会把被测操作的成本冲淡，而且两侧冲淡的程度还不一样。
function sample({ setup, run }) {
  const state = setup ? setup() : undefined;
  return elapsedMs(() => run(state));
}

// 配对交替测量。两个测量必须在同一时间窗口里交替进行，否则任何跨窗口的漂移
// （JIT 预热、GC、机器负载）都会被整份记进比值 —— 而不是被比值抵消掉。
// 每轮内部还要换序，这样两侧都不会系统性地替对方付预热/缓存的账。
// 比值取「每轮配对比值的中位数」，而不是「两个中位数的比值」：前者才真正抵消轮间波动。
function measurePairedRatio(numerator, denominator) {
  for (let i = 0; i < WARMUPS; i++) {
    sample(numerator);
    sample(denominator);
  }
  const numeratorSamples = [];
  const denominatorSamples = [];
  const ratios = [];
  for (let round = 0; round < ROUNDS; round++) {
    let top;
    let bottom;
    if (round % 2 === 0) {
      top = sample(numerator);
      bottom = sample(denominator);
    } else {
      bottom = sample(denominator);
      top = sample(numerator);
    }
    numeratorSamples.push(top);
    denominatorSamples.push(bottom);
    ratios.push(top / bottom);
  }
  return {
    ratio: median(ratios),
    ratios,
    numerator: { median: median(numeratorSamples), samples: numeratorSamples },
    denominator: { median: median(denominatorSamples), samples: denominatorSamples },
  };
}

function formatMeasurement(label, measurement) {
  return `${label}: median=${measurement.median.toFixed(1)}ms samples=[${measurement.samples.map((value) => value.toFixed(1)).join(', ')}]`;
}

function assertRatioBelow(result, threshold, description, numeratorLabel, denominatorLabel) {
  assert.ok(result.ratio < threshold, [
    `${description} exceeded threshold ${threshold.toFixed(2)} (actual ${result.ratio.toFixed(2)})`,
    `paired ratios=[${result.ratios.map((value) => value.toFixed(2)).join(', ')}]`,
    formatMeasurement(numeratorLabel, result.numerator),
    formatMeasurement(denominatorLabel, result.denominator),
  ].join('\n'));
}

const MAX_OPERATIONS_PER_REQUEST = 1000;

function seededDocument(totalParagraphs) {
  const doc = DocxDocument.create();
  let remaining = Math.max(0, totalParagraphs - INITIAL_PARAGRAPH_COUNT);
  while (remaining > 0) {
    const take = Math.min(MAX_OPERATIONS_PER_REQUEST, remaining);
    doc.applyOperations({
      operations: Array.from({ length: take }, (_, i) => ({ type: 'insertParagraph', text: `seed-${i}` })),
    });
    remaining -= take;
  }
  return doc;
}

function setTextRequest(start, count) {
  return { operations: Array.from({ length: count }, (_, i) => ({ type: 'setParagraphText', index: start + i, text: `更新-${start + i}` })) };
}

function singleInsert(count, seedCount = 0) {
  return {
    setup: () => seededDocument(seedCount),
    run: (doc) => {
      for (let i = 0; i < count; i++) doc.insertParagraph(`段落内容 ${i}`);
    },
  };
}

function batchInsert(count) {
  return {
    setup: () => ({
      doc: DocxDocument.create(),
      request: { operations: Array.from({ length: count }, (_, i) => ({ type: 'insertParagraph', text: `段落内容 ${i}` })) },
    }),
    run: ({ doc, request }) => doc.applyOperations(request),
  };
}

function singleSetParagraphText(count) {
  return {
    setup: () => seededDocument(count),
    run: (doc) => {
      for (let i = 0; i < count; i++) doc.setParagraphText(i, `改写 ${i}`);
    },
  };
}

// 在一份固定大小的文档上，把同样多的改写拆成 chunkSize 一批。
// 两侧每个样本做的工作量完全相同 —— 这样 JIT 与堆的轨迹一致，
// 比值才只反映批处理粒度本身，不掺进「小的那侧越跑越快」的假象。
function chunkedSetParagraphText(totalParagraphs, chunkSize) {
  return {
    setup: () => {
      const requests = [];
      for (let start = 0; start < totalParagraphs; start += chunkSize) {
        requests.push(setTextRequest(start, Math.min(chunkSize, totalParagraphs - start)));
      }
      return { doc: seededDocument(totalParagraphs), requests };
    },
    run: ({ doc, requests }) => {
      for (const request of requests) doc.applyOperations(request);
    },
  };
}

// 固定操作次数、只变文档规模，衡量「单次按下标操作的成本随文档规模怎么长」。
//
// 这一档必须开独立进程。同一个进程里，小规模那侧会因为 JIT 越跑越快、又会被大规模那侧
// 留下的垃圾拖慢 —— 配对交替抵消不了这种工作集大小的不对称：实测同进程的配对比值会在
// 3.8 到 11.0 之间跳，5 次里有 1 次误报。独立进程测出来则稳定在 3.7 上下。
function perOperationCostInChildProcess(documentParagraphs, operationCount, rounds = 5, mode = 'batch') {
  const output = execFileSync(process.execPath,
    [new URL('perf-worker.mjs', import.meta.url).pathname, String(documentParagraphs), String(operationCount), String(rounds), mode],
    { encoding: 'utf8' });
  return JSON.parse(output);
}

function directEditCostInChildProcess(documentParagraphs, operationCount, rounds = 3) {
  return perOperationCostInChildProcess(documentParagraphs, operationCount, rounds, 'direct');
}

function buildTableLookupDoc(tableCount) {
  const doc = DocxDocument.create();
  for (let i = 0; i < tableCount; i++) {
    doc.insertTable([
      [`${i}-0`, `${i}-1`, `${i}-2`],
      [`${i}-3`, `${i}-4`, `${i}-5`],
    ]);
    doc.insertParagraph(`after-${i}`);
  }
  return doc;
}

// 查的是缓存命中的成本，单次只有几十纳秒，所以必须查够多次：
// 被测区间若短到几十微秒，比值就只是计时噪声，断言会随机通过或失败。
const TABLE_CELL_LOOKUPS = 20_000;

function tableCellSwitches(tableCount) {
  return {
    setup: () => {
      const doc = buildTableLookupDoc(tableCount);
      const indices = doc.getBlocks()
        .filter((block) => block.type === 'table')
        .flatMap((table) => table.rows.flatMap((row) => row.cells.map((cell) => cell.blocks[0].paragraph.index)));
      doc.getTableCellAt(indices[0]);
      return { doc, indices };
    },
    run: ({ doc, indices }) => {
      for (let index = 0; index < TABLE_CELL_LOOKUPS; index++) doc.getTableCellAt(indices[index % indices.length]);
    },
  };
}

test('performance regression: insertParagraph stays within a calibrated multiple of setParagraphText', () => {
  const result = measurePairedRatio(singleInsert(200, 200), singleSetParagraphText(200));
  assertRatioBelow(result, 3, 'insertParagraph cost ratio',
    'single insert x200 into 200 seeded paragraphs', 'setParagraphText x200');
});

test('performance regression: batched inserts remain materially faster than repeated single inserts', () => {
  const result = measurePairedRatio(batchInsert(300), singleInsert(300));
  assertRatioBelow(result, 0.75, 'batch insert ratio',
    'batch insert x300', 'single insert x300');
});

test('performance regression: one large batch is no worse than several small ones', () => {
  // 同一份 1000 段文档，1000 次改写一批 vs 拆成两批 500。工作量相同，所以线性 ≈ 1.00。
  // 这条守的是 #15 的成果：批处理不得退化成「每次操作各提交一遍」。
  const result = measurePairedRatio(chunkedSetParagraphText(1000, 1000), chunkedSetParagraphText(1000, 500));
  assertRatioBelow(result, 1.5, 'single-batch vs split-batch ratio',
    'setParagraphText x1000 in one batch', 'setParagraphText x1000 in two batches');
});

test('performance regression: per-operation cost does not degrade further with document size', () => {
  // 固定 10 批、每批 1000 次操作，文档规模 4 倍。期望比值高于 1：
  // 每批都要 O(文档规模) 地构建一次段落缓存；阈值 2.5 区分这项固定成本与逐操作 O(文档规模) 退化。
  const large = perOperationCostInChildProcess(2000, 1000);
  const small = perOperationCostInChildProcess(500, 1000);
  assert.ok(large.median >= 30 && small.median >= 30, [
    'indexed edit measurement was shorter than the 30ms minimum; the instrumented setter may not have run',
    formatMeasurement('10000 ops on a 2000-paragraph document', large),
    formatMeasurement('10000 ops on a 500-paragraph document', small),
  ].join('\n'));
  const ratio = large.perOperation / small.perOperation;
  assert.ok(ratio < 2.5, [
    `per-operation cost ratio across a 4x document-size gap exceeded threshold 2.50 (actual ${ratio.toFixed(2)})`,
    formatMeasurement('10000 ops on a 2000-paragraph document', large),
    formatMeasurement('10000 ops on a 500-paragraph document', small),
  ].join('\n'));
});

test('performance regression: direct edits stay within the absolute history budget', () => {
  for (const paragraphs of [500, 2000]) {
    const direct = directEditCostInChildProcess(paragraphs, 200);
    const batch = perOperationCostInChildProcess(paragraphs, 200);
    assert.ok(direct.perOperation < 5, [
      `direct edits exceeded 5ms per operation at ${paragraphs} paragraphs`,
      `direct=${JSON.stringify(direct)}`,
      `batch=${JSON.stringify(batch)}`,
    ].join('\n'));
  }
});

test('performance regression: cached getTableCellAt lookups do not scale linearly with table count', () => {
  const result = measurePairedRatio(tableCellSwitches(36), tableCellSwitches(12));
  assertRatioBelow(result, 2.2, 'getTableCellAt lookup ratio',
    'cached getTableCellAt over 36 tables', 'cached getTableCellAt over 12 tables');
});
