import { test } from 'node:test';
import assert from 'node:assert/strict';
import { DocxDocument } from '../dist/document.js';

const ROUNDS = 5;
const WARMUPS = 1;

function elapsedMs(run) {
  const start = process.hrtime.bigint();
  run();
  return Number(process.hrtime.bigint() - start) / 1e6;
}

function median(values) {
  const sorted = [...values].sort((a, b) => a - b);
  return sorted[Math.floor(sorted.length / 2)];
}

function measureMedian(run) {
  for (let i = 0; i < WARMUPS; i++) run();
  const samples = Array.from({ length: ROUNDS }, () => elapsedMs(run));
  return { median: median(samples), samples };
}

function formatMeasurement(label, measurement) {
  return `${label}: median=${measurement.median.toFixed(1)}ms samples=[${measurement.samples.map((value) => value.toFixed(1)).join(', ')}]`;
}

function measureSingleInsert(count) {
  return measureMedian(() => {
    const doc = DocxDocument.create();
    for (let i = 0; i < count; i++) doc.insertParagraph(`段落内容 ${i}`);
  });
}

function measureBatchInsert(count) {
  return measureMedian(() => {
    const doc = DocxDocument.create();
    doc.applyOperations({
      operations: Array.from({ length: count }, (_, i) => ({ type: 'insertParagraph', text: `段落内容 ${i}` })),
    });
  });
}

function measureSetParagraphText(count) {
  return measureMedian(() => {
    const doc = DocxDocument.create();
    doc.applyOperations({
      operations: Array.from({ length: count - 1 }, (_, i) => ({ type: 'insertParagraph', text: `seed-${i}` })),
    });
    for (let i = 0; i < count; i++) doc.setParagraphText(i, `改写 ${i}`);
  });
}

function measureBatchSetParagraphText(count) {
  return measureMedian(() => {
    const doc = DocxDocument.create();
    doc.applyOperations({
      operations: Array.from({ length: count - 1 }, (_, i) => ({ type: 'insertParagraph', text: `seed-${i}` })),
    });
    doc.applyOperations({
      operations: Array.from({ length: count }, (_, i) => ({ type: 'setParagraphText', index: i, text: `更新-${i}` })),
    });
  });
}

test('performance regression: insertParagraph stays within a calibrated multiple of setParagraphText', () => {
  const inserts = measureSingleInsert(300);
  const sets = measureSetParagraphText(200);
  const perOpInsert = inserts.median / 300;
  const perOpSet = sets.median / 200;
  const ratio = perOpInsert / perOpSet;
  assert.ok(ratio < 4, [
    `insertParagraph cost ratio exceeded threshold 4.00 (actual ${ratio.toFixed(2)})`,
    formatMeasurement('single insert x300', inserts),
    formatMeasurement('setParagraphText x200', sets),
  ].join('\n'));
});

test('performance regression: batched inserts remain materially faster than repeated single inserts', () => {
  const single = measureSingleInsert(300);
  const batch = measureBatchInsert(300);
  const ratio = batch.median / single.median;
  assert.ok(ratio < 0.75, [
    `batch insert ratio exceeded threshold 0.75 (actual ${ratio.toFixed(2)})`,
    formatMeasurement('single insert x300', single),
    formatMeasurement('batch insert x300', batch),
  ].join('\n'));
});

test('performance regression: batched 1000-paragraph updates scale near-linearly', () => {
  const t500 = measureBatchSetParagraphText(500);
  const t1000 = measureBatchSetParagraphText(1000);
  const ratio = t1000.median / t500.median;
  assert.ok(ratio < 4, [
    `1000-paragraph update ratio exceeded threshold 4.00 (actual ${ratio.toFixed(2)})`,
    formatMeasurement('batch setParagraphText x500', t500),
    formatMeasurement('batch setParagraphText x1000', t1000),
  ].join('\n'));
});
