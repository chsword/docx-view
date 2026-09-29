import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { paginate } from '../dist/layout.js';

const section = (overrides = {}) => ({
  index: 0,
  startParagraph: 0,
  endParagraph: 99,
  type: 'nextPage',
  pageWidth: 1500,
  pageHeight: 1500,
  orientation: 'portrait',
  margins: { top: 0, right: 0, bottom: 0, left: 0, header: 0, footer: 0, gutter: 0 },
  columns: { count: 1, space: 0, equalWidth: true },
  titlePage: false,
  headers: {},
  footers: {},
  ...overrides,
});

const paragraph = (index, lines, format = {}) => ({
  index,
  text: '',
  runs: [],
  images: [],
  ...format,
  lines,
});

const measurer = {
  measureParagraph(p) {
    return p.lines.map((height, index) => ({ heightPx: height, startOffset: index, endOffset: index + 1 }));
  },
  measureTableRow() {
    return 10;
  },
};

const blocks = (...paragraphs) => paragraphs.map((p) => ({ type: 'paragraph', paragraph: p }));

test('paginates lines and honors explicit and paragraph page breaks', () => {
  const result = paginate(blocks(
    paragraph(0, [50, 50]),
    paragraph(1, [50], { pageBreakBefore: true }),
  ), [section()], measurer, { defaultTabStopTwips: 720 });
  assert.equal(result.length, 2);
  assert.deepEqual(result.map((page) => page.items.filter((item) => item.type === 'line').map((item) => item.paragraph)), [[0, 0], [1]]);
  assert.equal(paginate([{ type: 'pageBreak' }, ...blocks(paragraph(2, [10]))], [section()], measurer, { defaultTabStopTwips: 720 }).length, 2);
});

test('keeps a keepNext chain and keepLines paragraph together', () => {
  const result = paginate(blocks(
    paragraph(0, [50], { keepNext: true }),
    paragraph(1, [50], { keepNext: true }),
    paragraph(2, [50]),
    paragraph(3, [60, 60], { keepLines: true }),
  ), [section()], measurer, { defaultTabStopTwips: 720 });
  assert.deepEqual(result.map((page) => page.items.filter((item) => item.type === 'line').map((item) => item.paragraph)), [[0, 1, 2], [3, 3]]);
});

test('widow control defaults on but can be disabled', () => {
  const on = paginate(blocks(paragraph(9, [40]), paragraph(0, [60, 60])), [section()], measurer, { defaultTabStopTwips: 720 });
  const off = paginate(blocks(paragraph(9, [40]), paragraph(0, [60, 60], { widowControl: false })), [section()], measurer, { defaultTabStopTwips: 720 });
  assert.deepEqual(on.map((page) => page.items.filter((item) => item.type === 'line').map((item) => item.paragraph)), [[9], [0], [0]]);
  assert.deepEqual(off.map((page) => page.items.filter((item) => item.type === 'line').map((item) => item.paragraph)), [[9, 0], [0]]);
});

test('handles section numbering, parity and malformed page sizes without throwing', () => {
  const sections = [section({ pageNumbering: { start: 7 } }), section({ index: 1, pageHeight: 100, pageNumbering: { start: 20 } })];
  const result = paginate([
    ...blocks(paragraph(0, [100])),
    { type: 'sectionBreak', section: 1, breakType: 'oddPage' },
    ...blocks(paragraph(1, [10])),
  ], sections, measurer, { defaultTabStopTwips: 720 });
  assert.deepEqual(result.map((page) => page.number), [7, 8, 20]);
  assert.doesNotThrow(() => paginate(blocks(paragraph(0, [1000])), [section({ pageHeight: 0 })], measurer, { defaultTabStopTwips: 720 }));
  assert.deepEqual(paginate([], [section()], measurer, { defaultTabStopTwips: 720 }), []);
  assert.deepEqual(paginate(blocks(paragraph(0, [10])), [], measurer, { defaultTabStopTwips: 720 }), []);
});

test('layout has no DOM dependencies', () => {
  const source = readFileSync(new URL('../src/layout.ts', import.meta.url), 'utf8');
  assert.doesNotMatch(source, /\b(?:document|window)\.|HTMLElement/);
});
