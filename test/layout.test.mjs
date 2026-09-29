import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { paginate } from '../dist/layout.js';
import { DocxDocument } from '../dist/document.js';
import { WORD_NS } from '../dist/xml.js';

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
    const heights = p.lines ?? [10];
    return heights.map((height, index) => ({ heightPx: height, startOffset: index, endOffset: index + 1 }));
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
  ), [section({ pageHeight: 2250 })], measurer, { defaultTabStopTwips: 720 });
  assert.deepEqual(result.map((page) => page.items.filter((item) => item.type === 'line').map((item) => item.paragraph)), [[0, 1, 2], [3, 3]]);
});

test('breaks an oversized keepNext chain into independently paginated paragraphs', () => {
  const result = paginate(blocks(
    paragraph(0, [60], { keepNext: true }),
    paragraph(1, [60], { keepNext: true }),
    paragraph(2, [60]),
  ), [section()], measurer, { defaultTabStopTwips: 720 });
  assert.deepEqual(result.map((page) => page.items.filter((item) => item.type === 'line').map((item) => item.paragraph)), [[0], [1], [2]]);
});

test('does not create an empty page for pageBreakBefore plus keepNext', () => {
  const result = paginate(blocks(
    paragraph(0, [40]),
    paragraph(1, [40], { pageBreakBefore: true, keepNext: true }),
    paragraph(2, [40]),
  ), [section()], measurer, { defaultTabStopTwips: 720 });
  assert.deepEqual(result.map((page) => page.items.filter((item) => item.type === 'line').map((item) => item.paragraph)), [[0], [1, 2]]);
});

test('widow control defaults on but can be disabled', () => {
  const on = paginate(blocks(paragraph(9, [40]), paragraph(0, [60, 60])), [section()], measurer, { defaultTabStopTwips: 720 });
  const off = paginate(blocks(paragraph(9, [40]), paragraph(0, [60, 60], { widowControl: false })), [section()], measurer, { defaultTabStopTwips: 720 });
  assert.deepEqual(on.map((page) => page.items.filter((item) => item.type === 'line').map((item) => item.paragraph)), [[9], [0], [0]]);
  assert.deepEqual(off.map((page) => page.items.filter((item) => item.type === 'line').map((item) => item.paragraph)), [[9, 0], [0]]);
});

test('handles section numbering, parity and malformed page sizes without throwing', () => {
  const sections = [section({ pageNumbering: { start: 5 } }), section({ index: 1, pageHeight: 100, pageNumbering: { start: 20 } })];
  const result = paginate([
    ...blocks(paragraph(0, [100])),
    { type: 'sectionBreak', section: 0, breakType: 'oddPage' },
    ...blocks(paragraph(1, [10])),
  ], sections, measurer, { defaultTabStopTwips: 720 });
  assert.deepEqual(result.map((page) => page.number), [5, 6, 20]);
  assert.deepEqual(result.map((page) => page.section), [0, 0, 1]);
  assert.doesNotThrow(() => paginate(blocks(paragraph(0, [1000])), [section({ pageHeight: 0 })], measurer, { defaultTabStopTwips: 720 }));
  assert.deepEqual(paginate([], [section()], measurer, { defaultTabStopTwips: 720 }), []);
  assert.equal(paginate(blocks(paragraph(0, [10])), [], measurer, { defaultTabStopTwips: 720 }).length, 1);
});

test('paginates the real block and section shapes from DocxDocument', () => {
  const doc = DocxDocument.create();
  doc.setPartXml(doc.mainDocumentPath, `<w:document xmlns:w="${WORD_NS}"><w:body>
    <w:p><w:r><w:t>p0</w:t></w:r></w:p>
    <w:p><w:r><w:t>p1</w:t></w:r></w:p>
    <w:p><w:pPr><w:sectPr><w:pgSz w:w="1500" w:h="1500"/></w:sectPr></w:pPr></w:p>
    <w:p><w:r><w:t>p2</w:t></w:r></w:p>
    <w:p><w:r><w:t>p3</w:t></w:r></w:p>
    <w:sectPr><w:pgSz w:w="1500" w:h="1500"/></w:sectPr>
  </w:body></w:document>`);
  const actualBlocks = doc.getBlocks();
  const actualSections = doc.getSections();
  const result = paginate(actualBlocks, actualSections, measurer, { defaultTabStopTwips: 720 });
  const paragraphPages = result.flatMap((page) =>
    page.items.filter((item) => item.type === 'line').map((item) => ({ paragraph: item.paragraph, section: page.section })),
  );
  assert.deepEqual(actualBlocks.filter((block) => block.type === 'sectionBreak').map((block) => block.section), [0]);
  assert.deepEqual(paragraphPages.map((item) => item.paragraph), [0, 1, 2, 3, 4]);
  for (const item of paragraphPages) {
    const section = actualSections.find((candidate) => candidate.startParagraph <= item.paragraph && item.paragraph <= candidate.endParagraph);
    assert.equal(item.section, section?.index);
  }
});

test('layout has no DOM dependencies', () => {
  const source = readFileSync(new URL('../src/layout.ts', import.meta.url), 'utf8');
  assert.doesNotMatch(source, /\b(?:document|window)\.|HTMLElement/);
  assert.deepEqual([...source.matchAll(/from ['"]([^'"]+)['"]/g)].map((match) => match[1]), ['./types.js']);
});
