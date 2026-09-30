import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { columnWidthsPx, effectiveKinsoku, pageBoxPx, paginate, paragraphSpacingPx, snapLineHeightPx } from '../dist/layout.js';
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
  measureTableRow(_table, row) {
    return row.height ?? 10;
  },
};

const blocks = (...paragraphs) => paragraphs.map((p) => ({ type: 'paragraph', paragraph: p }));
const cell = (rowSpan = 1) => ({ blocks: [], colSpan: 1, rowSpan, isMergeContinuation: false });
const row = (height, format = {}, cells = [cell()]) => ({ cells, format, height });
const table = (...rows) => ({ type: 'table', index: 0, rows, grid: [1000] });

test('paginates lines and honors explicit and paragraph page breaks', () => {
  const result = paginate(blocks(
    paragraph(0, [50, 50]),
    paragraph(1, [50], { pageBreakBefore: true }),
  ), [section()], measurer, { defaultTabStopTwips: 720 });
  assert.equal(result.length, 2);
  assert.deepEqual(result.map((page) => page.items.filter((item) => item.type === 'line').map((item) => item.paragraph)), [[0, 0], [1]]);
  assert.equal(paginate([{ type: 'pageBreak' }, ...blocks(paragraph(2, [10]))], [section()], measurer, { defaultTabStopTwips: 720 }).length, 2);
});

test('snaps line heights to a valid section document grid', () => {
  assert.equal(snapLineHeightPx(13, undefined), 13);
  assert.equal(snapLineHeightPx(13, { type: 'default', linePitch: 300 }), 13);
  for (const type of ['lines', 'linesAndChars', 'snapToChars']) {
    const grid = { type, linePitch: 300 };
    assert.equal(snapLineHeightPx(10, grid), 20);
    assert.equal(snapLineHeightPx(20, grid), 20);
    assert.equal(snapLineHeightPx(20.01, grid), 40);
    assert.equal(snapLineHeightPx(20, { ...grid, linePitch: 0 }), 20);
    assert.equal(snapLineHeightPx(20, { ...grid, linePitch: -300 }), 20);
  }
});

test('uses collapsed paragraph spacing and contextual spacing consistently', () => {
  const first = paragraph(0, [10], { style: 'List', spacingAfter: 300, contextualSpacing: true });
  const second = paragraph(1, [10], { style: 'List', spacingAfter: 300, contextualSpacing: true });
  const different = paragraph(2, [10], { style: 'Body', spacingBefore: 100, spacingAfter: 0 });
  assert.equal(paragraphSpacingPx(undefined, first).afterPx, 20);
  assert.equal(paragraphSpacingPx(first, second).beforePx, 0);
  assert.equal(paragraphSpacingPx(second, different).beforePx, 20);
  const pages = paginate(blocks(first, second, different), [section({ pageHeight: 440 })], measurer,
    { defaultTabStopTwips: 720 });
  assert.deepEqual(pages.map((page) => page.items.filter((item) => item.type === 'line').map((item) => item.paragraph)), [[0, 1], [2]]);
});

test('document compatibility overrides auto spacing and East Asian break rules', () => {
  const current = paragraph(0, [10], { spacingBefore: 240, spacingBeforeAuto: true, kinsoku: true });
  assert.equal(paragraphSpacingPx(undefined, current).beforePx, 0);
  assert.equal(paragraphSpacingPx(undefined, current, { doNotUseHTMLParagraphAutoSpacing: true }).beforePx, 16);
  assert.equal(effectiveKinsoku(true), true);
  assert.equal(effectiveKinsoku(true, { doNotUseEastAsianBreakRules: true }), undefined);
});

test('doNotBreakWrappedTables keeps a wrapped table group together', () => {
  const wrapped = paragraph(0, [10], {
    images: [{ placement: 'floating', wrap: 'square', widthPx: 10, heightPx: 10 }],
  });
  const wrappedTable = table(row(10, {}, [cell()]), row(10, {}, [cell()]));
  wrappedTable.rows[0].cells[0].blocks = [{ type: 'paragraph', paragraph: wrapped }];
  const result = paginate([
    { type: 'paragraph', paragraph: paragraph(0, [10]) },
    wrappedTable,
  ], [section({ pageHeight: 15 })], measurer, {
    defaultTabStopTwips: 720,
    compatibilitySettings: { doNotBreakWrappedTables: true },
  });
  assert.equal(result.length, 2);
  assert.deepEqual(result[1].items.filter((item) => item.type === 'tableRow').map((item) => item.row), [0, 1]);
});

test('contextual spacing treats two default-style paragraphs as matching', () => {
  const previous = paragraph(0, [10], { spacingAfter: 300 });
  const current = paragraph(1, [10], { contextualSpacing: true, spacingBefore: 100 });
  assert.equal(paragraphSpacingPx(previous, current).beforePx, 100 * 96 / 1440);
});

test('contextual spacing only suppresses spacing between matching styles', () => {
  // 前一段是另一种样式时，它的段后距不能被本段的 contextualSpacing 吞掉
  const otherStyle = paragraph(0, [10], { style: 'Body', spacingAfter: 300 });
  const listItem = paragraph(1, [10], { style: 'List', contextualSpacing: true, spacingBefore: 100 });
  assert.equal(paragraphSpacingPx(otherStyle, listItem).beforePx, 300 * 96 / 1440);
  const sameStyle = paragraph(0, [10], { style: 'List', spacingAfter: 300 });
  assert.equal(paragraphSpacingPx(sameStyle, listItem).beforePx, 100 * 96 / 1440);
});

test('keepNext groups count the following paragraph spacing once', () => {
  const first = paragraph(0, [20], { keepNext: true, spacingAfter: 200 });
  const second = paragraph(1, [20], { spacingAfter: 200 });
  const following = paragraph(2, [20], { spacingBefore: 100 });
  const pages = paginate(blocks(first, second, following), [section({ pageHeight: 2000 })], measurer,
    { defaultTabStopTwips: 720 });
  assert.equal(pages[0].contentHeightPx, 20 * 3 + 2 * 200 * 96 / 1440);
});

test('paginates split paragraph lines using snapped heights', () => {
  const lines = Array.from({ length: 12 }, () => 15);
  const result = paginate(blocks(paragraph(0, lines)), [
    section({ pageHeight: 1500, docGrid: { type: 'linesAndChars', linePitch: 300, charSpace: 0 } }),
  ], measurer, { defaultTabStopTwips: 720 });
  assert.deepEqual(result.map((page) => page.items.filter((item) => item.type === 'line').length), [5, 5, 2]);
  assert.deepEqual(result.map((page) => page.contentHeightPx), [100, 100, 40]);
  assert.ok(result.flatMap((page) => page.items).filter((item) => item.type === 'line')
    .every((item) => item.line.heightPx === 20));
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

test('pageBreakBefore in the middle of a keepNext chain breaks the chain', () => {
  const result = paginate(blocks(
    paragraph(9, [40]),
    paragraph(0, [40], { keepNext: true }),
    paragraph(1, [40], { pageBreakBefore: true, keepNext: true }),
    paragraph(2, [40]),
  ), [section({ pageHeight: 1800 })], measurer, { defaultTabStopTwips: 720 });
  assert.deepEqual(result.map((page) => page.items.filter((item) => item.type === 'line').map((item) => item.paragraph)), [[9, 0], [1, 2]]);
});

test('splits tables by row and repeats all leading header rows', () => {
  const result = paginate([
    table(
      row(10, { header: true }),
      row(10, { header: true }),
      row(60),
      row(60, { cantSplit: true }),
      row(200),
    ),
  ], [section()], measurer, { defaultTabStopTwips: 720 });
  assert.deepEqual(result.map((page) =>
    page.items.filter((item) => item.type === 'tableRow').map((item) => item.row)), [
    [0, 1, 2],
    [0, 1, 3],
    [0, 1, 4],
  ]);
  assert.deepEqual(result.map((page) => page.contentHeightPx), [80, 80, 220]);
  assert.deepEqual(result.flatMap((page) =>
    page.items.filter((item) => item.type === 'tableRow' && item.row >= 2).map((item) => item.row)), [2, 3, 4]);
});

test('moves all rows covered by a rowSpan to the same page', () => {
  const result = paginate([
    ...blocks(paragraph(0, [30])),
    table(row(60, {}, [cell(2)]), row(60, {}, [{ ...cell(), isMergeContinuation: true, rowSpan: 0 }])),
  ], [section()], measurer, { defaultTabStopTwips: 720 });
  assert.deepEqual(result.map((page) =>
    page.items.filter((item) => item.type === 'tableRow').map((item) => item.row)), [[], [0, 1]]);
});

test('fills columns before creating a new page and measures at each column width', () => {
  const measuredWidths = [];
  const recordingMeasurer = {
    ...measurer,
    measureParagraph(p, area) {
      measuredWidths.push(area.widthPx);
      return measurer.measureParagraph(p);
    },
  };
  const result = paginate(blocks(
    paragraph(0, [60]),
    paragraph(1, [60]),
    paragraph(2, [60]),
  ), [section({ columns: { count: 2, space: 0, equalWidth: true } })], recordingMeasurer, { defaultTabStopTwips: 720 });
  assert.deepEqual(result.map((page) =>
    page.items.filter((item) => item.type === 'line').map((item) => [item.paragraph, item.column])), [
    [[0, 0], [1, 1]],
    [[2, 0]],
  ]);
  assert.ok(measuredWidths.every((width) => width === 50));
});

test('uses explicit unequal column widths and nextColumn stays on the page', () => {
  const measuredWidths = [];
  const recordingMeasurer = {
    ...measurer,
    measureParagraph(p, area) {
      measuredWidths.push(area.widthPx);
      return measurer.measureParagraph(p);
    },
  };
  const columns = { count: 2, space: 0, equalWidth: false, widths: [300, 600] };
  const result = paginate([
    ...blocks(paragraph(0, [20])),
    { type: 'sectionBreak', section: 0, breakType: 'nextColumn' },
    ...blocks(paragraph(1, [20])),
  ], [section({ columns }), section({ index: 1, columns })], recordingMeasurer, { defaultTabStopTwips: 720 });
  assert.equal(result.length, 1);
  assert.deepEqual(result[0].items.filter((item) => item.type === 'line').map((item) => item.column), [0, 1]);
  assert.deepEqual(measuredWidths, [20, 40]);
});

test('remeasures split table fragments at the destination column width', () => {
  const measuredWidths = [];
  const measuredKeys = [];
  const tableMeasurer = {
    ...measurer,
    measureTableRow(_table, row, rowIndex, widthPx) {
      measuredWidths.push(widthPx);
      measuredKeys.push(`${rowIndex}:${widthPx}`);
      return row.height;
    },
  };
  const result = paginate([
    table(row(10, { header: true }), row(60), row(60)),
  ], [section({ columns: { count: 2, space: 0, equalWidth: false, widths: [300, 600] } })],
  tableMeasurer, { defaultTabStopTwips: 720 });
  assert.equal(result.length, 1);
  assert.deepEqual(result[0].items.filter((item) => item.type === 'tableRow').map((item) => [item.row, item.column]), [
    [0, 0], [1, 0], [0, 1], [2, 1],
  ]);
  assert.ok(measuredWidths.includes(20));
  assert.ok(measuredWidths.includes(40));
  assert.equal(new Set(measuredKeys).size, measuredKeys.length);
});

test('passes floating wrap exclusions to paragraph measurement and carries their remaining height', () => {
  const areas = [];
  const wrappingMeasurer = {
    ...measurer,
    measureParagraph(p, area) {
      areas.push(area);
      const sideWidth = area.wraps
        .filter((wrap) => ['square', 'tight', 'through'].includes(wrap.wrap))
        .reduce((sum, wrap) => sum + wrap.widthPx, 0);
      return [{ heightPx: 20, startOffset: 0, endOffset: Math.max(1, Math.floor((area.widthPx - sideWidth) / 10)) }];
    },
  };
  const image = { placement: 'floating', wrap: 'square', widthPx: 30, heightPx: 50 };
  paginate(blocks(
    paragraph(0, [20], { images: [image] }),
    paragraph(1, [20]),
  ), [section()], wrappingMeasurer, { defaultTabStopTwips: 720 });
  assert.equal(areas[0].widthPx - areas[0].wraps[0].widthPx, 70);
  assert.deepEqual(areas[1].wraps.map((wrap) => [wrap.wrap, wrap.heightPx, wrap.carried]), [['square', 30, true]]);
});

test('does not duplicate an anchored float when its paragraph crosses a column', () => {
  const areas = [];
  const wrappingMeasurer = {
    ...measurer,
    measureParagraph(p, area) {
      areas.push({ paragraph: p.index, area });
      return measurer.measureParagraph(p);
    },
  };
  const image = { placement: 'floating', wrap: 'square', widthPx: 30, heightPx: 200 };
  const result = paginate(blocks(
    paragraph(0, [60, 60], { images: [image], widowControl: false }),
    paragraph(1, [10]),
  ), [section({ columns: { count: 2, space: 0, equalWidth: true } })],
  wrappingMeasurer, { defaultTabStopTwips: 720 });
  assert.deepEqual(result[0].items.filter((item) => item.type === 'line').map((item) => [item.paragraph, item.column]), [
    [0, 0], [0, 1], [1, 1],
  ]);
  assert.deepEqual(areas.find((entry) => entry.paragraph === 1).area.wraps, []);
});

test('topAndBottom exclusions occupy the full line and degenerate columns still make progress', () => {
  const areas = [];
  const recordingMeasurer = {
    ...measurer,
    measureParagraph(p, area) {
      areas.push(area);
      return measurer.measureParagraph(p);
    },
  };
  const image = { placement: 'floating', wrap: 'topAndBottom', widthPx: 20, heightPx: 50 };
  const result = paginate(blocks(paragraph(0, [200], { images: [image] })), [
    section({ pageWidth: 0, columns: { count: 3, space: 720, equalWidth: false, widths: [0] } }),
  ], recordingMeasurer, { defaultTabStopTwips: 720 });
  assert.equal(areas[0].wraps[0].wrap, 'topAndBottom');
  assert.equal(areas[0].widthPx, 0);
  assert.equal(result.length, 1);
  assert.equal(result[0].items.length, 1);
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

test('starts a new page when a continuous section changes column geometry', () => {
  const result = paginate([
    ...blocks(paragraph(0, [10])),
    { type: 'sectionBreak', section: 0, breakType: 'continuous' },
    ...blocks(paragraph(1, [10])),
  ], [
    section(),
    section({ index: 1, columns: { count: 2, space: 0, equalWidth: true } }),
  ], measurer, { defaultTabStopTwips: 720 });
  assert.deepEqual(result.map((page) =>
    page.items.filter((item) => item.type === 'line').map((item) => item.paragraph)), [[0], [1]]);
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

test('pageBoxPx converts section page size and all four margins from twips to px', () => {
  const box = pageBoxPx(section({
    pageWidth: 12240,
    pageHeight: 15840,
    margins: { top: 1440, right: 1800, bottom: 720, left: 1080, header: 360, footer: 360, gutter: 0 },
  }));
  assert.deepEqual(box, { widthPx: 816, heightPx: 1056, padding: { top: 96, right: 120, bottom: 48, left: 72 } });
  const landscape = pageBoxPx(section({
    pageWidth: 15840,
    pageHeight: 12240,
    margins: { top: 0, right: 15, bottom: 30, left: 45, header: 0, footer: 0, gutter: 0 },
  }));
  assert.deepEqual(landscape, { widthPx: 1056, heightPx: 816, padding: { top: 0, right: 1, bottom: 2, left: 3 } });
});

test('columnWidthsPx splits usable width across equal columns and honours explicit widths', () => {
  const page = { pageWidth: 12240, margins: { top: 0, right: 1440, bottom: 0, left: 1440, header: 0, footer: 0, gutter: 0 } };
  assert.deepEqual(columnWidthsPx(section({ ...page, columns: { count: 1, space: 0, equalWidth: true } })), [624]);
  assert.deepEqual(columnWidthsPx(section({ ...page, columns: { count: 2, space: 720, equalWidth: true } })), [288, 288]);
  assert.deepEqual(columnWidthsPx(section({ ...page, columns: { count: 2, space: 720, equalWidth: false, widths: [2880, 5760] } })), [192, 384]);
  assert.deepEqual(columnWidthsPx(section({ ...page, columns: { count: 3, space: 720, equalWidth: false, widths: [2880, 5760] } })), [176, 176, 176]);
});
