import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { columnWidthsPx, combineBracketChars, combinedTextLines, effectiveKinsoku, lineNumbersFor, pageBoxPx, paginate, paragraphSpacingPx, rubyAlignToCss, frameWrapExclusion, overstrikeLayers, snapLineHeightPx, tableWrapExclusion } from '../dist/layout.js';
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

test('numbers paginated lines with restart, suppression, and countBy rules', () => {
  const pages = [
    { index: 0, number: 1, section: 0, items: [
      { type: 'line', paragraph: 0 }, { type: 'line', paragraph: 1 }, { type: 'tableRow', table: 0, row: 0 },
    ] },
    { index: 1, number: 2, section: 0, items: [{ type: 'line', paragraph: 2 }] },
    { index: 2, number: 3, section: 1, items: [{ type: 'line', paragraph: 3 }] },
  ];
  const paragraphs = [
    { index: 0, effective: {} }, { index: 1, effective: { suppressLineNumbers: true } },
    { index: 2, effective: {} }, { index: 3, effective: {} },
  ];
  const sections = [
    section({ lineNumbering: { start: 1, countBy: 2, restart: 'newPage' } }),
    section({ index: 1, lineNumbering: { start: 5, restart: 'newSection' } }),
  ];
  assert.deepEqual(lineNumbersFor(pages, paragraphs, sections), [[null, null, null], [null], [5]]);
  assert.deepEqual(lineNumbersFor(
    pages.slice(0, 2), paragraphs, [section({ lineNumbering: { start: 0, countBy: 0, restart: 'continuous' } })],
  ), [[1, null, null], [2]]);
  assert.deepEqual(lineNumbersFor(
    pages.slice(0, 2), paragraphs, [section({ lineNumbering: { start: 1, restart: 'newSection' } })],
  ), [[1, null, null], [2]]);
});

test('line numbering advances in column and page item order and tolerates negative settings', () => {
  const pages = [{ index: 0, number: 1, section: 0, items: [
    { type: 'line', paragraph: 0, column: 1 }, { type: 'line', paragraph: 1, column: 0 },
  ] }];
  const paragraphs = [{ index: 0, effective: {} }, { index: 1, effective: {} }];
  assert.deepEqual(lineNumbersFor(pages, paragraphs, [
    section({ columns: { count: 2, space: 0, equalWidth: true }, lineNumbering: { start: -2, countBy: -1 } }),
  ]), [[1, 2]]);
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

test('combinedTextLines splits by code point, not by UTF-16 code unit', () => {
  assert.deepEqual(combinedTextLines('股份有限'), ['股份', '有限']);
  // 奇数个字时上一行多一个，和 Word 一致。
  assert.deepEqual(combinedTextLines('有限公司五'), ['有限公', '司五']);
  assert.deepEqual(combinedTextLines(''), ['', '']);
  assert.deepEqual(combinedTextLines('甲'), ['甲', '']);
  // 补充平面的字各占两个码元，按码元切会把代理对劈成两个无效半码。字数要是**奇数**：
  // 偶数个时中点恰好落在代理对边界上，按码元切也碰巧是对的，用例就没有区分力了。
  assert.deepEqual(combinedTextLines('𠀀𠀁𠀂'), ['𠀀𠀁', '𠀂']);
  // 按码点迭代：合法的代理对会得到一个 > 0xFFFF 的字符，孤立的半码会得到一个落在
  // 0xD800–0xDFFF 里的字符。
  for (const half of combinedTextLines('𠀀𠀁𠀂')) {
    for (const character of half) {
      const code = character.codePointAt(0);
      assert.ok(code < 0xd800 || code > 0xdfff, `孤立的代理码元 U+${code.toString(16)}`);
    }
  }
});

test('combineBracketChars and rubyAlignToCss only claim what they can actually draw', () => {
  assert.deepEqual(combineBracketChars('round'), ['（', '）']);
  assert.deepEqual(combineBracketChars('square'), ['［', '］']);
  assert.deepEqual(combineBracketChars('angle'), ['〈', '〉']);
  assert.deepEqual(combineBracketChars('curly'), ['｛', '｝']);
  assert.equal(combineBracketChars('none'), undefined);
  assert.equal(combineBracketChars(undefined), undefined);
  assert.equal(rubyAlignToCss('center'), 'center');
  assert.equal(rubyAlignToCss('distributeLetter'), 'space-between');
  assert.equal(rubyAlignToCss('distributeSpace'), 'space-around');
  assert.equal(rubyAlignToCss('left'), 'start');
  // CSS 的 ruby-align 只有 start / center / space-between / space-around，
  // right 和 rightVertical 没有对应项，返回 undefined 用浏览器默认值，不硬凑近似。
  assert.equal(rubyAlignToCss('right'), undefined);
  assert.equal(rubyAlignToCss('rightVertical'), undefined);
  assert.equal(rubyAlignToCss(undefined), undefined);
});

test('overstrikeLayers tells content apart from instruction-only glyphs', () => {
  // 「合并字符」：两组文字拼起来正好是域结果，所以两层都是内容。
  const combined = {
    switch: 'o',
    parts: [
      { switch: 's', options: ['up'], raisePoints: 9, parts: [{ text: '股份' }] },
      { switch: 's', options: ['do'], raisePoints: -3, parts: [{ text: '有限' }] },
    ],
  };
  assert.deepEqual(overstrikeLayers(combined, '股份有限'), [
    { text: '股份', content: true, raisePoints: 9 },
    { text: '有限', content: true, raisePoints: -3 },
  ]);

  // 「带圈字符」：圈只在指令里，域结果只有那个字，所以圈只能当装饰画——否则编辑正文会把
  // 它写回文档。
  const circled = {
    switch: 'o', options: ['ac'],
    parts: [{ switch: 's', options: ['up'], raisePoints: 10, parts: [{ text: '○' }] }, { text: '甲' }],
  };
  assert.deepEqual(overstrikeLayers(circled, '甲'), [
    { text: '○', content: false, raisePoints: 10 },
    { text: '甲', content: true, raisePoints: 0 },
  ]);

  // 各层拼不满域结果时退回 undefined：指令和缓存结果不一致（域是脏的），此时宁可不叠印，
  // 也不能把结果里的文字丢掉。
  assert.equal(overstrikeLayers(combined, '股份有限公司'), undefined);
  assert.equal(overstrikeLayers(combined, '完全不同'), undefined);
  // 不是 \o 的 EQ 不接管渲染。
  assert.equal(overstrikeLayers({ switch: 'f', parts: [{ text: '1' }, { text: '2' }] }, '1/2'), undefined);
  assert.equal(overstrikeLayers(undefined, '甲'), undefined);
  assert.equal(overstrikeLayers({ switch: 'o' }, ''), undefined);
});

test('frameWrapExclusion takes the drop cap box from measurement and the frame box from w:w / w:h', () => {
  const measured = { widthPx: 80, heightPx: 60 };
  // 下沉字的尺寸不从 w:lines 反推：Word 已经把它的 w:sz 调到正好跨 lines 行，量出来的字框
  // 就是排除区。
  assert.deepEqual(frameWrapExclusion({ dropCap: 'drop', lines: 3 }, measured),
    { widthPx: 80, heightPx: 60, wrap: 'square', carried: false });
  assert.deepEqual(frameWrapExclusion({ dropCap: 'margin', lines: 2 }, measured),
    { widthPx: 80, heightPx: 60, wrap: 'square', carried: false });
  // 文档明写了尺寸就照它来，下沉字也不例外——这条分支以前被 dropCap 的特判盖掉了，而特判
  // 本身没有理由：「文档说什么就读什么」。
  assert.deepEqual(frameWrapExclusion({ dropCap: 'drop', lines: 3, widthTwips: 1500 }, measured),
    { widthPx: 100, heightPx: 60, wrap: 'square', carried: false });
  assert.deepEqual(frameWrapExclusion({ dropCap: 'drop', lines: 3, heightTwips: 1500 }, measured),
    { widthPx: 80, heightPx: 100, wrap: 'square', carried: false });
  // 定位文本框优先用 w:w / w:h（缇，15 缇 = 1px）。
  assert.deepEqual(frameWrapExclusion({ widthTwips: 2880, heightTwips: 1440, wrap: 'around' }, measured),
    { widthPx: 192, heightPx: 96, wrap: 'square', carried: false });
  // 缺的那一维用量出来的。
  assert.deepEqual(frameWrapExclusion({ widthTwips: 1500, wrap: 'around' }, measured),
    { widthPx: 100, heightPx: 60, wrap: 'square', carried: false });

  // w:wrap 的取值正好对上现有的排除区类型。
  const wrapOf = (wrap) => frameWrapExclusion({ wrap, widthTwips: 1500 }, measured)?.wrap;
  assert.equal(wrapOf('tight'), 'tight');
  assert.equal(wrapOf('through'), 'through');
  assert.equal(wrapOf('around'), 'square');
  assert.equal(wrapOf('auto'), 'square');
  // notBeside 的字面意思就是「旁边不许有文字」。
  assert.equal(wrapOf('notBeside'), 'topAndBottom');
  // none 是正文不绕它排，所以没有排除区。
  assert.equal(frameWrapExclusion({ wrap: 'none', widthTwips: 1500 }, measured), undefined);
  // 没写 w:wrap 时 Word 的缺省是绕排。
  assert.equal(wrapOf(undefined), 'square');

  assert.equal(frameWrapExclusion(undefined, measured), undefined);
  assert.equal(frameWrapExclusion(null, measured), undefined);
  // 宽度量不到又没给 w:w 时不编一个出来：不环绕只是少个效果，编错了会把正文挤歪。
  assert.equal(frameWrapExclusion({ dropCap: 'drop' }, { heightPx: 60 }), undefined);
  assert.equal(frameWrapExclusion({ dropCap: 'drop' }, { widthPx: 0, heightPx: 60 }), undefined);
  assert.equal(frameWrapExclusion({ widthTwips: 1500 }, { heightPx: 0 }), undefined);
});

test('a framed paragraph floats out of the flow and leaves an exclusion for what follows', () => {
  // 带 framePr 的段落自己不占纵向高度，只给后面的内容留出一块排除区——这正是浏览器 float
  // 的行为，所以分页测量和实际渲染是同一套。
  const areas = [];
  const widthMeasurer = {
    measureParagraph(p, area) {
      areas.push({ paragraph: p.index, wraps: area.wraps.map((w) => ({ ...w })) });
      return (p.lines ?? [10]).map((height, index) => ({
        heightPx: height, startOffset: index, endOffset: index + 1, widthPx: p.widthPx ?? 0,
      }));
    },
    measureTableRow(_table, row) { return row.height ?? 10; },
  };
  const dropCap = { index: 0, text: '从', runs: [], images: [], lines: [60], widthPx: 80,
    frame: { dropCap: 'drop', lines: 3, wrap: 'around' } };
  const bodyText = { index: 1, text: '前有座山。', runs: [], images: [], lines: [20, 20, 20] };
  // 页高给足，这样「浮出流外」与「照常计入」的对比不被分页搅进来（默认页高只有 100px）。
  const tall = [section({ pageHeight: 6000 })];
  const pages = paginate([{ type: 'paragraph', paragraph: dropCap }, { type: 'paragraph', paragraph: bodyText }],
    tall, widthMeasurer, { defaultTabStopTwips: 720 });
  assert.equal(pages.length, 1);
  // 下沉段的行照样进页面（要画出来），但不计入栏高：只有正文那 3 行 × 20 算进去。
  assert.equal(pages[0].contentHeightPx, 60);
  assert.deepEqual(pages[0].items.map((item) => item.paragraph), [0, 1, 1, 1]);
  // 排除区带到了后面的段落。
  const bodyArea = areas.find((entry) => entry.paragraph === 1);
  assert.deepEqual(bodyArea.wraps, [{ widthPx: 80, heightPx: 60, wrap: 'square', carried: true }]);

  // 排除区放不进本栏剩余高度、而本栏已经有内容时先换栏：否则它会跨到下一栏去挤正文，
  // 而它自己又不占高度，错位不会有任何提示。
  const before = { index: 0, text: '已有内容', runs: [], images: [], lines: [60] };
  const framed = { index: 1, text: '从', runs: [], images: [], lines: [60], widthPx: 80,
    frame: { dropCap: 'drop', wrap: 'around' } };
  const moved = paginate([{ type: 'paragraph', paragraph: before }, { type: 'paragraph', paragraph: framed }],
    [section({ pageHeight: 1500 })], widthMeasurer, { defaultTabStopTwips: 720 });
  assert.equal(moved.length, 2, '页高 100px，60px 内容之后放不下 60px 的排除区');
  assert.deepEqual(moved.map((page) => page.items.map((item) => item.paragraph)), [[0], [1]]);

  // wrap="none" 的段落不脱离正常流，照常计入栏高。
  areas.length = 0;
  const plain = paginate([
    { type: 'paragraph', paragraph: { ...dropCap, frame: { dropCap: 'drop', wrap: 'none' } } },
    { type: 'paragraph', paragraph: bodyText },
  ], tall, widthMeasurer, { defaultTabStopTwips: 720 });
  assert.equal(plain.length, 1);
  assert.equal(plain[0].contentHeightPx, 120, '60 + 3 × 20');
  assert.deepEqual(areas.find((entry) => entry.paragraph === 1).wraps, []);
});

test('header rows repeat when the table style supplies tblHeader through a condition', () => {
  // 表头行可能来自表格样式的 firstRow 条件，而不是行自己写的 w:tblHeader；分页重复表头必须
  // 认这一种，否则用内置样式的长表格翻页后就没表头了。
  const cell = () => ({ blocks: [], colSpan: 1, rowSpan: 1, isMergeContinuation: false });
  const build = (headerFrom) => ({
    type: 'table',
    index: 0,
    grid: [1000],
    rows: [
      { cells: [cell()], height: 30, ...headerFrom },
      ...Array.from({ length: 6 }, () => ({ cells: [cell()], format: {}, height: 30 })),
    ],
  });
  const paginated = (table) => paginate([table], [section({ pageHeight: 1500 })], measurer,
    { defaultTabStopTwips: 720 });
  // 重复的表头没有专门的标志，就是第 0 行在每个片段上各出现一次，所以数它的出现次数。
  const headerPerPage = (pages) => pages.map((page) => page.items
    .filter((item) => item.type === 'tableRow' && item.row === 0).length);

  const fromStyle = paginated(build({ format: {}, effective: { header: true } }));
  const fromRow = paginated(build({ format: { header: true } }));
  const plain = paginated(build({ format: {} }));
  assert.ok(fromStyle.length > 1, '页高 100px、7 行 × 30px，必然跨页');
  assert.deepEqual(headerPerPage(fromStyle), headerPerPage(fromRow),
    '样式条件给的表头和行自己写的表头，重复行为必须一致');
  // 没有表头时第 0 行只出现一次；有表头时后续页上会再出现。
  assert.equal(headerPerPage(plain).reduce((sum, count) => sum + count, 0), 1);
  assert.ok(headerPerPage(fromStyle).reduce((sum, count) => sum + count, 0) > 1, '后续页上重复了表头');
});

test('a floating table leaves the flow and leaves an exclusion for what follows', () => {
  // 浮动表格和带 framePr 的段落是同一件事：整块脱离正常流，只给后面的内容留出排除区，
  // 环绕交给浏览器 float。
  assert.deepEqual(tableWrapExclusion({ leftFromText: 180 }, { widthPx: 192, heightPx: 60 }),
    { widthPx: 192, heightPx: 60, wrap: 'square', carried: false });
  // w:tblpPr 上没有 w:wrap —— 浮动表格一定绕排，所以类型固定是 square；空对象也是浮动。
  assert.equal(tableWrapExclusion({}, { widthPx: 100, heightPx: 50 })?.wrap, 'square');
  assert.equal(tableWrapExclusion(undefined, { widthPx: 100, heightPx: 50 }), undefined);
  assert.equal(tableWrapExclusion(null, { widthPx: 100, heightPx: 50 }), undefined);
  // 宽高算不出来时不编一个出来。
  assert.equal(tableWrapExclusion({}, { heightPx: 50 }), undefined);
  assert.equal(tableWrapExclusion({}, { widthPx: 0, heightPx: 50 }), undefined);
  assert.equal(tableWrapExclusion({}, { widthPx: 100, heightPx: 0 }), undefined);

  const areas = [];
  const watcher = {
    measureParagraph(p, area) {
      areas.push({ paragraph: p.index, wraps: area.wraps.map((w) => ({ ...w })) });
      return (p.lines ?? [10]).map((height, index) => ({ heightPx: height, startOffset: index, endOffset: index + 1 }));
    },
    measureTableRow(_table, row) { return row.height ?? 10; },
  };
  const cell = () => ({ blocks: [], colSpan: 1, rowSpan: 1, isMergeContinuation: false });
  const floatingTable = {
    type: 'table', index: 0, grid: [2880],
    format: { floatingPosition: { leftFromText: 180 }, width: { type: 'dxa', value: 2880 } },
    rows: [{ cells: [cell()], format: {}, height: 30 }, { cells: [cell()], format: {}, height: 30 }],
  };
  const body = { index: 1, text: '正文', runs: [], images: [], lines: [20, 20] };
  const pages = paginate([floatingTable, { type: 'paragraph', paragraph: body }],
    [section({ pageHeight: 6000 })], watcher, { defaultTabStopTwips: 720 });
  assert.equal(pages.length, 1);
  // 表格的行照样进页面（要画出来），但不计入栏高：只有正文那 2 行 × 20 算进去。
  assert.equal(pages[0].contentHeightPx, 40);
  assert.deepEqual(pages[0].items.filter((item) => item.type === 'tableRow').map((item) => item.row), [0, 1]);
  // 排除区带到了后面的段落：宽度来自 w:tblW（2880 缇 ÷ 15），高度是两行合计。
  assert.deepEqual(areas.find((entry) => entry.paragraph === 1).wraps,
    [{ widthPx: 192, heightPx: 60, wrap: 'square', carried: true }]);

  // 没有 w:tblW 时用网格列宽合计。
  areas.length = 0;
  paginate([{ ...floatingTable, grid: [1500, 1500], format: { floatingPosition: {} } },
    { type: 'paragraph', paragraph: body }], [section({ pageHeight: 6000 })], watcher,
    { defaultTabStopTwips: 720 });
  assert.equal(areas.find((entry) => entry.paragraph === 1).wraps[0].widthPx, 200, '3000 缇 ÷ 15');

  // 不浮动的表格照常计入栏高。
  areas.length = 0;
  const plain = paginate([{ ...floatingTable, format: { width: { type: 'dxa', value: 2880 } } },
    { type: 'paragraph', paragraph: body }], [section({ pageHeight: 6000 })], watcher,
    { defaultTabStopTwips: 720 });
  assert.equal(plain[0].contentHeightPx, 100, '60 + 2 × 20');
  assert.deepEqual(areas.find((entry) => entry.paragraph === 1).wraps, []);
});

test('paragraph snapToGrid=false opts that paragraph out of document-grid line snapping', () => {
  const grid = { type: 'linesAndChars', linePitch: 300, charSpace: 0 };
  // 开关只在显式 false 时生效；默认与显式 true 都照常吸附。
  assert.equal(snapLineHeightPx(15, grid), 20);
  assert.equal(snapLineHeightPx(15, grid, true), 20);
  assert.equal(snapLineHeightPx(15, grid, null), 20);
  assert.equal(snapLineHeightPx(15, grid, false), 15);

  // 分页走的是同一个函数：关掉的那段每行留在 15，没关的那段吸到 20。
  const result = paginate(blocks(
    paragraph(0, [15, 15], { snapToGrid: false }),
    paragraph(1, [15, 15]),
    // effective 优先于直接格式，和其它段落属性一致。
    paragraph(2, [15], { snapToGrid: true, effective: { snapToGrid: false } }),
  ), [section({ pageHeight: 3000, docGrid: grid })], measurer, { defaultTabStopTwips: 720 });
  const heights = result[0].items.filter((item) => item.type === 'line').map((item) => [item.paragraph, item.line.heightPx]);
  assert.deepEqual(heights, [[0, 15], [0, 15], [1, 20], [1, 20], [2, 15]]);
});
