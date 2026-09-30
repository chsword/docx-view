import { test } from 'node:test';
import assert from 'node:assert/strict';
import { DocxDocument } from '../dist/document.js';
import {
  DocxEditor, deriveLineBoxes, formatPageNumber, paginationInfoFromPages, replacePageFields,
  selectHeaderFooter, updateFieldsUntilStable,
} from '../dist/editor.js';
import { sanitizeTextWithInfo, WORD_NS } from '../dist/xml.js';

test('formats paginated page numbers and falls back to decimal', () => {
  assert.equal(formatPageNumber(4, 'upperRoman'), 'IV');
  assert.equal(formatPageNumber(27, 'lowerLetter'), 'aa');
  assert.equal(formatPageNumber(12, 'chineseCounting'), '一二');
  assert.equal(formatPageNumber(7, 'not-a-format'), '7');
});

test('field pagination adapter maps body and table paragraphs to displayed page numbers', () => {
  const paragraph = index => ({ type: 'paragraph', paragraph: { index } });
  const table = {
    type: 'table',
    rows: [{ cells: [{ blocks: [paragraph(2)] }] }],
    grid: [],
  };
  const pages = [
    { index: 0, number: 4, section: 0, items: [{ type: 'line', paragraph: 0 }] },
    { index: 1, number: 5, section: 0, items: [{ type: 'tableRow', table: 1, row: 0 }] },
  ];
  const pagination = paginationInfoFromPages(pages, [paragraph(0), table]);
  assert.equal(pagination.pageCount, 2);
  assert.equal(pagination.pageOfParagraph(0), 0);
  assert.equal(pagination.pageOfParagraph(2), 1);
  assert.equal(pagination.numberOfPage(pagination.pageOfParagraph(2)), 5);
  assert.equal(pagination.pageOfParagraph(1), undefined);
});

test('DocxEditor.updateFields passes computed pagination through to document writeback', () => {
  const block = { type: 'paragraph', paragraph: { index: 0 } };
  const pages = [{ index: 0, number: 8, section: 0, items: [{ type: 'line', paragraph: 0 }] }];
  const calls = [];
  const editor = Object.create(DocxEditor.prototype);
  editor.document = {
    getBlocks: () => [block],
    getSections: () => [],
    getSettings: () => ({ defaultTabStop: 720 }),
    updateFields: ({ pagination }) => {
      calls.push(pagination);
      return calls.length === 1;
    },
  };
  editor.paginateDocument = () => pages;
  let renders = 0;
  editor.render = () => { renders++; };
  assert.equal(editor.updateFields(), true);
  assert.equal(calls.length, 2);
  assert.equal(calls[0].pageCount, 1);
  assert.equal(calls[0].pageOfParagraph(0), 0);
  assert.equal(calls[0].numberOfPage(0), 8);
  assert.equal(renders, 1);
});

test('field update iteration stops when stable and caps non-convergent updates', () => {
  let stableCalls = 0;
  assert.deepEqual(updateFieldsUntilStable(() => ++stableCalls === 1), { updated: true, iterations: 2 });
  let changingCalls = 0;
  assert.deepEqual(updateFieldsUntilStable(() => { changingCalls++; return true; }),
    { updated: true, iterations: 5 });
  assert.equal(changingCalls, 5);
});

test('groups browser line rectangles into continuous character ranges', () => {
  assert.deepEqual(deriveLineBoxes([
    { top: 10, height: 18, start: 0, end: 1 },
    { top: 10.5, height: 18, start: 1, end: 3 },
    { top: 29, height: 18, start: 3, end: 5 },
  ]), [
    { heightPx: 18, startOffset: 0, endOffset: 3 },
    { heightPx: 18, startOffset: 3, endOffset: 5 },
  ]);
});

test('pagination paragraph slices retain zero-length inline content', () => {
  const editor = Object.create(DocxEditor.prototype);
  const paragraph = {
    index: 0,
    text: '前文',
    runs: [
      { index: 0, text: '前文', images: [] },
      { index: 1, text: '', images: [{ id: 'image-1' }] },
      { index: 2, text: '', noteReference: { kind: 'footnote', id: 1, number: 1, marker: '1' } },
    ],
    images: [],
  };
  const slice = editor.sliceParagraph(paragraph, 0, paragraph.text.length);
  assert.equal(slice.text, '前文');
  assert.equal(slice.runs[1].images[0].id, 'image-1');
  assert.equal(slice.runs[2].noteReference.marker, '1');
});

test('pagination paragraph slices concatenate back to the original without overlap or loss', () => {
  const editor = Object.create(DocxEditor.prototype);
  const paragraph = {
    index: 3,
    text: 'abcdefghij',
    runs: [
      { index: 0, text: '', images: [{ id: 'leading' }] },
      { index: 1, text: 'abc', images: [] },
      { index: 2, text: '', images: [{ id: 'between' }] },
      { index: 3, text: 'defgh', images: [] },
      { index: 4, text: 'ij', images: [] },
      { index: 5, text: '', noteReference: { kind: 'footnote', id: 1, number: 1, marker: '1' } },
    ],
    images: [],
  };
  const length = paragraph.text.length;
  for (let first = 1; first < length; first++) {
    for (let second = first; second < length; second++) {
      const bounds = [[0, first], [first, second], [second, length]].filter(([start, end]) => end > start);
      const slices = bounds.map(([start, end]) => editor.sliceParagraph(paragraph, start, end));
      const label = `cuts at ${first}/${second}`;
      for (const slice of slices) {
        assert.equal(slice.index, paragraph.index, label);
        assert.equal(slice.runs.map((run) => run.text).join(''), slice.text, label);
      }
      assert.equal(slices.map((slice) => slice.text).join(''), paragraph.text, label);
      const runs = slices.flatMap((slice) => slice.runs);
      assert.deepEqual(runs.filter((run) => !run.text).map((run) => run.index), [0, 2, 5], label);
      assert.deepEqual([...new Set(runs.map((run) => run.index))], [0, 1, 2, 3, 4, 5], label);
      const byRun = new Map();
      for (const run of runs) byRun.set(run.index, (byRun.get(run.index) ?? '') + run.text);
      for (const run of paragraph.runs) assert.equal(byRun.get(run.index), run.text, `${label} run ${run.index}`);
    }
  }
  assert.equal(paragraph.runs[3].text, 'defgh');
});

test('paginated header/footer kind follows title page, first physical page and parity', () => {
  const all = { default: 'word/header1.xml', first: 'word/header2.xml', even: 'word/header3.xml' };
  const pick = (options, parts = all) => selectHeaderFooter(parts, { selectedKind: 'default', ...options });
  assert.deepEqual(pick({ pageNumber: 1, firstPhysicalPage: true, titlePage: true }), { kind: 'first', part: all.first });
  assert.deepEqual(pick({ pageNumber: 1, firstPhysicalPage: true, titlePage: false }), { kind: 'default', part: all.default });
  assert.deepEqual(pick({ pageNumber: 2, firstPhysicalPage: false, titlePage: true }), { kind: 'even', part: all.even });
  assert.deepEqual(pick({ pageNumber: 3, firstPhysicalPage: false, titlePage: true }), { kind: 'default', part: all.default });
  // A section that starts on an even page still shows its first-page variant there.
  assert.deepEqual(pick({ pageNumber: 4, firstPhysicalPage: true, titlePage: true }), { kind: 'first', part: all.first });
  assert.deepEqual(pick({ pageNumber: 4, firstPhysicalPage: true, titlePage: false }), { kind: 'even', part: all.even });
  // Missing variants keep their kind but fall back to the default part.
  const defaultOnly = { default: 'word/header1.xml' };
  assert.deepEqual(pick({ pageNumber: 1, firstPhysicalPage: true, titlePage: true }, defaultOnly), { kind: 'first', part: defaultOnly.default });
  assert.deepEqual(pick({ pageNumber: 2, firstPhysicalPage: false, titlePage: false }, defaultOnly), { kind: 'even', part: defaultOnly.default });
  assert.deepEqual(pick({ pageNumber: 2, firstPhysicalPage: false, titlePage: false }, {}), { kind: 'even', part: undefined });
  // Without a physical page number the caller's selected kind is kept.
  assert.deepEqual(selectHeaderFooter(all, { selectedKind: 'even', firstPhysicalPage: true, titlePage: true }), { kind: 'even', part: all.even });
  assert.deepEqual(selectHeaderFooter(defaultOnly, { selectedKind: 'first' }), { kind: 'first', part: defaultOnly.default });
});

test('paginated header/footer replaces only PAGE and NUMPAGES field results', () => {
  const run = (index, text, field) => ({ index, text, images: [], ...(field ? { field } : {}) });
  const runs = [
    run(0, 'Page '),
    run(1, '', { index: 0, role: 'instruction', kind: 'PAGE', instruction: ' PAGE ' }),
    run(2, '1', { index: 0, role: 'result', kind: 'PAGE', instruction: ' PAGE ' }),
    run(3, ' of '),
    run(4, '', { index: 1, role: 'instruction', kind: 'NUMPAGES', instruction: ' NUMPAGES ' }),
    run(5, '9', { index: 1, role: 'result', kind: 'NUMPAGES', instruction: ' NUMPAGES ' }),
    run(6, '7', { index: 2, role: 'result', kind: 'SEQ', instruction: ' SEQ x ' }),
    run(7, '2', { index: 3, role: 'result', instruction: ' UNKNOWN ' }),
  ];
  const table = { type: 'table', rows: [], grid: [] };
  const blocks = [{ type: 'paragraph', paragraph: { index: 0, text: 'Page 1 of 972', runs, images: [] } }, table, { type: 'pageBreak' }];
  const replaced = replacePageFields(blocks, 3, 12);
  assert.deepEqual(replaced[0].paragraph.runs.map((entry) => entry.text), ['Page ', '', '3', ' of ', '', '12', '7', '2']);
  for (const index of [0, 1, 3, 4, 6, 7]) assert.equal(replaced[0].paragraph.runs[index], runs[index]);
  assert.deepEqual(replaced[0].paragraph.runs[2].field, runs[2].field);
  assert.equal(replaced[1], table);
  assert.equal(replaced[2], blocks[2]);
  assert.deepEqual(runs.map((entry) => entry.text), ['Page ', '', '1', ' of ', '', '9', '7', '2']);
  assert.deepEqual(replacePageFields(blocks, 4, 14, 'upperRoman')[0].paragraph.runs.map((entry) => entry.text)
    .filter((_, index) => index === 2 || index === 5), ['IV', 'XIV']);
  assert.equal(replacePageFields(blocks, 2, undefined)[0].paragraph.runs[5].text, '0');
});

function makeFlushEditor({ text, elementText = text, previous = '', options = {}, document = DocxDocument.create() }) {
  const editor = Object.create(DocxEditor.prototype);
  editor.destroyed = false;
  editor.document = document;
  editor.options = options;
  const element = { text: elementText };
  const content = { text };
  editor.paragraphs = new Map([[0, { element, content, text: previous, failed: false }]]);
  editor.readText = (entry) => entry.text;
  return { editor, element, content, document };
}

function makeSelectionEditor(cellByParagraph = {}) {
  const editor = Object.create(DocxEditor.prototype);
  const events = [];
  class FakeCustomEvent {
    constructor(type, init = {}) {
      this.type = type;
      this.detail = init.detail;
      this.bubbles = init.bubbles;
    }
  }
  editor.selected = null;
  editor.selectedImageInfo = null;
  editor.selectedTableCellInfo = null;
  editor.document = { getTableCellAt: (index) => cellByParagraph[index] ?? null };
  editor.root = {
    ownerDocument: { defaultView: { CustomEvent: FakeCustomEvent } },
    dispatchEvent: (event) => { events.push(event); return true; },
  };
  return { editor, events };
}

function makeRunRenderEditor({ showRevisions = true, revisionView = 'markup' } = {}) {
  const editor = Object.create(DocxEditor.prototype);
  editor.reviewFilter = { showRevisions, showComments: true, revisionView };
  editor.options = { showFormattingMarks: false };
  editor.commentParagraphIds = new Map();
  editor.commentRunIds = new Map();
  editor.revisionRunIds = new Map();
  editor.revisionParagraphIds = new Map();
  const createElement = (tagName) => {
    const element = {
      nodeType: 1,
      tagName: tagName.toUpperCase(),
      dataset: {},
      style: {},
      childNodes: [],
      className: '',
      textContent: '',
      attributes: new Map(),
      append(...children) { this.childNodes.push(...children); },
      appendChild(child) { this.childNodes.push(child); return child; },
      addEventListener() {},
      contains() { return false; },
      setAttribute(name, value) { this.attributes.set(name, value); },
      getBoundingClientRect() { return { height: this.tagName === 'TD' ? 27 : 0 }; },
      remove() {},
    };
    if (element.tagName === 'TABLE') {
      element.createTBody = () => {
        const body = createElement('tbody');
        element.append(body);
        return body;
      };
      Object.defineProperty(element, 'rows', {
        get: () => element.childNodes.flatMap((child) => child.tagName === 'TBODY' ? child.childNodes : []),
      });
    }
    if (element.tagName === 'TBODY') {
      element.insertRow = () => {
        const row = createElement('tr');
        element.append(row);
        return row;
      };
    }
    if (element.tagName === 'TR') {
      element.insertCell = () => {
        const cell = createElement('td');
        element.append(cell);
        return cell;
      };
      Object.defineProperty(element, 'cells', { get: () => element.childNodes });
    }
    element.classList = {
      add: (...names) => {
        for (const name of names) {
          if (!name) continue;
          element.className = element.className ? `${element.className} ${name}` : name;
        }
      },
    };
    return element;
  };
  editor.root = {
    ownerDocument: {
      createElement,
      createElementNS: (_namespace, tagName) => createElement(tagName),
      createTextNode: (text) => ({ nodeType: 3, textContent: text }),
    },
    append() {},
  };
  return editor;
}

test('makeShape renders SVG geometry beneath shape text on repeated renders', () => {
  const editor = makeRunRenderEditor();
  editor.document = { getShapeParagraphs: () => [{ text: 'Shape text', runs: [] }] };
  const shape = {
    id: 'shape-svg',
    paragraph: 0,
    run: 0,
    kind: 'textbox',
    form: 'drawingml',
    widthPx: 120,
    heightPx: 80,
    placement: 'inline',
    hasTextContent: true,
    geometry: 'diamond',
    fill: { type: 'solid', color: '#123456' },
    line: { color: '#abcdef', widthPx: 2, dash: '4 2' },
    rotation: 45,
    flipH: true,
    flipV: true,
  };
  const reviewContext = { deletedTextByRun: new Map(), revisionColors: new Map() };
  for (let render = 0; render < 2; render++) {
    const wrapper = editor.makeShape(shape, 720, reviewContext);
    const [svg, text] = wrapper.childNodes;
    const path = svg.childNodes.find((node) => node.tagName === 'PATH');
    assert.equal(svg.tagName, 'SVG');
    assert.equal(path.attributes.get('d'), 'M 60 0 L 120 40 L 60 80 L 0 40 Z');
    assert.equal(path.attributes.get('fill'), '#123456');
    assert.equal(path.attributes.get('stroke'), '#abcdef');
    assert.equal(path.attributes.get('stroke-width'), '2');
    assert.equal(path.attributes.get('stroke-dasharray'), '4 2');
    assert.match(path.attributes.get('transform'), /rotate\(45\).*scale\(-1 -1\)/);
    assert.equal(text.className, 'docx-shape-paragraph');
    assert.equal(text.textContent, 'Shape text');
    assert.ok(wrapper.childNodes.indexOf(svg) < wrapper.childNodes.indexOf(text));
  }
});

test('makeShape renders pre-rendered SmartArt children at their offsets with text', () => {
  const editor = makeRunRenderEditor();
  editor.document = {};
  const wrapper = editor.makeShape({
    id: 'smartart-svg',
    paragraph: 0,
    run: 0,
    kind: 'smartArt',
    form: 'drawingml',
    widthPx: 640,
    heightPx: 240,
    placement: 'inline',
    hasTextContent: false,
    children: [
      { offsetXPx: 12, offsetYPx: 8, widthPx: 100, heightPx: 40, geometry: 'rect', fill: { type: 'solid', color: '#123456' }, text: 'First' },
      { offsetXPx: 140, offsetYPx: 8, widthPx: 100, heightPx: 40, geometry: 'ellipse', fill: { type: 'solid', color: '#abcdef' }, text: 'Second' },
    ],
  }, 720, { deletedTextByRun: new Map(), revisionColors: new Map() });
  const svg = wrapper.childNodes[0];
  const children = svg.childNodes.filter((node) => node.attributes?.get('data-docx-shape-child') !== undefined);
  assert.equal(children.length, 2);
  assert.equal(children[0].attributes.get('fill'), '#123456');
  assert.equal(children[0].attributes.get('transform'), 'translate(12 8)');
  assert.deepEqual(svg.childNodes.filter((node) => node.tagName === 'TEXT').map((node) => node.textContent), ['First', 'Second']);
});

test('makeShape renders chart axes, ticks, and series paths', () => {
  const editor = makeRunRenderEditor();
  editor.document = {};
  const wrapper = editor.makeShape({
    id: 'chart-svg',
    paragraph: 0,
    run: 0,
    kind: 'chart',
    form: 'drawingml',
    widthPx: 320,
    heightPx: 180,
    placement: 'inline',
    hasTextContent: false,
    chart: {
      kind: 'line',
      title: 'Sales',
      categories: ['Q1', 'Q2', 'Q3'],
      series: [{ name: 'Actual', values: [1, null, 3], line: { color: '#123456' } }],
      axes: { category: { visible: true }, value: { visible: true, majorGridlines: true } },
    },
  }, 720, { deletedTextByRun: new Map(), revisionColors: new Map() });
  const svg = wrapper.childNodes[0];
  assert.equal(svg.tagName, 'SVG');
  assert.ok(svg.childNodes.some((node) => node.attributes?.get('data-docx-chart-axis') === 'value'));
  assert.ok(svg.childNodes.some((node) => node.attributes?.get('data-docx-chart-tick') === '1'));
  const series = svg.childNodes.find((node) => node.attributes?.get('data-docx-chart-series') === '0');
  assert.equal(series.tagName, 'PATH');
  assert.match(series.attributes.get('d'), /M .* M /);
});

test('makeShape renders category labels and uses cached scatter x values', () => {
  const editor = makeRunRenderEditor();
  editor.document = {};
  const wrapper = editor.makeShape({
    id: 'scatter-svg',
    paragraph: 0,
    run: 0,
    kind: 'chart',
    form: 'drawingml',
    widthPx: 320,
    heightPx: 180,
    placement: 'inline',
    hasTextContent: false,
    chart: {
      kind: 'scatter',
      categories: [],
      series: [{ values: [1, 2, 3], xValues: [1, 2, 100], line: { color: '#123456' } }],
      axes: { category: { visible: true }, value: { visible: true, majorGridlines: false } },
    },
  }, 720, { deletedTextByRun: new Map(), revisionColors: new Map() });
  const svg = wrapper.childNodes[0];
  const path = svg.childNodes.find((node) => node.attributes?.get('data-docx-chart-series') === '0');
  const coordinates = path.attributes.get('d').match(/M ([\d.]+) [\d.]+ L ([\d.]+) [\d.]+ L ([\d.]+) [\d.]+/).slice(1).map(Number);
  assert.ok(coordinates[1] - coordinates[0] < 10);
  assert.ok(coordinates[2] - coordinates[1] > 200);

  const categoryWrapper = editor.makeShape({
    ...wrapper,
    id: 'category-svg',
    chart: {
      kind: 'line',
      categories: ['Q1', 'Q2'],
      series: [{ values: [1, 2], line: { color: '#123456' } }],
      axes: { category: { visible: true }, value: { visible: true, majorGridlines: false } },
    },
  }, 720, { deletedTextByRun: new Map(), revisionColors: new Map() });
  assert.deepEqual(categoryWrapper.childNodes[0].childNodes.filter((node) => node.attributes?.get('data-docx-chart-category') === '1').map((node) => node.textContent), ['Q1', 'Q2']);
});

test('makeShape closes an area segment at its own ends and swaps horizontal bar axes', () => {
  const editor = makeRunRenderEditor();
  editor.document = {};
  const areaWrapper = editor.makeShape({
    id: 'area-chart',
    paragraph: 0,
    run: 0,
    kind: 'chart',
    form: 'drawingml',
    widthPx: 320,
    heightPx: 180,
    placement: 'inline',
    hasTextContent: false,
    chart: {
      kind: 'area',
      categories: ['甲', '乙', '丙'],
      series: [{ values: [null, 1, 4], fill: { color: '#123456' } }],
      axes: { category: { visible: true }, value: { visible: true, majorGridlines: false } },
    },
  }, 720, { deletedTextByRun: new Map(), revisionColors: new Map() });
  const areaSvg = areaWrapper.childNodes[0];
  const area = areaSvg.childNodes.find((node) => node.attributes?.get('data-docx-chart-series') === '0');
  assert.equal(area.attributes.get('fill'), '#123456');
  const closure = area.attributes.get('d').match(/L ([\d.]+) ([\d.]+) L ([\d.]+) ([\d.]+) Z\s*$/);
  assert.ok(closure);
  // 序列开头是 null，填充收在第一个有数据的点（x=173）而不是绘图区左边界（x=34）
  assert.deepEqual(closure.slice(1).map(Number), [312, 156, 173, 156]);

  const barWrapper = editor.makeShape({
    id: 'horizontal-bar-chart',
    paragraph: 0,
    run: 0,
    kind: 'chart',
    form: 'drawingml',
    widthPx: 320,
    heightPx: 180,
    placement: 'inline',
    hasTextContent: false,
    chart: {
      kind: 'bar',
      barDirection: 'bar',
      categories: ['甲', '乙'],
      series: [{ values: [1, 2], fill: { color: '#654321' } }],
      axes: { category: { visible: true }, value: { visible: true, majorGridlines: true } },
    },
  }, 720, { deletedTextByRun: new Map(), revisionColors: new Map() });
  const barSvg = barWrapper.childNodes[0];
  const categoryLabels = barSvg.childNodes.filter((node) => node.attributes?.get('data-docx-chart-category') === '1');
  const valueLabels = barSvg.childNodes.filter((node) => node.attributes?.get('data-docx-chart-tick') === '1');
  assert.deepEqual(categoryLabels.map((node) => node.textContent), ['甲', '乙']);
  assert.ok(categoryLabels.every((node) => Number(node.attributes.get('x')) < 34));
  assert.ok(Number(categoryLabels[0].attributes.get('y')) < Number(categoryLabels[1].attributes.get('y')));
  assert.ok(valueLabels.every((node) => Number(node.attributes.get('y')) > 150));
  const categoryAxis = barSvg.childNodes.find((node) => node.attributes?.get('data-docx-chart-axis') === 'category');
  const valueAxis = barSvg.childNodes.find((node) => node.attributes?.get('data-docx-chart-axis') === 'value');
  assert.equal(categoryAxis.attributes.get('x1'), categoryAxis.attributes.get('x2'));
  assert.equal(valueAxis.attributes.get('y1'), valueAxis.attributes.get('y2'));
  const bars = barSvg.childNodes.filter((node) => node.attributes?.get('data-docx-chart-series') === '0');
  assert.ok(bars.every((node) => Number(node.attributes.get('y')) + Number(node.attributes.get('height')) <= 156));
});

test('makeShape closes each area gap segment separately without filling the gap', () => {
  const editor = makeRunRenderEditor();
  editor.document = {};
  const areaPath = (values) => {
    const wrapper = editor.makeShape({
      id: 'area-gap',
      paragraph: 0,
      run: 0,
      kind: 'chart',
      form: 'drawingml',
      widthPx: 320,
      heightPx: 180,
      placement: 'inline',
      hasTextContent: false,
      chart: {
        kind: 'area',
        categories: ['甲', '乙', '丙', '丁'],
        series: [{ values, fill: { color: '#123456' } }],
        axes: { category: { visible: true }, value: { visible: true, majorGridlines: false } },
      },
    }, 720, { deletedTextByRun: new Map(), revisionColors: new Map() });
    return wrapper.childNodes[0].childNodes
      .find((node) => node.attributes?.get('data-docx-chart-series') === '0')
      .attributes.get('d');
  };

  // 数据完整时四个类目铺满绘图区 34..312，闭合落在零值线 156 上
  assert.equal(areaPath([1, 2, 3, 4]), 'M 34 118 L 126.66666666666666 80 L 219.33333333333331 42 L 312 4 L 312 156 L 34 156 Z');

  // 中间缺一点：拆成两段，每段各自闭合，缺口区间没有填充
  const interior = areaPath([1, null, 3, 4]);
  assert.equal((interior.match(/M /g) ?? []).length, (interior.match(/Z/g) ?? []).length);
  assert.equal((interior.match(/Z/g) ?? []).length, 2);
  assert.equal(interior, 'M 34 118 L 34 156 Z M 219.33333333333331 42 L 312 4 L 312 156 L 219.33333333333331 156 Z');

  // 结尾缺两点：闭合回到最后一个有数据的点，而不是绘图区右边界 312
  assert.equal(areaPath([1, 2, null, null]), 'M 34 80 L 126.66666666666666 4 L 126.66666666666666 156 L 34 156 Z');
});

test('makeShape keeps pie slice and category legend colors aligned', () => {
  const editor = makeRunRenderEditor();
  editor.document = {};
  const wrapper = editor.makeShape({
    id: 'pie-svg',
    paragraph: 0,
    run: 0,
    kind: 'chart',
    form: 'drawingml',
    widthPx: 320,
    heightPx: 180,
    placement: 'inline',
    hasTextContent: false,
    chart: {
      kind: 'pie',
      categories: ['Q1', 'Q2'],
      series: [{
        values: [1, 2],
        fill: { type: 'solid', color: '#111111' },
        pointFills: [{ type: 'solid', color: '#ff0000' }, { type: 'solid', color: '#00ff00' }],
      }],
      axes: { category: { visible: true }, value: { visible: true, majorGridlines: false } },
    },
  }, 720, { deletedTextByRun: new Map(), revisionColors: new Map() });
  const svg = wrapper.childNodes[0];
  const slices = svg.childNodes.filter((node) => node.attributes?.get('data-docx-chart-series') === '0');
  assert.deepEqual(slices.map((node) => node.attributes.get('fill')), ['#ff0000', '#00ff00']);
  assert.deepEqual(svg.childNodes.filter((node) => node.tagName === 'RECT').map((node) => node.attributes.get('fill')), ['#ff0000', '#00ff00']);
});

test('makeShape renders a single-category pie without a degenerate arc and wraps its legend', () => {
  const editor = makeRunRenderEditor();
  editor.document = {};
  const singlePieWrapper = editor.makeShape({
    id: 'single-pie',
    paragraph: 0,
    run: 0,
    kind: 'chart',
    form: 'drawingml',
    widthPx: 320,
    heightPx: 180,
    placement: 'inline',
    hasTextContent: false,
    chart: {
      kind: 'pie',
      categories: ['Only'],
      series: [{ values: [1], fill: { color: '#123456' } }],
      axes: { category: { visible: true }, value: { visible: true, majorGridlines: false } },
    },
  }, 720, { deletedTextByRun: new Map(), revisionColors: new Map() });
  const singleSlice = singlePieWrapper.childNodes[0].childNodes.find((node) => node.tagName === 'PATH');
  assert.equal((singleSlice.attributes.get('d').match(/ A /g) ?? []).length, 2);
  assert.match(singleSlice.attributes.get('d'), / A 74 74 0 0 1 173 154 A 74 74 0 0 1 173 6 Z$/);

  const categories = Array.from({ length: 8 }, (_, index) => `Category ${index + 1}`);
  const legendWrapper = editor.makeShape({
    id: 'wrapped-pie-legend',
    paragraph: 0,
    run: 0,
    kind: 'chart',
    form: 'drawingml',
    widthPx: 320,
    heightPx: 180,
    placement: 'inline',
    hasTextContent: false,
    chart: {
      kind: 'pie',
      categories,
      series: [{ values: categories.map(() => 1), fill: { color: '#123456' } }],
      axes: { category: { visible: true }, value: { visible: true, majorGridlines: false } },
    },
  }, 720, { deletedTextByRun: new Map(), revisionColors: new Map() });
  const legendSvg = legendWrapper.childNodes[0];
  const legendLabels = legendSvg.childNodes.filter((node) => categories.includes(node.textContent));
  const legendMarkers = legendSvg.childNodes.filter((node) => node.tagName === 'RECT');
  assert.equal(legendLabels.length, 8);
  assert.equal(legendMarkers.length, 8);
  assert.ok([...legendLabels, ...legendMarkers].every((node) => Number(node.attributes.get('x')) >= 0 && Number(node.attributes.get('x')) < 320));
  assert.ok(new Set(legendLabels.map((node) => node.attributes.get('y'))).size > 1);
});

test('makeShape shares the scatter x scale across series', () => {
  const editor = makeRunRenderEditor();
  editor.document = {};
  const wrapper = editor.makeShape({
    id: 'scatter-shared-scale',
    paragraph: 0,
    run: 0,
    kind: 'chart',
    form: 'drawingml',
    widthPx: 320,
    heightPx: 180,
    placement: 'inline',
    hasTextContent: false,
    chart: {
      kind: 'scatter',
      categories: [],
      series: [
        { values: [1, 2], xValues: [0, 10], line: { color: '#111111' } },
        { values: [1, 2], xValues: [0, 1000], line: { color: '#222222' } },
      ],
      axes: { category: { visible: false }, value: { visible: true, majorGridlines: false } },
    },
  }, 720, { deletedTextByRun: new Map(), revisionColors: new Map() });
  const paths = wrapper.childNodes[0].childNodes.filter((node) => node.attributes?.get('data-docx-chart-series') !== undefined);
  const firstEnd = Number(paths[0].attributes.get('d').match(/L ([\d.]+) /)[1]);
  const secondEnd = Number(paths[1].attributes.get('d').match(/L ([\d.]+) /)[1]);
  assert.ok(firstEnd < secondEnd);
});

test('makeShape leaves external picture fills as a local SVG placeholder', () => {
  const editor = makeRunRenderEditor();
  let partReads = 0;
  editor.document = { getPartBytes: () => { partReads++; throw new Error('must not read external image'); } };
  const wrapper = editor.makeShape({
    id: 'external-fill',
    paragraph: 0,
    run: 0,
    kind: 'shape',
    form: 'drawingml',
    widthPx: 100,
    heightPx: 50,
    placement: 'inline',
    hasTextContent: false,
    fill: { type: 'picture' },
  }, 720, { deletedTextByRun: new Map(), revisionColors: new Map() });
  const svg = wrapper.childNodes[0];
  const path = svg.childNodes.find((node) => node.tagName === 'RECT');
  assert.equal(path.attributes.get('fill'), '#f7f9fd');
  assert.equal(path.attributes.get('stroke-dasharray'), '6 4');
  assert.equal(svg.childNodes.some((node) => node.tagName === 'IMAGE'), false);
  assert.equal(partReads, 0);
});

test('makeShape renders package picture fills through an SVG pattern', () => {
  const editor = makeRunRenderEditor();
  let partReads = 0;
  editor.document = { getPartBytes: (path) => {
    assert.equal(path, 'word/media/shape.png');
    partReads++;
    return Uint8Array.from([1, 2, 3]);
  } };
  const wrapper = editor.makeShape({
    id: 'internal-fill',
    paragraph: 0,
    run: 0,
    kind: 'shape',
    form: 'drawingml',
    widthPx: 100,
    heightPx: 50,
    placement: 'inline',
    hasTextContent: false,
    fill: { type: 'picture', imagePartPath: 'word/media/shape.png' },
  }, 720, { deletedTextByRun: new Map(), revisionColors: new Map() });
  const svg = wrapper.childNodes[0];
  const pattern = svg.childNodes[0].childNodes[0];
  const image = pattern.childNodes[0];
  const shapePath = svg.childNodes.find((node) => node.tagName === 'PATH' || node.tagName === 'RECT');
  assert.equal(shapePath.attributes.get('fill'), 'url(#shape-pattern-internal-fill)');
  assert.match(image.attributes.get('href'), /^data:image\/png;base64,/);
  assert.equal(partReads, 1);
});

test('makeShape falls back to a styled rectangle for unsupported presets', () => {
  const editor = makeRunRenderEditor();
  editor.document = {};
  const wrapper = editor.makeShape({
    id: 'fallback-preset',
    paragraph: 0,
    run: 0,
    kind: 'shape',
    form: 'drawingml',
    widthPx: 100,
    heightPx: 40,
    placement: 'inline',
    hasTextContent: false,
    geometry: 'flowChartMagneticDisk',
    fill: { type: 'solid', color: '#654321' },
    line: { color: '#abcdef', widthPx: 3 },
  }, 720, { deletedTextByRun: new Map(), revisionColors: new Map() });
  const svg = wrapper.childNodes[0];
  const rect = svg.childNodes.find((node) => node.tagName === 'RECT');
  assert.ok(rect);
  assert.equal(rect.attributes.get('fill'), '#654321');
  assert.equal(rect.attributes.get('stroke'), '#abcdef');
  assert.equal(rect.attributes.get('stroke-width'), '3');
});

test('makeParagraph renders shapes only for matching renderShapeInfos', () => {
  const editor = makeRunRenderEditor();
  editor.paragraphs = new Map();
  editor.measuring = false;
  editor.composing = false;
  editor.renderAfterComposition = false;
  editor.readText = (content) => content.textContent ?? '';
  editor.document = { getShapeParagraphs: () => [] };
  const paragraph = { index: 0, text: 'text', runs: [{ index: 0, text: 'text' }] };
  const reviewContext = { deletedTextByRun: new Map(), revisionColors: new Map() };

  editor.renderShapeInfos = [{ paragraph: 0, run: 0, kind: 'textbox', hasTextContent: false, id: 'shape-1' }];
  const withShape = editor.makeParagraph(paragraph, 720, reviewContext);
  assert.equal(withShape.childNodes[0].childNodes.some((node) => node.className.includes('docx-shape')), true);

  editor.renderShapeInfos = [];
  const withoutShape = editor.makeParagraph(paragraph, 720, reviewContext);
  assert.equal(withoutShape.childNodes[0].childNodes.some((node) => node.className.includes('docx-shape')), false);
});

test('makeParagraph renders a math-only paragraph as MathML', () => {
  const editor = makeRunRenderEditor();
  editor.paragraphs = new Map();
  editor.measuring = false;
  editor.composing = false;
  editor.renderAfterComposition = false;
  editor.readText = (content) => content.textContent ?? '';
  editor.document = { getShapeParagraphs: () => [] };
  editor.renderShapeInfos = [];
  const paragraph = {
    index: 0,
    text: '',
    runs: [],
    images: [],
    math: [{
      runOffset: 0,
      display: 'block',
      linear: 'x+1',
      mathMl: { tag: 'math', attrs: { display: 'block' }, children: [{ tag: 'mi', text: 'x' }] },
    }],
  };
  const rendered = editor.makeParagraph(paragraph, 720, { deletedTextByRun: new Map(), revisionColors: new Map() });
  const math = rendered.childNodes[0].childNodes.find((node) => node.dataset?.docxMath === '1');
  assert.ok(math);
  assert.equal(math.tagName, 'MATH');
  assert.equal(math.attributes.get('display'), 'block');
  assert.equal(math.childNodes[0].tagName, 'MI');
});

test('pagination paragraph slices render each math exactly once', () => {
  const editor = makeRunRenderEditor();
  editor.paragraphs = new Map();
  editor.measuring = false;
  editor.composing = false;
  editor.renderAfterComposition = false;
  editor.readText = (content) => content.textContent ?? '';
  editor.document = { getShapeParagraphs: () => [] };
  editor.renderShapeInfos = [];
  const mathMl = { tag: 'math', children: [{ tag: 'mi', text: 'x' }] };
  const countMath = (paragraph) => {
    const element = editor.makeParagraph(paragraph, 720, { deletedTextByRun: new Map(), revisionColors: new Map() });
    const walk = (node, out = []) => {
      for (const child of node.childNodes ?? []) {
        if (child.tagName) out.push(child.tagName);
        walk(child, out);
      }
      return out;
    };
    return walk(element).filter((tag) => tag === 'MATH').length;
  };
  const renderedAcross = (paragraph, ranges) => ranges
    .map(([start, end]) => countMath(editor.sliceParagraph(paragraph, start, end)))
    .reduce((sum, count) => sum + count, 0);

  // 末尾公式：runOffset 等于 run 数，没有任何 run 的索引与它相同
  const trailing = {
    index: 0,
    text: '甲乙丙丁戊己',
    runs: [
      { index: 0, text: '甲乙', revisions: [] },
      { index: 1, text: '丙丁', revisions: [] },
      { index: 2, text: '戊己', revisions: [] },
    ],
    math: [{ runOffset: 3, display: 'inline', linear: 'x', mathMl }],
  };
  assert.equal(countMath(trailing), 1);
  assert.deepEqual([[0, 2], [2, 4], [4, 6]].map(([start, end]) =>
    editor.sliceParagraph(trailing, start, end).math?.length ?? 0), [0, 0, 1]);
  assert.equal(renderedAcross(trailing, [[0, 2], [2, 4], [4, 6]]), 1);

  // 中间公式：runOffset 同时是某个 run 的索引。该 run 单独成片时切片只有 1 个 run，
  // 若用 runs.length 判断末尾公式，这一片会渲染两次。
  const interior = {
    index: 0,
    text: '甲乙丙丁',
    runs: [
      { index: 0, text: '甲', revisions: [] },
      { index: 1, text: '乙', revisions: [] },
      { index: 2, text: '丙', revisions: [] },
      { index: 3, text: '丁', revisions: [] },
    ],
    math: [{ runOffset: 3, display: 'inline', linear: 'x', mathMl }],
  };
  assert.equal(countMath(interior), 1);
  assert.equal(countMath(editor.sliceParagraph(interior, 3, 4)), 1);
  assert.equal(renderedAcross(interior, [[0, 1], [1, 2], [2, 3], [3, 4]]), 1);
});

test('paginated page content renders row subsets in their assigned columns', () => {
  const editor = makeRunRenderEditor();
  editor.paragraphs = new Map();
  editor.measuring = false;
  editor.composing = false;
  editor.renderAfterComposition = false;
  editor.readText = (content) => content.textContent ?? '';
  editor.document = { getShapeParagraphs: () => [] };
  editor.renderShapeInfos = [];
  const paragraph = { index: 0, text: 'x', runs: [{ index: 0, text: 'x', images: [] }], images: [] };
  const makeRow = (header = false) => ({
    cells: [{ blocks: [], colSpan: 1, rowSpan: 1, isMergeContinuation: false }],
    format: header ? { header: true } : {},
  });
  const blocks = [
    { type: 'paragraph', paragraph },
    { type: 'table', rows: [makeRow(true), makeRow(), makeRow()], grid: [1000] },
  ];
  const page = {
    index: 0,
    number: 1,
    section: 0,
    contentHeightPx: 40,
    items: [
      { type: 'line', paragraph: 0, line: { heightPx: 10, startOffset: 0, endOffset: 1 }, column: 0 },
      { type: 'tableRow', table: 1, row: 0, heightPx: 10, column: 0 },
      { type: 'tableRow', table: 1, row: 2, heightPx: 10, column: 0 },
      { type: 'tableRow', table: 1, row: 0, heightPx: 10, column: 1 },
      { type: 'tableRow', table: 1, row: 1, heightPx: 10, column: 1 },
    ],
  };
  const section = {
    pageWidth: 1500,
    pageHeight: 1500,
    margins: { top: 0, right: 0, bottom: 0, left: 0 },
    columns: { count: 2, space: 0, equalWidth: true },
  };
  const reviewContext = { deletedTextByRun: new Map(), revisionColors: new Map() };
  const body = editor.makePageContent(page, section, blocks, [paragraph], 720, reviewContext);
  assert.equal(body.className, 'docx-page-content');
  assert.deepEqual(body.childNodes.map((column) => column.dataset.column), ['0', '1']);
  assert.deepEqual(body.childNodes[0].childNodes.map((node) => node.tagName), ['P', 'TABLE']);
  assert.deepEqual(body.childNodes[0].childNodes[1].rows.map((tr) => tr.childNodes[0].dataset.rowStart), ['0', '2']);
  assert.deepEqual(body.childNodes[1].childNodes[0].rows.map((tr) => tr.childNodes[0].dataset.rowStart), ['0', '1']);
  assert.equal(body.childNodes[0].childNodes[1].rows[0].dataset.header, 'true');
  assert.equal(body.childNodes[1].childNodes[0].rows[0].dataset.header, 'true');
});

test('makeTable renders only the requested row subset with row and column span datasets intact', () => {
  const editor = makeRunRenderEditor();
  editor.paragraphs = new Map();
  editor.measuring = false;
  editor.renderShapeInfos = [];
  const cell = (colSpan = 1, rowSpan = 1, isMergeContinuation = false) => ({ blocks: [], colSpan, rowSpan, isMergeContinuation });
  const block = {
    type: 'table',
    grid: [1000, 1000, 1000],
    rows: [
      { cells: [cell(2), cell(1, 2)], format: {} },
      { cells: [cell(), cell(), cell(1, 1, true)], format: {} },
      { cells: [cell(), cell(2, 2)], format: {} },
      { cells: [cell(), cell(2, 1, true)], format: {} },
    ],
  };
  const datasets = (table) => table.rows.map((tr) => tr.childNodes.map((td) => ({
    grid: `${td.dataset.gridStart}-${td.dataset.gridEnd}`,
    rows: `${td.dataset.rowStart}-${td.dataset.rowEnd}`,
    colSpan: td.colSpan,
    rowSpan: td.rowSpan,
  })));
  assert.deepEqual(datasets(editor.makeTable(block, [0, 2], 720, { deletedTextByRun: new Map(), revisionColors: new Map() })), [
    [
      { grid: '0-2', rows: '0-1', colSpan: 2, rowSpan: undefined },
      { grid: '2-3', rows: '0-2', colSpan: 1, rowSpan: 2 },
    ],
    [
      { grid: '0-1', rows: '2-3', colSpan: 1, rowSpan: undefined },
      { grid: '1-3', rows: '2-4', colSpan: 2, rowSpan: 2 },
    ],
  ]);
  assert.deepEqual(datasets(editor.makeTable(block, [1, 3], 720, { deletedTextByRun: new Map(), revisionColors: new Map() })), [
    [
      { grid: '0-1', rows: '1-2', colSpan: 1, rowSpan: undefined },
      { grid: '1-2', rows: '1-2', colSpan: 1, rowSpan: undefined },
    ],
    [{ grid: '0-1', rows: '3-4', colSpan: 1, rowSpan: undefined }],
  ]);
  assert.deepEqual(editor.makeTable(block, [7], 720, { deletedTextByRun: new Map(), revisionColors: new Map() }).rows, []);
});

test('makeParagraph keeps paragraphs with fields editable in markup view and read-only in preview views', () => {
  const run = (index, text, field) => ({ index, text, images: [], ...(field ? { field } : {}) });
  const paragraphs = {
    plain: { index: 0, text: 'plain', runs: [run(0, 'plain')], images: [] },
    field: {
      index: 1,
      text: 'Page 7 end',
      images: [],
      runs: [
        run(0, 'Page '),
        run(1, '', { index: 0, role: 'instruction', kind: 'PAGE', instruction: ' PAGE ' }),
        run(2, '7', { index: 0, role: 'result', kind: 'PAGE', instruction: ' PAGE ' }),
        run(3, ' end'),
      ],
    },
  };
  for (const revisionView of ['markup', 'final', 'original']) {
    const editor = makeRunRenderEditor({ revisionView });
    editor.paragraphs = new Map();
    editor.measuring = false;
    editor.composing = false;
    editor.renderAfterComposition = false;
    editor.readText = (content) => content.textContent ?? '';
    editor.document = { getShapeParagraphs: () => [] };
    editor.renderShapeInfos = [];
    for (const [name, paragraph] of Object.entries(paragraphs)) {
      const element = editor.makeParagraph(paragraph, 720, { deletedTextByRun: new Map(), revisionColors: new Map() });
      const content = element.childNodes.find((node) => node.className === 'docx-paragraph-content');
      assert.ok(content, `${revisionView}/${name}`);
      assert.equal(content.contentEditable, revisionView === 'markup' ? 'true' : 'false', `${revisionView}/${name}`);
      assert.equal(content.attributes.get('aria-readonly'), revisionView === 'markup' ? undefined : 'true', `${revisionView}/${name}`);
    }
  }
});

function makeViewModeRenderEditor(viewMode) {
  const editor = makeRunRenderEditor();
  const createElement = editor.root.ownerDocument.createElement;
  const output = { fragment: null, shapesDuringRender: null };
  Object.assign(editor, {
    viewMode,
    destroyed: false,
    measuring: false,
    composing: false,
    renderAfterComposition: false,
    paragraphs: new Map(),
    selected: null,
    selectedImageInfo: null,
    activeRevisionId: null,
    readText: (content) => content.textContent ?? '',
    flush: () => {},
    applyPageSetup: () => {},
    captureDocumentRange: () => null,
    updateRangeSelection: () => {},
    selectImage: () => {},
    setActiveRevision: () => {},
    makeHeaderFooter: (type) => {
      const area = createElement('div');
      area.className = `docx-${type}`;
      return area;
    },
    measureParagraphForPagination: (paragraph) => [{ heightPx: 18, startOffset: 0, endOffset: paragraph.text.length }],
  });
  const section = {
    index: 0,
    startParagraph: 0,
    endParagraph: 0,
    type: 'nextPage',
    pageWidth: 12240,
    pageHeight: 15840,
    orientation: 'portrait',
    margins: { top: 1440, right: 1800, bottom: 720, left: 1080, header: 0, footer: 0, gutter: 0 },
    columns: { count: 1, space: 0, equalWidth: true },
    titlePage: false,
    headers: {},
    footers: {},
  };
  const paragraph = { index: 0, text: 'shape anchor', runs: [{ index: 0, text: 'shape anchor', images: [] }], images: [] };
  editor.document = {
    mainDocumentPath: 'word/document.xml',
    getRevisions: () => [],
    getComments: () => [],
    getSettings: () => ({ defaultTabStop: 720 }),
    getShapes: () => [{ id: 'shape-1', paragraph: 0, run: 0, kind: 'textbox', placement: 'inline', hasTextContent: false }],
    getShapeParagraphs: () => [],
    getBlocks: () => [{ type: 'paragraph', paragraph }],
    getSections: () => [section],
    getSection: () => section,
    getImages: () => [],
  };
  editor.root.ownerDocument.createDocumentFragment = () => createElement('fragment');
  editor.root.replaceChildren = (fragment) => {
    output.fragment = fragment;
    output.shapesDuringRender = editor.renderShapeInfos.length;
  };
  return { editor, output };
}

function collectNodes(node, into = []) {
  into.push(node);
  for (const child of node.childNodes ?? []) collectNodes(child, into);
  return into;
}

test('shapes render in both continuous and paginated view modes', () => {
  for (const viewMode of ['continuous', 'paginated']) {
    const { editor, output } = makeViewModeRenderEditor(viewMode);
    editor.render();
    const nodes = collectNodes(output.fragment);
    const shapes = nodes.filter((node) => typeof node.className === 'string' && node.className.split(' ').includes('docx-shape'));
    assert.deepEqual(shapes.map((node) => node.dataset.docxShape), ['shape-1'], viewMode);
    assert.equal(output.shapesDuringRender, 1, viewMode);
    assert.deepEqual(editor.renderShapeInfos, [], viewMode);
    const pages = nodes.filter((node) => node.className === 'docx-page');
    if (viewMode === 'continuous') {
      assert.equal(pages.length, 0);
    } else {
      assert.equal(pages.length, 1);
      assert.equal(pages[0].style.width, '816px');
      assert.equal(pages[0].style.minHeight, '1056px');
      assert.equal(pages[0].style.padding, '96px 120px 48px 72px');
    }
  }
});

test('pagination table row measurement uses the tallest rendered cell', () => {
  const editor = makeRunRenderEditor();
  const table = {
    type: 'table',
    rows: [{
      cells: [
        { blocks: [], colSpan: 1, rowSpan: 1, isMergeContinuation: false },
        { blocks: [], colSpan: 1, rowSpan: 1, isMergeContinuation: false },
      ],
    }],
    grid: [1000, 1000],
  };
  assert.equal(editor.measureTableRowForPagination(table, 0, 100, { defaultTabStopTwips: 720 }), 27);
});

test('pagination table row measurement does not register temporary render state', () => {
  const editor = makeRunRenderEditor();
  editor.measuring = false;
  editor.paragraphs = new Map();
  editor.composing = false;
  editor.renderAfterComposition = false;
  editor.readText = (content) => content.textContent ?? '';
  editor.document = { getShapeParagraphs: () => [] };
  editor.renderShapeInfos = [];
  const paragraph = {
    index: 7,
    text: 'tracked',
    images: [],
    paragraphRevision: { id: 8, kind: 'insertion', author: 'Alice' },
    runs: [{ index: 0, text: 'tracked', images: [], revisions: [{ id: 9, kind: 'insertion', author: 'Alice' }] }],
  };
  const table = {
    type: 'table',
    rows: [{
      cells: [{ blocks: [{ type: 'paragraph', paragraph }], colSpan: 1, rowSpan: 1, isMergeContinuation: false }],
    }],
    grid: [1000],
  };
  assert.equal(editor.measureTableRowForPagination(table, 0, 100, { defaultTabStopTwips: 720 }), 27);
  assert.equal(editor.measuring, false);
  assert.equal(editor.paragraphs.size, 0);
  assert.equal(editor.revisionParagraphIds.size, 0);
  assert.equal(editor.revisionRunIds.size, 0);
});

test('withMeasuring restores nested state when temporary rendering throws', () => {
  const editor = makeRunRenderEditor();
  editor.measuring = false;
  assert.throws(() => editor.withMeasuring(() => {
    assert.equal(editor.measuring, true);
    throw new Error('render failed');
  }), /render failed/);
  assert.equal(editor.measuring, false);
  editor.measuring = true;
  assert.equal(editor.withMeasuring(() => editor.measuring), true);
  assert.equal(editor.measuring, true);
});

function appendRunToParagraph(editor, {
  paragraphIndex = 0,
  run,
  reviewContext = { deletedTextByRun: new Map(), revisionColors: new Map() },
} = {}) {
  const paragraphElement = editor.root.ownerDocument.createElement('span');
  const paragraph = { index: paragraphIndex };
  const offset = editor.appendRun(paragraphElement, paragraph, run, reviewContext, 720, 0);
  return { paragraphElement, runSpan: paragraphElement.childNodes[0], offset };
}

test('flush sanitizes disallowed control characters before commit', () => {
  const { editor, document } = makeFlushEditor({ text: 'a\u0001b' });
  assert.doesNotThrow(() => editor.flush());
  assert.equal(document.getParagraphs()[0].text, 'ab');
});

test('flush sanitizes isolated high surrogate before commit', () => {
  const { editor, document } = makeFlushEditor({ text: 'a\ud800b' });
  assert.doesNotThrow(() => editor.flush());
  assert.equal(document.getParagraphs()[0].text, 'ab');
});

test('flush sanitizes isolated low surrogate before commit', () => {
  const { editor, document } = makeFlushEditor({ text: 'a\udc00b' });
  assert.doesNotThrow(() => editor.flush());
  assert.equal(document.getParagraphs()[0].text, 'ab');
});

test('flush preserves tab and newline text', () => {
  const { editor, document } = makeFlushEditor({ text: 'a\tb\nc' });
  assert.doesNotThrow(() => editor.flush());
  assert.equal(document.getParagraphs()[0].text, 'a\tb\nc');
});

test('flush reads inner content text instead of outer paragraph wrapper text', () => {
  const { editor, document } = makeFlushEditor({ text: '用户文本', elementText: '1. 用户文本' });
  assert.doesNotThrow(() => editor.flush());
  assert.equal(document.getParagraphs()[0].text, '用户文本');
});

test('flush truncates overlong input, reports once, and commits legal text', () => {
  const calls = [];
  const overlong = `a${'x'.repeat(1_000_000)}`;
  const { editor, document } = makeFlushEditor({
    text: overlong,
    options: { onError: (error, context) => calls.push({ error, context }) },
  });
  assert.doesNotThrow(() => editor.flush());
  assert.equal(calls.length, 1);
  assert.equal(calls[0].context.paragraph, 0);
  assert.match(calls[0].error.message, /truncated at 1000000 characters/);
  assert.equal(document.getParagraphs()[0].text.length, 1_000_000);
  assert.equal(editor.paragraphs.get(0).text.length, 1_000_000);
});

test('flush does not repeatedly report unchanged overlong text after first truncation', () => {
  const calls = [];
  const overlong = `a${'x'.repeat(1_000_000)}`;
  const { editor } = makeFlushEditor({
    text: overlong,
    options: { onError: () => calls.push('error') },
  });
  editor.flush();
  editor.flush();
  assert.deepEqual(calls, ['error']);
});

test('flush falls back to console.error when onError is not provided', () => {
  const original = console.error;
  const calls = [];
  console.error = (...args) => calls.push(args);
  try {
    const overlong = `a${'x'.repeat(1_000_000)}`;
    const { editor } = makeFlushEditor({ text: overlong });
    assert.doesNotThrow(() => editor.flush());
    assert.equal(calls.length, 1);
  } finally {
    console.error = original;
  }
});

test('flush still logs original error when onError throws', () => {
  const original = console.error;
  const calls = [];
  console.error = (...args) => calls.push(args);
  try {
    const document = DocxDocument.create();
    document.setParagraphText = () => { throw new Error('commit failed'); };
    const { editor } = makeFlushEditor({
      text: 'changed',
      previous: '',
      document,
      options: { onError: () => { throw new Error('handler failed'); } },
    });
    assert.doesNotThrow(() => editor.flush());
    assert.equal(calls.length, 2);
    assert.match(String(calls[0][0]), /commit failed/);
    assert.match(String(calls[1][0]), /handler failed/);
  } finally {
    console.error = original;
  }
});

test('flush does not advance entry.text when commit fails', () => {
  const calls = [];
  const document = DocxDocument.create();
  document.setParagraphText = () => { throw new Error('commit failed'); };
  const { editor } = makeFlushEditor({
    text: 'new text',
    previous: 'old text',
    document,
    options: { onError: () => calls.push('error') },
  });
  assert.doesNotThrow(() => editor.flush());
  assert.equal(editor.paragraphs.get(0).text, 'old text');
  assert.equal(editor.paragraphs.get(0).failed, true);
  assert.deepEqual(calls, ['error']);
});

test('render completes even when flush commit throws', () => {
  const document = DocxDocument.create();
  document.setParagraphText = () => { throw new Error('commit failed'); };
  const { editor } = makeFlushEditor({ text: 'changed', previous: '', document, options: { onError: () => {} } });
  editor.composing = false;
  editor.renderAfterComposition = false;
  editor.selected = null;
  editor.root = {
    ownerDocument: {
      activeElement: null,
      getSelection: () => null,
      createDocumentFragment: () => ({ tag: 'fragment' }),
    },
    contains: () => false,
    replaceChildren: () => {},
  };
  editor.document.getBlocks = () => [];
  assert.doesNotThrow(() => editor.render());
});

test('setDocument and destroy remain usable after flush commit failures', () => {
  const document = DocxDocument.create();
  document.setParagraphText = () => { throw new Error('commit failed'); };
  const { editor } = makeFlushEditor({ text: 'changed', previous: '', document, options: { onError: () => {} } });
  let rendered = 0;
  editor.render = () => { rendered++; };
  const removed = [];
  let rootRemoved = false;
  editor.handleSelection = () => {};
  editor.handleRootKeydown = () => {};
  editor.root = {
    ownerDocument: { removeEventListener: (...args) => removed.push(args) },
    removeEventListener: (...args) => removed.push(args),
    remove: () => { rootRemoved = true; },
  };
  const next = DocxDocument.create();
  assert.doesNotThrow(() => editor.setDocument(next));
  assert.equal(editor.document, next);
  assert.equal(rendered, 1);
  assert.doesNotThrow(() => editor.destroy());
  assert.equal(editor.destroyed, true);
  assert.equal(editor.paragraphs.size, 0);
  assert.equal(rootRemoved, true);
  assert.equal(removed.length, 2);
  assert.equal(removed[0][0], 'selectionchange');
  assert.equal(removed[1][0], 'keydown');
});

test('sanitizeTextWithInfo preserves surrogate pairs when truncating at max length', () => {
  const input = `${'a'.repeat(999_999)}😀`;
  const sanitized = sanitizeTextWithInfo(input);
  assert.equal(sanitized.text.length, 999_999);
  assert.equal(sanitized.text, 'a'.repeat(999_999));
  assert.equal(sanitized.truncated, true);
  assert.equal(sanitized.truncatedAt, 999_999);
});

test('insertText sanitizes invalid paste-like input before insertion', () => {
  const editor = Object.create(DocxEditor.prototype);
  let inserted = '';
  const range = {
    startContainer: {},
    endContainer: {},
    deleteContents: () => {},
    insertNode: (node) => { inserted = node.textContent; },
    setStartAfter: () => {},
    collapse: () => {},
  };
  const selection = {
    rangeCount: 1,
    getRangeAt: () => range,
    removeAllRanges: () => {},
    addRange: () => {},
  };
  editor.root = {
    ownerDocument: {
      getSelection: () => selection,
      createTextNode: (text) => ({ textContent: text }),
    },
  };
  const element = { contains: () => true };
  assert.doesNotThrow(() => editor.insertText(element, 'a\u0001b\ud800c\r\nd'));
  assert.equal(inserted, 'abc\nd');
});

test('selectedTableCell getter and event update when selection enters a table cell', () => {
  const cell = { table: 0, row: 1, col: 2, rowSpan: 1, colSpan: 3, nested: false };
  const { editor, events } = makeSelectionEditor({ 4: cell });
  editor.selectParagraph(4);
  assert.equal(editor.selectedParagraph, 4);
  assert.deepEqual(editor.selectedTableCell, cell);
  assert.deepEqual(events.map((event) => [event.type, event.detail]), [
    ['docx-selectionchange', { index: 4 }],
    ['docx-tablecellchange', { cell }],
  ]);
});

test('table cell change dispatch is debounced within the same table cell while exposed state refreshes', () => {
  const first = { table: 0, row: 0, col: 0, rowSpan: 1, colSpan: 1, nested: false };
  const second = { table: 0, row: 0, col: 0, rowSpan: 2, colSpan: 1, nested: false };
  const { editor, events } = makeSelectionEditor({ 1: first, 2: second });
  editor.selectParagraph(1);
  editor.selectParagraph(2);
  assert.deepEqual(editor.selectedTableCell, second);
  assert.deepEqual(events.map((event) => event.type), [
    'docx-selectionchange',
    'docx-tablecellchange',
    'docx-selectionchange',
  ]);
});

test('selectedTableCell refreshes span changes for the same paragraph without redispatching table cell change', () => {
  const first = { table: 0, row: 0, col: 0, rowSpan: 1, colSpan: 1, nested: false };
  const second = { table: 0, row: 0, col: 0, rowSpan: 1, colSpan: 2, nested: false };
  let current = first;
  const { editor, events } = makeSelectionEditor();
  editor.document = { getTableCellAt: () => current };
  editor.selectParagraph(1);
  current = second;
  editor.selectParagraph(1);
  assert.equal(editor.selectedTableCell.colSpan, 2);
  assert.deepEqual(events.map((event) => event.type), [
    'docx-selectionchange',
    'docx-tablecellchange',
  ]);
});

test('table cell change dispatch fires when selection crosses table cells and when it leaves tables', () => {
  const a = { table: 0, row: 0, col: 0, rowSpan: 1, colSpan: 1, nested: false };
  const b = { table: 0, row: 0, col: 1, rowSpan: 1, colSpan: 1, nested: false };
  const { editor, events } = makeSelectionEditor({ 0: a, 1: b, 2: null });
  editor.selectParagraph(0);
  editor.selectParagraph(1);
  editor.selectParagraph(2);
  assert.equal(editor.selectedTableCell, null);
  assert.deepEqual(events.map((event) => [event.type, event.detail]), [
    ['docx-selectionchange', { index: 0 }],
    ['docx-tablecellchange', { cell: a }],
    ['docx-selectionchange', { index: 1 }],
    ['docx-tablecellchange', { cell: b }],
    ['docx-selectionchange', { index: 2 }],
    ['docx-tablecellchange', null],
  ]);
});

test('readText skips non-editable decorations unless explicitly marked as document content', () => {
  const editor = Object.create(DocxEditor.prototype);
  const text = (value) => ({ nodeType: 3, textContent: value });
  const element = (tagName, { dataset = {}, contentEditable = 'inherit' } = {}, childNodes = []) => ({
    nodeType: 1,
    tagName,
    dataset,
    contentEditable,
    childNodes,
  });
  const root = element('SPAN', {}, [
    element('SPAN', { contentEditable: 'false' }, [text('drop')]),
    element('SPAN', { contentEditable: 'false', dataset: { docxContent: '1' } }, [text('keep')]),
    element('SPAN', { contentEditable: 'false', dataset: { docxMark: '1' } }, [text('drop')]),
  ]);
  assert.equal(editor.readText(root), 'keep');
});

test('applyHistory flushes first and re-renders after undo/redo', () => {
  const calls = [];
  const editor = Object.create(DocxEditor.prototype);
  editor.destroyed = false;
  editor.selected = 0;
  editor.paragraphs = new Map([[0, { content: { focus: () => {} } }]]);
  editor.focusContent = () => {};
  editor.flush = () => calls.push('flush');
  editor.render = () => calls.push('render');
  editor.options = { onChange: () => calls.push('change') };
  editor.document = {
    revision: 1,
    undo: () => ({ revision: 2, paragraphs: [{ index: 0 }] }),
    redo: () => ({ revision: 3, paragraphs: [{ index: 0 }] }),
  };
  editor.applyHistory('undo');
  editor.applyHistory('redo');
  assert.deepEqual(calls, ['flush', 'render', 'change', 'flush', 'render', 'change']);
});

test('applyHistory is a no-op when revision does not change', () => {
  const calls = [];
  const editor = Object.create(DocxEditor.prototype);
  editor.destroyed = false;
  editor.selected = null;
  editor.paragraphs = new Map();
  editor.flush = () => calls.push('flush');
  editor.render = () => calls.push('render');
  editor.options = { onChange: () => calls.push('change') };
  editor.document = {
    revision: 2,
    undo: () => ({ revision: 2, paragraphs: [] }),
    redo: () => ({ revision: 2, paragraphs: [] }),
  };
  editor.applyHistory('undo');
  assert.deepEqual(calls, ['flush']);
});

test('handleHistoryShortcut maps Ctrl/Cmd+Z to undo', () => {
  const editor = Object.create(DocxEditor.prototype);
  const calls = [];
  editor.composing = false;
  editor.applyHistory = (direction) => calls.push(direction);
  const content = {};
  editor.root = { contains: (value) => value === content };
  const event = {
    isComposing: false,
    ctrlKey: true,
    metaKey: false,
    altKey: false,
    shiftKey: false,
    key: 'z',
    preventDefault: () => calls.push('prevent'),
    target: { nodeType: 1, closest: () => content },
  };
  assert.equal(editor.handleHistoryShortcut(event), true);
  assert.deepEqual(calls, ['prevent', 'undo']);
});

test('handleHistoryShortcut maps Cmd+Shift+Z and Ctrl+Y to redo', () => {
  const editor = Object.create(DocxEditor.prototype);
  editor.composing = false;
  const calls = [];
  editor.applyHistory = (direction) => calls.push(direction);
  const content = {};
  editor.root = { contains: (value) => value === content };
  const metaShiftZ = {
    isComposing: false,
    ctrlKey: false,
    metaKey: true,
    altKey: false,
    shiftKey: true,
    key: 'z',
    preventDefault: () => calls.push('prevent:z'),
    target: { nodeType: 1, closest: () => content },
  };
  const ctrlY = {
    isComposing: false,
    ctrlKey: true,
    metaKey: false,
    altKey: false,
    shiftKey: false,
    key: 'y',
    preventDefault: () => calls.push('prevent:y'),
    target: { nodeType: 1, closest: () => content },
  };
  assert.equal(editor.handleHistoryShortcut(metaShiftZ), true);
  assert.equal(editor.handleHistoryShortcut(ctrlY), true);
  assert.deepEqual(calls, ['prevent:z', 'redo', 'prevent:y', 'redo']);
});

test('parseClipboardFragment validates payload shape', () => {
  const editor = Object.create(DocxEditor.prototype);
  const valid = editor.parseClipboardFragment('{"version":1,"text":"x","paragraphs":[{"runs":[{"text":"x"}]}]}');
  const invalid = editor.parseClipboardFragment('{"version":2}');
  assert.equal(valid.version, 1);
  assert.equal(invalid, null);
});

test('writeClipboardFragment writes custom mime, plain html text, and plain text', () => {
  const editor = Object.create(DocxEditor.prototype);
  const written = new Map();
  const transfer = { setData: (type, value) => written.set(type, value) };
  const ok = editor.writeClipboardFragment(transfer, { version: 1, text: 'hello', paragraphs: [{ runs: [{ text: 'hello' }] }] });
  assert.equal(ok, true);
  assert.equal(written.get('application/x-docx-view+json').includes('"version":1'), true);
  assert.equal(written.get('text/plain'), 'hello');
  assert.equal(written.get('text/html'), 'hello');
});

test('mapExternalHtmlFragment strips unsafe href and external images', () => {
  const editor = Object.create(DocxEditor.prototype);
  const text = (value) => ({ nodeType: 3, textContent: value, childNodes: [] });
  const element = (tagName, attrs = {}, childNodes = []) => ({
    nodeType: 1,
    tagName,
    childNodes,
    children: childNodes.filter((child) => child.nodeType === 1),
    getAttribute: (name) => attrs[name] ?? null,
    querySelectorAll: () => [],
  });
  editor.root = {
    ownerDocument: {
      defaultView: {
        DOMParser: class {
          parseFromString() {
            return {
              body: {
                childNodes: [
                  element('DIV', {}, [
                    element('A', { href: 'javascript:alert(1)' }, [text('bad link')]),
                    element('IMG', { src: 'https://example.com/x.png', alt: 'ext' }, []),
                  ]),
                ],
                textContent: 'bad link',
              },
            };
          }
        },
      },
    },
  };
  const fragment = editor.mapExternalHtmlFragment('<div/>', 'fallback');
  assert.equal(fragment.paragraphs[0].runs.some((run) => run.hyperlink?.url), false);
  assert.equal(fragment.paragraphs[0].runs.some((run) => (run.images?.length ?? 0) > 0), false);
});

test('mapExternalHtmlFragment keeps data-image URIs as images', () => {
  const editor = Object.create(DocxEditor.prototype);
  const element = (tagName, attrs = {}, childNodes = []) => ({
    nodeType: 1,
    tagName,
    childNodes,
    children: childNodes.filter((child) => child.nodeType === 1),
    getAttribute: (name) => attrs[name] ?? null,
    querySelectorAll: () => [],
  });
  editor.root = {
    ownerDocument: {
      defaultView: {
        DOMParser: class {
          parseFromString() {
            return {
              body: {
                childNodes: [
                  element('DIV', {}, [
                    element('IMG', { src: 'data:image/png;base64,AAAA', alt: 'ok' }, []),
                  ]),
                ],
                textContent: '',
              },
            };
          }
        },
      },
    },
  };
  const fragment = editor.mapExternalHtmlFragment('<div/>', '');
  assert.equal(fragment.paragraphs[0].runs.some((run) => (run.images?.length ?? 0) === 1), true);
});

test('mapExternalHtmlFragment does not inject list marker text into document runs', () => {
  const editor = Object.create(DocxEditor.prototype);
  const text = (value) => ({ nodeType: 3, textContent: value, childNodes: [] });
  const element = (tagName, attrs = {}, childNodes = []) => ({
    nodeType: 1,
    tagName,
    childNodes,
    children: childNodes.filter((child) => child.nodeType === 1),
    getAttribute: (name) => attrs[name] ?? null,
    querySelectorAll: () => [],
  });
  editor.root = {
    ownerDocument: {
      defaultView: {
        DOMParser: class {
          parseFromString() {
            const li = element('LI', {}, [text('item')]);
            return {
              body: {
                childNodes: [element('UL', {}, [li])],
                textContent: 'item',
              },
            };
          }
        },
      },
    },
  };
  const fragment = editor.mapExternalHtmlFragment('<ul><li>item</li></ul>', 'item');
  assert.equal(fragment.paragraphs[0].runs.map((run) => run.text ?? '').join(''), 'item');
  assert.deepEqual(fragment.paragraphs[0].numbering, { kind: 'bullet', level: 0, listId: 1 });
});

test('mapExternalHtmlFragment maps html tables into table blocks', () => {
  const editor = Object.create(DocxEditor.prototype);
  const text = (value) => ({ nodeType: 3, textContent: value, childNodes: [] });
  const element = (tagName, attrs = {}, childNodes = []) => ({
    nodeType: 1,
    tagName,
    childNodes,
    children: childNodes.filter((child) => child.nodeType === 1),
    getAttribute: (name) => attrs[name] ?? null,
    querySelectorAll: () => [],
  });
  editor.root = {
    ownerDocument: {
      defaultView: {
        DOMParser: class {
          parseFromString() {
            return {
              body: {
                childNodes: [element('TABLE', {}, [
                  element('TR', {}, [
                    element('TD', {}, [text('A')]),
                    element('TD', {}, [text('B')]),
                  ]),
                ])],
                textContent: 'A B',
              },
            };
          }
        },
      },
    },
  };
  const fragment = editor.mapExternalHtmlFragment('<table><tr><td>A</td><td>B</td></tr></table>', '');
  assert.equal(fragment.blocks[0].type, 'table');
  assert.equal(fragment.blocks[0].table.rows[0][0].runs[0].text, 'A');
  assert.equal(fragment.blocks[0].table.rows[0][1].runs[0].text, 'B');
});

test('handleClipboardPaste prefers internal rich fragment and re-renders once', () => {
  const editor = Object.create(DocxEditor.prototype);
  const calls = [];
  editor.root = { ownerDocument: {}, contains: () => true };
  editor.captureDocumentRange = () => ({ start: { paragraph: 0, offset: 0 }, end: { paragraph: 0, offset: 0 } });
  editor.documentRange = (range) => range;
  editor.document = {
    revision: 1,
    pasteClipboardFragment: function () { this.revision = 2; return true; },
    getSnapshot: () => ({ revision: 2 }),
  };
  editor.parseClipboardFragment = () => ({ version: 1, text: 'x', paragraphs: [{ runs: [{ text: 'x' }] }] });
  editor.render = () => calls.push('render');
  editor.options = { onChange: () => calls.push('change') };
  const event = {
    clipboardData: { getData: () => '' },
    preventDefault: () => calls.push('prevent'),
  };
  editor.handleClipboardPaste(event, {});
  assert.deepEqual(calls, ['prevent', 'render', 'change']);
});

test('handleClipboardPaste does not trust HTML marker without custom mime payload', () => {
  const editor = Object.create(DocxEditor.prototype);
  const calls = [];
  editor.root = { ownerDocument: {}, contains: () => true };
  editor.captureDocumentRange = () => ({ start: { paragraph: 0, offset: 0 }, end: { paragraph: 0, offset: 0 } });
  editor.documentRange = (range) => range;
  editor.document = { revision: 1, pasteClipboardFragment: () => false, getSnapshot: () => ({ revision: 1 }) };
  editor.insertText = (_content, text) => calls.push(`plain:${text}`);
  editor.parseClipboardFragment = () => null;
  editor.mapExternalHtmlFragment = (html) => {
    calls.push(`html:${html.includes('data-docx-clip="1"')}`);
    return { version: 1, text: 'safe', paragraphs: [{ runs: [{ text: 'safe' }] }] };
  };
  const event = {
    clipboardData: {
      getData: (type) => type === 'text/plain' ? 'plain' : type === 'text/html' ? '<div data-docx-clip="1" data-docx-payload="..."></div>' : '',
    },
    preventDefault: () => calls.push('prevent'),
  };
  editor.handleClipboardPaste(event, {});
  assert.deepEqual(calls, ['html:true', 'prevent', 'plain:plain']);
});

test('handleClipboardCut wraps deletion in one history group', () => {
  const editor = Object.create(DocxEditor.prototype);
  const calls = [];
  editor.root = { ownerDocument: {}, contains: () => true };
  editor.captureDocumentRange = () => ({ start: { paragraph: 0, offset: 1 }, end: { paragraph: 0, offset: 2 } });
  editor.documentRange = (range) => range;
  editor.document = {
    revision: 3,
    copyClipboardFragment: () => ({ version: 1, text: 'x', paragraphs: [{ runs: [{ text: 'x' }] }] }),
    beginHistoryGroup: (label) => calls.push(`begin:${label}`),
    pasteClipboardFragment: function () { this.revision = 4; return true; },
    endHistoryGroup: () => calls.push('end'),
    getSnapshot: () => ({ revision: 4 }),
  };
  editor.writeClipboardFragment = () => true;
  editor.render = () => calls.push('render');
  editor.options = { onChange: () => calls.push('change') };
  editor.handleClipboardCut({ clipboardData: {}, preventDefault: () => calls.push('prevent') }, {});
  assert.deepEqual(calls, ['begin:cut', 'end', 'prevent', 'render', 'change']);
});

test('handleClipboardCut prevents default for unsupported multi-paragraph ranges', () => {
  const editor = Object.create(DocxEditor.prototype);
  let prevented = false;
  editor.captureDocumentRange = () => ({ start: { paragraph: 0, offset: 0 }, end: { paragraph: 1, offset: 0 } });
  editor.documentRange = (range) => range;
  editor.handleClipboardCut({ preventDefault: () => { prevented = true; } }, {});
  assert.equal(prevented, true);
});

test('setReviewFilter updates render state without triggering onChange', () => {
  const editor = Object.create(DocxEditor.prototype);
  let renderCalls = 0;
  let changes = 0;
  editor.destroyed = false;
  editor.reviewFilter = { showRevisions: true, showComments: true, revisionView: 'markup' };
  editor.options = { onChange: () => { changes++; } };
  editor.flush = () => {};
  editor.render = () => { renderCalls++; };
  editor.setReviewFilter({ authors: [{ kind: 'named', author: 'Alice' }], revisionView: 'final', showComments: false });
  assert.equal(renderCalls, 1);
  assert.equal(changes, 0);
  assert.deepEqual(editor.reviewFilter, {
    authors: [{ kind: 'named', author: 'Alice' }],
    showRevisions: true,
    showComments: false,
    revisionView: 'final',
  });
});

test('setReviewFilter is a no-op when filter is unchanged', () => {
  const editor = Object.create(DocxEditor.prototype);
  let renders = 0;
  editor.destroyed = false;
  editor.reviewFilter = { authors: [{ kind: 'named', author: 'Alice' }], showRevisions: true, showComments: true, revisionView: 'markup' };
  editor.render = () => { renders++; };
  editor.setReviewFilter({ authors: [{ kind: 'named', author: 'Alice' }], showRevisions: true, showComments: true, revisionView: 'markup' });
  assert.equal(renders, 0);
});

test('setReviewFilter merges partial updates without clearing existing fields', () => {
  const editor = Object.create(DocxEditor.prototype);
  editor.destroyed = false;
  editor.reviewFilter = { authors: [{ kind: 'named', author: 'Alice' }], showRevisions: true, showComments: true, revisionView: 'markup' };
  editor.flush = () => {};
  editor.render = () => {};
  editor.setReviewFilter({ showComments: false });
  assert.deepEqual(editor.reviewFilter, {
    authors: [{ kind: 'named', author: 'Alice' }],
    showRevisions: true,
    showComments: false,
    revisionView: 'markup',
  });
});

test('setReviewFilter validates authors as bounded text list', () => {
  const editor = Object.create(DocxEditor.prototype);
  editor.destroyed = false;
  editor.reviewFilter = { showRevisions: true, showComments: true, revisionView: 'markup' };
  editor.flush = () => {};
  editor.render = () => {};
  assert.throws(() => editor.setReviewFilter({ authors: 'Alice' }), /authors must be an array/);
  assert.throws(() => editor.setReviewFilter({ authors: [{ kind: 'named', author: 'A\u0000' }] }), /reviewFilter\.authors\[\]\.author/);
  assert.throws(() => editor.setReviewFilter({ authors: Array.from({ length: 1001 }, (_, index) => ({ kind: 'named', author: String(index) })) }), /at most 1000 items/);
});

test('setReviewFilter validates author kind and raw author consistency', () => {
  const editor = Object.create(DocxEditor.prototype);
  editor.destroyed = false;
  editor.reviewFilter = { showRevisions: true, showComments: true, revisionView: 'markup' };
  editor.flush = () => {};
  editor.render = () => {};
  assert.throws(() => editor.setReviewFilter({ authors: [{ kind: 'unattributed', author: 'Alice' }] }), /must be omitted for unattributed kind/);
  assert.throws(() => editor.setReviewFilter({ authors: [{ kind: 'named' }] }), /is required for named kind/);
  assert.throws(() => editor.setReviewFilter({ authors: [{ kind: 'empty', author: '  ' }] }), /must be an empty string for empty kind/);
  assert.throws(() => editor.setReviewFilter({ authors: [{ kind: 'blank', author: 'Alice' }] }), /must be whitespace-only for blank kind/);
});

test('setReviewFilter accepts getReviewers() author buckets without reshaping', () => {
  const doc = DocxDocument.create();
  doc.setPartXml(doc.mainDocumentPath, `<w:document xmlns:w="${WORD_NS}"><w:body><w:p><w:ins w:id="1"><w:r><w:t>u</w:t></w:r></w:ins><w:ins w:id="2" w:author=""><w:r><w:t>e</w:t></w:r></w:ins><w:ins w:id="3" w:author="   "><w:r><w:t>b</w:t></w:r></w:ins><w:ins w:id="4" w:author="Alice"><w:r><w:t>n</w:t></w:r></w:ins></w:p><w:sectPr/></w:body></w:document>`);
  const authors = doc.getReviewers().map((reviewer) =>
    (reviewer.author === undefined ? { kind: reviewer.kind } : { kind: reviewer.kind, author: reviewer.author }));
  const editor = Object.create(DocxEditor.prototype);
  editor.destroyed = false;
  editor.reviewFilter = { showRevisions: true, showComments: true, revisionView: 'markup' };
  editor.flush = () => {};
  editor.render = () => {};
  assert.doesNotThrow(() => editor.setReviewFilter({ authors }));
  assert.deepEqual(new Set(editor.reviewFilter.authors.map((item) => item.kind)), new Set(['unattributed', 'empty', 'blank', 'named']));
});

test('setReviewFilter validates show flags and revisionView enum', () => {
  const editor = Object.create(DocxEditor.prototype);
  editor.destroyed = false;
  editor.reviewFilter = { showRevisions: true, showComments: true, revisionView: 'markup' };
  editor.flush = () => {};
  editor.render = () => {};
  assert.throws(() => editor.setReviewFilter({ showComments: 'yes' }), /showComments must be boolean/);
  assert.throws(() => editor.setReviewFilter({ showRevisions: 1 }), /showRevisions must be boolean/);
  assert.throws(() => editor.setReviewFilter({ revisionView: 'other' }), /revisionView must be one of/);
});

test('setReviewFilter flushes pending edits across all filter dimensions', () => {
  const cases = [
    { name: 'showComments', filter: { showComments: false } },
    { name: 'showRevisions', filter: { showRevisions: false } },
    { name: 'authors', filter: { authors: [{ kind: 'named', author: 'Alice' }] } },
    { name: 'revisionView original', filter: { revisionView: 'original' } },
    { name: 'revisionView final', filter: { revisionView: 'final' } },
  ];
  for (const sample of cases) {
    const doc = DocxDocument.create();
    const { editor } = makeFlushEditor({
      text: `pending-${sample.name}`,
      previous: '',
      document: doc,
    });
    editor.reviewFilter = { showRevisions: true, showComments: true, revisionView: 'markup' };
    editor.options = {};
    editor.root = { ownerDocument: {} };
    editor.render = function () { this.flush(); };
    const before = doc.revision;
    editor.setReviewFilter(sample.filter);
    assert.equal(doc.getParagraphs()[0].text, `pending-${sample.name}`);
    assert.equal(doc.revision, before + 1);
  }
});

test('flush is a no-op in original review view to avoid projected-text writeback', () => {
  const doc = DocxDocument.create();
  doc.setTrackChanges(true);
  doc.setRevisionAuthor('Alice');
  doc.setParagraphText(0, 'ABXDEF');
  const beforeRevision = doc.revision;
  const beforeRevisions = doc.getRevisions().map((revision) => ({ ...revision }));
  const { editor } = makeFlushEditor({
    text: 'ABCDEFZ',
    previous: 'ABCDEF',
    document: doc,
  });
  editor.reviewFilter = { showRevisions: true, showComments: true, revisionView: 'original' };
  editor.options = {};
  editor.flush();
  assert.equal(doc.revision, beforeRevision);
  assert.deepEqual(doc.getRevisions(), beforeRevisions);
  assert.equal(doc.getParagraphs()[0].text, 'ABXDEF');
});

test('switching from non-markup back to markup does not flush preview DOM text', () => {
  const doc = DocxDocument.create();
  const { editor } = makeFlushEditor({
    text: '',
    previous: '',
    document: doc,
  });
  editor.reviewFilter = { showRevisions: true, showComments: true, revisionView: 'original' };
  editor.options = {};
  editor.root = { ownerDocument: {} };
  editor.render = function () { this.flush(); };
  const before = doc.revision;
  editor.setReviewFilter({ revisionView: 'markup' });
  assert.equal(doc.getParagraphs()[0].text, '');
  assert.equal(doc.revision, before);
});

test('toggling reviewFilter fields does not mutate document revision, text, or XML bytes', () => {
  const doc = DocxDocument.create();
  doc.setTrackChanges(true);
  doc.setRevisionAuthor('Alice');
  doc.setParagraphText(0, 'Alpha');
  const beforeRevision = doc.revision;
  const beforeText = doc.getParagraphs().map((paragraph) => paragraph.text);
  const beforeXml = doc.listParts().map((path) => [path, doc.getPartXml(path)]);
  const editor = Object.create(DocxEditor.prototype);
  editor.destroyed = false;
  editor.reviewFilter = { showRevisions: true, showComments: true, revisionView: 'markup' };
  editor.flush = () => {};
  editor.render = () => {};
  let changes = 0;
  editor.options = { onChange: () => { changes++; } };
  editor.setReviewFilter({ showComments: false });
  editor.setReviewFilter({ showComments: true, revisionView: 'original' });
  editor.setReviewFilter({ showRevisions: false, revisionView: 'final' });
  editor.setReviewFilter({ revisionView: 'markup' });
  assert.equal(doc.revision, beforeRevision);
  assert.deepEqual(doc.getParagraphs().map((paragraph) => paragraph.text), beforeText);
  assert.deepEqual(doc.listParts().map((path) => [path, doc.getPartXml(path)]), beforeXml);
  assert.equal(changes, 0);
});

test('revision acceptance mutates the document in final and original preview views', async (t) => {
  for (const revisionView of ['final', 'original']) {
    await t.test(revisionView, () => {
      const doc = DocxDocument.create();
      doc.setTrackChanges(true);
      doc.setParagraphText(0, 'Accepted text');
      const [revision] = doc.getRevisions();
      assert.ok(revision);

      const editor = Object.create(DocxEditor.prototype);
      let changes = 0;
      editor.reviewFilter = { showRevisions: true, showComments: true, revisionView };
      editor.document = doc;
      editor.flush = () => {};
      editor.render = () => {};
      editor.options = { onChange: () => { changes++; } };
      editor.setActiveRevision = () => {};

      assert.equal(editor.acceptRevision(revision.id), true);
      assert.equal(doc.getRevisions().length, 0);
      assert.equal(doc.getParagraphs()[0].text, 'Accepted text');
      assert.equal(changes, 1);
    });
  }
});

test('reviewFilter author narrowing does not change unfiltered getRevisions output', () => {
  const doc = DocxDocument.create();
  doc.setTrackChanges(true);
  doc.setRevisionAuthor('Alice');
  doc.setParagraphText(0, 'Alice text');
  doc.setRevisionAuthor('Bob');
  doc.setParagraphText(0, 'Bob text');
  const before = doc.getRevisions().map((item) => item.id);
  const editor = Object.create(DocxEditor.prototype);
  editor.destroyed = false;
  editor.reviewFilter = { showRevisions: true, showComments: true, revisionView: 'markup' };
  editor.flush = () => {};
  editor.render = () => {};
  editor.setReviewFilter({ authors: [{ kind: 'named', author: 'Alice' }] });
  assert.deepEqual(doc.getRevisions().map((item) => item.id), before);
});

test('reviewScopedRun applies revisionView final/original semantics without mutating source run', () => {
  const editor = Object.create(DocxEditor.prototype);
  editor.reviewFilter = { showRevisions: true, showComments: true, revisionView: 'original' };
  const context = { deletedTextByRun: new Map([['0:0', 'deleted text']]) };
  const deletedRun = { index: 0, text: '', revisions: [{ id: 1, kind: 'deletion' }] };
  const insertionRun = { index: 1, text: 'inserted', revisions: [{ id: 2, kind: 'insertion' }] };
  const moveFromRun = { index: 2, text: '', revisions: [{ id: 3, kind: 'move', move: { name: 'm', side: 'from', pairedId: 4 } }] };
  const moveToRun = { index: 3, text: 'moved', revisions: [{ id: 4, kind: 'move', move: { name: 'm', side: 'to', pairedId: 3 } }] };
  const originalDeleted = editor.reviewScopedRun(0, deletedRun, context);
  const originalInserted = editor.reviewScopedRun(0, insertionRun, context);
  const originalMoveFrom = editor.reviewScopedRun(0, moveFromRun, { deletedTextByRun: new Map([['0:2', 'moved']]) });
  const originalMoveTo = editor.reviewScopedRun(0, moveToRun, context);
  assert.equal(originalDeleted.text, 'deleted text');
  assert.equal(originalInserted.text, '');
  assert.equal(originalMoveFrom.text, 'moved');
  assert.equal(originalMoveTo.text, '');
  assert.equal(deletedRun.text, '');
  assert.equal(insertionRun.text, 'inserted');
  editor.reviewFilter = { showRevisions: true, showComments: true, revisionView: 'final' };
  const finalDeleted = editor.reviewScopedRun(0, { index: 0, text: 'old', revisions: [{ id: 3, kind: 'deletion' }] }, context);
  const finalMoveFrom = editor.reviewScopedRun(0, moveFromRun, context);
  const finalMoveTo = editor.reviewScopedRun(0, moveToRun, context);
  assert.equal(finalDeleted.text, '');
  assert.equal(finalMoveFrom.text, '');
  assert.equal(finalMoveTo.text, 'moved');
});

test('reviewScopedRun author filtering does not collide named and unattributed buckets', () => {
  const editor = Object.create(DocxEditor.prototype);
  editor.reviewFilter = {
    authors: [{ kind: 'named', author: '(unattributed)' }],
    showRevisions: true,
    showComments: true,
    revisionView: 'markup',
  };
  const context = { authors: new Set(['named:(unattributed)']), deletedTextByRun: new Map() };
  const run = {
    index: 0,
    text: 'text',
    revisions: [
      { id: 1, kind: 'insertion', author: '(unattributed)' },
      { id: 2, kind: 'insertion' },
    ],
  };
  const filtered = editor.reviewScopedRun(0, run, context);
  assert.deepEqual(filtered.revisions?.map((revision) => revision.id), [1]);
});

test('revisionAriaDescription includes move label', () => {
  const editor = Object.create(DocxEditor.prototype);
  assert.equal(editor.revisionAriaDescription([{ id: 1, kind: 'move', author: 'Alice', move: { name: 'm', side: 'to' } }]).includes('移动'), true);
});

test('readText skips deleted-text visualization nodes', () => {
  const editor = Object.create(DocxEditor.prototype);
  const text = (value) => ({ nodeType: 3, textContent: value });
  const element = (tagName, { dataset = {}, contentEditable = 'inherit' } = {}, childNodes = []) => ({
    nodeType: 1,
    tagName,
    dataset,
    contentEditable,
    childNodes,
  });
  const root = element('SPAN', {}, [
    text('AB'),
    element('SPAN', { dataset: { docxDeleted: '1' }, contentEditable: 'false' }, [text('C')]),
    text('XDEF'),
  ]);
  assert.equal(editor.readText(root), 'ABXDEF');
});

test('readText skips hidden-run preservation nodes', () => {
  const editor = Object.create(DocxEditor.prototype);
  const content = {
    nodeType: 1, tagName: 'SPAN', dataset: {}, childNodes: [
      { nodeType: 3, textContent: 'visible ' },
      {
        nodeType: 1, tagName: 'SPAN', dataset: { docxHidden: '1', docxHiddenPreserved: '1' }, contentEditable: 'false',
        childNodes: [{ nodeType: 3, textContent: 'hidden' }],
      },
      { nodeType: 3, textContent: ' text' },
    ],
  };
  assert.equal(editor.readText(content), 'visible  text');
  assert.deepEqual(editor.readTextSegments(content), ['visible ', ' text']);
});

test('focusRevision validates revision id input', () => {
  const editor = Object.create(DocxEditor.prototype);
  editor.filteredRevisionIds = () => [1];
  editor.setActiveRevision = () => {};
  editor.revisionRunIds = new Map();
  editor.revisionParagraphIds = new Map();
  assert.throws(() => editor.focusRevision(-1), /non-negative integer/);
});

test('focusRevision returns false for filtered-out revision', () => {
  const editor = Object.create(DocxEditor.prototype);
  editor.filteredRevisionIds = () => [2];
  editor.setActiveRevision = () => {};
  editor.revisionRunIds = new Map();
  editor.revisionParagraphIds = new Map();
  assert.equal(editor.focusRevision(1), false);
});

test('focusRevision activates and focuses mapped revision node', () => {
  const editor = Object.create(DocxEditor.prototype);
  const calls = [];
  editor.filteredRevisionIds = () => [3];
  editor.setActiveRevision = (id) => calls.push(`active:${id}`);
  editor.revisionRunIds = new Map([['3', [{ focus: () => calls.push('focus'), scrollIntoView: () => calls.push('scroll') }]]]);
  editor.revisionParagraphIds = new Map();
  assert.equal(editor.focusRevision(3), true);
  assert.deepEqual(calls, ['active:3', 'focus', 'scroll']);
});

test('focusNextRevision cycles through filtered revision ids', () => {
  const editor = Object.create(DocxEditor.prototype);
  const focused = [];
  editor.filteredRevisionIds = () => [10, 20];
  editor.setActiveRevision = (id) => { editor.activeRevisionId = id; };
  editor.revisionRunIds = new Map([
    ['10', [{ focus: () => focused.push(10), scrollIntoView: () => {} }]],
    ['20', [{ focus: () => focused.push(20), scrollIntoView: () => {} }]],
  ]);
  editor.revisionParagraphIds = new Map();
  editor.activeRevisionId = null;
  assert.equal(editor.focusNextRevision(), 10);
  assert.equal(editor.focusNextRevision(), 20);
  assert.deepEqual(focused, [10, 20]);
});

test('focusPreviousRevision cycles backwards through filtered revision ids', () => {
  const editor = Object.create(DocxEditor.prototype);
  const focused = [];
  editor.filteredRevisionIds = () => [10, 20];
  editor.setActiveRevision = (id) => { editor.activeRevisionId = id; };
  editor.revisionRunIds = new Map([
    ['10', [{ focus: () => focused.push(10), scrollIntoView: () => {} }]],
    ['20', [{ focus: () => focused.push(20), scrollIntoView: () => {} }]],
  ]);
  editor.revisionParagraphIds = new Map();
  editor.activeRevisionId = null;
  assert.equal(editor.focusPreviousRevision(), 20);
  assert.equal(editor.focusPreviousRevision(), 10);
  assert.deepEqual(focused, [20, 10]);
});

test('acceptRevision flushes, applies, rerenders, and emits onChange once', () => {
  const editor = Object.create(DocxEditor.prototype);
  const calls = [];
  editor.flush = () => calls.push('flush');
  editor.document = {
    revision: 1,
    acceptRevision: () => { editor.document.revision = 2; },
    getSnapshot: () => ({ revision: 2 }),
  };
  editor.render = () => calls.push('render');
  editor.options = { onChange: () => calls.push('change') };
  editor.setActiveRevision = () => calls.push('clear');
  assert.equal(editor.acceptRevision(1), true);
  assert.deepEqual(calls, ['flush', 'render', 'change', 'clear']);
});

test('acceptRevision returns false when document revision is unchanged', () => {
  const editor = Object.create(DocxEditor.prototype);
  const calls = [];
  editor.flush = () => calls.push('flush');
  editor.document = {
    revision: 3,
    acceptRevision: () => {},
    getSnapshot: () => ({ revision: 3 }),
  };
  editor.render = () => calls.push('render');
  editor.options = { onChange: () => calls.push('change') };
  editor.setActiveRevision = () => calls.push('clear');
  assert.equal(editor.acceptRevision(1), false);
  assert.deepEqual(calls, ['flush']);
});

test('rejectRevision clears active revision after successful mutation', () => {
  const editor = Object.create(DocxEditor.prototype);
  const calls = [];
  editor.flush = () => calls.push('flush');
  editor.document = {
    revision: 4,
    rejectRevision: () => { editor.document.revision = 5; },
    getSnapshot: () => ({ revision: 5 }),
  };
  editor.render = () => calls.push('render');
  editor.options = { onChange: () => calls.push('change') };
  editor.setActiveRevision = (id) => calls.push(`active:${id}`);
  assert.equal(editor.rejectRevision(2), true);
  assert.deepEqual(calls, ['flush', 'render', 'change', 'active:null']);
});

test('acceptAllRevisions and rejectAllRevisions are no-ops when revision is unchanged', () => {
  const editor = Object.create(DocxEditor.prototype);
  const calls = [];
  editor.flush = () => calls.push('flush');
  editor.document = {
    revision: 7,
    acceptAllRevisions: () => {},
    rejectAllRevisions: () => {},
    getSnapshot: () => ({ revision: 7 }),
  };
  editor.render = () => calls.push('render');
  editor.options = { onChange: () => calls.push('change') };
  editor.setActiveRevision = () => calls.push('clear');
  assert.equal(editor.acceptAllRevisions({}), false);
  assert.equal(editor.rejectAllRevisions({}), false);
  assert.deepEqual(calls, ['flush', 'flush']);
});

test('acceptAllRevisions honors non-named reviewer filters via atomic applyOperations', () => {
  const editor = Object.create(DocxEditor.prototype);
  editor.flush = () => {};
  const calls = [];
  editor.document = {
    revision: 1,
    getRevisions: () => [{ id: 2 }, { id: 3, author: '' }],
    applyOperations: ({ operations }) => {
      calls.push(operations.map((operation) => `${operation.type}:${operation.id}`).join(','));
      editor.document.revision = 2;
    },
    getSnapshot: () => ({ revision: 2 }),
  };
  editor.render = () => {};
  editor.options = {};
  editor.setActiveRevision = () => {};
  assert.equal(editor.acceptAllRevisions({ authors: [{ kind: 'unattributed' }, { kind: 'empty', author: '' }] }), true);
  assert.deepEqual(calls, ['acceptRevision:2,acceptRevision:3']);
});

test('rejectAllRevisions with reviewer filter skips mutation when nothing matches', () => {
  const editor = Object.create(DocxEditor.prototype);
  editor.flush = () => {};
  const calls = [];
  editor.document = {
    revision: 9,
    getRevisions: () => [{ id: 1, author: 'Alice' }],
    applyOperations: () => calls.push('apply'),
    getSnapshot: () => ({ revision: 9 }),
  };
  editor.render = () => calls.push('render');
  editor.options = { onChange: () => calls.push('change') };
  editor.setActiveRevision = () => calls.push('active');
  assert.equal(editor.rejectAllRevisions({ authors: [{ kind: 'named', author: 'Bob' }] }), false);
  assert.deepEqual(calls, []);
});

test('setActiveRevision toggles docx-revision-active class by id membership', () => {
  const editor = Object.create(DocxEditor.prototype);
  const states = [];
  const node = {
    dataset: { docxRevisionIds: '1,2' },
    classList: { toggle: (_name, active) => states.push(active) },
  };
  editor.root = { querySelectorAll: () => [node] };
  editor.setActiveRevision(2);
  editor.setActiveRevision(3);
  assert.deepEqual(states, [true, false]);
});

test('appendRun writes data-docx-run for runs without revisions', () => {
  const editor = makeRunRenderEditor();
  const { runSpan } = appendRunToParagraph(editor, { run: { index: 2, text: 'plain' } });
  assert.equal(runSpan.dataset.docxRun, '2');
  assert.equal('docxRevisionIds' in runSpan.dataset, false);
});

test('appendRun hides hidden text by default and marks it with dashed underline when enabled', () => {
  for (const key of ['hidden', 'webHidden']) {
    for (const showHiddenText of [undefined, false, true]) {
      const editor = makeRunRenderEditor();
      editor.options.showHiddenText = showHiddenText;
      const shown = showHiddenText === true;
      let measurements = 0;
      editor.measure = () => { measurements++; return 12; };
      const { paragraphElement, runSpan, offset } = appendRunToParagraph(editor, {
        run: { index: 0, text: 'hidden', [key]: true, effective: { [key]: true } },
      });
      assert.equal(runSpan.dataset.docxHidden, '1');
      if (shown) {
        assert.equal(runSpan.className.includes('docx-hidden-text'), true);
        assert.equal(runSpan.style.textDecorationStyle, 'dashed');
        assert.equal(editor.readText(paragraphElement), 'hidden');
        assert.equal(measurements, 1);
        assert.equal(offset, 12);
      } else {
        assert.equal(runSpan.style.display, 'none');
        assert.equal(runSpan.contentEditable, 'false');
        assert.equal(editor.readText(paragraphElement), '');
        assert.equal(measurements, 0);
        assert.equal(offset, 0);
      }
    }
  }
});

test('makeParagraph omits hidden runs from pagination measurement and renders math consistently', () => {
  for (const showHiddenText of [undefined, false, true]) {
    const shown = showHiddenText === true;
    const editor = makeRunRenderEditor();
    editor.options.showHiddenText = showHiddenText;
    editor.paragraphs = new Map();
    editor.measuring = true;
    editor.composing = false;
    editor.renderAfterComposition = false;
    editor.document = { getShapeParagraphs: () => [] };
    editor.renderShapeInfos = [];
    const measured = [];
    editor.measure = (text) => { measured.push(text); return 8; };
    const paragraph = {
      index: 0,
      text: 'draftvisible',
      runs: [
        {
          index: 0, text: 'draft', hidden: true, effective: { hidden: true },
          revisions: [{ id: 1, kind: 'insertion' }],
          field: { index: 5, role: 'result', kind: 'PAGE' },
        },
        { index: 1, text: 'visible' },
      ],
      math: [{
        runOffset: 1, display: 'inline', linear: 'x',
        mathMl: { tag: 'math', children: [{ tag: 'mi', text: 'x' }] },
      }],
    };
    const reviewContext = { deletedTextByRun: new Map(), revisionColors: new Map() };
    for (let renderCount = 0; renderCount < 2; renderCount++) {
      const rendered = editor.makeParagraph(paragraph, 720, reviewContext);
      const content = rendered.childNodes.find((node) => node.className === 'docx-paragraph-content');
      const hidden = content.childNodes.find((node) => node.dataset.docxHidden === '1');
      const math = content.childNodes.filter((node) => node.dataset.docxMath === '1');
      assert.ok(hidden);
      assert.equal(hidden.style.display === 'none', !shown);
      assert.equal(hidden.dataset.docxFieldRole, shown ? 'result' : undefined);
      assert.equal(math.length, 1);
      assert.equal(hidden.dataset.docxRevisionIds, shown ? '1' : undefined);
      assert.equal(editor.readText(content), shown ? 'draftvisible' : 'visible');
    }
    assert.deepEqual(measured, shown ? ['draft', 'visible', 'draft', 'visible'] : ['visible', 'visible']);
  }
});

test('field runs retain roles, hide instruction text, and shade results only when enabled', () => {
  for (const showFieldShading of [undefined, false]) {
    const editor = makeRunRenderEditor();
    editor.options.showFieldShading = showFieldShading;
    const instruction = appendRunToParagraph(editor, {
      run: { index: 0, text: 'SHOULD NOT APPEAR', field: { index: 3, role: 'instruction' } },
    }).runSpan;
    const result = appendRunToParagraph(editor, {
      run: { index: 1, text: 'value', field: { index: 3, role: 'result' },
        revisions: [{ id: 4, kind: 'insertion' }] },
    }).runSpan;
    assert.equal(instruction.dataset.docxField, '3');
    assert.equal(instruction.dataset.docxFieldRole, 'instruction');
    assert.equal(instruction.childNodes.length, 0);
    assert.equal(instruction.contentEditable, 'false');
    assert.equal(result.dataset.docxFieldRole, 'result');
    assert.equal(result.dataset.docxContent, '1');
    assert.equal(result.contentEditable, 'false');
    assert.equal(result.dataset.docxRevisionIds, '4');
    assert.equal(result.childNodes[0].textContent, 'value');
    assert.equal(result.className.includes('docx-field-shading'), showFieldShading !== false);
  }
});

test('field result keeps comment and revision markers through repeated rendering', () => {
  const editor = makeRunRenderEditor();
  editor.commentRunIds.set('0:1', [5]);
  for (let index = 0; index < 2; index++) {
    const { runSpan } = appendRunToParagraph(editor, {
      run: { index: 1, text: 'cached', field: { index: 2, role: 'result' },
        revisions: [{ id: 10 + index, kind: 'insertion' }] },
    });
    assert.equal(runSpan.dataset.docxField, '2');
    assert.equal(runSpan.dataset.docxCommentIds, '5');
    assert.equal(runSpan.dataset.docxRevisionIds, String(10 + index));
  }
});

test('collapsed backspace/delete at a field boundary is blocked but adjacent text remains editable', () => {
  const text = (value) => ({ nodeType: 3, textContent: value });
  const span = (children, role) => ({
    nodeType: 1, tagName: 'SPAN', childNodes: children,
    dataset: role ? { docxField: '0', docxFieldRole: role, ...(role === 'result' ? { docxContent: '1' } : {}) } : {},
    contentEditable: role ? 'false' : 'inherit',
  });
  const before = text('before ');
  const after = text(' after');
  const content = span([
    span([before]), span([], 'instruction'), span([], 'instruction'),
    span([text('7')], 'result'), span([], 'instruction'), span([after]),
  ]);
  const attach = (node, parent) => {
    node.parentNode = parent;
    for (const child of node.childNodes ?? []) attach(child, node);
  };
  attach(content, null);
  content.contains = (node) => {
    for (; node; node = node.parentNode) if (node === content) return true;
    return false;
  };
  const range = { startContainer: after, startOffset: 0, collapsed: true };
  const editor = Object.create(DocxEditor.prototype);
  editor.root = { ownerDocument: { getSelection: () => ({ rangeCount: 1, getRangeAt: () => range }) } };
  const event = (inputType) => ({ inputType, prevented: false, preventDefault() { this.prevented = true; } });
  const backspace = event('deleteContentBackward');
  assert.equal(editor.preventFieldDeletion(content, backspace), true);
  assert.equal(backspace.prevented, true);
  for (const inputType of ['deleteWordBackward', 'deleteSoftLineBackward', 'deleteHardLineBackward']) {
    const wordDelete = event(inputType);
    assert.equal(editor.preventFieldDeletion(content, wordDelete), true);
    assert.equal(wordDelete.prevented, true);
  }
  range.startContainer = before;
  range.startOffset = before.textContent.length;
  const deleteForward = event('deleteContentForward');
  assert.equal(editor.preventFieldDeletion(content, deleteForward), true);
  assert.equal(deleteForward.prevented, true);
  for (const inputType of ['deleteWordForward', 'deleteSoftLineForward', 'deleteHardLineForward']) {
    const wordDelete = event(inputType);
    assert.equal(editor.preventFieldDeletion(content, wordDelete), true);
    assert.equal(wordDelete.prevented, true);
  }
  const safeBackward = event('deleteContentBackward');
  assert.equal(editor.preventFieldDeletion(content, safeBackward), false);
  assert.equal(safeBackward.prevented, false);
  range.startContainer = content;
  range.startOffset = 4;
  assert.equal(editor.preventFieldDeletion(content, event('deleteContentBackward')), true);
  range.startOffset = 3;
  assert.equal(editor.preventFieldDeletion(content, event('deleteContentForward')), true);
  range.startContainer = after;
  range.startOffset = 1;
  assert.equal(editor.preventFieldDeletion(content, event('deleteContentBackward')), false);
  range.startOffset = 0;
  range.collapsed = false;
  assert.equal(editor.preventFieldDeletion(content, event('deleteContentBackward')), false);
  range.collapsed = true;
  assert.equal(editor.preventFieldDeletion(content, event('insertText')), false);
});

test('flush edits ordinary text around complex and simple fields without changing fields', () => {
  for (const form of ['complex', 'simple']) {
    const doc = DocxDocument.create();
    const field = form === 'simple'
      ? `<w:fldSimple w:instr=" SEQ test "><w:r><w:t>7</w:t></w:r></w:fldSimple>`
      : `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> SEQ test </w:instrText></w:r>` +
        `<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>7</w:t></w:r>` +
        `<w:r><w:fldChar w:fldCharType="end"/></w:r>`;
    doc.setPartXml(doc.mainDocumentPath,
      `<w:document xmlns:w="${WORD_NS}"><w:body><w:p><w:r><w:t xml:space="preserve">前 </w:t></w:r>` +
      field + `<w:r><w:t xml:space="preserve"> 后</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`);
    const editor = Object.create(DocxEditor.prototype);
    editor.destroyed = false;
    editor.document = doc;
    editor.options = {};
    const content = { text: '前 7 后' };
    editor.readText = (node) => node.text;
    editor.paragraphs = new Map([[0, { content, text: content.text, failed: false }]]);
    const revision = doc.revision;
    content.text = '前缀 7 后';
    editor.flush();
    content.text = '前缀 7 后续';
    editor.flush();
    assert.equal(doc.revision, revision + 2, form);
    assert.equal(doc.getParagraphs()[0].text, content.text, form);
    assert.equal(doc.getFields()[0].result, '7', form);
    assert.equal(doc.getFields()[0].instruction, ' SEQ test ', form);
    const xml = doc.getPartXml(doc.mainDocumentPath);
    assert.match(xml, form === 'simple' ? /<w:fldSimple\b/ : /<w:fldChar\b/, form);
    if (form === 'complex') assert.match(xml, /<w:instrText> SEQ test <\/w:instrText>/);
  }
});

test('flush edits visible text on both sides of hidden runs without changing their XML', () => {
  const doc = DocxDocument.create();
  doc.setPartXml(doc.mainDocumentPath,
    `<w:document xmlns:w="${WORD_NS}"><w:body><w:p>` +
    '<w:r><w:t xml:space="preserve">before </w:t></w:r>' +
    '<w:ins w:id="1" w:author="Alice"><w:r><w:rPr><w:vanish/></w:rPr><w:t>secret</w:t></w:r></w:ins>' +
    '<w:r><w:t xml:space="preserve"> after</w:t></w:r>' +
    '</w:p><w:sectPr/></w:body></w:document>');

  const editor = Object.create(DocxEditor.prototype);
  const before = { nodeType: 3, textContent: 'updated before ' };
  const hidden = {
    nodeType: 1, tagName: 'SPAN', dataset: { docxHidden: '1', docxHiddenPreserved: '1' }, contentEditable: 'false',
    childNodes: [{ nodeType: 3, textContent: 'secret' }],
  };
  const after = { nodeType: 3, textContent: ' after changed' };
  const content = { nodeType: 1, tagName: 'SPAN', dataset: {}, childNodes: [before, hidden, after] };
  editor.destroyed = false;
  editor.options = {};
  editor.document = doc;
  editor.paragraphs = new Map([[0, { content, text: 'before  after', failed: false }]]);

  editor.flush();
  const firstRevision = doc.revision;
  assert.equal(firstRevision, 2);
  assert.equal(doc.getParagraphs()[0].text, 'updated before secret after changed');
  assert.equal(doc.getParagraphs()[0].visibleText, 'updated before  after changed');
  assert.equal(doc.getParagraphs()[0].runs[1].text, 'secret');
  assert.equal(doc.getParagraphs()[0].runs[1].revisions[0].kind, 'insertion');
  assert.match(doc.getPartXml(doc.mainDocumentPath), /<w:ins w:id="1" w:author="Alice"><w:r><w:rPr><w:vanish\/><\/w:rPr><w:t>secret<\/w:t><\/w:r><\/w:ins>/);

  editor.flush();
  assert.equal(doc.revision, firstRevision);
  before.textContent = 'again ';
  after.textContent = ' changed';
  editor.flush();
  assert.equal(doc.getParagraphs()[0].text, 'again secret changed');
  assert.equal(doc.getParagraphs()[0].runs[1].text, 'secret');
  assert.equal(doc.getParagraphs()[0].runs[1].revisions[0].kind, 'insertion');
  assert.match(doc.getPartXml(doc.mainDocumentPath), /<w:ins w:id="1" w:author="Alice"><w:r><w:rPr><w:vanish\/><\/w:rPr><w:t>secret<\/w:t><\/w:r><\/w:ins>/);
});

test('readText includes read-only field results but not instructions', () => {
  const editor = Object.create(DocxEditor.prototype);
  const text = (value) => ({ nodeType: 3, textContent: value });
  const field = (role, value) => ({
    nodeType: 1, tagName: 'SPAN', contentEditable: 'false',
    dataset: { docxField: '0', docxFieldRole: role, ...(role === 'result' ? { docxContent: '1' } : {}) },
    childNodes: [text(value)],
  });
  const content = { nodeType: 1, tagName: 'SPAN', dataset: {}, childNodes: [
    text('before '), field('instruction', ' HIDDEN '), field('result', '1'), text(' after'),
  ] };
  assert.equal(editor.readText(content), 'before 1 after');
});

test('cut and paste on a field result never reach the document mutation APIs', () => {
  const editor = Object.create(DocxEditor.prototype);
  editor.selectionTouchesField = () => true;
  editor.currentDocumentRange = () => { throw new Error('field selection must not become an edit range'); };
  const event = { prevented: false, preventDefault() { this.prevented = true; } };
  editor.handleClipboardCut(event, {});
  assert.equal(event.prevented, true);
  event.prevented = false;
  editor.handleClipboardPaste(event, {});
  assert.equal(event.prevented, true);
});

test('selection inside a field result expands to all result runs, including repeated selection', () => {
  const editor = Object.create(DocxEditor.prototype);
  const nodes = [
    { start: 2, end: 4, dataset: { docxField: '0' } },
    { start: 4, end: 6, dataset: { docxField: '1' } },
    { start: 6, end: 8, dataset: { docxField: '0' } },
  ];
  editor.root = { querySelectorAll: () => nodes };
  const range = (start, end) => ({
    startContainer: 'text', endContainer: 'text', startOffset: start, endOffset: end,
    collapsed: start === end,
    intersectsNode: (node) => start >= node.start && start < node.end || end > node.start && end <= node.end,
    cloneRange() { return range(this.startOffset, this.endOffset); },
    selectNode(node) { this.startOffset = node.start; this.endOffset = node.end; },
    compareBoundaryPoints(how, other) { return (how === 0 ? this.startOffset - other.startOffset : this.endOffset - other.endOffset); },
    setStartBefore(node) { this.startOffset = node.start; },
    setEndAfter(node) { this.endOffset = node.end; },
  });
  const selection = {
    anchorNode: { nodeType: 3, parentElement: { closest: () => nodes[0] } },
    current: range(3, 3), rangeCount: 1,
    getRangeAt() { return this.current; },
    removeAllRanges() {},
    addRange(value) { this.current = value; },
  };
  const previousRange = globalThis.Range;
  globalThis.Range = { START_TO_START: 0, END_TO_END: 2 };
  try {
    editor.expandFieldSelection(selection);
    assert.deepEqual([selection.current.startOffset, selection.current.endOffset], [2, 8]);
    editor.expandFieldSelection(selection);
    assert.deepEqual([selection.current.startOffset, selection.current.endOffset], [2, 8]);
    selection.anchorNode.parentElement.closest = () => nodes[1];
    selection.current = range(5, 5);
    editor.expandFieldSelection(selection);
    assert.deepEqual([selection.current.startOffset, selection.current.endOffset], [4, 6]);
  } finally {
    if (previousRange === undefined) delete globalThis.Range;
    else globalThis.Range = previousRange;
  }
});

test('appendRun omits data-docx-revision-ids for empty revision arrays', () => {
  const editor = makeRunRenderEditor();
  const { runSpan } = appendRunToParagraph(editor, { run: { index: 1, text: 'plain', revisions: [] } });
  assert.equal(runSpan.dataset.docxRun, '1');
  assert.equal('docxRevisionIds' in runSpan.dataset, false);
});

test('appendRun preserves hyperlink and comment datasets alongside run dataset', () => {
  const editor = makeRunRenderEditor();
  editor.commentParagraphIds.set(0, [5]);
  editor.commentRunIds.set('0:3', [7, 5]);
  const { runSpan } = appendRunToParagraph(editor, {
    run: {
      index: 3,
      text: 'link',
      hyperlink: { url: 'https://example.com' },
    },
  });
  assert.equal(runSpan.dataset.docxRun, '3');
  assert.equal(runSpan.dataset.docxLink, '1');
  assert.equal(runSpan.dataset.docxUrl, 'https://example.com');
  assert.equal(runSpan.dataset.docxCommentIds, '5,7');
});

test('appendRun writes distinct revision ids for different runs in the same paragraph', () => {
  const editor = makeRunRenderEditor();
  const paragraphElement = editor.root.ownerDocument.createElement('span');
  const paragraph = { index: 0 };
  const context = { deletedTextByRun: new Map(), revisionColors: new Map() };
  editor.appendRun(paragraphElement, paragraph, { index: 0, text: 'A', revisions: [{ id: 11, kind: 'insertion' }] }, context, 720, 0);
  editor.appendRun(paragraphElement, paragraph, { index: 1, text: 'B', revisions: [{ id: 22, kind: 'deletion' }] }, context, 720, 0);
  assert.equal(paragraphElement.childNodes[0].dataset.docxRevisionIds, '11');
  assert.equal(paragraphElement.childNodes[1].dataset.docxRevisionIds, '22');
  assert.equal(paragraphElement.childNodes[0].dataset.docxRun, '0');
  assert.equal(paragraphElement.childNodes[1].dataset.docxRun, '1');
});

test('appendRun writes all revision ids for a run with multiple revisions', () => {
  const editor = makeRunRenderEditor();
  const { runSpan } = appendRunToParagraph(editor, {
    run: {
      index: 4,
      text: 'AB',
      revisions: [
        { id: 31, kind: 'insertion', author: 'Alice' },
        { id: 32, kind: 'move', author: 'Bob', move: { name: 'm', side: 'to', pairedId: 33 } },
      ],
    },
  });
  assert.equal(runSpan.dataset.docxRevisionIds, '31,32');
  assert.equal(runSpan.dataset.docxRun, '4');
});

test('appendRun omits revision ids when revisions are hidden', () => {
  const editor = makeRunRenderEditor({ showRevisions: false });
  const reviewContext = { deletedTextByRun: new Map(), revisionColors: new Map() };
  const visibleRun = editor.reviewScopedRun(0, { index: 0, text: 'A', revisions: [{ id: 41, kind: 'insertion' }] }, reviewContext);
  const { runSpan } = appendRunToParagraph(editor, { run: visibleRun, reviewContext });
  assert.equal('docxRevisionIds' in runSpan.dataset, false);
  assert.equal(editor.revisionRunIds.size, 0);
});

test('appendRun omits revision ids when revisions are filtered out by author', () => {
  const editor = makeRunRenderEditor();
  const reviewContext = {
    authors: new Set(['named:Bob']),
    deletedTextByRun: new Map(),
    revisionColors: new Map(),
  };
  const visibleRun = editor.reviewScopedRun(0, { index: 0, text: 'A', revisions: [{ id: 42, kind: 'insertion', author: 'Alice' }] }, reviewContext);
  const { runSpan } = appendRunToParagraph(editor, { run: visibleRun, reviewContext });
  assert.equal('docxRevisionIds' in runSpan.dataset, false);
  assert.equal(editor.revisionRunIds.size, 0);
});

test('appendRun exposes revision ids in final view', () => {
  const editor = makeRunRenderEditor({ revisionView: 'final' });
  const reviewContext = { deletedTextByRun: new Map(), revisionColors: new Map() };
  const visibleRun = editor.reviewScopedRun(0, { index: 6, text: 'A', revisions: [{ id: 51, kind: 'deletion' }] }, reviewContext);
  const { runSpan } = appendRunToParagraph(editor, { run: visibleRun, reviewContext });
  assert.equal(runSpan.dataset.docxRevisionIds, '51');
  assert.equal(runSpan.dataset.docxRun, '6');
});

test('appendRun refreshes data-docx-run after rerendered run insertion', () => {
  const editor = makeRunRenderEditor();
  const initial = appendRunToParagraph(editor, { run: { index: 0, text: 'A' } });
  const rerenderedParagraph = editor.root.ownerDocument.createElement('span');
  const paragraph = { index: 0 };
  const context = { deletedTextByRun: new Map(), revisionColors: new Map() };
  editor.appendRun(rerenderedParagraph, paragraph, { index: 0, text: 'X' }, context, 720, 0);
  editor.appendRun(rerenderedParagraph, paragraph, { index: 1, text: 'A' }, context, 720, 0);
  assert.equal(initial.runSpan.dataset.docxRun, '0');
  assert.equal(rerenderedParagraph.childNodes[0].dataset.docxRun, '0');
  assert.equal(rerenderedParagraph.childNodes[1].dataset.docxRun, '1');
});

test('appendRun refreshes data-docx-run after rerendered run deletion', () => {
  const editor = makeRunRenderEditor();
  const initialParagraph = editor.root.ownerDocument.createElement('span');
  const paragraph = { index: 0 };
  const context = { deletedTextByRun: new Map(), revisionColors: new Map() };
  editor.appendRun(initialParagraph, paragraph, { index: 0, text: 'A' }, context, 720, 0);
  editor.appendRun(initialParagraph, paragraph, { index: 1, text: 'B' }, context, 720, 0);
  const rerendered = appendRunToParagraph(editor, { run: { index: 0, text: 'B' } });
  assert.equal(initialParagraph.childNodes[1].dataset.docxRun, '1');
  assert.equal(rerendered.runSpan.dataset.docxRun, '0');
});

test('appendDeletedRunVisualization + readText keeps deleted visualization text out of writeback text', () => {
  const doc = DocxDocument.create();
  doc.setTrackChanges(true);
  doc.setRevisionAuthor('Alice');
  doc.setParagraphText(0, 'ABXDEF');
  const editor = Object.create(DocxEditor.prototype);
  editor.reviewFilter = { showRevisions: true, showComments: true, revisionView: 'markup' };
  editor.revisionRunIds = new Map();
  editor.root = {
    ownerDocument: {
      createElement: (tagName) => ({
        nodeType: 1,
        tagName: tagName.toUpperCase(),
        dataset: {},
        contentEditable: 'inherit',
        style: {},
        childNodes: [],
        append(child) { this.childNodes.push(child); },
        setAttribute: () => {},
      }),
      createTextNode: (value) => ({ nodeType: 3, textContent: value }),
    },
  };
  const content = {
    nodeType: 1,
    tagName: 'SPAN',
    dataset: {},
    contentEditable: 'inherit',
    childNodes: [
      { nodeType: 3, textContent: 'AB' },
      { nodeType: 3, textContent: 'DEF' },
    ],
    append(node) { this.childNodes.splice(1, 0, node); },
  };
  const run = { index: 0, revisions: [{ id: 1, kind: 'deletion', author: 'Alice' }] };
  const context = { deletedTextByRun: new Map([['0:0', 'X']]), revisionColors: new Map() };
  editor.appendDeletedRunVisualization(content, 0, run, context);
  assert.equal(editor.readText(content), 'ABDEF');
  editor.reviewFilter = { showRevisions: true, showComments: true, revisionView: 'final' };
  editor.appendDeletedRunVisualization(content, 0, run, context);
  assert.equal(content.childNodes.length, 3);
});

test('filteredRevisionIds respects showRevisions and author filters', () => {
  const editor = Object.create(DocxEditor.prototype);
  editor.document = {
    getRevisions: () => [
      { id: 1, author: 'Alice' },
      { id: 2, author: 'Bob' },
    ],
  };
  editor.reviewFilter = { showRevisions: false, showComments: true, revisionView: 'markup' };
  assert.deepEqual(editor.filteredRevisionIds(), []);
  editor.reviewFilter = { showRevisions: true, showComments: true, revisionView: 'markup', authors: [{ kind: 'named', author: 'Bob' }] };
  assert.deepEqual(editor.filteredRevisionIds(), [2]);
});

test('filteredRevisionIds treats named author filter as trimmed identity', () => {
  const editor = Object.create(DocxEditor.prototype);
  editor.document = {
    getRevisions: () => [
      { id: 1, author: 'Alice' },
      { id: 2, author: ' Alice' },
      { id: 3, author: 'Alice ' },
    ],
  };
  editor.reviewFilter = { showRevisions: true, showComments: true, revisionView: 'markup', authors: [{ kind: 'named', author: 'Alice' }] };
  assert.deepEqual(editor.filteredRevisionIds(), [1, 2, 3]);
});

test('reviewColor stays stable for each author after accept/reject changes reviewer counts', () => {
  const doc = DocxDocument.create();
  doc.setTrackChanges(true);
  doc.setRevisionAuthor('Bob');
  doc.setParagraphText(0, 'B1');
  doc.setParagraphText(0, 'B2');
  doc.setRevisionAuthor('Alice');
  doc.setParagraphText(0, 'A1');
  const editor = Object.create(DocxEditor.prototype);
  const beforeContext = { deletedTextByRun: new Map(), revisionColors: new Map() };
  const beforeBob = editor.reviewColor('Bob', beforeContext);
  const beforeAlice = editor.reviewColor('Alice', beforeContext);
  const bobRevision = doc.getRevisions().find((revision) => revision.author === 'Bob');
  assert.ok(bobRevision);
  doc.acceptRevision(bobRevision.id);
  const afterContext = { deletedTextByRun: new Map(), revisionColors: new Map() };
  assert.equal(editor.reviewColor('Bob', afterContext), beforeBob);
  assert.equal(editor.reviewColor('Alice', afterContext), beforeAlice);
});

test('appendDeletedRunVisualization emits a non-editable deleted marker only in markup view', () => {
  const editor = Object.create(DocxEditor.prototype);
  const appended = [];
  editor.reviewFilter = { showRevisions: true, showComments: true, revisionView: 'markup' };
  editor.root = {
    ownerDocument: {
      createElement: () => ({
        dataset: {},
        style: {},
        className: '',
        setAttribute: () => {},
      }),
    },
  };
  editor.registerRevisionNode = () => {};
  const parent = { append: (node) => appended.push(node) };
  editor.appendDeletedRunVisualization(parent, 0, { index: 0, revisions: [{ id: 1, kind: 'deletion', author: 'Alice' }] }, {
    deletedTextByRun: new Map([['0:0', 'old']]),
    revisionColors: new Map([['named:Alice', '#2E75B6']]),
  });
  assert.equal(appended.length, 1);
  editor.reviewFilter = { showRevisions: true, showComments: true, revisionView: 'final' };
  editor.appendDeletedRunVisualization(parent, 0, { index: 0, revisions: [{ id: 1, kind: 'deletion', author: 'Alice' }] }, {
    deletedTextByRun: new Map([['0:0', 'old']]),
    revisionColors: new Map([['named:Alice', '#2E75B6']]),
  });
  assert.equal(appended.length, 1);
});

test('appendDeletedRunVisualization marks move-from text distinctly in markup view', () => {
  const editor = Object.create(DocxEditor.prototype);
  const appended = [];
  editor.reviewFilter = { showRevisions: true, showComments: true, revisionView: 'markup' };
  editor.makeMark = (text, label) => ({ dataset: { docxMark: '1' }, textContent: text, title: label });
  editor.root = {
    ownerDocument: {
      createElement: () => ({
        dataset: {},
        style: {},
        className: '',
        children: [],
        setAttribute: () => {},
        append(node) { this.children.push(node); },
      }),
    },
  };
  editor.registerRevisionNode = () => {};
  const parent = { append: (node) => appended.push(node) };
  editor.appendDeletedRunVisualization(parent, 0, {
    index: 0,
    revisions: [{ id: 1, kind: 'move', author: 'Alice', move: { name: 'm', side: 'from', pairedId: 2 } }],
  }, {
    deletedTextByRun: new Map([['0:0', 'old']]),
    revisionColors: new Map([['named:Alice', '#2E75B6']]),
  });
  assert.equal(appended.length, 1);
  assert.equal(appended[0].style.textDecoration.includes('underline'), true);
});
