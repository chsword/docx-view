import { test } from 'node:test';
import assert from 'node:assert/strict';
import { DocxDocument } from '../dist/document.js';
import { WORD_NS } from '../dist/xml.js';
import { AGENT_OPERATION_SCHEMA, validateCellFormat, validateRowFormat, validateTableFormat } from '../dist/operations.js';
import { PROPERTY_ORDER } from '../dist/internal/elements.js';

function withBody(xml) {
  const doc = DocxDocument.create();
  doc.setPartXml(doc.mainDocumentPath, `<w:document xmlns:w="${WORD_NS}"><w:body>${xml}<w:sectPr/></w:body></w:document>`);
  return doc;
}

function tableXml(inner) {
  return `<w:tbl>${inner}</w:tbl>`;
}

test('table parser folds horizontal and vertical merges into spans', () => {
  const doc = withBody(tableXml(`
    <w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>
    <w:tr>
      <w:tc><w:tcPr><w:gridSpan w:val="2"/><w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>
      <w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc>
    </w:tr>
    <w:tr>
      <w:tc><w:tcPr><w:gridSpan w:val="2"/><w:vMerge/></w:tcPr><w:p><w:r><w:t>C</w:t></w:r></w:p></w:tc>
      <w:tc><w:p><w:r><w:t>D</w:t></w:r></w:p></w:tc>
    </w:tr>
  `));
  const table = doc.getBlocks()[0];
  assert.equal(table.type, 'table');
  assert.deepEqual(table.grid, [2400, 2400, 2400]);
  assert.equal(table.rows[0].cells[0].colSpan, 2);
  assert.equal(table.rows[0].cells[0].rowSpan, 2);
  assert.equal(table.rows[1].cells[0].isMergeContinuation, true);
  assert.equal(table.rows[1].cells[0].rowSpan, 0);
});

test('table parser keeps nested tables recursively', () => {
  const doc = withBody(tableXml(`
    <w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid>
    <w:tr><w:tc>
      <w:p><w:r><w:t>outer</w:t></w:r></w:p>
      <w:tbl><w:tblGrid><w:gridCol w:w="1200"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>inner</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
    </w:tc></w:tr>
  `));
  const table = doc.getBlocks()[0];
  assert.equal(table.rows[0].cells[0].blocks[1].type, 'table');
  assert.equal(table.rows[0].cells[0].blocks[1].rows[0].cells[0].blocks[0].paragraph.text, 'inner');
});

test('table parser reads table, row, and cell formats', () => {
  const doc = withBody(tableXml(`
    <w:tblPr>
      <w:tblW w:type="pct" w:w="2500"/>
      <w:jc w:val="center"/>
      <w:tblLayout w:type="fixed"/>
      <w:tblBorders><w:insideH w:val="single" w:sz="8" w:color="FF0000"/></w:tblBorders>
      <w:shd w:fill="00FF00"/>
      <w:tblCaption w:val="caption"/>
      <w:tblDescription w:val="description"/>
    </w:tblPr>
    <w:tblGrid><w:gridCol w:w="3600"/></w:tblGrid>
    <w:tr><w:trPr><w:trHeight w:val="480" w:hRule="exact"/><w:tblHeader/></w:trPr><w:tc><w:tcPr><w:shd w:fill="FFFF00"/><w:vAlign w:val="center"/></w:tcPr><w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc></w:tr>
  `));
  const table = doc.getBlocks()[0];
  assert.equal(table.format.width.type, 'pct');
  assert.equal(table.format.alignment, 'center');
  assert.equal(table.format.layout, 'fixed');
  assert.equal(table.format.caption, 'caption');
  assert.equal(table.rows[0].format.height.rule, 'exact');
  assert.equal(table.rows[0].format.header, true);
  assert.equal(table.rows[0].cells[0].format.verticalAlign, 'center');
  assert.equal(table.rows[0].cells[0].format.shading.fill, 'FFFF00');
});

test('getTable returns parsed table info by top-level index', () => {
  const doc = DocxDocument.create();
  doc.insertParagraph('before');
  doc.insertTable([['A']]);
  doc.insertTable([['B', 'C']]);
  const table = doc.getTable(1);
  assert.equal(table.index, 1);
  assert.equal(table.rows[0].cells[1].blocks[0].paragraph.text, 'C');
  assert.equal(table.grid.length, 2);
});

test('insertTableAt inserts before a body block and applies format', () => {
  const doc = DocxDocument.create();
  doc.setParagraphText(0, 'tail');
  doc.insertTableAt(2, 2, 0, { layout: 'fixed', width: { type: 'dxa', value: 7200 } });
  const blocks = doc.getBlocks();
  assert.equal(blocks[0].type, 'table');
  assert.equal(blocks[1].paragraph.text, 'tail');
  // w:tblLayout 的值在 w:type 上，不是 w:val —— 原来这里断言 w:val，把一个 Word 不认的
  // 写法钉成了期望。
  assert.match(doc.getPartXml(doc.mainDocumentPath), /w:tblLayout w:type="fixed"/);
});

test('insertTable fills all cells atomically in one revision', () => {
  const doc = DocxDocument.create();
  doc.insertTable([['A', 'B'], ['C', 'D']]);
  assert.equal(doc.revision, 1);
  assert.equal(doc.getTable(0).rows[1].cells[1].blocks[0].paragraph.text, 'D');
  assert.match(doc.getPartXml(doc.mainDocumentPath), /<w:tbl><w:tblPr\/><w:tblGrid>/);
});

test('insertTableAt keeps a paragraph after a table inserted before another table', () => {
  const doc = DocxDocument.create();
  doc.insertTable([['A']]);
  doc.insertTableAt(1, 1, 1);
  assert.match(doc.getPartXml(doc.mainDocumentPath), /<w:tbl>[\s\S]*?<\/w:tbl><w:p>/);
});

test('insertTableRow appends a blank row with the current grid width count', () => {
  const doc = DocxDocument.create();
  doc.insertTable([['A', 'B']]);
  doc.insertTableRow(0, 1);
  const table = doc.getTable(0);
  assert.equal(table.rows.length, 2);
  assert.equal(table.rows[1].cells.length, 2);
  assert.equal(table.rows[1].cells[0].blocks[0].paragraph.text, '');
});

test('insertTableRow rejects negative positions', () => {
  const doc = DocxDocument.create();
  doc.insertTable([['A']]);
  assert.throws(() => doc.insertTableRow(0, -1), /non-negative/);
});

test('deleteTableRow rejects deleting the only row', () => {
  const doc = DocxDocument.create();
  doc.insertTable([['A']]);
  assert.throws(() => doc.deleteTableRow(0, 0), /only table row/);
});

test('deleteTableRow promotes the next vertical merge continuation to restart', () => {
  const doc = withBody(tableXml(`
    <w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid>
    <w:tr><w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc></w:tr>
    <w:tr><w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc></w:tr>
  `));
  doc.deleteTableRow(0, 0);
  assert.match(doc.getPartXml(doc.mainDocumentPath), /<w:vMerge w:val="restart"\/>/);
});

test('deleteTableRow repairs vertical merge chains after removing a middle continuation row', () => {
  const doc = withBody(tableXml(`
    <w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid>
    <w:tr><w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc></w:tr>
    <w:tr><w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc></w:tr>
    <w:tr><w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p><w:r><w:t>C</w:t></w:r></w:p></w:tc></w:tr>
  `));
  doc.deleteTableRow(0, 1);
  assert.equal(doc.getTable(0).rows[0].cells[0].rowSpan, 2);
});

test('insertTableColumn updates tblGrid and can expand merged cells', () => {
  const doc = withBody(tableXml(`
    <w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>
    <w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc></w:tr>
  `));
  doc.insertTableColumn(0, 1);
  const table = doc.getTable(0);
  assert.equal(table.grid.length, 3);
  assert.equal(table.rows[0].cells[0].colSpan, 3);
});

test('insertTableColumn rejects negative positions', () => {
  const doc = DocxDocument.create();
  doc.insertTable([['A']]);
  assert.throws(() => doc.insertTableColumn(0, -1), /non-negative/);
});

test('insertTableColumn preserves vertical merge geometry', () => {
  const doc = withBody(tableXml(`
    <w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>
    <w:tr><w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc></w:tr>
    <w:tr><w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p><w:r><w:t>C</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>D</w:t></w:r></w:p></w:tc></w:tr>
  `));
  doc.insertTableColumn(0, 0);
  const table = doc.getTable(0);
  assert.equal(table.grid.length, 3);
  assert.equal(table.rows[0].cells[0].colSpan, 2);
  assert.equal(table.rows[0].cells[0].rowSpan, 2);
});

test('deleteTableColumn shrinks grid spans and rejects deleting the only column', () => {
  const doc = withBody(tableXml(`
    <w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>
    <w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc></w:tr>
  `));
  doc.deleteTableColumn(0, 0);
  assert.equal(doc.getTable(0).grid.length, 1);
  assert.equal(doc.getTable(0).rows[0].cells[0].colSpan, 1);
  const single = DocxDocument.create();
  single.insertTable([['A']]);
  assert.throws(() => single.deleteTableColumn(0, 0), /only table column/);
});

test('deleteTableColumn preserves vertical merge geometry', () => {
  const doc = withBody(tableXml(`
    <w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>
    <w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/><w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc></w:tr>
    <w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/><w:vMerge/></w:tcPr><w:p><w:r><w:t>C</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>D</w:t></w:r></w:p></w:tc></w:tr>
  `));
  doc.deleteTableColumn(0, 0);
  const table = doc.getTable(0);
  assert.equal(table.grid.length, 2);
  assert.equal(table.rows[0].cells[0].colSpan, 1);
  assert.equal(table.rows[0].cells[0].rowSpan, 2);
});

test('mergeCells creates horizontal and vertical merge markup', () => {
  const doc = DocxDocument.create();
  doc.insertTable([['A', 'B'], ['C', 'D']]);
  doc.mergeCells(0, { row: 0, col: 0, rowSpan: 2, colSpan: 2 });
  const table = doc.getTable(0);
  assert.equal(table.rows[0].cells[0].colSpan, 2);
  assert.equal(table.rows[0].cells[0].rowSpan, 2);
  assert.match(doc.getPartXml(doc.mainDocumentPath), /<w:gridSpan w:val="2"\/>/);
  assert.match(doc.getPartXml(doc.mainDocumentPath), /<w:vMerge w:val="restart"\/>/);
});

test('mergeCells does not absorb cells to the right of the merge range', () => {
  const doc = DocxDocument.create();
  doc.insertTable([['A', 'B', 'C', 'D']]);
  doc.mergeCells(0, { row: 0, col: 1, rowSpan: 1, colSpan: 2 });
  const table = doc.getTable(0);
  assert.equal(table.rows[0].cells.length, 3);
  assert.equal(table.rows[0].cells[1].blocks[0].paragraph.text, 'B');
  assert.equal(table.rows[0].cells[2].blocks[0].paragraph.text, 'D');
  assert.equal(table.rows[0].cells[1].colSpan, 2);
});

test('splitCell restores a merged 2x2 cell to the original grid', () => {
  const doc = DocxDocument.create();
  doc.insertTable([['A', 'B'], ['C', 'D']]);
  doc.mergeCells(0, { row: 0, col: 0, rowSpan: 2, colSpan: 2 });
  doc.splitCell(0, 0, 0, 2, 2);
  const table = doc.getTable(0);
  assert.equal(table.rows[0].cells[0].colSpan, 1);
  assert.equal(table.rows[0].cells.length, 2);
  assert.equal(table.rows[1].cells.length, 2);
});

test('splitCell restores missing cells before trailing cells', () => {
  const doc = withBody(tableXml(`
    <w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>
    <w:tr>
      <w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p><w:r><w:t>M</w:t></w:r></w:p></w:tc>
      <w:tc><w:p><w:r><w:t>C</w:t></w:r></w:p></w:tc>
    </w:tr>
  `));
  doc.splitCell(0, 0, 0, 1, 2);
  const table = doc.getTable(0);
  assert.equal(table.rows[0].cells.length, 3);
  assert.equal(table.rows[0].cells[2].blocks[0].paragraph.text, 'C');
});

test('setCellText updates the visible cell text without removing the cell paragraph', () => {
  const doc = DocxDocument.create();
  doc.insertTable([['A']]);
  doc.setCellText(0, 0, 0, 'updated');
  assert.equal(doc.getTable(0).rows[0].cells[0].blocks[0].paragraph.text, 'updated');
  assert.match(doc.getPartXml(doc.mainDocumentPath), /updated/);
});

test('setCellText updates the outer cell paragraph instead of nested tables', () => {
  const doc = withBody(tableXml(`
    <w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid>
    <w:tr><w:tc>
      <w:tbl><w:tblGrid><w:gridCol w:w="1200"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>INNER</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
      <w:p><w:r><w:t>OUTER</w:t></w:r></w:p>
    </w:tc></w:tr>
  `));
  doc.setCellText(0, 0, 0, 'NEWTEXT');
  const xml = doc.getPartXml(doc.mainDocumentPath);
  assert.match(xml, /INNER/);
  assert.match(xml, /NEWTEXT/);
  assert.doesNotMatch(xml, /<w:t[^>]*>INNER<\/w:t>[\s\S]*<w:t[^>]*>OUTER<\/w:t>/);
});

test('formatTable writes layout and border properties', () => {
  const doc = DocxDocument.create();
  doc.insertTable([['A']]);
  doc.formatTable(0, { layout: 'fixed', borders: { top: { style: 'single', size: 8, color: '112233' } } });
  const xml = doc.getPartXml(doc.mainDocumentPath);
  assert.match(xml, /w:tblLayout w:type="fixed"/);
  assert.match(xml, /w:top w:val="single" w:sz="8" w:color="112233"/);
});

test('formatTable and formatCell keep schema-valid property order', () => {
  const doc = DocxDocument.create();
  doc.insertTable([['A']]);
  doc.formatTable(0, {
    style: 'TableGrid',
    width: { type: 'dxa', value: 9000 },
    layout: 'fixed',
    cellMargin: { top: { type: 'dxa', value: 50 } },
    look: '04A0',
  });
  doc.formatCell(0, 0, 0, {
    width: { type: 'dxa', value: 3000 },
    verticalAlign: 'center',
    textDirection: 'tbRl',
    noWrap: true,
    hideMark: true,
    vMerge: 'restart',
  });
  const xml = doc.getPartXml(doc.mainDocumentPath);
  assert.match(xml, /<w:tblPr><w:tblStyle[\s\S]*<w:tblW[\s\S]*<w:tblLayout[\s\S]*<w:tblCellMar[\s\S]*<w:tblLook/);
  assert.match(xml, /<w:tcPr><w:tcW[\s\S]*<w:vMerge[\s\S]*<w:noWrap[\s\S]*<w:textDirection[\s\S]*<w:vAlign[\s\S]*<w:hideMark/);
});

test('column edits materialize missing tblGrid after tblPr', () => {
  const doc = withBody(tableXml(`
    <w:tblPr><w:tblW w:type="dxa" w:w="9000"/></w:tblPr>
    <w:tr><w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc></w:tr>
  `));
  doc.insertTableColumn(0, 2);
  const xml = doc.getPartXml(doc.mainDocumentPath);
  assert.match(xml, /<w:tbl>[\s\S]*<w:tblPr[\s\S]*<\/w:tblPr>\s*<w:tblGrid><w:gridCol[^>]*\/><w:gridCol[^>]*\/><w:gridCol[^>]*\/><\/w:tblGrid>\s*<w:tr>/);
});

test('formatTableRow writes row height and header properties', () => {
  const doc = DocxDocument.create();
  doc.insertTable([['A']]);
  doc.formatTableRow(0, 0, { height: { value: 480, rule: 'exact' }, header: true });
  const xml = doc.getPartXml(doc.mainDocumentPath);
  assert.match(xml, /w:trHeight w:val="480" w:hRule="exact"/);
  assert.match(xml, /<w:tblHeader\/>/);
});

test('formatTableRow writes tracked row revision markup with required id and author attributes', async () => {
  const doc = DocxDocument.create();
  doc.insertTable([['A']]);
  doc.formatTableRow(0, 0, { inserted: true, deleted: true });
  const document = doc.getPartDocument(doc.mainDocumentPath);
  const markers = [
    ...Array.from(document.getElementsByTagNameNS(WORD_NS, 'ins')),
    ...Array.from(document.getElementsByTagNameNS(WORD_NS, 'del')),
  ];
  assert.equal(markers.length, 2);
  for (const marker of markers) {
    assert.match(marker.getAttributeNS(WORD_NS, 'id') ?? '', /^\d+$/);
    assert.equal(marker.getAttributeNS(WORD_NS, 'author'), 'docx-view');
  }
  const reopened = await DocxDocument.load(await doc.toUint8Array());
  assert.equal(reopened.getTable(0).rows[0].format.inserted, true);
  assert.equal(reopened.getTable(0).rows[0].format.deleted, true);
});

test('formatTableRow can write an explicit revision author and date', () => {
  const doc = DocxDocument.create();
  doc.insertTable([['A']]);
  doc.formatTableRow(0, 0, { deleted: true, revision: { author: 'Alice', date: '2026-09-28T00:00:00Z' } });
  const xml = doc.getPartXml(doc.mainDocumentPath);
  assert.match(xml, /<w:del w:id="\d+" w:author="Alice" w:date="2026-09-28T00:00:00Z"\/>/);
});

test('formatTableRow(false) removes tracked row markup instead of writing boolean val', () => {
  const doc = withBody(tableXml(`
    <w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid>
    <w:tr><w:trPr><w:ins w:id="4" w:author="docx-view"/><w:del w:id="5" w:author="docx-view"/></w:trPr><w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc></w:tr>
  `));
  doc.formatTableRow(0, 0, { inserted: false, deleted: false });
  const xml = doc.getPartXml(doc.mainDocumentPath);
  assert.doesNotMatch(xml, /w:ins/);
  assert.doesNotMatch(xml, /w:del/);
  assert.doesNotMatch(xml, /w:val="0"/);
});

test('applyOperations accepts row revision metadata', () => {
  const doc = DocxDocument.create();
  doc.insertTable([['A']]);
  doc.applyOperations({
    operations: [{ type: 'formatTableRow', table: 0, row: 0, format: { inserted: true, revision: { author: 'Agent' } } }],
  });
  assert.match(doc.getPartXml(doc.mainDocumentPath), /<w:ins w:id="\d+" w:author="Agent"\/>/);
});

test('formatCell writes cell formatting and preserves merge metadata fields', () => {
  const doc = DocxDocument.create();
  doc.insertTable([['A']]);
  doc.formatCell(0, 0, 0, { verticalAlign: 'bottom', shading: { fill: 'ABCDEF' }, noWrap: true });
  const xml = doc.getPartXml(doc.mainDocumentPath);
  assert.match(xml, /w:vAlign w:val="bottom"/);
  assert.match(xml, /w:shd w:fill="ABCDEF"/);
  assert.match(xml, /<w:noWrap\/>/);
});

test('table operations are available through atomic agent batches', () => {
  const doc = DocxDocument.create();
  const snapshot = doc.applyOperations({
    expectedRevision: 0,
    operations: [
      { type: 'insertTableAt', rows: 2, cols: 2 },
      { type: 'mergeCells', table: 0, range: { row: 0, col: 0, rowSpan: 2, colSpan: 2 } },
      { type: 'formatCell', table: 0, row: 0, col: 0, format: { shading: { fill: 'CCCCCC' } } },
    ],
  });
  assert.equal(snapshot.revision, 1);
  assert.equal(doc.getTable(0).rows[0].cells[0].rowSpan, 2);
  assert.match(doc.getPartXml(doc.mainDocumentPath), /CCCCCC/);
});

test('tblLayout round-trips through w:type, the attribute Word actually writes', () => {
  const doc = withBody(tableXml('<w:tblPr><w:tblLayout w:type="fixed"/></w:tblPr>'));
  assert.equal(doc.getTable(0).format.layout, 'fixed');
  doc.formatTable(0, { layout: 'autofit' });
  const xml = doc.getPartXml(doc.mainDocumentPath);
  assert.match(xml, /<w:tblLayout w:type="autofit"\/>/);
  assert.doesNotMatch(xml, /w:tblLayout w:val=/);
  assert.equal(doc.getTable(0).format.layout, 'autofit');
});

test('gridBefore shifts the row grid so vertical merges still line up', () => {
  const tc = (text, props = '') => `<w:tc>${props ? `<w:tcPr>${props}</w:tcPr>` : ''}`
    + `<w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
  // 三列。第 0 行最后一格起跨行合并；第 1 行用 gridBefore 跳过前两列，只剩那一格续接。
  const doc = withBody(
    `<w:tbl><w:tblPr/><w:tblGrid>${'<w:gridCol w:w="1000"/>'.repeat(3)}</w:tblGrid>
      <w:tr>${tc('a0')}${tc('a1')}${tc('a2', '<w:vMerge w:val="restart"/>')}</w:tr>
      <w:tr><w:trPr><w:gridBefore w:val="2"/><w:wBefore w:w="2000" w:type="dxa"/></w:trPr>
        ${tc('', '<w:vMerge/>')}</w:tr>
    </w:tbl>`);
  const rows = doc.getTable(0).rows;
  assert.deepEqual(rows[1].format.gridBefore, 2);
  assert.deepEqual(rows[1].format.widthBefore, { type: 'dxa', value: 2000 });
  // 跨行合并是靠网格起始列匹配的。gridBefore 不算进起始列，第 1 行那一格会被当成第 0 列，
  // 对不上第 2 列登记的 restart，合并就断了。
  assert.equal(rows[0].cells[2].rowSpan, 2);
  assert.equal(rows[1].cells[0].isMergeContinuation, true);
});

test('row grid skips read, reject nonsense and round-trip', () => {
  const row = (trPr) => withBody(
    `<w:tbl><w:tblPr/><w:tblGrid>${'<w:gridCol w:w="1000"/>'.repeat(3)}</w:tblGrid>
      <w:tr><w:trPr>${trPr}</w:trPr><w:tc><w:p/></w:tc></w:tr></w:tbl>`).getTable(0).rows[0].format;
  assert.equal(row('<w:gridAfter w:val="1"/><w:wAfter w:w="1000" w:type="dxa"/>').gridAfter, 1);
  assert.deepEqual(row('<w:gridAfter w:val="1"/><w:wAfter w:w="1000" w:type="dxa"/>').widthAfter,
    { type: 'dxa', value: 1000 });
  // 0、负数、非数字都按未设置处理——跳过 0 列和没写是一回事。
  for (const nonsense of ['<w:gridBefore w:val="0"/>', '<w:gridBefore w:val="-1"/>', '<w:gridBefore w:val="abc"/>']) {
    assert.equal(row(nonsense)?.gridBefore, undefined, nonsense);
  }

  const doc = withBody(
    `<w:tbl><w:tblPr/><w:tblGrid>${'<w:gridCol w:w="1000"/>'.repeat(3)}</w:tblGrid>
      <w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>`);
  doc.formatTableRow(0, 0, { gridBefore: 2, widthBefore: { type: 'dxa', value: 2000 } });
  const xml = doc.getPartXml(doc.mainDocumentPath);
  assert.match(xml, /<w:gridBefore w:val="2"\/>/);
  assert.match(xml, /<w:wBefore w:type="dxa" w:w="2000"\/>/);
  assert.equal(doc.getTable(0).rows[0].format.gridBefore, 2);
  // 传 0 表示清除。
  doc.formatTableRow(0, 0, { gridBefore: 0 });
  assert.doesNotMatch(doc.getPartXml(doc.mainDocumentPath), /gridBefore/);

  // 校验在 agent 操作那条路上（直接的 TS 方法对所有行格式字段都不校验，这是既有设计），
  // 所以直接测导出的校验器。
  for (const invalid of [{ gridBefore: -1 }, { gridBefore: 1.5 }, { gridAfter: 'x' },
    { widthBefore: { type: 'nope', value: 1 } }, { gridNope: 1 }]) {
    assert.throws(() => validateRowFormat(invalid), undefined, JSON.stringify(invalid));
  }
  assert.doesNotThrow(() => validateRowFormat({ gridBefore: 2, widthBefore: { type: 'dxa', value: 2000 } }));
});

test('tblpPr reads, writes and clears floating table positioning', () => {
  const floating = (tblPr) => withBody(
    `<w:tbl><w:tblPr>${tblPr}</w:tblPr><w:tblGrid><w:gridCol w:w="2880"/></w:tblGrid>
      <w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>`).getTable(0).format?.floatingPosition;

  assert.deepEqual(floating('<w:tblpPr w:leftFromText="180" w:rightFromText="180" w:vertAnchor="text"'
    + ' w:horzAnchor="margin" w:tblpXSpec="right" w:tblpY="1"/>'), {
    leftFromText: 180, rightFromText: 180, verticalAnchor: 'text',
    horizontalAnchor: 'margin', xSpec: 'right', y: 1,
  });
  // w:tblpPr 在场就是浮动表格，即便一个属性都没给——空对象也不能退化成 undefined，
  // 那会把「浮动」这件事本身丢掉。
  assert.deepEqual(floating('<w:tblpPr/>'), {});
  // 非法枚举与非数字按未设置处理，但仍然是浮动表格。
  assert.deepEqual(floating('<w:tblpPr w:vertAnchor="nope" w:tblpXSpec="sideways" w:tblpX="abc"/>'), {});
  assert.equal(floating(''), undefined);

  const doc = withBody(
    `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2880"/></w:tblGrid>
      <w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>`);
  doc.formatTable(0, { floatingPosition: { leftFromText: 180, horizontalAnchor: 'margin', xSpec: 'right', y: 1 } });
  assert.match(doc.getPartXml(doc.mainDocumentPath),
    /<w:tblpPr w:leftFromText="180" w:horzAnchor="margin" w:tblpXSpec="right" w:tblpY="1"\/>/);
  assert.deepEqual(doc.getTable(0).format.floatingPosition,
    { leftFromText: 180, horizontalAnchor: 'margin', xSpec: 'right', y: 1 });
  // null 清除浮动定位，表格回到正常流。
  doc.formatTable(0, { floatingPosition: null });
  assert.doesNotMatch(doc.getPartXml(doc.mainDocumentPath), /tblpPr/);
  assert.equal(doc.getTable(0).format?.floatingPosition, undefined);

  for (const invalid of [{ verticalAnchor: 'nope' }, { xSpec: 'sideways' }, { leftFromText: -1 },
    { leftFromText: 1.5 }, { vertAnchor: 'margin' }]) {
    assert.throws(() => validateTableFormat({ floatingPosition: invalid }), undefined, JSON.stringify(invalid));
  }
  // x / y 是坐标，挪到页边距外时是负数，不能按间距去卡。
  assert.doesNotThrow(() => validateTableFormat({ floatingPosition: { x: -720, y: -360 } }));
});

function firstTable(doc) {
  return doc.getBlocks().find((block) => block.type === 'table');
}

/** 一个 td 的子元素名序列，用来核对 schema 顺序。 */
function childNames(element) {
  return [...element.childNodes].filter((node) => node.nodeType === 1).map((node) => node.localName);
}

test('diagonal cell borders read, write back and stay inside tcBorders order', () => {
  const doc = withBody(tableXml(`
    <w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid>
    <w:tr><w:tc><w:tcPr><w:tcBorders>
      <w:top w:val="single" w:sz="4"/>
      <w:tl2br w:val="single" w:sz="8" w:color="FF0000"/>
      <w:tr2bl w:val="dashed" w:sz="12" w:color="00FF00"/>
    </w:tcBorders></w:tcPr><w:p/></w:tc></w:tr>`));
  const borders = firstTable(doc).rows[0].cells[0].format.borders;
  assert.deepEqual(borders.tl2br, { style: 'single', size: 8, space: undefined, color: 'FF0000', none: false });
  assert.deepEqual(borders.tr2bl, { style: 'dashed', size: 12, space: undefined, color: '00FF00', none: false });

  // 第 5 条：读出来的格式对象原样写回，对角线不能丢。setBorders 是整块重建 tcBorders 的，
  // 漏写哪条边，一次忠实的 read-modify-write 就会把那条边从文档里抹掉。
  doc.formatCell(0, 0, 0, firstTable(doc).rows[0].cells[0].format);
  const tcBorders = doc.getPartDocument(doc.mainDocumentPath).getElementsByTagNameNS(WORD_NS, 'tcBorders')[0];
  assert.deepEqual(childNames(tcBorders), ['top', 'tl2br', 'tr2bl']);
  const written = firstTable(doc).rows[0].cells[0].format.borders;
  assert.deepEqual([written.tl2br, written.tr2bl], [borders.tl2br, borders.tr2bl]);

  // 顺序：对角线排在六条常规边之后。
  doc.formatCell(0, 0, 0, { borders: { tl2br: { style: 'single' }, insideV: { style: 'single' }, left: { style: 'single' } } });
  const reordered = doc.getPartDocument(doc.mainDocumentPath).getElementsByTagNameNS(WORD_NS, 'tcBorders')[0];
  assert.deepEqual(childNames(reordered), ['left', 'insideV', 'tl2br']);
});

test('diagonals belong to w:tcBorders only, never to w:tblBorders', () => {
  // w:tblBorders 的 schema 里没有对角线，所以表格级必须拒，单元格级必须收。
  assert.throws(() => validateTableFormat({ borders: { tl2br: { style: 'single' } } }), /tl2br/);
  assert.throws(() => validateTableFormat({ borders: { tr2bl: { style: 'single' } } }), /tr2bl/);
  assert.doesNotThrow(() => validateCellFormat({ borders: { tl2br: { style: 'single' }, tr2bl: { style: 'single' } } }));
  assert.throws(() => validateCellFormat({ borders: { tl2br: { size: -1 } } }), /borders\.tl2br\.size/);

  // 写入侧也要拒：formatTable 拿到对角线时不能悄悄落到 tblBorders 里。
  const doc = withBody(tableXml('<w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid><w:tr><w:tc><w:p/></w:tc></w:tr>'));
  assert.throws(() => doc.formatTable(0, { borders: { tl2br: { style: 'single' } } }));
  assert.equal(/tl2br/.test(doc.getPartXml(doc.mainDocumentPath)), false);
});

test('tblCellSpacing reads on table and row and writes in schema order', () => {
  const doc = withBody(tableXml(`
    <w:tblPr><w:tblCellSpacing w:w="30" w:type="dxa"/></w:tblPr>
    <w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid>
    <w:tr><w:trPr><w:tblCellSpacing w:w="45" w:type="dxa"/></w:trPr><w:tc><w:p/></w:tc></w:tr>`));
  const table = firstTable(doc);
  assert.deepEqual(table.format.cellSpacing, { type: 'dxa', value: 30 });
  assert.deepEqual(table.rows[0].format.cellSpacing, { type: 'dxa', value: 45 });

  // w:trPr 里 tblCellSpacing 排在 tblHeader 之后、jc 之前。PROPERTY_ORDER 漏掉它的话
  // property() 会静默追加到末尾，落到 w:jc / w:del 后面，顺序就不合法了。
  doc.formatTableRow(0, 0, { header: true, alignment: 'center', cellSpacing: { type: 'dxa', value: 60 }, cantSplit: true });
  const trPr = doc.getPartDocument(doc.mainDocumentPath).getElementsByTagNameNS(WORD_NS, 'trPr')[0];
  assert.deepEqual(childNames(trPr), ['cantSplit', 'tblHeader', 'tblCellSpacing', 'jc']);
  assert.deepEqual(firstTable(doc).rows[0].format.cellSpacing, { type: 'dxa', value: 60 });

  // 表格级：排在 w:jc 之后、w:tblInd 之前。
  doc.formatTable(0, { cellSpacing: { type: 'dxa', value: 15 }, alignment: 'center', indent: 120 });
  const tblPr = doc.getPartDocument(doc.mainDocumentPath).getElementsByTagNameNS(WORD_NS, 'tblPr')[0];
  const written = childNames(tblPr);
  assert.deepEqual(written.filter((name) => ['jc', 'tblCellSpacing', 'tblInd'].includes(name)),
    ['jc', 'tblCellSpacing', 'tblInd']);
});

test('tblOverlap and tcFitText read, round-trip and reject nonsense', () => {
  const doc = withBody(tableXml(`
    <w:tblPr><w:tblOverlap w:val="never"/></w:tblPr>
    <w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid>
    <w:tr><w:tc><w:tcPr><w:tcFitText/></w:tcPr><w:p/></w:tc></w:tr>`));
  assert.equal(firstTable(doc).format.overlap, 'never');
  assert.equal(firstTable(doc).rows[0].cells[0].format.fitText, true);

  doc.formatTable(0, { overlap: 'overlap' });
  doc.formatCell(0, 0, 0, { fitText: false });
  assert.equal(firstTable(doc).format.overlap, 'overlap');
  assert.equal(firstTable(doc).rows[0].cells[0].format.fitText, false);
  // w:tcFitText 排在 w:textDirection 之后、w:vAlign 之前。
  doc.formatCell(0, 0, 0, { fitText: true, verticalAlign: 'center', textDirection: 'tbRl' });
  const tcPr = doc.getPartDocument(doc.mainDocumentPath).getElementsByTagNameNS(WORD_NS, 'tcPr')[0];
  assert.deepEqual(childNames(tcPr).filter((name) => ['textDirection', 'tcFitText', 'vAlign'].includes(name)),
    ['textDirection', 'tcFitText', 'vAlign']);

  assert.throws(() => validateTableFormat({ overlap: 'sometimes' }), /overlap must be never or overlap/);
  assert.throws(() => validateCellFormat({ fitText: 'yes' }), /fitText must be boolean/);
  // 未知枚举值按未设置处理，不抛错（畸形输入降级）。
  const degraded = withBody(tableXml(`
    <w:tblPr><w:tblOverlap w:val="maybe"/></w:tblPr>
    <w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid><w:tr><w:tc><w:p/></w:tc></w:tr>`));
  assert.equal(firstTable(degraded).format?.overlap, undefined);
});

test('every table property the writers emit is in PROPERTY_ORDER', () => {
  // property() 对不在表里的名字是静默追加到末尾的——没有任何信号，而写出来的 XML 顺序
  // 不合法。所以这里直接按 PROPERTY_ORDER 这一份真相核对，不另抄一份名单。
  const doc = withBody(tableXml('<w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid><w:tr><w:tc><w:p/></w:tc></w:tr>'));
  doc.formatTable(0, {
    width: { type: 'dxa', value: 4800 }, alignment: 'center', indent: 120, layout: 'fixed',
    borders: { top: { style: 'single' }, insideV: { style: 'single' } }, shading: { fill: 'EEEEEE' },
    cellMargin: { top: { type: 'dxa', value: 60 } }, cellSpacing: { type: 'dxa', value: 30 },
    overlap: 'never', style: 'Grid', look: '04A0', caption: 'c', description: 'd', bidiVisual: true,
    floatingPosition: { leftFromText: 180, xSpec: 'right' },
  });
  // 单元格先写：gridBefore 会把这一行的网格整体右移，之后 cellAt(0, 0) 就找不到单元格了。
  doc.formatCell(0, 0, 0, {
    width: { type: 'dxa', value: 2400 },
    borders: { top: { style: 'single' }, tl2br: { style: 'single' }, tr2bl: { style: 'single' } },
    shading: { fill: 'DDDDDD' }, margin: { left: { type: 'dxa', value: 60 } }, verticalAlign: 'center',
    textDirection: 'tbRl', noWrap: true, fitText: true, hideMark: true,
  });
  doc.formatTableRow(0, 0, {
    height: { value: 400, rule: 'atLeast' }, cellSpacing: { type: 'dxa', value: 45 }, cantSplit: true,
    header: true, alignment: 'center', gridBefore: 1, widthBefore: { type: 'dxa', value: 600 },
    gridAfter: 1, widthAfter: { type: 'dxa', value: 600 },
  });
  const xml = doc.getPartDocument(doc.mainDocumentPath);
  for (const container of ['tblPr', 'trPr', 'tcPr', 'tblBorders', 'tcBorders', 'tblCellMar', 'tcMar']) {
    const element = xml.getElementsByTagNameNS(WORD_NS, container)[0];
    assert.ok(element, `${container} was not written`);
    const order = PROPERTY_ORDER[container];
    const positions = childNames(element).map((name) => [name, order.indexOf(name)]);
    assert.deepEqual(positions.filter(([, position]) => position < 0), [],
      `${container} wrote names that PROPERTY_ORDER does not know`);
    assert.deepEqual(positions.map(([, position]) => position),
      [...positions.map(([, position]) => position)].sort((left, right) => left - right),
      `${container} children are out of schema order`);
  }
});

test('table format fields accepted at runtime are also in the agent schema', () => {
  // 第 13 条：schema 比运行时校验窄，等于这些字段对 agent 不存在。
  const operations = AGENT_OPERATION_SCHEMA.properties.operations.items.oneOf;
  const formatSchema = (name) => operations.find((operation) => operation.properties.type.const === name)
    .properties.format.properties;
  for (const key of ['cellSpacing', 'overlap', 'floatingPosition']) assert.ok(key in formatSchema('formatTable'), key);
  for (const key of ['cellSpacing', 'gridBefore', 'widthBefore', 'gridAfter', 'widthAfter']) {
    assert.ok(key in formatSchema('formatTableRow'), key);
  }
  for (const key of ['fitText']) assert.ok(key in formatSchema('formatCell'), key);
  for (const key of ['tl2br', 'tr2bl']) {
    assert.ok(key in formatSchema('formatCell').borders.properties, key);
    assert.equal(key in formatSchema('formatTable').borders.properties, false, `${key} must not be on tblBorders`);
  }

  const doc = withBody(tableXml('<w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid><w:tr><w:tc><w:p/></w:tc></w:tr>'));
  assert.doesNotThrow(() => doc.applyOperations({ operations: [
    { type: 'formatTable', table: 0, format: { cellSpacing: { type: 'dxa', value: 30 }, overlap: 'never', floatingPosition: { leftFromText: 180 } } },
    { type: 'formatCell', table: 0, row: 0, col: 0, format: { fitText: true, borders: { tl2br: { style: 'single' } } } },
    { type: 'formatTableRow', table: 0, row: 0, format: { cellSpacing: { type: 'dxa', value: 45 }, gridBefore: 1 } },
  ] }));
});

test('tblPrEx reads as a row table exception, writes before trPr, and round-trips', () => {
  // Word 合并两张表时下半张的边框落在各行的 tblPrEx 上，形状照 Word 的输出写。
  const doc = withBody(`<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>
    <w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>
    <w:tr><w:tc><w:p/></w:tc><w:tc><w:p/></w:tc></w:tr>
    <w:tr><w:tblPrEx><w:tblBorders><w:top w:val="single" w:sz="12" w:space="0" w:color="FF0000"/>
      <w:insideV w:val="dashed" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders>
      <w:shd w:val="clear" w:color="auto" w:fill="FFF2CC"/>
      <w:tblCellMar><w:left w:w="57" w:type="dxa"/></w:tblCellMar></w:tblPrEx>
      <w:trPr><w:trHeight w:val="400"/></w:trPr><w:tc><w:p/></w:tc><w:tc><w:p/></w:tc></w:tr>
    <w:tr><w:tblPrEx/><w:tc><w:p/></w:tc><w:tc><w:p/></w:tc></w:tr></w:tbl>`);
  const rows = doc.getBlocks().find((block) => block.type === 'table').rows;
  assert.equal(rows[0].format?.tableException, undefined);
  const exception = rows[1].format.tableException;
  assert.deepEqual(exception.borders.top, { style: 'single', size: 12, space: 0, color: 'FF0000', none: false });
  assert.equal(exception.borders.insideV.style, 'dashed');
  assert.equal(exception.shading.fill, 'FFF2CC');
  assert.deepEqual(exception.cellMargin, { left: { type: 'dxa', value: 57 } });
  assert.equal(rows[1].format.height.value, 400, 'trPr 照常读');
  assert.equal(rows[2].format?.tableException, undefined, '空的 tblPrEx 不覆盖任何东西');

  // 第 5 条：读出来原样写回不抛错，再读出来等价。
  doc.formatTableRow(0, 1, { tableException: exception });
  assert.deepEqual(doc.getBlocks().find((block) => block.type === 'table').rows[1].format.tableException, exception);

  // 给只有 tblPrEx、没有 trPr 的行加行高：trPr 必须排在 tblPrEx 后面。
  doc.formatTableRow(0, 2, { tableException: { alignment: 'center', shading: { value: 'clear', fill: 'DDDDDD' } }, height: { value: 300 } });
  const xml = new TextDecoder().decode(doc.getPartBytes('word/document.xml'));
  const third = xml.split('<w:tr>')[3];
  assert.match(third, /^<w:tblPrEx><w:jc w:val="center"\/><w:shd [^>]*\/><\/w:tblPrEx><w:trPr><w:trHeight w:val="300"\/><\/w:trPr><w:tc>/);
  assert.equal(doc.getBlocks().find((block) => block.type === 'table').rows[2].format.tableException.alignment, 'center');

  doc.formatTableRow(0, 1, { tableException: null });
  assert.equal(doc.getBlocks().find((block) => block.type === 'table').rows[1].format.tableException, undefined);
  assert.doesNotMatch(new TextDecoder().decode(doc.getPartBytes('word/document.xml')).split('<w:tr>')[2], /tblPrEx/);
});

test('formatTableRow and formatCell validate on the public API, not only in agent batches', () => {
  const doc = withBody('<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>');
  const revision = doc.revision;
  assert.throws(() => doc.formatTableRow(0, 0, { height: { value: -1 } }), /height/);
  assert.throws(() => doc.formatTableRow(0, 0, { tableException: { style: 'Grid' } }), /style/, 'tblPrEx 里没有 tblStyle');
  assert.throws(() => doc.formatTableRow(0, 0, { tableException: { shading: { value: 'clear', fill: 'nope' } } }));
  assert.throws(() => doc.formatCell(0, 0, 0, { verticalAlign: 'middle' }), /verticalAlign/);
  assert.equal(doc.revision, revision);
});
