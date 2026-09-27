import { test } from 'node:test';
import assert from 'node:assert/strict';
import { DocxDocument } from '../dist/document.js';
import { WORD_NS } from '../dist/xml.js';

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
      <w:tblLayout w:val="fixed"/>
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
  assert.match(doc.getPartXml(doc.mainDocumentPath), /w:tblLayout w:val="fixed"/);
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

test('setCellText updates the visible cell text without removing the cell paragraph', () => {
  const doc = DocxDocument.create();
  doc.insertTable([['A']]);
  doc.setCellText(0, 0, 0, 'updated');
  assert.equal(doc.getTable(0).rows[0].cells[0].blocks[0].paragraph.text, 'updated');
  assert.match(doc.getPartXml(doc.mainDocumentPath), /updated/);
});

test('formatTable writes layout and border properties', () => {
  const doc = DocxDocument.create();
  doc.insertTable([['A']]);
  doc.formatTable(0, { layout: 'fixed', borders: { top: { style: 'single', size: 8, color: '112233' } } });
  const xml = doc.getPartXml(doc.mainDocumentPath);
  assert.match(xml, /w:tblLayout w:val="fixed"/);
  assert.match(xml, /w:top w:val="single" w:sz="8" w:color="112233"/);
});

test('formatTableRow writes row height and header properties', () => {
  const doc = DocxDocument.create();
  doc.insertTable([['A']]);
  doc.formatTableRow(0, 0, { height: { value: 480, rule: 'exact' }, header: true });
  const xml = doc.getPartXml(doc.mainDocumentPath);
  assert.match(xml, /w:trHeight w:val="480" w:hRule="exact"/);
  assert.match(xml, /<w:tblHeader\/>/);
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
