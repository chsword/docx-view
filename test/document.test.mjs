import { test } from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import { DocxDocument } from '../dist/document.js';
import { AGENT_OPERATION_SCHEMA } from '../dist/operations.js';
import {
  A_NS, OFFICE_REL_NS, PIC_NS, V_NS, WP_NS,
  dataUrlForBytes, decodeBase64, emuToPx, pxToEmu,
} from '../dist/index.js';
import { REL_NS, WORD_NS } from '../dist/xml.js';

const RELS_TYPE = 'application/vnd.openxmlformats-package.relationships+xml';
const PNG_BYTES = decodeBase64('iVBORw0KGgoAAAANSUhEUgAAAAIAAAABCAQAAAD8fJRsAAAAC0lEQVR42mP8/x8AAusB9WnM0iEAAAAASUVORK5CYII=');
const GIF_BYTES = decodeBase64('R0lGODlhAgABAIABAAAAAP///yH5BAEAAAEALAAAAAACAAEAAAICRAEAOw==');
const BMP_BYTES = Uint8Array.from([0x42, 0x4d, 58, 0, 0, 0, 0, 0, 0, 0, 54, 0, 0, 0, 40, 0, 0, 0, 2, 0, 0, 0, 1, 0, 0, 0, 1, 0, 24, 0, 0, 0, 0, 0, 4, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 255, 0, 0, 0]);
const encoder = new TextEncoder();

const STYLES_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml';
const THEME_TYPE = 'application/vnd.openxmlformats-officedocument.theme+xml';

function withBody(xml) {
  const doc = DocxDocument.create();
  doc.setPartXml(doc.mainDocumentPath, `<w:document xmlns:w="${WORD_NS}"><w:body>${xml}<w:sectPr/></w:body></w:document>`);
  return doc;
}

function paragraphIndicesFromBlocks(blocks) {
  const indices = [];
  const walk = items => {
    for (const block of items) {
      if (block.type === 'paragraph') indices.push(block.paragraph.index);
      else for (const row of block.rows) for (const cell of row.cells) walk(cell.blocks);
    }
  };
  walk(blocks);
  return indices;
}

function withStyles(bodyXml, stylesXml, themeXml) {
  const doc = withBody(bodyXml);
  doc.addPart('word/styles.xml', new TextEncoder().encode(stylesXml), STYLES_TYPE);
  if (themeXml) doc.addPart('word/theme/theme1.xml', new TextEncoder().encode(themeXml), THEME_TYPE);
  return doc;
}

function withImageDoc(body, relationships, media = [{ path: 'word/media/image1.png', bytes: PNG_BYTES, type: 'image/png' }]) {
  const doc = DocxDocument.create();
  doc.setPartXml(doc.mainDocumentPath, `<w:document xmlns:w="${WORD_NS}" xmlns:r="${OFFICE_REL_NS}" xmlns:wp="${WP_NS}" xmlns:a="${A_NS}" xmlns:pic="${PIC_NS}" xmlns:v="${V_NS}"><w:body>${body}<w:sectPr/></w:body></w:document>`);
  doc.addPart('word/_rels/document.xml.rels', encoder.encode(`<Relationships xmlns="${REL_NS}">${relationships}</Relationships>`), RELS_TYPE);
  for (const part of media) doc.addPart(part.path, part.bytes, part.type);
  return doc;
}

test('create, edit, export and reopen a DOCX in Node without browser globals', async () => {
  const doc = DocxDocument.create();
  doc.setParagraphText(0, '你好 DOCX & <world> 😀');
  doc.insertParagraph('second\tline\nthird');
  doc.formatRun(0, 0, { bold: true, italic: true, underline: true, fontSize: 14, color: 'FF0000', fontFamily: 'Arial' });
  doc.formatParagraph(0, { alignment: 'center', style: 'Title' });
  const reopened = await DocxDocument.load(await doc.toUint8Array());
  assert.equal(reopened.getParagraphs()[0].text, '你好 DOCX & <world> 😀');
  assert.equal(reopened.getParagraphs()[1].text, 'second\tline\nthird');
  assert.equal(reopened.getParagraphs()[0].alignment, 'center');
  assert.equal(reopened.getParagraphs()[0].runs[0].bold, true);
  assert.equal(reopened.getParagraphs()[0].runs[0].fontSize, 14);
  assert.equal((await doc.toBlob()).type, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
  assert.equal((await DocxDocument.load(await doc.toBlob())).getParagraphs()[0].text, reopened.getParagraphs()[0].text);
});

test('unmodified binary and XML parts survive an unrelated paragraph edit byte-for-byte', async () => {
  const doc = DocxDocument.create();
  const bytes = Uint8Array.from([0, 1, 2, 255, 23]);
  const customXml = new TextEncoder().encode('<custom xmlns="urn:custom" data="keep"><item/></custom>');
  doc.addPart('word/media/image1.png', bytes, 'image/png');
  doc.addPart('customXml/item1.xml', customXml, 'application/xml');
  const loaded = await DocxDocument.load(await doc.toUint8Array());
  loaded.setParagraphText(0, 'changed');
  const saved = await DocxDocument.load(await loaded.toUint8Array());
  assert.deepEqual(saved.getPartBytes('word/media/image1.png'), bytes);
  assert.deepEqual(saved.getPartBytes('customXml/item1.xml'), customXml);
  const copy = saved.getPartBytes('word/media/image1.png');
  copy[0] = 99;
  assert.equal(saved.getPartBytes('word/media/image1.png')[0], 0);
});

test('text edits retain runs, bookmarks, drawings, and paragraph properties', () => {
  const doc = withBody('<w:p><w:pPr><w:spacing w:after="100"/></w:pPr><w:bookmarkStart w:id="1" w:name="keep"/><w:r><w:rPr><w:b/></w:rPr><w:t>Hello </w:t><w:drawing/></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>world</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>');
  doc.setParagraphText(0, 'Hello DOCX world!');
  const xml = doc.getPartXml(doc.mainDocumentPath);
  assert.match(xml, /w:drawing/);
  assert.match(xml, /w:bookmarkStart/);
  assert.match(xml, /w:spacing/);
  assert.equal(doc.getParagraphs()[0].text, 'Hello DOCX world!');
  assert.equal(doc.getParagraphs()[0].runs.length, 2);
  assert.equal(doc.getParagraphs()[0].runs[0].bold, true);
  assert.equal(doc.getParagraphs()[0].runs[1].italic, true);
});

test('replacement spans runs and repeated matches without losing non-text content', () => {
  const doc = withBody('<w:p><w:r><w:t>hel</w:t></w:r><w:r><w:drawing/><w:t>lo hello</w:t></w:r></w:p>');
  doc.replaceText('hello', '你好\tDOCX\n');
  assert.equal(doc.getParagraphs()[0].text, '你好\tDOCX\n 你好\tDOCX\n');
  assert.match(doc.getPartXml(doc.mainDocumentPath), /w:drawing/);
  assert.throws(() => doc.replaceText('', 'x'), /empty/);
});

test('replacement uses the affected run formatting and preserves preceding page breaks', () => {
  for (const edit of [
    doc => doc.replaceText('world', 'Earth'),
    doc => doc.setParagraphText(0, 'Hello Earth'),
  ]) {
    const doc = withBody('<w:p><w:r><w:t>Hello </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>world</w:t></w:r></w:p>');
    edit(doc);
    assert.equal(doc.getParagraphs()[0].runs[0].text, 'Hello ');
    assert.equal(doc.getParagraphs()[0].runs[1].text, 'Earth');
    assert.equal(doc.getParagraphs()[0].runs[1].bold, true);
  }
  const doc = withBody('<w:p><w:r><w:br w:type="page"/></w:r><w:r><w:t>world</w:t></w:r></w:p>');
  doc.replaceText('world', 'Earth');
  assert.match(doc.getPartXml(doc.mainDocumentPath), /<w:br w:type="page"\/>/);
  doc.setParagraphText(0, '\nnew Earth');
  assert.match(doc.getPartXml(doc.mainDocumentPath), /<w:br w:type="page"\/>/);
  assert.equal(doc.getParagraphs()[0].text, '\nnew Earth');
});

test('Node Buffer inputs and outputs cannot mutate internal parts without a revision', () => {
  const doc = DocxDocument.create();
  const input = Buffer.from([1, 2, 3]);
  doc.addPart('custom.bin', input, 'application/octet-stream');
  const revision = doc.revision;
  input[0] = 99;
  doc.getPartBytes('custom.bin')[1] = 99;
  assert.deepEqual(doc.getPartBytes('custom.bin'), Uint8Array.from([1, 2, 3]));
  assert.equal(doc.revision, revision);
  const replacement = Buffer.from([4, 5, 6]);
  doc.setPartBytes('custom.bin', replacement);
  replacement[0] = 99;
  assert.deepEqual(doc.getPartBytes('custom.bin'), Uint8Array.from([4, 5, 6]));
});

test('Unicode edits do not split surrogate pairs and invalid XML text is rejected', () => {
  const doc = DocxDocument.create();
  for (const text of ['😀', '😁', 'abc😁', 'abc😀', '😀😁', '😁', '', 'plain']) {
    doc.setParagraphText(0, text);
    assert.equal(doc.getParagraphs()[0].text, text);
  }
  assert.throws(() => doc.setParagraphText(0, '\ud800'), /valid XML/);
  assert.throws(() => doc.setParagraphText(0, '\u0000'), /valid XML/);
});

test('tables and cell editing retain a valid final paragraph', () => {
  const doc = DocxDocument.create();
  doc.insertTable([['A', 'B'], ['C']]);
  const table = doc.getBlocks().find(block => block.type === 'table');
  assert.equal(table.rows.length, 2);
  assert.equal(table.rows[1].cells.length, 2);
  const cellIndex = table.rows[0].cells[0].blocks[0].paragraph.index;
  doc.setParagraphText(cellIndex, 'updated cell');
  assert.equal(doc.getParagraphs()[cellIndex].text, 'updated cell');
  doc.deleteParagraph(cellIndex);
  assert.equal(doc.getBlocks().find(block => block.type === 'table').rows[0].cells[0].blocks[0].paragraph.text, '');
  assert.throws(() => doc.insertTable([]), /Table/);
});

test('getBlocks reads table rows wrapped by w:sdt and preserves paragraph indexing', () => {
  const doc = DocxDocument.create();
  doc.insertTable([['A', 'B'], ['C', 'D']]);
  let xml = doc.getPartXml(doc.mainDocumentPath);
  const rows = xml.match(/<w:tr>[\s\S]*?<\/w:tr>/g);
  xml = xml.replace(rows[1], `<w:sdt><w:sdtPr/><w:sdtContent>${rows[1]}</w:sdtContent></w:sdt>`);
  doc.setPartXml(doc.mainDocumentPath, xml);

  const table = doc.getBlocks().find(block => block.type === 'table');
  assert.equal(table.rows.length, 2);
  assert.equal(table.rows[1].cells[0].blocks[0].paragraph.text, 'C');
  assert.equal(table.rows[1].cells[1].blocks[0].paragraph.text, 'D');
});

test('setParagraphText can edit a row wrapped by w:sdt', () => {
  const doc = DocxDocument.create();
  doc.insertTable([['A', 'B'], ['C', 'D']]);
  let xml = doc.getPartXml(doc.mainDocumentPath);
  const rows = xml.match(/<w:tr>[\s\S]*?<\/w:tr>/g);
  xml = xml.replace(rows[1], `<w:sdt><w:sdtPr/><w:sdtContent>${rows[1]}</w:sdtContent></w:sdt>`);
  doc.setPartXml(doc.mainDocumentPath, xml);

  const cellParagraph = doc.getBlocks().find(block => block.type === 'table').rows[1].cells[0].blocks[0].paragraph.index;
  doc.setParagraphText(cellParagraph, 'C-updated');
  assert.equal(doc.getParagraphs()[cellParagraph].text, 'C-updated');
});

test('getBlocks pierces w:sdt wrappers around table cells including nested wrappers', () => {
  const doc = DocxDocument.create();
  doc.insertTable([['A', 'B'], ['C', 'D']]);
  let xml = doc.getPartXml(doc.mainDocumentPath);
  const firstCell = xml.match(/<w:tc>[\s\S]*?<\/w:tc>/)?.[0];
  xml = xml.replace(firstCell, `<w:sdt><w:sdtPr/><w:sdtContent><w:sdt><w:sdtPr/><w:sdtContent>${firstCell}</w:sdtContent></w:sdt></w:sdtContent></w:sdt>`);
  doc.setPartXml(doc.mainDocumentPath, xml);

  const table = doc.getBlocks().find(block => block.type === 'table');
  assert.equal(table.rows[0].cells[0].blocks[0].paragraph.text, 'A');
  const index = table.rows[0].cells[0].blocks[0].paragraph.index;
  doc.setParagraphText(index, 'A-updated');
  assert.equal(doc.getParagraphs()[index].text, 'A-updated');
});

test('run-level w:sdt wrappers are included in ownRuns/getParagraphs', () => {
  const doc = withBody('<w:p><w:r><w:t>head</w:t></w:r><w:sdt><w:sdtPr/><w:sdtContent><w:r><w:rPr><w:b/></w:rPr><w:t>tail</w:t></w:r></w:sdtContent></w:sdt></w:p>');
  const paragraph = doc.getParagraphs()[0];
  assert.deepEqual(paragraph.runs.map(run => run.text), ['head', 'tail']);
  assert.equal(paragraph.runs[1].bold, true);
});

test('round-trip keeps w:sdt and w:sdtPr while editing wrapped content', async () => {
  const doc = DocxDocument.create();
  doc.insertTable([['A', 'B'], ['C', 'D']]);
  let xml = doc.getPartXml(doc.mainDocumentPath);
  const rows = xml.match(/<w:tr>[\s\S]*?<\/w:tr>/g);
  xml = xml.replace(
    rows[1],
    `<w:sdt w:id="9"><w:sdtPr><w:alias w:val="row-wrap"/><w:tag w:val="meta"/></w:sdtPr><w:sdtContent>${rows[1]}</w:sdtContent></w:sdt>`,
  );
  doc.setPartXml(doc.mainDocumentPath, xml);
  const idx = doc.getBlocks().find(block => block.type === 'table').rows[1].cells[0].blocks[0].paragraph.index;
  doc.setParagraphText(idx, 'C2');

  const reopened = await DocxDocument.load(await doc.toUint8Array());
  const out = reopened.getPartXml(reopened.mainDocumentPath);
  assert.match(out, /<w:sdt w:id="9">/);
  assert.match(out, /<w:sdtPr><w:alias w:val="row-wrap"\/><w:tag w:val="meta"\/><\/w:sdtPr>/);
  assert.equal(reopened.getParagraphs()[idx].text, 'C2');
});

test('getBlocks and getParagraphs contain the same paragraph set with wrapped table content', () => {
  const doc = DocxDocument.create();
  doc.insertTable([['A', 'B'], ['C', 'D']]);
  let xml = doc.getPartXml(doc.mainDocumentPath);
  const rows = xml.match(/<w:tr>[\s\S]*?<\/w:tr>/g);
  const cells = xml.match(/<w:tc>[\s\S]*?<\/w:tc>/g);
  xml = xml
    .replace(rows[1], `<w:sdt><w:sdtPr/><w:sdtContent><w:customXml>${rows[1]}</w:customXml></w:sdtContent></w:sdt>`)
    .replace(cells[0], `<w:sdt><w:sdtPr/><w:sdtContent>${cells[0]}</w:sdtContent></w:sdt>`);
  doc.setPartXml(doc.mainDocumentPath, xml);

  const blockIndices = paragraphIndicesFromBlocks(doc.getBlocks()).sort((a, b) => a - b);
  const paragraphIndices = doc.getParagraphs().map(paragraph => paragraph.index).sort((a, b) => a - b);
  assert.deepEqual(blockIndices, paragraphIndices);
});

test('deleteParagraph keeps wrapped table cells structurally valid', () => {
  const doc = withBody('<w:tbl><w:tr><w:tc><w:sdt><w:sdtPr/><w:sdtContent><w:p><w:r><w:t>cellp</w:t></w:r></w:p></w:sdtContent></w:sdt></w:tc></w:tr></w:tbl>');
  doc.deleteParagraph(0);
  assert.equal(doc.getParagraphs().length, 1);
  assert.match(doc.getPartXml(doc.mainDocumentPath), /<w:tc>[\s\S]*<w:p(?:>|\/>)/);

  const viaOps = withBody('<w:tbl><w:tr><w:tc><w:sdt><w:sdtPr/><w:sdtContent><w:p><w:r><w:t>cellp</w:t></w:r></w:p></w:sdtContent></w:sdt></w:tc></w:tr></w:tbl>');
  viaOps.applyOperations({ operations: [{ type: 'deleteParagraph', index: 0 }] });
  assert.equal(viaOps.getParagraphs().length, 1);
});

test('deleteParagraph does not add blank paragraphs when wrapped cell content remains', () => {
  const doc = withBody('<w:tbl><w:tr><w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p><w:sdt><w:sdtPr/><w:sdtContent><w:p><w:r><w:t>B</w:t></w:r></w:p></w:sdtContent></w:sdt></w:tc></w:tr></w:tbl>');
  doc.deleteParagraph(0);
  assert.deepEqual(doc.getParagraphs().map(p => p.text), ['B']);
});

test('wrapped table rows remain addressable after structural edits and cell text updates', () => {
  const doc = DocxDocument.create();
  doc.insertTable([['A', 'B'], ['C', 'D']]);
  let xml = doc.getPartXml(doc.mainDocumentPath);
  const rows = xml.match(/<w:tr>[\s\S]*?<\/w:tr>/g);
  xml = xml.replace(rows[1], `<w:sdt><w:sdtPr/><w:sdtContent><w:customXml>${rows[1]}</w:customXml></w:sdtContent></w:sdt>`);
  doc.setPartXml(doc.mainDocumentPath, xml);

  doc.insertTableColumn(0, 1);
  doc.setCellText(0, 1, 2, 'D2');

  const table = doc.getTable(0);
  assert.equal(table.rows.length, 2);
  assert.equal(table.rows[1].cells.length, 3);
  assert.equal(table.rows[1].cells[2].blocks[0].paragraph.text, 'D2');
});

test('insertions precede section properties and deletion protects section breaks', () => {
  const doc = DocxDocument.create();
  doc.insertParagraph('first', 0);
  doc.insertParagraph('last');
  assert.deepEqual(doc.getParagraphs().map(p => p.text), ['first', '', 'last']);
  assert.match(doc.getPartXml(doc.mainDocumentPath), /last[\s\S]*<w:sectPr/);
  const sectionDoc = withBody('<w:p><w:pPr><w:sectPr/></w:pPr></w:p>');
  assert.throws(() => sectionDoc.deleteParagraph(0), /section-break/);
  const emptyDoc = DocxDocument.create();
  emptyDoc.deleteParagraph(0);
  assert.deepEqual(emptyDoc.getParagraphs().map(p => p.text), ['']);
});

test('deleteParagraph clears deleting the only body paragraph inside w:sdt', () => {
  const doc = withBody('<w:sdt><w:sdtPr/><w:sdtContent><w:p><w:r><w:t>x</w:t></w:r></w:p></w:sdtContent></w:sdt>');
  doc.deleteParagraph(0);
  assert.deepEqual(doc.getParagraphs().map(p => p.text), ['']);
});

test('deleteParagraph clears deleting the only body paragraph inside w:customXml', () => {
  const doc = withBody('<w:customXml><w:p><w:r><w:t>x</w:t></w:r></w:p></w:customXml>');
  doc.deleteParagraph(0);
  assert.deepEqual(doc.getParagraphs().map(p => p.text), ['']);
});

test('deleteParagraph clears deleting the only body paragraph inside nested w:sdt', () => {
  const doc = withBody('<w:sdt><w:sdtPr/><w:sdtContent><w:sdt><w:sdtPr/><w:sdtContent><w:p><w:r><w:t>x</w:t></w:r></w:p></w:sdtContent></w:sdt></w:sdtContent></w:sdt>');
  doc.deleteParagraph(0);
  assert.deepEqual(doc.getParagraphs().map(p => p.text), ['']);
});

test('deleteParagraph clears deleting the trailing paragraph after final table in body', () => {
  const doc = DocxDocument.create();
  doc.insertTable([['A', 'B']]);
  doc.deleteParagraph(3);
  assert.deepEqual(doc.getParagraphs().map(p => p.text), ['', 'A', 'B', '']);
});

test('deleteParagraph allows deleting one of multiple paragraphs after a table', () => {
  const doc = DocxDocument.create();
  doc.insertTable([['A']]);
  doc.insertParagraph('tail-1');
  doc.insertParagraph('tail-2');
  doc.deleteParagraph(3);
  assert.deepEqual(doc.getParagraphs().map(p => p.text), ['', 'A', '', 'tail-2']);
});

test('deleteParagraph protects unique table-cell paragraph wrapped by w:sdt', () => {
  const doc = withBody('<w:tbl><w:tr><w:tc><w:sdt><w:sdtPr/><w:sdtContent><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:sdtContent></w:sdt></w:tc></w:tr></w:tbl><w:p/>');
  doc.deleteParagraph(0);
  assert.equal(doc.getBlocks().find(block => block.type === 'table').rows[0].cells[0].blocks[0].paragraph.text, '');
});

test('deleteParagraph protects unique table-cell paragraph wrapped by w:customXml', () => {
  const doc = withBody('<w:tbl><w:tr><w:tc><w:customXml><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:customXml></w:tc></w:tr></w:tbl><w:p/>');
  doc.deleteParagraph(0);
  assert.equal(doc.getBlocks().find(block => block.type === 'table').rows[0].cells[0].blocks[0].paragraph.text, '');
});

test('deleteParagraph keeps normal body deletions working', () => {
  const doc = withBody('<w:p><w:r><w:t>first</w:t></w:r></w:p><w:p><w:r><w:t>second</w:t></w:r></w:p><w:p><w:r><w:t>third</w:t></w:r></w:p>');
  doc.deleteParagraph(1);
  assert.deepEqual(doc.getParagraphs().map(p => p.text), ['first', 'third']);
});

test('paragraph indexes stay table-interleaved and cell deletes use tc protection', () => {
  const doc = withBody('<w:p><w:r><w:t>before</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p><w:r><w:t>after</w:t></w:r></w:p>');
  assert.deepEqual(doc.getParagraphs().map(p => p.text), ['before', 'cell', 'after']);
  doc.deleteParagraph(1);
  assert.deepEqual(doc.getParagraphs().map(p => p.text), ['before', '', 'after']);
});

test('deleteParagraph keeps w:tcPr first when clearing required cell paragraph', () => {
  const doc = withBody('<w:tbl><w:tr><w:tc><w:tcPr><w:tcW w:w="100"/></w:tcPr><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/>');
  doc.deleteParagraph(0);
  assert.match(doc.getPartXml(doc.mainDocumentPath), /<w:tc><w:tcPr>[\s\S]*?<\/w:tcPr><w:p(?:>|\/>)/);
});

test('deleteParagraph clears table separator paragraph between two body tables', () => {
  const doc = withBody('<w:tbl><w:tr><w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p><w:r><w:t>sep</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p><w:r><w:t>tail</w:t></w:r></w:p>');
  doc.deleteParagraph(1);
  assert.deepEqual(doc.getBlocks().map(block => block.type), ['table', 'paragraph', 'table', 'paragraph']);
  assert.deepEqual(doc.getParagraphs().map(p => p.text), ['A', '', 'B', 'tail']);
});

test('deleteParagraph guard works with style+numbering and wrapped table row', () => {
  const doc = DocxDocument.create();
  doc.defineStyle({ id: 'BodyCenter', type: 'paragraph', name: 'BodyCenter', paragraph: { alignment: 'center' } });
  const numId = doc.createNumbering('bullet');
  doc.setParagraphText(0, 'intro');
  doc.formatParagraph(0, { style: 'BodyCenter' });
  doc.setParagraphNumbering(0, numId);
  doc.insertTable([['cell']]);
  let xml = doc.getPartXml(doc.mainDocumentPath);
  const row = xml.match(/<w:tr>[\s\S]*?<\/w:tr>/)?.[0];
  xml = xml.replace(row, `<w:sdt><w:sdtPr/><w:sdtContent>${row}</w:sdtContent></w:sdt>`);
  doc.setPartXml(doc.mainDocumentPath, xml);
  const cellIndex = doc.getBlocks().find(block => block.type === 'table').rows[0].cells[0].blocks[0].paragraph.index;
  doc.deleteParagraph(cellIndex);
  assert.equal(doc.getBlocks().find(block => block.type === 'table').rows[0].cells[0].blocks[0].paragraph.text, '');
  const intro = doc.getParagraphs()[0];
  assert.equal(intro.style, 'BodyCenter');
  assert.equal(intro.numbering?.numId, numId);
});

test('format toggles explicitly disable formatting and retain OOXML property order', () => {
  const doc = DocxDocument.create();
  doc.formatRun(0, 0, { color: 'AABBCC', underline: true, fontSize: 12 });
  doc.formatRun(0, 0, { bold: false, italic: false, underline: false, fontFamily: 'Arial' });
  doc.formatParagraph(0, { alignment: 'right' });
  doc.formatParagraph(0, { style: 'Normal' });
  const xml = doc.getPartXml(doc.mainDocumentPath);
  assert.match(xml, /<w:rFonts[^>]+\/><w:b w:val="0"\/><w:i w:val="0"\/><w:color/);
  assert.match(xml, /<w:pStyle w:val="Normal"\/><w:jc/);
  assert.equal(doc.getParagraphs()[0].runs[0].underline, false);
  assert.throws(() => doc.formatRun(0, 0, { fontSize: NaN }), /fontSize/);
  assert.throws(() => doc.formatRun(0, 0, { color: 'url(evil)' }), /color/);
  assert.throws(() => doc.formatRun(0, 0, { bold: 'true' }), /boolean/);
});

test('OOXML DOM edits are namespace-aware and detached until explicitly committed', () => {
  const doc = DocxDocument.create();
  const detached = doc.getPartDocument(doc.mainDocumentPath);
  detached.documentElement.setAttribute('custom', 'no');
  assert.doesNotMatch(doc.getPartXml(doc.mainDocumentPath), /custom=/);
  doc.updatePartXml(doc.mainDocumentPath, xml => {
    xml.getElementsByTagNameNS(WORD_NS, 'pgSz')[0].setAttributeNS(WORD_NS, 'w:w', '15000');
  });
  assert.match(doc.getPartXml(doc.mainDocumentPath), /w:w="15000"/);
  const before = doc.getPartXml(doc.mainDocumentPath);
  assert.throws(() => doc.updatePartXml(doc.mainDocumentPath, () => { throw new Error('cancel'); }), /cancel/);
  assert.equal(doc.getPartXml(doc.mainDocumentPath), before);
});

test('agent batches are atomic, revision checked and increment once per transaction', () => {
  const doc = DocxDocument.create();
  const result = doc.applyOperations({ expectedRevision: 0, operations: [
    { type: 'setParagraphText', index: 0, text: 'agent' },
    { type: 'insertParagraph', text: 'second' },
    { type: 'formatRun', paragraph: 0, run: 0, format: { bold: true } },
  ] });
  assert.equal(result.revision, 1);
  assert.equal(result.paragraphs[0].text, 'agent');
  assert.throws(() => doc.applyOperations({ expectedRevision: 0, operations: [] }), /Revision conflict/);
  assert.throws(() => doc.applyOperations({ expectedRevision: 1, operations: [
    { type: 'setParagraphText', index: 0, text: 'must rollback' },
    { type: 'deleteParagraph', index: 99 },
  ] }), /does not exist/);
  assert.equal(doc.revision, 1);
  assert.equal(doc.getParagraphs()[0].text, 'agent');
  assert.equal(doc.applyOperations({ operations: [] }).revision, 1);
  assert.equal(AGENT_OPERATION_SCHEMA.properties.operations.items.oneOf.length, 37);
});

test('agent JSON validates unknown methods, shapes and fields without executing code', () => {
  const doc = DocxDocument.create();
  for (const request of [
    null, [], {}, { operations: null }, { operations: [{ type: 'eval', code: 'alert(1)' }] },
    { operations: [{ type: 'deleteParagraph', index: -1 }] },
    { operations: [{ type: 'insertParagraph', text: 3 }] },
    { operations: [{ type: 'formatParagraph', index: 0, format: { alignment: 'evil' } }] },
    { operations: [], extra: 1 },
    { operations: [{ type: 'formatRun', paragraph: 0, run: 0, format: { extra: true } }] },
  ]) assert.throws(() => doc.applyOperations(request));
  assert.equal(doc.revision, 0);
});

test('inline drawing images are parsed and exposed on runs and paragraphs', () => {
  const doc = withImageDoc(
    `<w:p><w:r><w:drawing><wp:inline><wp:extent cx="190500" cy="95250"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="1" name="logo" descr="封面图" title="标题图"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="${PIC_NS}"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="logo"/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="190500" cy="95250"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`,
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>`,
  );
  const paragraph = doc.getParagraphs()[0];
  assert.equal(paragraph.images.length, 1);
  assert.equal(paragraph.runs[0].image?.alt, '封面图');
  assert.equal(paragraph.runs[0].images?.length, 1);
  assert.equal(paragraph.images[0].widthPx, emuToPx(190500));
  assert.deepEqual(doc.getImageBytes(paragraph.images[0]), PNG_BYTES);
  assert.match(doc.getImageDataUrl(paragraph.images[0]), /^data:image\/png;base64,/);
});

test('a run exposes and preserves multiple images in document order', () => {
  const doc = withImageDoc(
    `<w:p><w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="457200"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="1" name="one"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="${PIC_NS}"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="one"/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing><w:drawing><wp:inline><wp:extent cx="457200" cy="457200"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="2" name="two"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="${PIC_NS}"><pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="two"/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="457200" cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`,
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image2.gif"/>`,
    [{ path: 'word/media/image1.png', bytes: PNG_BYTES, type: 'image/png' }, { path: 'word/media/image2.gif', bytes: GIF_BYTES, type: 'image/gif' }],
  );
  const run = doc.getParagraphs()[0].runs[0];
  assert.deepEqual(run.images?.map((image) => image.relationshipId), ['rId1', 'rId2']);
  assert.deepEqual(doc.getImages().map((image) => image.relationshipId), ['rId1', 'rId2']);
});

test('floating drawing metadata such as wrap, rotation, flip and crop is parsed', () => {
  const doc = withImageDoc(
    `<w:p><w:r><w:drawing><wp:anchor behindDoc="1"><wp:extent cx="952500" cy="476250"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapSquare wrapText="bothSides"/><wp:docPr id="2" name="hero" descr="海报"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="${PIC_NS}"><pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="hero"/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId1"/><a:srcRect l="10000" t="0" r="2500" b="5000"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm rot="5400000" flipH="1" flipV="true"><a:off x="0" y="0"/><a:ext cx="952500" cy="476250"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>`,
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>`,
  );
  const image = doc.getImages()[0];
  assert.equal(image.placement, 'floating');
  assert.equal(image.wrap, 'square');
  assert.equal(image.rotation, 90);
  assert.equal(image.flipH, true);
  assert.equal(image.flipV, true);
  assert.equal(image.crop.left, 0.1);
  assert.equal(image.crop.right, 0.025);
  assert.equal(image.behindDoc, true);
});

test('VML images are parsed for display', () => {
  const doc = withImageDoc(
    `<w:p><w:r><w:pict><v:shape id="shape1" style="width:48pt;height:24pt"><v:imagedata r:id="rId1"/></v:shape></w:pict></w:r></w:p>`,
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>`,
  );
  const image = doc.getImages()[0];
  assert.equal(Math.round(image.widthPx), 64);
  assert.equal(Math.round(image.heightPx), 32);
});

test('external linked images are not loaded and use placeholders', () => {
  const doc = withImageDoc(
    `<w:p><w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="457200"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="3" name="remote"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="${PIC_NS}"><pic:pic><pic:nvPicPr><pic:cNvPr id="2" name="remote"/></pic:nvPicPr><pic:blipFill><a:blip r:link="rId9"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`,
    `<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="https://example.com/remote.png" TargetMode="External"/>`,
    [],
  );
  const image = doc.getImages()[0];
  assert.equal(image.isExternal, true);
  assert.match(doc.getImageDataUrl(image), /^data:image\/svg\+xml;base64,/);
  assert.throws(() => doc.getImageBytes(image), /External/);
  assert.doesNotMatch(doc.getImageDataUrl(image), /example\.com/);
});

test('unsupported but known image types get placeholders instead of errors', () => {
  const doc = withImageDoc(
    `<w:p><w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="457200"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="1" name="vector"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="${PIC_NS}"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="vector"/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`,
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.emf"/>`,
    [{ path: 'word/media/image1.emf', bytes: Uint8Array.from([1, 2, 3]), type: 'image/x-emf' }],
  );
  assert.match(doc.getImageDataUrl(doc.getImages()[0]), /^data:image\/svg\+xml;base64,/);
});

test('insertImage creates media, relationships, content type entries and round-trips', async () => {
  const doc = DocxDocument.create();
  const image = doc.insertImage({ bytes: PNG_BYTES, contentType: 'image/png', alt: '徽标' });
  assert.equal(image.alt, '徽标');
  assert.equal(doc.getImages().length, 1);
  assert.match(doc.getPartXml('word/_rels/document.xml.rels'), /Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/image"/);
  assert.match(doc.getPartXml(doc.mainDocumentPath), /wp:inline/);
  assert.deepEqual((await DocxDocument.load(await doc.toUint8Array())).getImageBytes(image), PNG_BYTES);
});

test('floating insertImage writes anchor attributes required by Word', () => {
  const doc = DocxDocument.create();
  doc.insertImage({ bytes: PNG_BYTES, contentType: 'image/png', placement: 'floating' });
  const xml = doc.getPartXml(doc.mainDocumentPath);
  assert.match(xml, /<wp:anchor[^>]*locked="0"/);
  assert.match(xml, /<wp:anchor[^>]*layoutInCell="1"/);
  assert.match(xml, /<wp:simplePos x="0" y="0"\/>/);
});

test('insertImage before a nested run uses the paragraph child container instead of throwing DOMException', () => {
  const doc = withBody(`<w:p xmlns:r="${OFFICE_REL_NS}"><w:hyperlink r:id="rIdX"><w:r><w:t>link</w:t></w:r></w:hyperlink></w:p>`);
  const image = doc.insertImage({ bytes: PNG_BYTES, contentType: 'image/png', paragraph: 0, run: 0 });
  assert.equal(doc.getImages().length, 1);
  assert.equal(image.run, 0);
});

test('insertImage infers intrinsic dimensions from PNG bytes when size is omitted', () => {
  const doc = DocxDocument.create();
  const image = doc.insertImage({ bytes: PNG_BYTES, contentType: 'image/png' });
  assert.equal(Math.round(image.widthPx), 2);
  assert.equal(Math.round(image.heightPx), 1);
});

test('insertImage via applyOperations decodes strict base64 and remains atomic on failure', () => {
  const doc = DocxDocument.create();
  const bytes = dataUrlForBytes(GIF_BYTES, 'image/gif').split(',')[1];
  const snapshot = doc.applyOperations({ operations: [{ type: 'insertImage', bytes, contentType: 'image/gif', alt: '动图' }] });
  assert.equal(snapshot.revision, 1);
  assert.equal(doc.getImages()[0].contentType, 'image/gif');
  assert.throws(() => doc.applyOperations({ expectedRevision: 1, operations: [{ type: 'insertImage', bytes: '***', contentType: 'image/png' }] }), /base64/);
  assert.equal(doc.revision, 1);
  assert.equal(doc.getImages().length, 1);
});

test('replaceImageBytes swaps image bytes without breaking the relationship', () => {
  const doc = DocxDocument.create();
  const image = doc.insertImage({ bytes: PNG_BYTES, contentType: 'image/png' });
  doc.replaceImageBytes(image, GIF_BYTES, 'image/gif');
  const updated = doc.getImages()[0];
  assert.deepEqual(doc.getImageBytes(image.relationshipId), GIF_BYTES);
  assert.match(doc.getImageDataUrl(image.relationshipId), /^data:image\/gif;base64,/);
  assert.match(updated.partPath, /\.gif$/);
  assert.match(doc.getPartXml('word/_rels/document.xml.rels'), /image\d+\.gif/);
});

test('replaceImageBytes infers content type from bytes when omitted', () => {
  const doc = DocxDocument.create();
  const image = doc.insertImage({ bytes: PNG_BYTES, contentType: 'image/png' });
  doc.replaceImageBytes(image, GIF_BYTES);
  const updated = doc.getImages()[0];
  assert.equal(updated.contentType, 'image/gif');
  assert.match(updated.partPath, /\.gif$/);
});

test('getImageDataUrl cache invalidates after replaceImageBytes', () => {
  const doc = DocxDocument.create();
  const image = doc.insertImage({ bytes: PNG_BYTES, contentType: 'image/png' });
  const before = doc.getImageDataUrl(image);
  doc.replaceImageBytes(image, GIF_BYTES);
  const after = doc.getImageDataUrl(doc.getImages()[0]);
  assert.notEqual(after, before);
  assert.match(after, /^data:image\/gif;base64,/);
});

test('resizeImage updates stored extents and can keep aspect ratio', () => {
  const doc = DocxDocument.create();
  const image = doc.insertImage({ bytes: PNG_BYTES, contentType: 'image/png' });
  doc.resizeImage(image, { widthEmu: pxToEmu(20), keepAspect: true });
  const resized = doc.getImages()[0];
  assert.equal(Math.round(resized.widthPx), 20);
  assert.equal(Math.round(resized.heightPx), 10);
});

test('setImageAlt updates alt and title metadata', () => {
  const doc = DocxDocument.create();
  const image = doc.insertImage({ bytes: PNG_BYTES, contentType: 'image/png' });
  doc.setImageAlt(image, '替代文本', '标题');
  const updated = doc.getImages()[0];
  assert.equal(updated.alt, '替代文本');
  assert.equal(updated.title, '标题');
  assert.match(doc.getPartXml(doc.mainDocumentPath), /descr="替代文本"/);
});

test('VML images can be resized and retitled instead of silently no-oping', () => {
  const doc = withImageDoc(
    `<w:p><w:r><w:pict><v:shape id="shape1" style="width:48pt;height:24pt"><v:imagedata r:id="rId1"/></v:shape></w:pict></w:r></w:p>`,
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>`,
  );
  const image = doc.getImages()[0];
  const revision = doc.revision;
  doc.resizeImage(image, { widthEmu: pxToEmu(300), heightEmu: pxToEmu(150) });
  doc.setImageAlt(image, 'new alt', 'new title');
  const updated = doc.getImages()[0];
  assert.equal(Math.round(updated.widthPx), 300);
  assert.equal(Math.round(updated.heightPx), 150);
  assert.equal(updated.alt, 'new alt');
  assert.match(doc.getPartXml(doc.mainDocumentPath), /width:225pt/);
  assert.match(doc.getPartXml(doc.mainDocumentPath), /height:112\.5pt/);
  assert.match(doc.getPartXml(doc.mainDocumentPath), /alt="new alt"/);
  assert.equal(doc.revision, revision + 2);
});

test('deleteImage removes its drawing, empty run and unique media part', () => {
  const doc = DocxDocument.create();
  const image = doc.insertImage({ bytes: PNG_BYTES, contentType: 'image/png' });
  doc.deleteImage(image);
  assert.equal(doc.getImages().length, 0);
  assert.throws(() => doc.getPartBytes('word/media/image1.png'));
  assert.doesNotMatch(doc.getPartXml(doc.mainDocumentPath), /w:drawing/);
});

test('deleteImage keeps the media part when another relationship still targets it', () => {
  const doc = withImageDoc(
    `<w:p><w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="457200"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="1" name="one"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="${PIC_NS}"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="one"/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r><w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="457200"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="2" name="two"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="${PIC_NS}"><pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="two"/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`,
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/shared.png"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/shared.png"/>`,
    [{ path: 'word/media/shared.png', bytes: PNG_BYTES, type: 'image/png' }],
  );
  doc.deleteImage(doc.getImages()[0]);
  assert.deepEqual(doc.getPartBytes('word/media/shared.png'), PNG_BYTES);
  assert.equal(doc.getImages().length, 1);
});

test('images sharing a relationship within one run remain individually addressable', () => {
  const drawing = `<w:drawing><wp:inline><wp:extent cx="914400" cy="457200"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="1" name="one"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="${PIC_NS}"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="one"/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>`;
  const doc = withImageDoc(
    `<w:p><w:r>${drawing}${drawing}</w:r></w:p>`,
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>`,
  );
  const [first, second] = doc.getImages();
  assert.notEqual(first.ordinal, second.ordinal);
  doc.deleteImage(second);
  assert.equal(doc.getImages().length, 1);
  assert.equal((doc.getPartXml(doc.mainDocumentPath).match(/<w:drawing>/g) ?? []).length, 1);
});

test('images with different relationships and mixed drawing order remain individually addressable', () => {
  const doc = withImageDoc(
    `<w:p><w:r><w:pict><v:shape style="width:24pt;height:24pt"><v:imagedata r:id="rId1"/></v:shape></w:pict><w:drawing><wp:inline><wp:extent cx="914400" cy="457200"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="2" name="two"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="${PIC_NS}"><pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="two"/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`,
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image2.gif"/>`,
    [{ path: 'word/media/image1.png', bytes: PNG_BYTES, type: 'image/png' }, { path: 'word/media/image2.gif', bytes: GIF_BYTES, type: 'image/gif' }],
  );
  const [, second] = doc.getImages();
  doc.setImageAlt(second, 'second');
  doc.resizeImage(second, { widthEmu: pxToEmu(120), keepAspect: true });
  assert.equal(Math.round(doc.getImages()[1].widthPx), 120);
  assert.match(doc.getPartXml(doc.mainDocumentPath), /descr="second"/);
});

test('broken relationships, missing media parts and invalid extents do not crash image reads', () => {
  const doc = withImageDoc(
    `<w:p><w:r><w:drawing><wp:inline><wp:extent cx="-1" cy="0"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="1" name="broken"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="${PIC_NS}"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="broken"/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rMissing"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="-1" cy="0"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`,
    '',
    [],
  );
  const image = doc.getImages()[0];
  assert.equal(image.widthEmu, 0);
  assert.equal(image.heightEmu, 0);
  assert.match(doc.getImageDataUrl(image), /^data:image\/svg\+xml;base64,/);
  assert.throws(() => doc.getImageBytes(image), /not found/);
});

test('operations schema includes the image operations', () => {
  assert.equal(AGENT_OPERATION_SCHEMA.properties.operations.items.oneOf.length, 37);
  const resize = AGENT_OPERATION_SCHEMA.properties.operations.items.oneOf.find((entry) => entry.properties.type.const === 'resizeImage');
  assert.equal(resize.properties.size.anyOf.length, 2);
});

test('malformed XML, DTD, broken package targets and oversized parts are rejected atomically', async () => {
  const doc = DocxDocument.create();
  const original = doc.getPartXml(doc.mainDocumentPath);
  for (const xml of ['<unclosed>', '<root/>', '<!DOCTYPE r [<!ENTITY e SYSTEM "file:///etc/passwd">]><r>&e;</r>']) {
    assert.throws(() => doc.setPartXml(doc.mainDocumentPath, xml));
    assert.equal(doc.getPartXml(doc.mainDocumentPath), original);
  }
  assert.throws(() => doc.setPartXml('_rels/.rels', doc.getPartXml('_rels/.rels').replace('word/document.xml', '../secret.xml')));
  assert.throws(() => doc.addPart('../bad.xml', new Uint8Array(), 'application/xml'), /path/);
  assert.throws(() => doc.addPart('large.bin', new Uint8Array(16 * 1024 * 1024 + 1), 'application/octet-stream'), /size/);
  assert.throws(() => doc.setPartBytes('[Content_Types].xml', new TextEncoder().encode('<Types/>')));
  await assert.rejects(() => DocxDocument.load(new Uint8Array([1, 2, 3])));
  assert.equal(doc.revision, 0);
});

test('setPartXml rejects malformed external XML input', () => {
  const doc = DocxDocument.create();
  const before = doc.getPartXml(doc.mainDocumentPath);
  assert.throws(() => doc.setPartXml(doc.mainDocumentPath, '<w:document>'));
  assert.equal(doc.getPartXml(doc.mainDocumentPath), before);
  assert.equal(doc.revision, 0);
});

test('updatePartXml rejects malformed DOM output and rolls back atomically', () => {
  const doc = DocxDocument.create();
  const before = doc.getPartXml(doc.mainDocumentPath);
  assert.throws(() => doc.updatePartXml(doc.mainDocumentPath, (document) => {
    const body = document.getElementsByTagNameNS(WORD_NS, 'body')[0];
    body.appendChild(document.createComment('a--b'));
  }), /Invalid XML/);
  assert.equal(doc.getPartXml(doc.mainDocumentPath), before);
  assert.equal(doc.revision, 0);
});

test('nested setPartXml in updatePartXml keeps outer callback edits', () => {
  const doc = DocxDocument.create();
  const main = doc.mainDocumentPath;
  doc.updatePartXml(main, (document) => {
    const body = document.getElementsByTagNameNS(WORD_NS, 'body')[0];
    const paragraph = document.createElementNS(WORD_NS, 'w:p');
    const run = document.createElementNS(WORD_NS, 'w:r');
    const text = document.createElementNS(WORD_NS, 'w:t');
    text.appendChild(document.createTextNode('OUTER'));
    run.appendChild(text);
    paragraph.appendChild(run);
    body.insertBefore(paragraph, body.getElementsByTagNameNS(WORD_NS, 'sectPr')[0] ?? null);
    doc.setPartXml(main, `<w:document xmlns:w="${WORD_NS}"><w:body><w:p><w:r><w:t>NESTED</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`);
  });
  assert.equal(doc.revision, 2);
  assert.ok(doc.getParagraphs().some((paragraph) => paragraph.text === 'OUTER'));
  assert.doesNotMatch(doc.getPartXml(main), /NESTED/);
});

test('performance regression: single-op loops and batched ops stay within acceptance thresholds', () => {
  const elapsed = (run) => {
    const start = process.hrtime.bigint();
    run();
    return Number(process.hrtime.bigint() - start) / 1e6;
  };
  const single = DocxDocument.create();
  const insertMs = elapsed(() => { for (let i = 0; i < 300; i++) single.insertParagraph(`段落内容 ${i}`); });
  const setMs = elapsed(() => { for (let i = 0; i < 200; i++) single.setParagraphText(i, `改写 ${i}`); });
  const batch = DocxDocument.create();
  const batchMs = elapsed(() => batch.applyOperations({
    operations: Array.from({ length: 300 }, (_, i) => ({ type: 'insertParagraph', text: `x${i}` })),
  }));
  assert.ok(insertMs < 1000, `300 insert took ${insertMs.toFixed(1)}ms`);
  assert.ok(setMs < 1000, `200 set took ${setMs.toFixed(1)}ms`);
  assert.ok(batchMs < 500, `batch 300 ops took ${batchMs.toFixed(1)}ms`);
});

test('performance regression: repeated single inserts should not grow quadratically', () => {
  const measure = (count) => {
    const doc = DocxDocument.create();
    const start = process.hrtime.bigint();
    for (let i = 0; i < count; i++) doc.insertParagraph(`p${i}`);
    return Number(process.hrtime.bigint() - start) / 1e6;
  };
  const t600 = measure(600);
  assert.ok(t600 < 1500, `insert performance regressed: 600=${t600.toFixed(1)}ms`);
});

test('performance regression: 1000 paragraph document handles 1000 operations quickly', () => {
  const doc = DocxDocument.create();
  doc.applyOperations({
    operations: Array.from({ length: 999 }, (_, i) => ({ type: 'insertParagraph', text: `seed-${i}` })),
  });
  const start = process.hrtime.bigint();
  doc.applyOperations({
    operations: Array.from({ length: 1000 }, (_, i) => ({ type: 'setParagraphText', index: i, text: `更新-${i}` })),
  });
  const elapsedMs = Number(process.hrtime.bigint() - start) / 1e6;
  assert.ok(elapsedMs < 5000, `1000 ops took ${elapsedMs.toFixed(1)}ms`);
});

test('integration: table cell paragraph supports style, numbering and image together', () => {
  const doc = DocxDocument.create();
  doc.defineStyle({ id: 'CellStyle', type: 'paragraph', name: 'CellStyle', paragraph: { alignment: 'center' }, run: { bold: true } });
  const numId = doc.createNumbering('decimal');
  doc.insertTable([['cell']]);
  const table = doc.getBlocks().find((block) => block.type === 'table');
  const paragraphIndex = table.rows[0].cells[0].blocks[0].paragraph.index;
  doc.formatParagraph(paragraphIndex, { style: 'CellStyle' });
  doc.setParagraphNumbering(paragraphIndex, numId, 0);
  doc.insertImage({ bytes: PNG_BYTES, contentType: 'image/png', paragraph: paragraphIndex, alt: 'cell-image' });
  const paragraph = doc.getParagraphs()[paragraphIndex];
  assert.equal(paragraph.style, 'CellStyle');
  assert.equal(paragraph.numbering?.numId, numId);
  assert.ok(paragraph.images.length >= 1);
  assert.equal(doc.getBlocks().find((block) => block.type === 'table').rows[0].cells[0].blocks[0].paragraph.images.length >= 1, true);
});

test('ZIP traversal and decompression bombs are bounded', async () => {
  const base = await DocxDocument.create().toUint8Array();
  const zip = await JSZip.loadAsync(base);
  zip.file('../bad.xml', '<bad/>', { createFolders: false });
  await assert.rejects(() => zip.generateAsync({ type: 'uint8array' }).then(DocxDocument.load), /Unsafe ZIP/);
  zip.remove('../bad.xml');
  zip.file('large.bin', new Uint8Array(16 * 1024 * 1024 + 1));
  await assert.rejects(() => zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }).then(DocxDocument.load), /size limit/);
});

test('main document relationship is resolved rather than hard-coded', async () => {
  const doc = DocxDocument.create();
  doc.setParagraphText(0, 'alternate path');
  const zip = await JSZip.loadAsync(await doc.toUint8Array());
  zip.file('custom/main.xml', doc.getPartBytes(doc.mainDocumentPath));
  zip.remove(doc.mainDocumentPath);
  zip.file('_rels/.rels', doc.getPartXml('_rels/.rels').replace('word/document.xml', 'custom/main.xml'));
  zip.file('[Content_Types].xml', doc.getPartXml('[Content_Types].xml').replace('/word/document.xml', '/custom/main.xml'));
  const loaded = await DocxDocument.load(await zip.generateAsync({ type: 'uint8array' }));
  assert.equal(loaded.mainDocumentPath, 'custom/main.xml');
  assert.equal(loaded.getParagraphs()[0].text, 'alternate path');
});

test('UTF-16 XML input is decoded and edited output declares UTF-8', async () => {
  const doc = DocxDocument.create();
  const xml = `<?xml version="1.0" encoding="UTF-16"?>${doc.getPartXml(doc.mainDocumentPath)}`;
  const bytes = Buffer.concat([Buffer.from([0xff, 0xfe]), Buffer.from(xml, 'utf16le')]);
  doc.setPartBytes(doc.mainDocumentPath, bytes);
  doc.setParagraphText(0, '编码');
  assert.match(doc.getPartXml(doc.mainDocumentPath), /encoding="UTF-8"/);
  assert.equal((await DocxDocument.load(await doc.toUint8Array())).getParagraphs()[0].text, '编码');
});

test('snapshot exposes parsed styles and default effective formatting', () => {
  const doc = withStyles(
    '<w:p><w:r><w:t>Styled</w:t></w:r></w:p>',
    `<w:styles xmlns:w="${WORD_NS}">
      <w:docDefaults>
        <w:pPrDefault><w:pPr><w:spacing w:after="120"/></w:pPr></w:pPrDefault>
        <w:rPrDefault><w:rPr><w:sz w:val="22"/><w:color w:val="112233"/></w:rPr></w:rPrDefault>
      </w:docDefaults>
      <w:style w:type="paragraph" w:styleId="Normal" w:default="1"><w:name w:val="Normal"/><w:qFormat/></w:style>
      <w:style w:type="character" w:styleId="DefaultParagraphFont" w:default="1"><w:name w:val="Default Paragraph Font"/></w:style>
    </w:styles>`,
  );
  const snapshot = doc.getSnapshot();
  assert.equal(snapshot.styles[0].id, 'Normal');
  assert.equal(snapshot.paragraphs[0].effective.spacingAfter, 120);
  assert.equal(snapshot.paragraphs[0].runs[0].effective.fontSize, 11);
  assert.equal(snapshot.paragraphs[0].runs[0].effective.color, '112233');
});

test('paragraph basedOn chains merge from root to leaf', () => {
  const doc = withStyles(
    '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Hello</w:t></w:r></w:p>',
    `<w:styles xmlns:w="${WORD_NS}">
      <w:style w:type="paragraph" w:styleId="Base"><w:name w:val="Base"/><w:rPr><w:sz w:val="28"/><w:color w:val="224488"/></w:rPr><w:pPr><w:spacing w:after="160"/></w:pPr></w:style>
      <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="Heading 1"/><w:basedOn w:val="Base"/><w:rPr><w:b/></w:rPr><w:pPr><w:jc w:val="center"/></w:pPr></w:style>
    </w:styles>`,
  );
  const paragraph = doc.getParagraphs()[0];
  assert.equal(paragraph.effective.alignment, 'center');
  assert.equal(paragraph.effective.spacingAfter, 160);
  assert.equal(paragraph.runs[0].effective.bold, true);
  assert.equal(paragraph.runs[0].effective.fontSize, 14);
  assert.equal(paragraph.runs[0].effective.color, '224488');
});

test('basedOn cycles are truncated without throwing', () => {
  const doc = withStyles(
    '<w:p><w:pPr><w:pStyle w:val="A"/></w:pPr><w:r><w:t>Loop</w:t></w:r></w:p>',
    `<w:styles xmlns:w="${WORD_NS}">
      <w:style w:type="paragraph" w:styleId="A"><w:name w:val="A"/><w:basedOn w:val="B"/><w:rPr><w:b/></w:rPr></w:style>
      <w:style w:type="paragraph" w:styleId="B"><w:name w:val="B"/><w:basedOn w:val="A"/><w:rPr><w:i/></w:rPr></w:style>
    </w:styles>`,
  );
  const run = doc.getParagraphs()[0].runs[0];
  assert.equal(run.effective.bold, true);
  assert.equal(run.effective.italic, true);
});

test('style toggle off overrides inherited bold formatting', () => {
  const doc = withStyles(
    '<w:p><w:pPr><w:pStyle w:val="Child"/></w:pPr><w:r><w:t>Off</w:t></w:r></w:p>',
    `<w:styles xmlns:w="${WORD_NS}">
      <w:style w:type="paragraph" w:styleId="Base"><w:name w:val="Base"/><w:rPr><w:b/></w:rPr></w:style>
      <w:style w:type="paragraph" w:styleId="Child"><w:name w:val="Child"/><w:basedOn w:val="Base"/><w:rPr><w:b w:val="0"/></w:rPr></w:style>
    </w:styles>`,
  );
  assert.equal(doc.getParagraphs()[0].runs[0].effective.bold, false);
});

test('character style chains merge onto runs', () => {
  const doc = withStyles(
    '<w:p><w:r><w:rPr><w:rStyle w:val="Emphasis"/></w:rPr><w:t>Chain</w:t></w:r></w:p>',
    `<w:styles xmlns:w="${WORD_NS}">
      <w:style w:type="character" w:styleId="Strong"><w:name w:val="Strong"/><w:rPr><w:b/></w:rPr></w:style>
      <w:style w:type="character" w:styleId="Emphasis"><w:name w:val="Emphasis"/><w:basedOn w:val="Strong"/><w:rPr><w:i/></w:rPr></w:style>
    </w:styles>`,
  );
  const run = doc.getParagraphs()[0].runs[0];
  assert.equal(run.style, 'Emphasis');
  assert.equal(run.effective.bold, true);
  assert.equal(run.effective.italic, true);
});

test('theme fonts and shaded theme colors are resolved from theme1.xml', () => {
  const doc = withStyles(
    '<w:p><w:pPr><w:pStyle w:val="ThemeStyle"/></w:pPr><w:r><w:t>Theme</w:t></w:r></w:p>',
    `<w:styles xmlns:w="${WORD_NS}">
      <w:style w:type="paragraph" w:styleId="ThemeStyle"><w:name w:val="Theme Style"/><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia"/><w:color w:themeColor="accent1" w:themeShade="80"/></w:rPr></w:style>
    </w:styles>`,
    `<?xml version="1.0" encoding="UTF-8"?>
    <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <a:themeElements>
        <a:clrScheme name="Custom"><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="111111"/></a:dk2><a:lt2><a:srgbClr val="EEEEEE"/></a:lt2><a:accent1><a:srgbClr val="FF0000"/></a:accent1><a:accent2><a:srgbClr val="00FF00"/></a:accent2><a:accent3><a:srgbClr val="0000FF"/></a:accent3><a:accent4><a:srgbClr val="888888"/></a:accent4><a:accent5><a:srgbClr val="999999"/></a:accent5><a:accent6><a:srgbClr val="AAAAAA"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme>
        <a:fontScheme name="Custom"><a:majorFont><a:latin typeface="Major Font"/><a:ea typeface="Major EA"/><a:cs typeface="Major CS"/></a:majorFont><a:minorFont><a:latin typeface="Minor Font"/><a:ea typeface="Minor EA"/><a:cs typeface="Minor CS"/></a:minorFont></a:fontScheme>
      </a:themeElements>
    </a:theme>`,
  );
  const run = doc.getParagraphs()[0].runs[0];
  assert.equal(run.effective.fontFamily, 'Minor Font');
  assert.equal(run.effective.fontFamilyEastAsia, 'Minor EA');
  assert.equal(run.effective.color, '800000');
});

test('theme tint moves colors toward white with Word-compatible math', () => {
  const doc = withStyles(
    '<w:p><w:pPr><w:pStyle w:val="Tinted"/></w:pPr><w:r><w:t>Tint</w:t></w:r></w:p>',
    `<w:styles xmlns:w="${WORD_NS}">
      <w:style w:type="paragraph" w:styleId="Tinted"><w:name w:val="Tinted"/><w:rPr><w:color w:themeColor="accent1" w:themeTint="99"/></w:rPr></w:style>
    </w:styles>`,
  );
  assert.equal(doc.getParagraphs()[0].runs[0].effective.color, '8EAADB');
});

test('non-table paragraphs do not inherit the default table style', () => {
  const doc = withStyles(
    '<w:p><w:r><w:t>Plain</w:t></w:r></w:p>',
    `<w:styles xmlns:w="${WORD_NS}">
      <w:style w:type="table" w:default="1" w:styleId="TableNormal">
        <w:name w:val="Table Normal"/>
        <w:pPr><w:jc w:val="center"/></w:pPr>
        <w:rPr><w:color w:val="FF0000"/><w:sz w:val="40"/></w:rPr>
      </w:style>
    </w:styles>`,
  );
  const paragraph = doc.getParagraphs()[0];
  assert.equal(paragraph.effective.alignment, undefined);
  assert.equal(paragraph.runs[0].effective.color, undefined);
  assert.equal(paragraph.runs[0].effective.fontSize, undefined);
});

test('missing theme1.xml falls back to built-in theme tables', () => {
  const doc = withStyles(
    '<w:p><w:pPr><w:pStyle w:val="FallbackTheme"/></w:pPr><w:r><w:t>Fallback</w:t></w:r></w:p>',
    `<w:styles xmlns:w="${WORD_NS}">
      <w:style w:type="paragraph" w:styleId="FallbackTheme"><w:name w:val="Fallback Theme"/><w:rPr><w:rFonts w:asciiTheme="minorHAnsi"/><w:color w:themeColor="accent1"/></w:rPr></w:style>
    </w:styles>`,
  );
  const run = doc.getParagraphs()[0].runs[0];
  assert.equal(run.effective.fontFamily, 'Calibri');
  assert.equal(run.effective.color, '4472C4');
});

test('unknown or missing styles do not throw, but optional validation can reject them', () => {
  const doc = DocxDocument.create();
  doc.formatParagraph(0, { style: 'MissingStyle' });
  assert.equal(doc.getParagraphs()[0].style, 'MissingStyle');
  assert.doesNotThrow(() => doc.getEffectiveParagraphFormat(0));
  assert.throws(() => doc.formatParagraph(0, { style: 'StillMissing' }, { validateStyle: true }), /style not found/i);
});

test('getEffectiveParagraphFormat and getEffectiveRunFormat expose merged values', () => {
  const doc = withStyles(
    '<w:p><w:pPr><w:pStyle w:val="Styled"/></w:pPr><w:r><w:t>APIs</w:t></w:r></w:p>',
    `<w:styles xmlns:w="${WORD_NS}">
      <w:style w:type="paragraph" w:styleId="Styled"><w:name w:val="Styled"/><w:pPr><w:spacing w:before="120" w:after="240"/><w:jc w:val="distribute"/></w:pPr><w:rPr><w:highlight w:val="yellow"/></w:rPr></w:style>
    </w:styles>`,
  );
  assert.equal(doc.getEffectiveParagraphFormat(0).alignment, 'distribute');
  assert.equal(doc.getEffectiveParagraphFormat(0).spacingAfter, 240);
  assert.equal(doc.getEffectiveRunFormat(0, 0).highlight, 'yellow');
});

test('read-only round trips preserve styles.xml bytes exactly', async () => {
  const stylesXml = `<?xml version="1.0" encoding="UTF-8"?><w:styles xmlns:w="${WORD_NS}">\n  <w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style>\n</w:styles>`;
  const doc = withStyles('<w:p><w:r><w:t>Keep</w:t></w:r></w:p>', stylesXml);
  const loaded = await DocxDocument.load(await doc.toUint8Array());
  loaded.setParagraphText(0, 'Changed');
  const reopened = await DocxDocument.load(await loaded.toUint8Array());
  assert.deepEqual(reopened.getPartBytes('word/styles.xml'), new TextEncoder().encode(stylesXml));
});

test('defineStyle creates styles.xml and styles can be read back after reload', async () => {
  const doc = DocxDocument.create();
  doc.defineStyle({
    id: 'MyHeading',
    name: 'My Heading',
    type: 'paragraph',
    quickFormat: true,
    paragraph: { spacingAfter: 240, alignment: 'center' },
    run: { bold: true, fontSize: 18, color: '3355AA' },
  });
  doc.formatParagraph(0, { style: 'MyHeading' }, { validateStyle: true });
  const reopened = await DocxDocument.load(await doc.toUint8Array());
  assert.equal(reopened.getStyle('MyHeading').name, 'My Heading');
  assert.equal(reopened.getParagraphs()[0].effective.alignment, 'center');
  assert.equal(reopened.getParagraphs()[0].runs[0].effective.bold, true);
  assert.match(reopened.getPartXml('word/styles.xml'), /MyHeading/);
});

test('defineStyle preserves CT_Style child order when updating an existing style', () => {
  const doc = withStyles(
    '<w:p><w:r><w:t>Heading</w:t></w:r></w:p>',
    `<w:styles xmlns:w="${WORD_NS}">
      <w:style w:type="paragraph" w:styleId="Heading1">
        <w:name w:val="heading 1"/>
        <w:uiPriority w:val="9"/>
        <w:semiHidden/>
        <w:unhideWhenUsed/>
      </w:style>
    </w:styles>`,
  );
  doc.defineStyle({
    id: 'Heading1',
    name: 'heading 1',
    type: 'paragraph',
    basedOn: 'Normal',
    aliases: ['H1'],
    quickFormat: true,
    run: { fontSize: 18 },
  });
  const xml = doc.getPartXml('word/styles.xml');
  const order = ['<w:name', '<w:aliases', '<w:basedOn', '<w:uiPriority', '<w:semiHidden', '<w:unhideWhenUsed', '<w:qFormat', '<w:rPr']
    .map((token) => xml.indexOf(token));
  assert.deepEqual(order.slice().sort((a, b) => a - b), order);
});

test('defineStyle picks the first unused relationship id for styles.xml', () => {
  const doc = DocxDocument.create();
  doc.addPart('word/_rels/document.xml.rels', new TextEncoder().encode(
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="urn:one" Target="settings.xml"/>
      <Relationship Id="rId3" Type="urn:three" Target="webSettings.xml"/>
    </Relationships>`,
  ), 'application/vnd.openxmlformats-package.relationships+xml');
  doc.defineStyle({ id: 'Gap', name: 'Gap', type: 'paragraph' });
  const rels = doc.getPartXml('word/_rels/document.xml.rels');
  assert.match(rels, /Id="rId2"[^>]+styles\.xml/);
  assert.equal((rels.match(/Id="rId3"/g) ?? []).length, 1);
});

test('format APIs can clear direct formatting with null to fall back to inherited values', () => {
  const doc = withStyles(
    '<w:p><w:pPr><w:pStyle w:val="Styled"/></w:pPr><w:r><w:t>Clear</w:t></w:r></w:p>',
    `<w:styles xmlns:w="${WORD_NS}">
      <w:style w:type="paragraph" w:styleId="Styled"><w:name w:val="Styled"/><w:pPr><w:spacing w:after="240"/></w:pPr><w:rPr><w:color w:val="336699"/></w:rPr></w:style>
    </w:styles>`,
  );
  doc.formatParagraph(0, { spacingAfter: 120 });
  doc.formatRun(0, 0, { color: 'AA5500', highlight: 'yellow', fontFamily: 'Arial' });
  doc.formatParagraph(0, { spacingAfter: null });
  doc.formatRun(0, 0, { color: null, highlight: 'none', fontFamily: null });
  const xml = doc.getPartXml(doc.mainDocumentPath);
  assert.doesNotMatch(xml, /w:after="120"/);
  assert.doesNotMatch(xml, /w:highlight/);
  assert.doesNotMatch(xml, /w:rFonts/);
  assert.equal(doc.getParagraphs()[0].effective.spacingAfter, 240);
  assert.equal(doc.getParagraphs()[0].runs[0].effective.color, '336699');
});

test('clearing paragraph indents removes imported start/end attributes too', () => {
  const doc = withBody('<w:p><w:pPr><w:ind w:start="720" w:end="360"/></w:pPr><w:r><w:t>Indent</w:t></w:r></w:p>');
  assert.equal(doc.getParagraphs()[0].indentLeft, 720);
  assert.equal(doc.getParagraphs()[0].indentRight, 360);
  doc.formatParagraph(0, { indentLeft: null, indentRight: null });
  const xml = doc.getPartXml(doc.mainDocumentPath);
  assert.doesNotMatch(xml, /w:start=/);
  assert.doesNotMatch(xml, /w:end=/);
  assert.equal(doc.getParagraphs()[0].indentLeft, undefined);
  assert.equal(doc.getParagraphs()[0].indentRight, undefined);
});

test('setting paragraph indents rewrites imported start/end attributes', () => {
  const doc = withBody('<w:p><w:pPr><w:ind w:start="720" w:end="360"/></w:pPr><w:r><w:t>Indent</w:t></w:r></w:p>');
  doc.formatParagraph(0, { indentLeft: 100, indentRight: 50 });
  const xml = doc.getPartXml(doc.mainDocumentPath);
  assert.doesNotMatch(xml, /w:start=/);
  assert.doesNotMatch(xml, /w:end=/);
  assert.match(xml, /w:left="100"/);
  assert.match(xml, /w:right="50"/);
});

test('eastAsia-only font families are surfaced in effective formatting', () => {
  const doc = withBody('<w:p><w:r><w:rPr><w:rFonts w:eastAsia="SimSun"/></w:rPr><w:t>字体</w:t></w:r></w:p>');
  const run = doc.getParagraphs()[0].runs[0];
  assert.equal(run.fontFamily, 'SimSun');
  assert.equal(run.fontFamilyEastAsia, 'SimSun');
});

test('defineStyle updates existing styles without leaving stale metadata behind', () => {
  const doc = DocxDocument.create();
  doc.defineStyle({
    id: 'Mutable',
    name: 'Mutable',
    type: 'paragraph',
    quickFormat: true,
    isDefault: true,
    paragraph: { spacingAfter: 240 },
    run: { bold: true },
  });
  doc.defineStyle({ id: 'Mutable', name: 'Mutable 2', type: 'paragraph' });
  const xml = doc.getPartXml('word/styles.xml');
  assert.doesNotMatch(xml, /w:default=/);
  assert.doesNotMatch(xml, /w:qFormat/);
  assert.doesNotMatch(xml, /<w:pPr/);
  assert.doesNotMatch(xml, /<w:rPr/);
  assert.match(xml, /Mutable 2/);
});

test('defineStyle rejects null patch-style values', () => {
  const doc = DocxDocument.create();
  assert.throws(() => doc.defineStyle({
    id: 'Invalid',
    name: 'Invalid',
    type: 'paragraph',
    run: { color: null },
  }), /cannot be null/);
});

test('paragraph validation rejects unsigned twips underflow and outline overflow', () => {
  const doc = DocxDocument.create();
  assert.throws(() => doc.formatParagraph(0, { spacingBefore: -50 }), /unsigned twips/i);
  assert.throws(() => doc.formatParagraph(0, { indentHanging: -20 }), /unsigned twips/i);
  assert.throws(() => doc.formatParagraph(0, { outlineLevel: 5000 }), /0 to 9/);
});

test('table style conditions apply first-row and horizontal band run formatting', () => {
  const doc = withStyles(
    `<w:tbl>
      <w:tblPr><w:tblStyle w:val="FancyTable"/><w:tblLook w:firstRow="1" w:noHBand="0"/></w:tblPr>
      <w:tr><w:tc><w:p><w:r><w:t>H1</w:t></w:r></w:p></w:tc></w:tr>
      <w:tr><w:tc><w:p><w:r><w:t>R2</w:t></w:r></w:p></w:tc></w:tr>
      <w:tr><w:tc><w:p><w:r><w:t>R3</w:t></w:r></w:p></w:tc></w:tr>
    </w:tbl>`,
    `<w:styles xmlns:w="${WORD_NS}">
      <w:style w:type="table" w:styleId="FancyTable"><w:name w:val="Fancy Table"/>
        <w:tblStylePr w:type="firstRow"><w:rPr><w:b/></w:rPr></w:tblStylePr>
        <w:tblStylePr w:type="band1Horz"><w:rPr><w:color w:val="008800"/></w:rPr></w:tblStylePr>
        <w:tblStylePr w:type="band2Horz"><w:rPr><w:color w:val="AA5500"/></w:rPr></w:tblStylePr>
      </w:style>
    </w:styles>`,
  );
  const paragraphs = doc.getParagraphs();
  assert.equal(paragraphs[0].runs[0].effective.bold, true);
  assert.equal(paragraphs[1].runs[0].effective.color, '008800');
  assert.equal(paragraphs[2].runs[0].effective.color, 'AA5500');
});

test('table style firstRow/lastRow conditions stay aligned when last row is wrapped by w:sdt', () => {
  const doc = withStyles(
    `<w:tbl>
      <w:tblPr><w:tblStyle w:val="TS"/><w:tblLook w:firstRow="1" w:lastRow="1" w:noHBand="1"/></w:tblPr>
      <w:tr><w:tc><w:p><w:r><w:t>h1</w:t></w:r></w:p></w:tc></w:tr>
      <w:tr><w:tc><w:p><w:r><w:t>a</w:t></w:r></w:p></w:tc></w:tr>
      <w:tr><w:tc><w:p><w:r><w:t>c</w:t></w:r></w:p></w:tc></w:tr>
    </w:tbl>`,
    `<w:styles xmlns:w="${WORD_NS}">
      <w:style w:type="table" w:styleId="TS"><w:name w:val="TS"/>
        <w:tblStylePr w:type="firstRow"><w:rPr><w:color w:val="FF0000"/></w:rPr></w:tblStylePr>
        <w:tblStylePr w:type="lastRow"><w:rPr><w:color w:val="0000FF"/></w:rPr></w:tblStylePr>
      </w:style>
    </w:styles>`,
  );
  let xml = doc.getPartXml(doc.mainDocumentPath);
  const rows = xml.match(/<w:tr>[\s\S]*?<\/w:tr>/g);
  xml = xml.replace(rows[2], `<w:sdt><w:sdtPr/><w:sdtContent>${rows[2]}</w:sdtContent></w:sdt>`);
  doc.setPartXml(doc.mainDocumentPath, xml);

  const paragraphs = doc.getParagraphs();
  assert.equal(paragraphs[0].runs[0].effective.color, 'FF0000');
  assert.equal(paragraphs[1].runs[0].effective.color, undefined);
  assert.equal(paragraphs[2].runs[0].effective.color, '0000FF');
});

test('table style firstCol can be explicitly disabled by tblLook', () => {
  const doc = withStyles(
    `<w:tbl>
      <w:tblPr><w:tblStyle w:val="Cols"/><w:tblLook w:firstColumn="0" w:lastColumn="0"/></w:tblPr>
      <w:tr><w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc></w:tr>
    </w:tbl>`,
    `<w:styles xmlns:w="${WORD_NS}">
      <w:style w:type="table" w:styleId="Cols"><w:name w:val="Cols"/>
        <w:tblStylePr w:type="firstCol"><w:rPr><w:b/></w:rPr></w:tblStylePr>
      </w:style>
    </w:styles>`,
  );
  assert.equal(doc.getParagraphs()[0].runs[0].effective.bold, undefined);
});

test('theme colors read sysClr lastClr fallbacks', () => {
  const doc = withStyles(
    '<w:p><w:pPr><w:pStyle w:val="Text2"/></w:pPr><w:r><w:t>Theme</w:t></w:r></w:p>',
    `<w:styles xmlns:w="${WORD_NS}">
      <w:style w:type="paragraph" w:styleId="Text2"><w:name w:val="Text2"/><w:rPr><w:color w:themeColor="text2"/></w:rPr></w:style>
    </w:styles>`,
    `<?xml version="1.0" encoding="UTF-8"?>
    <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <a:themeElements>
        <a:clrScheme name="Custom">
          <a:dk1><a:srgbClr val="000000"/></a:dk1>
          <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
          <a:dk2><a:sysClr val="windowText" lastClr="3A3A3A"/></a:dk2>
          <a:lt2><a:srgbClr val="EEEEEE"/></a:lt2>
          <a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
        </a:clrScheme>
      </a:themeElements>
    </a:theme>`,
  );
  assert.equal(doc.getParagraphs()[0].runs[0].effective.color, '3A3A3A');
});

test('underline patches keep color when only the style is cleared', () => {
  const doc = withBody('<w:p><w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t>U</w:t></w:r></w:p>');
  doc.formatRun(0, 0, { underlineStyle: null, underlineColor: 'FF0000' });
  const xml = doc.getPartXml(doc.mainDocumentPath);
  assert.match(xml, /<w:u w:color="FF0000"\/>/);
  assert.equal(doc.getParagraphs()[0].runs[0].underlineColor, 'FF0000');
});

test('insertFootnote creates note part and marker without changing paragraph text', () => {
  const doc = DocxDocument.create();
  doc.setParagraphText(0, 'body');
  const inserted = doc.insertFootnote(0, 1, 'note');
  assert.equal(doc.getParagraphs()[0].text, 'body');
  const noteRun = doc.getParagraphs()[0].runs.find((run) => run.noteReference?.id === inserted.id);
  assert.equal(noteRun?.noteReference?.marker, '1');
  assert.match(doc.getPartXml('word/footnotes.xml'), /note/);
});

test('footnote numbering follows reference order, not id order', () => {
  const doc = withBody('<w:p><w:r><w:footnoteReference w:id="9"/></w:r></w:p><w:p><w:r><w:footnoteReference w:id="3"/></w:r></w:p>');
  doc.addPart('word/footnotes.xml', encoder.encode(`<w:footnotes xmlns:w="${WORD_NS}">
    <w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>
    <w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>
    <w:footnote w:id="3"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t>a</w:t></w:r></w:p></w:footnote>
    <w:footnote w:id="9"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t>b</w:t></w:r></w:p></w:footnote>
  </w:footnotes>`), 'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml');
  assert.deepEqual(doc.getFootnotes().map((item) => item.id), [9, 3]);
  assert.deepEqual(doc.getFootnotes().map((item) => item.number), [1, 2]);
});

test('table-cell footnote references are recognized', () => {
  const doc = withBody('<w:tbl><w:tr><w:tc><w:p><w:r><w:t>x</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r></w:p></w:tc></w:tr></w:tbl>');
  doc.addPart('word/footnotes.xml', encoder.encode(`<w:footnotes xmlns:w="${WORD_NS}">
    <w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>
    <w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>
    <w:footnote w:id="1"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t>cell</w:t></w:r></w:p></w:footnote>
  </w:footnotes>`), 'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml');
  assert.equal(doc.getFootnotes()[0].reference.paragraph, 0);
});

test('separator placeholders are excluded from visible footnotes', () => {
  const doc = DocxDocument.create();
  doc.setParagraphText(0, 'body');
  doc.insertFootnote(0, 1, 'visible');
  assert.equal(doc.getFootnotes().length, 1);
});

test('setNoteText replaces non-marker note body content', () => {
  const doc = withBody('<w:p><w:r><w:footnoteReference w:id="1"/></w:r></w:p>');
  doc.addPart('word/footnotes.xml', encoder.encode(`<w:footnotes xmlns:w="${WORD_NS}">
    <w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>
    <w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>
    <w:footnote w:id="1"><w:p><w:r><w:t>Old note body</w:t></w:r></w:p></w:footnote>
  </w:footnotes>`), 'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml');
  doc.setNoteText('footnote', 1, 'NEW');
  assert.match(doc.getPartXml('word/footnotes.xml'), /NEW/);
  assert.doesNotMatch(doc.getPartXml('word/footnotes.xml'), /Old note body NEW/);
});

test('setNoteSettings writes correct settings content type', () => {
  const doc = DocxDocument.create();
  doc.setNoteSettings({ footnote: { numFmt: 'decimal' } });
  assert.match(doc.getPartXml('[Content_Types].xml'), /wordprocessingml.settings\+xml/);
  assert.doesNotMatch(doc.getPartXml('[Content_Types].xml'), /wordprocessingml.document.settings\+xml/);
});

test('section note settings are applied to the paragraph owning sectPr', () => {
  const doc = withBody('<w:p><w:pPr><w:sectPr><w:footnotePr><w:numFmt w:val="lowerRoman"/></w:footnotePr></w:sectPr></w:pPr><w:r><w:footnoteReference w:id="1"/></w:r></w:p><w:p><w:r><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:footnotePr><w:numFmt w:val="upperLetter"/></w:footnotePr></w:sectPr>');
  doc.addPart('word/footnotes.xml', encoder.encode(`<w:footnotes xmlns:w="${WORD_NS}">
    <w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>
    <w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>
    <w:footnote w:id="1"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t>x</w:t></w:r></w:p></w:footnote>
    <w:footnote w:id="2"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t>y</w:t></w:r></w:p></w:footnote>
  </w:footnotes>`), 'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml');
  assert.deepEqual(doc.getFootnotes().map((item) => item.marker), ['i', 'B']);
});

test('non-standard footnotes target path is resolved without duplicate relationships', () => {
  const doc = DocxDocument.create();
  doc.insertFootnote(0, 1, 'base');
  doc.addPart('word/fn.xml', encoder.encode(doc.getPartXml('word/footnotes.xml')), 'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml');
  doc.setPartXml('word/_rels/document.xml.rels', doc.getPartXml('word/_rels/document.xml.rels').replace('footnotes.xml', 'fn.xml'));
  doc.insertFootnote(0, 1, 'next');
  assert.equal((doc.getPartXml('word/_rels/document.xml.rels').match(/relationships\/footnotes/g) ?? []).length, 1);
});

test('insertFootnote uses nested run parent as anchor', () => {
  const doc = withBody('<w:p><w:hyperlink w:anchor="x"><w:r><w:t>link</w:t></w:r></w:hyperlink><w:r><w:t>tail</w:t></w:r></w:p>');
  assert.doesNotThrow(() => doc.insertFootnote(0, 0, 'note'));
});

test('failed direct note operations are atomic', () => {
  const doc = withBody('<w:p><w:r><w:t>x</w:t></w:r></w:p>');
  const revision = doc.revision;
  const parts = doc.listParts();
  assert.throws(() => doc.insertFootnote(0, 99, 'x'), /Run 99/);
  assert.equal(doc.revision, revision);
  assert.deepEqual(doc.listParts(), parts);
});

test('next note id also considers dangling references in body', () => {
  const doc = withBody('<w:p><w:r><w:footnoteReference w:id="1"/></w:r></w:p><w:p><w:r><w:t>x</w:t></w:r></w:p>');
  const note = doc.insertFootnote(1, 1, 'new');
  assert.equal(note.id, 2);
});

test('note body paragraph indexes are sentinel values outside main namespace', () => {
  const doc = DocxDocument.create();
  doc.setParagraphText(0, 'body');
  const inserted = doc.insertFootnote(0, 1, 'one');
  const note = doc.getFootnotes().find((item) => item.id === inserted.id);
  const noteParagraphIndex = note?.blocks[0]?.paragraph.index;
  assert.equal(noteParagraphIndex, -1);
  assert.throws(() => doc.setParagraphText(noteParagraphIndex, 'x'), /non-negative/);
});

test('convertNote updates reference styles to target kind', () => {
  const doc = DocxDocument.create();
  const note = doc.insertFootnote(0, 1, 'convert me');
  doc.convertNote('footnote', note.id);
  assert.match(doc.getPartXml(doc.mainDocumentPath), /EndnoteReference/);
  assert.match(doc.getPartXml('word/endnotes.xml'), /EndnoteReference/);
});

test('setNoteSettings rejects unknown fields and does not bump revision on no-op', () => {
  const doc = DocxDocument.create();
  const baseRevision = doc.revision;
  assert.throws(() => doc.setNoteSettings({ footnote: { bogusKey: 1 } }), /Unknown note setting property/);
  assert.equal(doc.revision, baseRevision);
  doc.setNoteSettings({ footnote: {} });
  assert.equal(doc.revision, baseRevision);
  doc.setNoteSettings({});
  assert.equal(doc.revision, baseRevision);
});

test('paragraph tabs/borders/shading read shape can be written back', () => {
  const doc = withBody('<w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="720"/></w:tabs><w:pBdr><w:top w:val="single" w:sz="8" w:space="0" w:color="FF0000"/></w:pBdr><w:shd w:val="clear" w:fill="AABBCC"/></w:pPr><w:r><w:t>A</w:t></w:r></w:p>');
  const paragraph = doc.getParagraphs()[0];
  doc.setParagraphTabs(0, paragraph.tabs);
  doc.setParagraphBorders(0, paragraph.borders);
  doc.setParagraphShading(0, paragraph.shading);
  const next = doc.getParagraphs()[0];
  assert.equal(next.tabs?.[0]?.leader, undefined);
  assert.equal(next.borders?.top?.shadow, undefined);
  assert.equal(next.shading?.color, undefined);
});

test('setParagraphTabs normalizes order, dedup, clear, and empty removal', () => {
  const doc = withBody('<w:p><w:r><w:t>A</w:t></w:r></w:p>');
  doc.setParagraphTabs(0, [
    { position: 2880, alignment: 'left' },
    { position: 720, alignment: 'left' },
    { position: 2880, alignment: 'right' },
    { position: 720, alignment: 'clear' },
  ]);
  const xml = doc.getPartXml(doc.mainDocumentPath);
  assert.match(xml, /w:tab w:val="right" w:pos="2880"/);
  assert.doesNotMatch(xml, /w:pos="720"/);
  doc.setParagraphTabs(0, []);
  assert.doesNotMatch(doc.getPartXml(doc.mainDocumentPath), /<w:tabs>/);
});

test('insertSymbol rejects XML-illegal code points', () => {
  const doc = withBody('<w:p><w:r><w:t>A</w:t></w:r></w:p>');
  assert.throws(() => doc.insertSymbol(0, 0, 'Wingdings', 0x0001), /XML-valid BMP code point/);
  assert.throws(() => doc.insertSymbol(0, 0, 'Wingdings', 0xD800), /XML-valid BMP code point/);
  doc.insertSymbol(0, 0, 'Wingdings', 0xF04A);
  assert.match(doc.getPartXml(doc.mainDocumentPath), /<w:sym w:font="Wingdings" w:char="F04A"\/>/);
});

test('insertBreak writes OOXML default form for textWrapping', () => {
  const doc = withBody('<w:p><w:r><w:t>A</w:t></w:r></w:p>');
  doc.insertBreak(0, 0, 'textWrapping');
  doc.insertBreak(0, 0, 'page');
  const xml = doc.getPartXml(doc.mainDocumentPath);
  assert.match(xml, /<w:br\/>/);
  assert.match(xml, /<w:br w:type="page"\/>/);
});

test('getSettings resolves related settings.xml with defaults', () => {
  const doc = withBody('<w:p><w:r><w:t>A</w:t></w:r></w:p>');
  doc.addPart('word/_rels/document.xml.rels', encoder.encode(`<Relationships xmlns="${REL_NS}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>`), RELS_TYPE);
  doc.addPart('word/settings.xml', encoder.encode(`<w:settings xmlns:w="${WORD_NS}"><w:defaultTabStop w:val="1440"/><w:evenAndOddHeaders/></w:settings>`), 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml');
  assert.deepEqual(doc.getSettings(), { defaultTabStop: 1440, evenAndOddHeaders: true });
});

test('applyOperations validates new formatting operations', () => {
  const doc = withBody('<w:p><w:r><w:t>A</w:t></w:r></w:p>');
  const snapshot = doc.applyOperations({ operations: [
    { type: 'setParagraphTabs', index: 0, tabs: [{ position: 720, alignment: 'left' }] },
    { type: 'setParagraphBorders', index: 0, borders: { top: { style: 'single', size: 8, space: 0, color: 'auto' } } },
    { type: 'setParagraphShading', index: 0, shading: { pattern: 'clear', fill: 'AABBCC' } },
    { type: 'insertBreak', paragraph: 0, run: 0, breakType: 'page' },
    { type: 'insertSymbol', paragraph: 0, run: 0, font: 'Wingdings', charCode: 0xF04A },
  ] });
  assert.equal(snapshot.revision, doc.revision);
});
