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
const COMMENTS_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml';
const COMMENTS_EXTENDED_TYPE = 'application/vnd.ms-word.commentsExtended+xml';
const W14_NS = 'http://schemas.microsoft.com/office/word/2010/wordml';
const W15_NS = 'http://schemas.microsoft.com/office/word/2012/wordml';
const SETTINGS_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml';

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

function withCommentsDoc(bodyXml, commentsXml, commentsExtendedXml = '', target = 'comments.xml', extendedTarget = 'commentsExtended.xml') {
  const doc = DocxDocument.create();
  doc.setPartXml(doc.mainDocumentPath, `<w:document xmlns:w="${WORD_NS}" xmlns:r="${OFFICE_REL_NS}" xmlns:w14="${W14_NS}" xmlns:w15="${W15_NS}"><w:body>${bodyXml}<w:sectPr/></w:body></w:document>`);
  const rels = [`<Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="${target}"/>`];
  doc.addPart('word/_rels/document.xml.rels', encoder.encode(`<Relationships xmlns="${REL_NS}">${rels.join('')}</Relationships>`), RELS_TYPE);
  doc.addPart(`word/${target}`, encoder.encode(commentsXml), COMMENTS_TYPE);
  if (commentsExtendedXml) {
    doc.updatePartXml('word/_rels/document.xml.rels', (document) => {
      const relation = document.createElementNS(REL_NS, 'Relationship');
      relation.setAttribute('Id', 'rId8');
      relation.setAttribute('Type', 'http://schemas.microsoft.com/office/2011/relationships/commentsExtended');
      relation.setAttribute('Target', extendedTarget);
      document.documentElement.appendChild(relation);
    });
    doc.addPart(`word/${extendedTarget}`, encoder.encode(commentsExtendedXml), COMMENTS_EXTENDED_TYPE);
  }
  return doc;
}

function withSettingsXml(settingsXml, target = 'settings.xml') {
  const doc = withBody('<w:p><w:r><w:t>A</w:t></w:r></w:p>');
  doc.addPart('word/_rels/document.xml.rels', encoder.encode(`<Relationships xmlns="${REL_NS}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="${target}"/></Relationships>`), RELS_TYPE);
  doc.addPart(`word/${target}`, encoder.encode(settingsXml), SETTINGS_TYPE);
  return doc;
}

function trackedDoc(bodyXml = '<w:p><w:r><w:t>A</w:t></w:r></w:p>') {
  const doc = withBody(bodyXml);
  doc.setTrackChanges(true);
  return doc;
}

function paragraphTexts(doc) {
  return doc.getParagraphs().map((paragraph) => paragraph.text);
}

async function compareRoundTrip(base, revised, options) {
  const compared = DocxDocument.compare(base, revised, options);
  const bytes = await compared.toUint8Array();
  const accepted = await DocxDocument.load(bytes);
  accepted.acceptAllRevisions();
  const rejected = await DocxDocument.load(bytes);
  rejected.rejectAllRevisions();
  return { compared, bytes, accepted, rejected };
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

test('formatRange updates only the selected character span and keeps surrounding run format', () => {
  const doc = withBody('<w:p><w:r><w:rPr><w:color w:val="112233"/><w:foo w:bar="1"/></w:rPr><w:t>abcdef</w:t></w:r></w:p>');
  doc.formatRange({ paragraph: 0, start: 2, end: 4 }, { bold: true });
  const paragraph = doc.getParagraphs()[0];
  assert.deepEqual(paragraph.runs.map((run) => run.text), ['ab', 'cd', 'ef']);
  assert.equal(paragraph.runs[0].color, '112233');
  assert.equal(paragraph.runs[2].color, '112233');
  assert.equal(paragraph.runs[1].bold, true);
  assert.match(doc.getPartXml(doc.mainDocumentPath), /<w:foo w:bar="1"\/>/);
});

test('formatRange split is idempotent across repeated formatting calls', () => {
  const doc = withBody('<w:p><w:r><w:t>abcdef</w:t></w:r></w:p>');
  doc.formatRange({ paragraph: 0, start: 1, end: 5 }, { bold: true });
  const once = doc.getParagraphs()[0].runs.map((run) => run.text);
  doc.formatRange({ paragraph: 0, start: 1, end: 5 }, { italic: true });
  const twice = doc.getParagraphs()[0].runs.map((run) => run.text);
  assert.deepEqual(once, ['a', 'bcde', 'f']);
  assert.deepEqual(twice, ['a', 'bcde', 'f']);
});

test('formatRange handles tab and line-break boundaries without corrupting text', () => {
  const doc = withBody('<w:p><w:r><w:t>a</w:t><w:tab/><w:t>b</w:t><w:br/><w:t>c</w:t></w:r></w:p>');
  doc.formatRange({ paragraph: 0, start: 1, end: 4 }, { underline: true });
  const paragraph = doc.getParagraphs()[0];
  assert.equal(paragraph.text, 'a\tb\nc');
  assert.equal(paragraph.runs.some((run) => run.underline), true);
});

test('formatRange does not cut surrogate pairs', () => {
  const doc = withBody('<w:p><w:r><w:t>A😀B</w:t></w:r></w:p>');
  doc.formatRange({ paragraph: 0, start: 1, end: 2 }, { color: 'FF0000' });
  assert.equal(doc.getParagraphs()[0].text, 'A😀B');
  assert.equal(doc.getParagraphs()[0].runs.some((run) => run.text === '😀'), true);
});

test('formatRange across hyperlink text keeps hyperlink metadata readable', () => {
  const doc = withBody('<w:p xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:hyperlink w:anchor="bk"><w:r><w:t>hello</w:t></w:r></w:hyperlink><w:r><w:t> world</w:t></w:r></w:p>');
  doc.formatRange({ paragraph: 0, start: 1, end: 7 }, { italic: true });
  const links = doc.getHyperlinks();
  assert.equal(links.length, 1);
  assert.equal(links[0].anchor, 'bk');
  assert.equal(links[0].text, 'hello');
});

test('getRangeFormat returns common values and undefined for mixed values', () => {
  const doc = withBody('<w:p><w:r><w:rPr><w:b/><w:color w:val="112233"/></w:rPr><w:t>ab</w:t></w:r><w:r><w:rPr><w:b w:val="0"/><w:color w:val="112233"/></w:rPr><w:t>cd</w:t></w:r></w:p>');
  const format = doc.getRangeFormat({ paragraph: 0, start: 0, end: 4 });
  assert.equal(format.bold, undefined);
  assert.equal(format.color, '112233');
});

test('clearRangeFormat clears only requested run-format fields', () => {
  const doc = withBody('<w:p><w:r><w:rPr><w:b/><w:i/><w:color w:val="112233"/></w:rPr><w:t>abcd</w:t></w:r></w:p>');
  doc.clearRangeFormat({ paragraph: 0, start: 1, end: 3 }, ['bold']);
  const middle = doc.getParagraphs()[0].runs.find((run) => run.text === 'bc');
  assert.equal(middle?.bold, undefined);
  assert.equal(middle?.italic, true);
  assert.equal(middle?.color, '112233');
});

test('clearRangeFormat without fields clears all modeled run direct formatting', () => {
  const doc = withBody('<w:p><w:r><w:rPr><w:b/><w:i/><w:color w:val="112233"/></w:rPr><w:t>abcd</w:t></w:r></w:p>');
  doc.clearRangeFormat({ paragraph: 0, start: 0, end: 4 });
  const run = doc.getParagraphs()[0].runs[0];
  assert.equal(run.bold, undefined);
  assert.equal(run.italic, undefined);
  assert.equal(run.color, undefined);
});

test('formatDocumentRange applies partial/whole/partial spans with one revision bump', () => {
  const doc = withBody('<w:p><w:r><w:t>first</w:t></w:r></w:p><w:p><w:r><w:t>middle</w:t></w:r></w:p><w:p><w:r><w:t>last</w:t></w:r></w:p>');
  const before = doc.revision;
  doc.formatDocumentRange({
    start: { paragraph: 0, offset: 2 },
    end: { paragraph: 2, offset: 2 },
  }, { bold: true });
  assert.equal(doc.revision, before + 1);
  const paragraphs = doc.getParagraphs();
  assert.equal(paragraphs[0].runs.some((run) => run.bold && run.text === 'rst'), true);
  assert.equal(paragraphs[1].runs.every((run) => run.bold === true), true);
  assert.equal(paragraphs[2].runs.some((run) => run.bold && run.text.startsWith('la')), true);
});

test('formatDocumentRange is atomic when range validation fails', () => {
  const doc = withBody('<w:p><w:r><w:t>first</w:t></w:r></w:p><w:p><w:r><w:t>second</w:t></w:r></w:p>');
  const beforeXml = doc.getPartXml(doc.mainDocumentPath);
  const beforeRevision = doc.revision;
  assert.throws(() => doc.formatDocumentRange({
    start: { paragraph: 0, offset: 0 },
    end: { paragraph: 1, offset: 999 },
  }, { bold: true }), /out of bounds/);
  assert.equal(doc.revision, beforeRevision);
  assert.equal(doc.getPartXml(doc.mainDocumentPath), beforeXml);
});

test('getDocumentRangeFormat returns undefined for mixed values across paragraphs', () => {
  const doc = withBody('<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>first</w:t></w:r></w:p><w:p><w:r><w:t>second</w:t></w:r></w:p>');
  const format = doc.getDocumentRangeFormat({
    start: { paragraph: 0, offset: 0 },
    end: { paragraph: 1, offset: 6 },
  });
  assert.equal(format.bold, undefined);
});

test('formatRange skips image-only offsets safely', () => {
  const doc = withBody(`<w:p xmlns:r="${OFFICE_REL_NS}" xmlns:wp="${WP_NS}" xmlns:a="${A_NS}" xmlns:pic="${PIC_NS}"><w:r><w:drawing><wp:inline><wp:extent cx="190500" cy="95250"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="1" name="logo"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="${PIC_NS}"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="logo"/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="190500" cy="95250"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`);
  const revision = doc.revision;
  doc.formatRange({ paragraph: 0, start: 0, end: 0 }, { italic: true });
  assert.equal(doc.revision, revision);
});

test('agent validates and applies formatRange operation', () => {
  const doc = withBody('<w:p><w:r><w:t>abcdef</w:t></w:r></w:p>');
  assert.throws(() => doc.applyOperations({ operations: [{ type: 'formatRange', range: { paragraph: 0, start: 4, end: 1 }, format: { bold: true } }] }), /range.end/);
  doc.applyOperations({ operations: [{ type: 'formatRange', range: { paragraph: 0, start: 2, end: 4 }, format: { bold: true } }] });
  assert.equal(doc.getParagraphs()[0].runs.some((run) => run.text === 'cd' && run.bold), true);
});

test('agent validates and applies clearRangeFormat operation', () => {
  const doc = withBody('<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>abcd</w:t></w:r></w:p>');
  assert.throws(() => doc.applyOperations({ operations: [{ type: 'clearRangeFormat', range: { paragraph: 0, start: 0, end: 2 }, fields: ['unknown'] }] }), /Unsupported/);
  doc.applyOperations({ operations: [{ type: 'clearRangeFormat', range: { paragraph: 0, start: 0, end: 2 }, fields: ['bold'] }] });
  assert.equal(doc.getParagraphs()[0].runs.some((run) => run.text === 'ab' && run.bold === undefined), true);
});

test('agent validates and applies formatDocumentRange operation', () => {
  const doc = withBody('<w:p><w:r><w:t>first</w:t></w:r></w:p><w:p><w:r><w:t>second</w:t></w:r></w:p>');
  assert.throws(() => doc.applyOperations({ operations: [{
    type: 'formatDocumentRange',
    range: { start: { paragraph: 1, offset: 0 }, end: { paragraph: 0, offset: 1 } },
    format: { italic: true },
  }] }), /must not be after/);
  doc.applyOperations({ operations: [{
    type: 'formatDocumentRange',
    range: { start: { paragraph: 0, offset: 2 }, end: { paragraph: 1, offset: 3 } },
    format: { italic: true },
  }] });
  assert.equal(doc.getParagraphs()[1].runs.some((run) => run.italic), true);
});

test('getRangeFormat on collapsed range returns nearby run direct format', () => {
  const doc = withBody('<w:p><w:r><w:rPr><w:color w:val="AA0000"/></w:rPr><w:t>ab</w:t></w:r><w:r><w:rPr><w:color w:val="00AA00"/></w:rPr><w:t>cd</w:t></w:r></w:p>');
  const left = doc.getRangeFormat({ paragraph: 0, start: 1, end: 1 });
  const right = doc.getRangeFormat({ paragraph: 0, start: 2, end: 2 });
  assert.equal(left.color, 'AA0000');
  assert.equal(right.color, '00AA00');
});

test('clearRangeFormat rejects unsupported fields at runtime API', () => {
  const doc = DocxDocument.create();
  assert.throws(() => doc.clearRangeFormat({ paragraph: 0, start: 0, end: 0 }, ['notAField']), /Unsupported/);
});

test('formatDocumentRange rejects cross-container ranges without corrupting text', () => {
  const doc = DocxDocument.create();
  doc.setParagraphText(0, 'head');
  doc.insertTable([['cell']]);
  const cellParagraph = doc.getBlocks().find((block) => block.type === 'table').rows[0].cells[0].blocks[0].paragraph.index;
  doc.insertParagraph('tail');
  const before = doc.getPartXml(doc.mainDocumentPath);
  assert.throws(() => doc.formatDocumentRange({
    start: { paragraph: 0, offset: 1 },
    end: { paragraph: cellParagraph, offset: doc.getParagraphs()[cellParagraph].text.length },
  }, { underline: true }), /Cross-container/);
  assert.equal(doc.getPartXml(doc.mainDocumentPath), before);
  assert.equal(doc.getParagraphs()[0].text, 'head');
  assert.equal(doc.getParagraphs()[cellParagraph].text, 'cell');
});

test('agent schema includes new range formatting operations', () => {
  const names = AGENT_OPERATION_SCHEMA.properties.operations.items.oneOf
    .map((entry) => entry.properties.type.const)
    .sort();
  assert.equal(names.includes('formatRange'), true);
  assert.equal(names.includes('clearRangeFormat'), true);
  assert.equal(names.includes('formatDocumentRange'), true);
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
  assert.equal(AGENT_OPERATION_SCHEMA.properties.operations.items.oneOf.length, 62);
});

test('undo and redo share one stack with monotonic revision', () => {
  const doc = DocxDocument.create();
  doc.setParagraphText(0, 'a');
  doc.setParagraphText(0, 'ab');
  doc.setParagraphText(0, 'abc');
  assert.equal(doc.getHistory().undo.length, 1);
  assert.equal(doc.revision, 3);
  doc.undo();
  assert.equal(doc.getParagraphs()[0].text, '');
  assert.equal(doc.revision, 4);
  doc.redo();
  assert.equal(doc.getParagraphs()[0].text, 'abc');
  assert.equal(doc.revision, 5);
});

test('setParagraphText merge is paragraph-local', () => {
  const doc = DocxDocument.create();
  doc.insertParagraph('x');
  doc.setParagraphText(0, 'A');
  doc.setParagraphText(1, 'B');
  assert.equal(doc.getHistory().undo.length, 3);
  doc.undo();
  assert.equal(doc.getParagraphs()[0].text, 'A');
  assert.equal(doc.getParagraphs()[1].text, 'x');
});

test('applyOperations batch is one undo step and revision still advances on undo', () => {
  const doc = DocxDocument.create();
  doc.applyOperations({
    operations: [
      { type: 'setParagraphText', index: 0, text: 'A' },
      { type: 'insertParagraph', text: 'B' },
      { type: 'formatRun', paragraph: 0, run: 0, format: { bold: true } },
      { type: 'formatParagraph', index: 0, format: { alignment: 'center' } },
      { type: 'replaceText', search: 'B', replacement: 'C' },
    ],
  });
  assert.equal(doc.getHistory().undo.length, 1);
  const afterBatch = doc.revision;
  doc.undo();
  assert.equal(doc.getParagraphs().length, 1);
  assert.equal(doc.getParagraphs()[0].text, '');
  assert.equal(doc.revision, afterBatch + 1);
});

test('beginHistoryGroup/endHistoryGroup merges multiple edits into one step', () => {
  const doc = DocxDocument.create();
  doc.beginHistoryGroup('format painter');
  doc.formatRun(0, 0, { bold: true });
  doc.formatParagraph(0, { alignment: 'center' });
  doc.endHistoryGroup();
  const history = doc.getHistory();
  assert.equal(history.undo.length, 1);
  assert.equal(history.undo[0].label, 'format painter');
  doc.undo();
  const paragraph = doc.getParagraphs()[0];
  assert.equal(paragraph.alignment, undefined);
  assert.equal(paragraph.runs[0].bold, undefined);
});

test('nested history groups still produce a single undo step', () => {
  const doc = DocxDocument.create();
  doc.beginHistoryGroup('nested');
  doc.setParagraphText(0, 'A');
  doc.beginHistoryGroup();
  doc.formatRun(0, 0, { italic: true });
  doc.endHistoryGroup();
  doc.endHistoryGroup();
  assert.equal(doc.getHistory().undo.length, 1);
  doc.undo();
  assert.equal(doc.getParagraphs()[0].text, '');
});

test('history snapshot materializes dirty XML before recording next step', () => {
  const doc = DocxDocument.create();
  doc.setParagraphText(0, 'A');
  doc.formatRun(0, 0, { bold: true });
  doc.undo();
  const paragraph = doc.getParagraphs()[0];
  assert.equal(paragraph.text, 'A');
  assert.equal(paragraph.runs[0].bold, undefined);
});

test('undo/redo restores binary image bytes', () => {
  const doc = DocxDocument.create();
  const image = doc.insertImage({ bytes: PNG_BYTES, contentType: 'image/png', alt: 'x' });
  doc.replaceImageBytes(image, GIF_BYTES, 'image/gif');
  doc.undo();
  assert.deepEqual(doc.getImageBytes(doc.getImages()[0]), PNG_BYTES);
  doc.redo();
  assert.deepEqual(doc.getImageBytes(doc.getImages()[0]), GIF_BYTES);
});

test('undo across insertImage keeps document.xml/media/rels consistent without probe reads', () => {
  const doc = DocxDocument.create();
  doc.setParagraphText(0, '一');
  doc.insertImage({ bytes: PNG_BYTES, contentType: 'image/png', paragraph: 0, widthEmu: 914400, heightEmu: 914400 });
  doc.insertParagraph('二');

  doc.undo();
  const afterFirstUndoImage = doc.getImages()[0];
  assert.ok(afterFirstUndoImage);
  assert.deepEqual(doc.getImageBytes(afterFirstUndoImage), PNG_BYTES);

  doc.undo();
  assert.equal(doc.getImages().length, 0);
  assert.equal(doc.listParts().filter((path) => path.startsWith('word/media/')).length, 0);
  assert.equal(doc.listParts().includes('word/_rels/document.xml.rels'), false);
  assert.doesNotMatch(doc.getPartXml(doc.mainDocumentPath), /<w:drawing[\s>]/);
});

test('redo across insertImage restores image bytes and relationships exactly', () => {
  const doc = DocxDocument.create();
  doc.setParagraphText(0, '一');
  doc.insertImage({ bytes: PNG_BYTES, contentType: 'image/png', paragraph: 0, widthEmu: 914400, heightEmu: 914400 });
  doc.insertParagraph('二');

  doc.undo();
  doc.undo();
  doc.redo();
  doc.redo();

  const image = doc.getImages()[0];
  assert.ok(image);
  assert.deepEqual(doc.getImageBytes(image), PNG_BYTES);
  assert.ok(doc.listParts().includes('word/_rels/document.xml.rels'));
  assert.equal(doc.listParts().filter((path) => path.startsWith('word/media/')).length, 1);
});

test('history cap drops oldest steps', () => {
  const doc = DocxDocument.create();
  for (let i = 0; i < 55; i++) doc.formatRun(0, 0, { bold: i % 2 === 0 });
  assert.equal(doc.getHistory().undo.length, 50);
  let undos = 0;
  while (doc.canUndo()) {
    doc.undo();
    undos++;
  }
  assert.equal(undos, 50);
});

test('clearHistory empties undo/redo stacks', () => {
  const doc = DocxDocument.create();
  doc.setParagraphText(0, 'a');
  assert.equal(doc.canUndo(), true);
  doc.clearHistory();
  assert.equal(doc.canUndo(), false);
  assert.equal(doc.canRedo(), false);
  assert.deepEqual(doc.getHistory(), { undo: [], redo: [] });
});

test('undo keeps expectedRevision conflict detection correct', () => {
  const doc = DocxDocument.create();
  doc.setParagraphText(0, 'A');
  const staleRevision = doc.revision;
  doc.undo();
  assert.throws(() => doc.applyOperations({
    expectedRevision: staleRevision,
    operations: [{ type: 'setParagraphText', index: 0, text: 'B' }],
  }), /Revision conflict/);
  assert.doesNotThrow(() => doc.applyOperations({
    expectedRevision: doc.revision,
    operations: [{ type: 'setParagraphText', index: 0, text: 'B' }],
  }));
});

test('applyOperations undo/redo uses the same document history stack', () => {
  const doc = DocxDocument.create();
  doc.setParagraphText(0, 'A');
  doc.applyOperations({ operations: [{ type: 'undo' }] });
  assert.equal(doc.getParagraphs()[0].text, '');
  doc.applyOperations({ operations: [{ type: 'redo' }] });
  assert.equal(doc.getParagraphs()[0].text, 'A');
});

test('mixed applyOperations batch with undo and edits still records one transaction step', () => {
  const doc = DocxDocument.create();
  doc.setParagraphText(0, 'A');
  doc.applyOperations({
    operations: [
      { type: 'undo' },
      { type: 'setParagraphText', index: 0, text: 'B' },
    ],
  });
  assert.equal(doc.getParagraphs()[0].text, 'B');
  doc.undo();
  assert.equal(doc.getParagraphs()[0].text, 'A');
});

test('history entry revision stores the revision at completion time', () => {
  const doc = DocxDocument.create();
  doc.setParagraphText(0, 'A');
  const [entry] = doc.getHistory().undo;
  assert.equal(entry.revision, 1);
  assert.equal(typeof entry.at, 'number');
});

test('history snapshot buffers are immutable across later edits', () => {
  const doc = DocxDocument.create();
  doc.setParagraphText(0, 'A');
  const historyPart = doc.undoHistory[0].parts.get(doc.mainDocumentPath);
  assert.ok(historyPart);
  const before = Uint8Array.from(historyPart);
  doc.setParagraphText(0, 'B');
  assert.deepEqual(Uint8Array.from(historyPart), before);
});

test('merged setParagraphText edits skip repeated history snapshot materialization', () => {
  const doc = DocxDocument.create();
  doc.setParagraphText(0, 'A');
  const original = doc.materializeAllParts;
  let calls = 0;
  doc.materializeAllParts = function materializeProxy() {
    calls++;
    return original.call(this);
  };
  doc.setParagraphText(0, 'AB');
  doc.setParagraphText(0, 'ABC');
  assert.equal(calls, 0);
  assert.equal(doc.getHistory().undo.length, 1);
  doc.undo();
  assert.equal(doc.getParagraphs()[0].text, '');
});

test('undo/redo without history is a no-op', () => {
  const doc = DocxDocument.create();
  const revision = doc.revision;
  doc.undo();
  doc.redo();
  assert.equal(doc.revision, revision);
  assert.equal(doc.getParagraphs()[0].text, '');
});

test('endHistoryGroup throws when no history group is active', () => {
  const doc = DocxDocument.create();
  assert.throws(() => doc.endHistoryGroup(), /not active/);
});

test('failed mutation aborts active history group commit', () => {
  const doc = DocxDocument.create();
  doc.beginHistoryGroup('batch');
  doc.setParagraphText(0, 'A');
  assert.throws(() => doc.formatRun(0, 99, { bold: true }), /Run 99 does not exist/);
  doc.endHistoryGroup();
  assert.equal(doc.getHistory().undo.length, 0);
  assert.equal(doc.getParagraphs()[0].text, 'A');
});

test('applyOperations batches do not merge with neighboring direct edits', () => {
  const doc = DocxDocument.create();
  doc.setParagraphText(0, 'A');
  doc.applyOperations({ operations: [{ type: 'setParagraphText', index: 0, text: 'B' }] });
  assert.equal(doc.getHistory().undo.length, 2);
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
  assert.equal(AGENT_OPERATION_SCHEMA.properties.operations.items.oneOf.length, 62);
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

test('getSections returns basic page setup from body sectPr', () => {
  const doc = DocxDocument.create();
  const section = doc.getSection(0);
  assert.equal(section.index, 0);
  assert.equal(section.startParagraph, 0);
  assert.equal(section.type, 'nextPage');
  assert.equal(section.orientation, 'portrait');
  assert.equal(section.pageWidth, 11906);
});

test('getSections synthesizes an implicit default section when sectPr is missing', () => {
  const doc = DocxDocument.create();
  doc.setPartXml(doc.mainDocumentPath, `<w:document xmlns:w="${WORD_NS}"><w:body><w:p><w:r><w:t>x</w:t></w:r></w:p></w:body></w:document>`);
  const section = doc.getSection(0);
  assert.equal(section.isImplicit, true);
  assert.equal(section.startParagraph, 0);
  assert.equal(section.endParagraph, 0);
  doc.setPageSetup(0, { pageWidth: 15000 });
  assert.match(doc.getPartXml(doc.mainDocumentPath), /<w:sectPr>/);
});

test('paragraph-scoped sectPr produces multiple sections with correct ranges', () => {
  const doc = withBody('<w:p><w:pPr><w:sectPr><w:type w:val="continuous"/></w:sectPr></w:pPr><w:r><w:t>A</w:t></w:r></w:p><w:p><w:r><w:t>B</w:t></w:r></w:p>');
  const sections = doc.getSections();
  assert.equal(sections.length, 2);
  assert.equal(sections[0].startParagraph, 0);
  assert.equal(sections[0].endParagraph, 0);
  assert.equal(sections[0].type, 'continuous');
  assert.equal(sections[1].startParagraph, 1);
});

test('section ranges stay aligned when body contains table paragraphs', () => {
  const doc = withBody('<w:tbl><w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p><w:pPr><w:sectPr/></w:pPr><w:r><w:t>end</w:t></w:r></w:p><w:p><w:r><w:t>tail</w:t></w:r></w:p>');
  const sections = doc.getSections();
  assert.equal(doc.getParagraphs().length, 3);
  assert.equal(sections[0].endParagraph, 1);
  assert.equal(sections[1].startParagraph, 2);
});

test('section ranges stay aligned with descendants order including txbxContent paragraphs', () => {
  const doc = withBody('<w:p><w:r><w:pict><w:txbxContent><w:p><w:r><w:t>inside box</w:t></w:r></w:p></w:txbxContent></w:pict><w:t>A</w:t></w:r></w:p><w:p><w:pPr><w:sectPr/></w:pPr><w:r><w:t>B</w:t></w:r></w:p><w:p><w:r><w:t>C</w:t></w:r></w:p>');
  const sections = doc.getSections();
  assert.equal(doc.getParagraphs().length, 4);
  assert.equal(sections[0].endParagraph, 2);
  assert.equal(sections[1].startParagraph, 3);
});

test('setPageSetup updates known fields and keeps unknown sectPr children', () => {
  const doc = withBody('<w:p><w:r><w:t>x</w:t></w:r></w:p><w:sectPr><w:docGrid w:linePitch="360"/></w:sectPr>');
  doc.setPageSetup(0, { pageWidth: 20000, orientation: 'landscape', margins: { left: 700 }, columns: { count: 2, space: 360 } });
  const xml = doc.getPartXml(doc.mainDocumentPath);
  assert.match(xml, /w:pgSz[^>]*w:w="20000"[^>]*w:orient="landscape"/);
  assert.match(xml, /w:pgMar[^>]*w:left="700"/);
  assert.match(xml, /w:cols[^>]*w:num="2"[^>]*w:space="360"/);
  assert.match(xml, /w:docGrid/);
});

test('setPageSetup with only orientation swaps page width and height when needed', () => {
  const doc = DocxDocument.create();
  doc.setPageSetup(0, { orientation: 'landscape' });
  let section = doc.getSection(0);
  assert.equal(section.orientation, 'landscape');
  assert.ok(section.pageWidth > section.pageHeight);
  assert.match(doc.getPartXml(doc.mainDocumentPath), /<w:pgSz[^>]*w:w="16838"[^>]*w:h="11906"[^>]*w:orient="landscape"/);
  doc.setPageSetup(0, { orientation: 'portrait' });
  section = doc.getSection(0);
  assert.equal(section.orientation, 'portrait');
  assert.ok(section.pageWidth < section.pageHeight);
});

test('insertSectionBreak and deleteSectionBreak update section structure explicitly', () => {
  const doc = withBody('<w:p><w:r><w:t>one</w:t></w:r></w:p><w:p><w:r><w:t>two</w:t></w:r></w:p>');
  doc.insertSectionBreak(0, 'continuous');
  assert.equal(doc.getSections().length, 2);
  assert.equal(doc.getSection(0).type, 'continuous');
  doc.deleteSectionBreak(0);
  assert.equal(doc.getSections().length, 1);
});

test('deleteSectionBreak keeps following section properties after merge', () => {
  const doc = withBody('<w:p><w:pPr><w:sectPr><w:pgSz w:w="14000" w:h="10000"/></w:sectPr></w:pPr><w:r><w:t>a</w:t></w:r></w:p><w:p><w:r><w:t>b</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="18000" w:h="12000"/></w:sectPr>');
  doc.deleteSectionBreak(0);
  assert.equal(doc.getSections().length, 1);
  assert.equal(doc.getSection(0).pageWidth, 18000);
});

test('deleteSectionBreak rejects final section and insertSectionBreak rejects duplicate break paragraph', () => {
  const doc = DocxDocument.create();
  assert.throws(() => doc.deleteSectionBreak(0), /final section break/);
  doc.insertSectionBreak(0, 'nextPage');
  assert.throws(() => doc.insertSectionBreak(0, 'nextPage'), /already ends with a section break/);
});

test('createHeader registers content type, relationship and section reference order', () => {
  const doc = DocxDocument.create();
  const path = doc.createHeader(0);
  assert.equal(path, 'word/header1.xml');
  const rels = doc.getPartXml('word/_rels/document.xml.rels');
  assert.match(rels, /relationships\/header/);
  const main = doc.getPartXml(doc.mainDocumentPath);
  assert.match(main, /<w:headerReference[\s\S]*<w:pgSz/);
  assert.match(doc.getPartXml('[Content_Types].xml'), /word\/header1.xml/);
});

test('creating even header enables evenAndOddHeaders in settings.xml', () => {
  const doc = DocxDocument.create();
  doc.setHeaderText(0, 'Even page header', 'even');
  assert.equal(doc.getHeaderBlocks(0, 'even')[0].paragraph.text, 'Even page header');
  assert.match(doc.getPartXml('word/settings.xml'), /<w:evenAndOddHeaders\/>/);
});

test('setHeaderText and setFooterText create parts and keep editable text', () => {
  const doc = DocxDocument.create();
  doc.setHeaderText(0, 'Header');
  doc.setFooterText(0, 'Footer');
  assert.equal(doc.getHeaderBlocks(0)[0].paragraph.text, 'Header');
  assert.equal(doc.getFooterBlocks(0)[0].paragraph.text, 'Footer');
});

test('setHeaderText and setFooterText do not bump revision on no-op writes', () => {
  const doc = DocxDocument.create();
  doc.setHeaderText(0, 'Header');
  doc.setFooterText(0, 'Footer');
  const revision = doc.revision;
  const headerXml = doc.getPartXml(doc.getSection(0).headers.default);
  const footerXml = doc.getPartXml(doc.getSection(0).footers.default);
  doc.setHeaderText(0, 'Header');
  doc.setFooterText(0, 'Footer');
  assert.equal(doc.revision, revision);
  assert.equal(doc.getPartXml(doc.getSection(0).headers.default), headerXml);
  assert.equal(doc.getPartXml(doc.getSection(0).footers.default), footerXml);
});

test('getHeaderBlocks reads nested table structure from header part', () => {
  const doc = DocxDocument.create();
  const path = doc.createHeader(0);
  doc.setPartXml(path, `<w:hdr xmlns:w="${WORD_NS}"><w:tbl><w:tr><w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:hdr>`);
  const blocks = doc.getHeaderBlocks(0);
  assert.equal(blocks[0].type, 'table');
  assert.equal(blocks[0].rows[0].cells[0].blocks[0].paragraph.text, 'A');
});

test('header and footer kind fall back to default when specific kind is missing', () => {
  const doc = DocxDocument.create();
  doc.setHeaderText(0, 'Default header');
  assert.equal(doc.getHeaderBlocks(0, 'first')[0].paragraph.text, 'Default header');
  doc.setFooterText(0, 'Default footer');
  assert.equal(doc.getFooterBlocks(0, 'even')[0].paragraph.text, 'Default footer');
});

test('missing header/footer parts or dangling references do not crash block reads', () => {
  const doc = withBody('<w:p><w:r><w:t>x</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rId9" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/></w:sectPr>');
  assert.deepEqual(doc.getHeaderBlocks(0), []);
});

test('insertPageNumberField writes PAGE placeholder and renders text', () => {
  const doc = DocxDocument.create();
  const path = doc.createHeader(0);
  doc.insertPageNumberField(path);
  assert.match(doc.getPartXml(path), /w:fldSimple[^>]+PAGE/);
  const texts = doc.getHeaderBlocks(0).filter(block => block.type === 'paragraph').map(block => block.paragraph.text);
  assert.ok(texts.includes('1'));
});

test('insertPageNumberField writes NUMPAGES placeholder and renders text', () => {
  const doc = DocxDocument.create();
  const path = doc.createFooter(0);
  doc.insertPageNumberField(path, { total: true, format: 'ROMAN' });
  assert.ok(doc.getPartXml(path).includes('NUMPAGES \\* ROMAN'));
  const texts = doc.getFooterBlocks(0).filter(block => block.type === 'paragraph').map(block => block.paragraph.text);
  assert.ok(texts.includes('?'));
});

test('setParagraphText preserves fldSimple and keeps runs valid', () => {
  const doc = DocxDocument.create();
  doc.setPartXml(doc.mainDocumentPath,
    `<w:document xmlns:w="${WORD_NS}"><w:body>` +
    `<w:p><w:r><w:t xml:space="preserve">Page </w:t></w:r>` +
    `<w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple>` +
    `<w:r><w:t xml:space="preserve"> end</w:t></w:r></w:p>` +
    `<w:sectPr/></w:body></w:document>`);
  doc.setParagraphText(0, 'Page 9 end');
  const xml = doc.getPartXml(doc.mainDocumentPath);
  assert.match(xml, /<w:fldSimple/);
  assert.doesNotMatch(xml, /<w:p><w:r>[\s\S]*<\/w:r><w:t/);
});

test('section parser tolerates zero page size and mismatched column declaration', () => {
  const doc = withBody('<w:p><w:r><w:t>x</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="0" w:h="0"/><w:cols w:num="3"><w:col w:w="1000"/></w:cols></w:sectPr>');
  const section = doc.getSection(0);
  assert.equal(section.pageWidth, 0);
  assert.equal(section.pageHeight, 0);
  assert.equal(section.columns.count, 3);
  assert.deepEqual(section.columns.widths, [1000]);
});

test('getBlocks includes visible page and section break markers', () => {
  const doc = withBody('<w:p><w:r><w:t>a</w:t></w:r><w:r><w:br w:type="page"/></w:r></w:p><w:p><w:pPr><w:sectPr><w:type w:val="oddPage"/></w:sectPr></w:pPr><w:r><w:t>b</w:t></w:r></w:p>');
  const kinds = doc.getBlocks().map(block => block.type);
  assert.ok(kinds.includes('pageBreak'));
  assert.ok(kinds.includes('sectionBreak'));
});

test('page-break markers ignore nested textbox breaks and keep count', () => {
  const doc = withBody('<w:p><w:r><w:pict><w:txbxContent><w:p><w:r><w:br w:type="page"/></w:r></w:p></w:txbxContent></w:pict><w:br w:type="page"/><w:t>A</w:t><w:br w:type="page"/></w:r></w:p>');
  const blocks = doc.getBlocks();
  assert.equal(blocks.filter(block => block.type === 'pageBreak').length, 2);
});

test('even/odd header references are disabled when settings flag is absent', () => {
  const doc = DocxDocument.create();
  doc.addPart('word/_rels/document.xml.rels', new TextEncoder().encode(
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>`,
  ), 'application/vnd.openxmlformats-package.relationships+xml');
  doc.setPartXml(doc.mainDocumentPath,
    `<w:document xmlns:w="${WORD_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t>a</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="even" r:id="rId2"/></w:sectPr></w:body></w:document>`);
  doc.addPart('word/header1.xml', new TextEncoder().encode(`<w:hdr xmlns:w="${WORD_NS}"><w:p><w:r><w:t>even</w:t></w:r></w:p></w:hdr>`), 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml');
  assert.equal(doc.getSection(0).headers.even, undefined);
});

test('setPageSetup can persist explicit column widths and numbering start', () => {
  const doc = DocxDocument.create();
  doc.setPageSetup(0, { columns: { count: 2, equalWidth: false, widths: [3000, 5000] }, pageNumbering: { start: 5, format: 'decimal' }, titlePage: true });
  const xml = doc.getPartXml(doc.mainDocumentPath);
  assert.match(xml, /w:cols[^>]*w:equalWidth="0"/);
  assert.match(xml, /<w:col w:w="3000"/);
  assert.match(xml, /<w:pgNumType[^>]*w:start="5"[^>]*w:fmt="decimal"/);
  assert.match(xml, /<w:titlePg\/>/);
});

test('setPageSetup can clear stale column widths and page numbering attributes', () => {
  const doc = DocxDocument.create();
  doc.setPageSetup(0, { columns: { count: 2, equalWidth: false, widths: [3000, 5000] }, pageNumbering: { start: 3, format: 'decimal' } });
  doc.setPageSetup(0, { columns: { equalWidth: true, widths: [] }, pageNumbering: {} });
  const xml = doc.getPartXml(doc.mainDocumentPath);
  assert.doesNotMatch(xml, /<w:col /);
  assert.doesNotMatch(xml, /w:start="/);
  assert.doesNotMatch(xml, /w:fmt="/);
});

test('relationship target decoding errors are tolerated when resolving section references', () => {
  const doc = DocxDocument.create();
  doc.addPart('word/_rels/document.xml.rels', new TextEncoder().encode(
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdBad" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="%E0%A4%A"/></Relationships>`,
  ), 'application/vnd.openxmlformats-package.relationships+xml');
  doc.setPartXml(doc.mainDocumentPath,
    `<w:document xmlns:w="${WORD_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t>a</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rIdBad"/></w:sectPr></w:body></w:document>`);
  assert.doesNotThrow(() => doc.getSections());
  assert.equal(Object.values(doc.getSection(0).headers).filter(Boolean).length, 0);
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

test('getStyleGallery returns only quick-format styles sorted by uiPriority', () => {
  const doc = withStyles(
    '<w:p><w:r><w:t>Gallery</w:t></w:r></w:p>',
    `<w:styles xmlns:w="${WORD_NS}">
      <w:style w:type="paragraph" w:styleId="Body"><w:name w:val="Body"/><w:qFormat/><w:uiPriority w:val="20"/></w:style>
      <w:style w:type="character" w:styleId="Strong"><w:name w:val="Strong"/><w:qFormat/><w:uiPriority w:val="10"/></w:style>
      <w:style w:type="paragraph" w:styleId="Alpha"><w:name w:val="Alpha"/><w:qFormat/></w:style>
      <w:style w:type="paragraph" w:styleId="Hidden"><w:name w:val="Hidden"/></w:style>
    </w:styles>`,
  );
  assert.deepEqual(doc.getStyleGallery().map((style) => style.id), ['Strong', 'Body', 'Alpha']);
});

test('applyParagraphStyle rejects missing styles without changing legacy formatParagraph behavior', () => {
  const doc = DocxDocument.create();
  assert.throws(() => doc.applyParagraphStyle(0, 'Missing'), /Paragraph style not found/i);
  assert.doesNotThrow(() => doc.formatParagraph(0, { style: 'Missing' }));
});

test('applyParagraphStyle can clear conflicting direct paragraph and run formatting', () => {
  const doc = withStyles(
    '<w:p><w:pPr><w:jc w:val="right"/><w:spacing w:before="80" w:after="120"/></w:pPr><w:r><w:rPr><w:color w:val="AA5500"/><w:u w:val="single"/></w:rPr><w:t>Clear</w:t></w:r></w:p>',
    `<w:styles xmlns:w="${WORD_NS}">
      <w:style w:type="paragraph" w:styleId="Styled">
        <w:name w:val="Styled"/>
        <w:pPr><w:jc w:val="center"/><w:spacing w:after="240"/></w:pPr>
        <w:rPr><w:color w:val="336699"/></w:rPr>
      </w:style>
    </w:styles>`,
  );
  doc.applyParagraphStyle(0, 'Styled', { clearDirectFormat: true });
  const xml = doc.getPartXml(doc.mainDocumentPath);
  const paragraph = doc.getParagraphs()[0];
  assert.equal(paragraph.style, 'Styled');
  assert.equal(paragraph.effective.alignment, 'center');
  assert.equal(paragraph.effective.spacingAfter, 240);
  assert.equal(paragraph.spacingBefore, 80);
  assert.equal(paragraph.runs[0].effective.color, '336699');
  assert.equal(paragraph.runs[0].underline, true);
  assert.doesNotMatch(xml, /w:jc w:val="right"/);
  assert.doesNotMatch(xml, /w:after="120"/);
  assert.match(xml, /w:before="80"/);
  assert.doesNotMatch(xml, /w:color w:val="AA5500"/);
  assert.match(xml, /w:u w:val="single"/);
});

test('applyParagraphStyle clears direct formatting that would otherwise override inherited defaults', () => {
  const doc = withStyles(
    '<w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:t>Clear</w:t></w:r></w:p>',
    `<w:styles xmlns:w="${WORD_NS}">
      <w:docDefaults><w:pPrDefault><w:pPr><w:jc w:val="left"/></w:pPr></w:pPrDefault></w:docDefaults>
      <w:style w:type="paragraph" w:styleId="Styled"><w:name w:val="Styled"/></w:style>
    </w:styles>`,
  );
  doc.applyParagraphStyle(0, 'Styled', { clearDirectFormat: true });
  assert.equal(doc.getParagraphs()[0].effective.alignment, 'left');
  assert.doesNotMatch(doc.getPartXml(doc.mainDocumentPath), /w:jc w:val="right"/);
});

test('applyParagraphStyle is a no-op when the style and conflicting direct format already match', () => {
  const doc = withStyles(
    '<w:p><w:pPr><w:pStyle w:val="Styled"/><w:spacing w:after="240"/></w:pPr><w:r><w:rPr><w:color w:val="336699"/></w:rPr><w:t>A</w:t></w:r></w:p>',
    `<w:styles xmlns:w="${WORD_NS}">
      <w:style w:type="paragraph" w:styleId="Styled">
        <w:name w:val="Styled"/>
        <w:pPr><w:spacing w:after="240"/></w:pPr>
        <w:rPr><w:color w:val="336699"/></w:rPr>
      </w:style>
    </w:styles>`,
  );
  const revision = doc.revision;
  doc.applyParagraphStyle(0, 'Styled', { clearDirectFormat: true });
  assert.equal(doc.revision, revision);
});

test('applyCharacterStyle rejects missing character styles', () => {
  const doc = DocxDocument.create();
  assert.throws(() => doc.applyCharacterStyle({ paragraph: 0, start: 0, end: 1 }, 'Missing'), /Character style not found/i);
});

test('applyCharacterStyle can clear conflicting direct run formatting', () => {
  const doc = withStyles(
    '<w:p><w:r><w:rPr><w:b/><w:color w:val="AA5500"/><w:u w:val="single"/></w:rPr><w:t>Word</w:t></w:r></w:p>',
    `<w:styles xmlns:w="${WORD_NS}">
      <w:style w:type="character" w:styleId="Emphasis">
        <w:name w:val="Emphasis"/>
        <w:rPr><w:i/><w:color w:val="224466"/></w:rPr>
      </w:style>
    </w:styles>`,
  );
  doc.applyCharacterStyle({ paragraph: 0, start: 0, end: 4 }, 'Emphasis', { clearDirectFormat: true });
  const run = doc.getParagraphs()[0].runs[0];
  assert.equal(run.style, 'Emphasis');
  assert.equal(run.effective.italic, true);
  assert.equal(run.effective.color, '224466');
  assert.equal(run.effective.bold, true);
  assert.equal(run.effective.underline, true);
  assert.doesNotMatch(doc.getPartXml(doc.mainDocumentPath), /w:color w:val="AA5500"/);
});

test('applyCharacterStyle clears direct formatting that would otherwise override inherited defaults', () => {
  const doc = withStyles(
    '<w:p><w:r><w:rPr><w:color w:val="AA5500"/></w:rPr><w:t>Word</w:t></w:r></w:p>',
    `<w:styles xmlns:w="${WORD_NS}">
      <w:docDefaults><w:rPrDefault><w:rPr><w:color w:val="224466"/></w:rPr></w:rPrDefault></w:docDefaults>
      <w:style w:type="character" w:styleId="Emphasis"><w:name w:val="Emphasis"/></w:style>
    </w:styles>`,
  );
  doc.applyCharacterStyle({ paragraph: 0, start: 0, end: 4 }, 'Emphasis', { clearDirectFormat: true });
  assert.equal(doc.getParagraphs()[0].runs[0].effective.color, '224466');
  assert.doesNotMatch(doc.getPartXml(doc.mainDocumentPath), /w:color w:val="AA5500"/);
});

test('createStyleFromSelection creates a paragraph style that reloads with ordered children', async () => {
  const doc = withBody('<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:before="80"/></w:pPr><w:r><w:rPr><w:b/><w:color w:val="445566"/></w:rPr><w:t>Pick</w:t></w:r></w:p>');
  const created = doc.createStyleFromSelection(
    { start: { paragraph: 0, offset: 0 }, end: { paragraph: 0, offset: 4 } },
    { id: 'FromSelection', name: 'From Selection' },
  );
  assert.equal(created.type, 'paragraph');
  assert.equal(created.paragraph.alignment, 'center');
  assert.equal(created.paragraph.spacingBefore, 80);
  assert.equal(created.run.bold, true);
  assert.equal(created.run.color, '445566');
  const xml = doc.getPartXml('word/styles.xml');
  const order = ['<w:name', '<w:pPr', '<w:rPr'].map((token) => xml.indexOf(token));
  assert.deepEqual(order.slice().sort((a, b) => a - b), order);
  const reopened = await DocxDocument.load(await doc.toUint8Array());
  assert.equal(reopened.getStyle('FromSelection').run.color, '445566');
});

test('createStyleFromSelection preserves effective run formatting from character styles', async () => {
  const doc = withStyles(
    '<w:p><w:r><w:rPr><w:rStyle w:val="Emphasis"/></w:rPr><w:t>Pick</w:t></w:r></w:p>',
    `<w:styles xmlns:w="${WORD_NS}">
      <w:style w:type="character" w:styleId="Emphasis"><w:name w:val="Emphasis"/><w:rPr><w:b/><w:color w:val="445566"/></w:rPr></w:style>
    </w:styles>`,
  );
  doc.createStyleFromSelection(
    { start: { paragraph: 0, offset: 0 }, end: { paragraph: 0, offset: 4 } },
    { id: 'FromSelection', name: 'From Selection' },
  );
  const reopened = await DocxDocument.load(await doc.toUint8Array());
  const style = reopened.getStyle('FromSelection');
  assert.equal(style.run.bold, true);
  assert.equal(style.run.color, '445566');
});

test('createStyleFromSelection rejects duplicate ids', () => {
  const doc = DocxDocument.create();
  doc.defineStyle({ id: 'Taken', name: 'Taken', type: 'paragraph' });
  assert.throws(() => doc.createStyleFromSelection(
    { start: { paragraph: 0, offset: 0 }, end: { paragraph: 0, offset: 0 } },
    { id: 'Taken', name: 'Taken Again' },
  ), /already exists/i);
});

test('createStyleFromSelection rejects empty or blank ids and names without writing styles.xml', () => {
  for (const style of [
    { id: '', name: 'Valid Name', message: /style\.id must not be empty or whitespace\./i },
    { id: '   ', name: 'Valid Name', message: /style\.id must not be empty or whitespace\./i },
    { id: 'ValidId', name: '', message: /style\.name must not be empty or whitespace\./i },
    { id: 'ValidId', name: '   ', message: /style\.name must not be empty or whitespace\./i },
  ]) {
    const doc = withBody('<w:p><w:r><w:t>Pick</w:t></w:r></w:p>');
    assert.throws(() => doc.createStyleFromSelection(
      { start: { paragraph: 0, offset: 0 }, end: { paragraph: 0, offset: 4 } },
      style,
    ), style.message);
    assert.equal(doc.listParts().includes('word/styles.xml'), false);
  }
});

test('getOutline infers levels from style names instead of style ids', () => {
  const doc = withStyles(
    '<w:p><w:pPr><w:pStyle w:val="Foo"/></w:pPr><w:r><w:t>Heading</w:t></w:r></w:p>',
    `<w:styles xmlns:w="${WORD_NS}">
      <w:style w:type="paragraph" w:styleId="Foo"><w:name w:val="heading 2"/></w:style>
    </w:styles>`,
  );
  assert.deepEqual(doc.getOutline(), [{ paragraph: 0, level: 1, text: 'Heading', styleId: 'Foo', children: [] }]);
});

test('getOutline prefers direct outlineLvl over heading style names', () => {
  const doc = withStyles(
    '<w:p><w:pPr><w:pStyle w:val="HeadingLike"/><w:outlineLvl w:val="4"/></w:pPr><w:r><w:t>Override</w:t></w:r></w:p>',
    `<w:styles xmlns:w="${WORD_NS}">
      <w:style w:type="paragraph" w:styleId="HeadingLike"><w:name w:val="heading 1"/></w:style>
    </w:styles>`,
  );
  assert.equal(doc.getOutline()[0].level, 4);
});

test('getOutline follows basedOn chains when a derived style has no heading name', () => {
  const doc = withStyles(
    '<w:p><w:pPr><w:pStyle w:val="Derived"/></w:pPr><w:r><w:t>Chain</w:t></w:r></w:p>',
    `<w:styles xmlns:w="${WORD_NS}">
      <w:style w:type="paragraph" w:styleId="BaseHeading"><w:name w:val="heading 3"/></w:style>
      <w:style w:type="paragraph" w:styleId="Derived"><w:name w:val="Custom Title"/><w:basedOn w:val="BaseHeading"/></w:style>
    </w:styles>`,
  );
  assert.equal(doc.getOutline()[0].level, 2);
});

test('getOutline caches builds until the revision changes', () => {
  const doc = withStyles(
    '<w:p><w:pPr><w:pStyle w:val="Heading"/></w:pPr><w:r><w:t>A</w:t></w:r></w:p>',
    `<w:styles xmlns:w="${WORD_NS}">
      <w:style w:type="paragraph" w:styleId="Heading"><w:name w:val="heading 1"/></w:style>
    </w:styles>`,
  );
  const original = doc.buildOutline;
  let calls = 0;
  doc.buildOutline = function (...args) {
    calls++;
    return original.apply(this, args);
  };
  assert.equal(doc.getOutline()[0].text, 'A');
  assert.equal(doc.getOutline()[0].text, 'A');
  assert.equal(calls, 1);
  doc.insertParagraph('B');
  doc.formatParagraph(1, { style: 'Heading' });
  assert.equal(doc.getOutline().at(-1).text, 'B');
  assert.equal(calls, 2);
});

test('setOutlineLevel writes and clears explicit outlineLvl', () => {
  const doc = withBody('<w:p><w:r><w:t>Outline</w:t></w:r></w:p>');
  doc.setOutlineLevel(0, 3);
  assert.equal(doc.getParagraphs()[0].outlineLevel, 3);
  assert.match(doc.getPartXml(doc.mainDocumentPath), /w:outlineLvl w:val="3"/);
  doc.setOutlineLevel(0, null);
  assert.equal(doc.getParagraphs()[0].outlineLevel, undefined);
  assert.doesNotMatch(doc.getPartXml(doc.mainDocumentPath), /w:outlineLvl/);
});

test('moveOutlineSection moves a heading subtree together with following tables in one revision', async () => {
  const doc = withStyles(
    [
      '<w:p><w:pPr><w:pStyle w:val="Heading"/></w:pPr><w:r><w:t>Alpha</w:t></w:r></w:p>',
      '<w:p><w:r><w:t>Alpha body</w:t></w:r></w:p>',
      '<w:p><w:pPr><w:pStyle w:val="Heading"/></w:pPr><w:r><w:t>Beta</w:t></w:r></w:p>',
      '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>Beta table</w:t></w:r></w:p></w:tc></w:tr></w:tbl>',
      '<w:p><w:r><w:t>Beta body</w:t></w:r></w:p>',
    ].join(''),
    `<w:styles xmlns:w="${WORD_NS}">
      <w:style w:type="paragraph" w:styleId="Heading"><w:name w:val="heading 1"/></w:style>
    </w:styles>`,
  );
  const revision = doc.revision;
  doc.moveOutlineSection(2, 0);
  assert.equal(doc.revision, revision + 1);
  assert.deepEqual(doc.getBlocks().map((block) => block.type === 'paragraph' ? block.paragraph.text : 'TABLE'), [
    'Beta',
    'TABLE',
    'Beta body',
    'Alpha',
    'Alpha body',
  ]);
  const reopened = await DocxDocument.load(await doc.toUint8Array());
  assert.deepEqual(reopened.getBlocks().map((block) => block.type === 'paragraph' ? block.paragraph.text : 'TABLE'), [
    'Beta',
    'TABLE',
    'Beta body',
    'Alpha',
    'Alpha body',
  ]);
});

test('moveOutlineSection is a no-op when moving within the same section span', () => {
  const doc = withStyles(
    '<w:p><w:pPr><w:pStyle w:val="Heading"/></w:pPr><w:r><w:t>Alpha</w:t></w:r></w:p><w:p><w:r><w:t>Body</w:t></w:r></w:p><w:p><w:pPr><w:pStyle w:val="Heading"/></w:pPr><w:r><w:t>Beta</w:t></w:r></w:p>',
    `<w:styles xmlns:w="${WORD_NS}">
      <w:style w:type="paragraph" w:styleId="Heading"><w:name w:val="heading 1"/></w:style>
    </w:styles>`,
  );
  const revision = doc.revision;
  doc.moveOutlineSection(0, 1);
  assert.equal(doc.revision, revision);
});

test('moveOutlineSection stops before the next sibling heading even when following sections contain deeper headings', () => {
  const doc = withStyles(
    [
      '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Root</w:t></w:r></w:p>',
      '<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>Beta</w:t></w:r></w:p>',
      '<w:p><w:r><w:t>Beta body</w:t></w:r></w:p>',
      '<w:p><w:pPr><w:pStyle w:val="Heading4"/></w:pPr><w:r><w:t>Beta child</w:t></w:r></w:p>',
      '<w:p><w:r><w:t>Beta child body</w:t></w:r></w:p>',
      '<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>Gamma</w:t></w:r></w:p>',
      '<w:p><w:r><w:t>Gamma body</w:t></w:r></w:p>',
      '<w:p><w:pPr><w:pStyle w:val="Heading3"/></w:pPr><w:r><w:t>Gamma child</w:t></w:r></w:p>',
      '<w:p><w:r><w:t>Gamma child body</w:t></w:r></w:p>',
    ].join(''),
    `<w:styles xmlns:w="${WORD_NS}">
      <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/></w:style>
      <w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/></w:style>
      <w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/></w:style>
      <w:style w:type="paragraph" w:styleId="Heading4"><w:name w:val="heading 4"/></w:style>
    </w:styles>`,
  );
  doc.moveOutlineSection(1, 9);
  assert.deepEqual(doc.getBlocks().map((block) => block.type === 'paragraph' ? block.paragraph.text : 'TABLE'), [
    'Root',
    'Gamma',
    'Gamma body',
    'Gamma child',
    'Gamma child body',
    'Beta',
    'Beta body',
    'Beta child',
    'Beta child body',
  ]);
});

test('applyOperations accepts new style and outline operations', () => {
  const doc = withStyles(
    '<w:p><w:r><w:t>A</w:t></w:r></w:p><w:p><w:r><w:t>B</w:t></w:r></w:p>',
    `<w:styles xmlns:w="${WORD_NS}">
      <w:style w:type="paragraph" w:styleId="Heading"><w:name w:val="heading 1"/></w:style>
      <w:style w:type="character" w:styleId="Emphasis"><w:name w:val="Emphasis"/><w:rPr><w:i/></w:rPr></w:style>
    </w:styles>`,
  );
  const snapshot = doc.applyOperations({
    operations: [
      { type: 'applyParagraphStyle', index: 0, styleId: 'Heading' },
      { type: 'applyCharacterStyle', range: { paragraph: 0, start: 0, end: 1 }, styleId: 'Emphasis' },
      { type: 'setOutlineLevel', index: 1, level: 2 },
      { type: 'moveOutlineSection', from: 1, to: 0 },
    ],
  });
  assert.equal(snapshot.paragraphs[0].text, 'B');
  assert.equal(snapshot.paragraphs[1].text, 'A');
  assert.equal(doc.getParagraphs()[1].runs[0].style, 'Emphasis');
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

test('read path marks anchor + javascript relationship hyperlinks as unsafe', () => {
  const doc = withBody(`<w:p xmlns:r="${OFFICE_REL_NS}"><w:hyperlink r:id="rId7" w:anchor="top"><w:r><w:t>x</w:t></w:r></w:hyperlink></w:p>`);
  doc.addPart('word/_rels/document.xml.rels', encoder.encode(`<Relationships xmlns="${REL_NS}">
    <Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="javascript:alert(1)" TargetMode="External"/>
  </Relationships>`), RELS_TYPE);
  const link = doc.getHyperlinks()[0];
  assert.equal(link.unsafe, true);
  assert.equal(doc.getParagraphs()[0].runs[0].hyperlink?.unsafe, true);
});

test('insertHyperlink supports runs inside hyperlink containers', () => {
  const doc = withBody(`<w:p xmlns:r="${OFFICE_REL_NS}"><w:hyperlink w:anchor="a"><w:r><w:t>linked text</w:t></w:r></w:hyperlink></w:p>`);
  const link = doc.insertHyperlink({ paragraph: 0, start: 0, end: 6 }, { url: 'https://example.com' });
  assert.equal(link.text, 'linked');
  assert.equal(doc.getHyperlinks().some((item) => item.url === 'https://example.com' && item.text === 'linked'), true);
});

test('insertHyperlink full-cover inside existing hyperlink rewrites target without nesting hyperlinks', () => {
  const doc = withBody(`<w:p xmlns:r="${OFFICE_REL_NS}"><w:hyperlink w:anchor="a"><w:r><w:t>linked</w:t></w:r></w:hyperlink></w:p>`);
  const link = doc.insertHyperlink({ paragraph: 0, start: 0, end: 6 }, { url: 'https://example.com' });
  const links = doc.getHyperlinks();
  assert.equal(links.length, 1);
  assert.equal(links[0]?.url, 'https://example.com');
  assert.equal(links[0]?.text, 'linked');
  assert.equal(links.some((item) => item.url === link.url && item.text === link.text), true);
  assert.doesNotMatch(doc.getPartXml(doc.mainDocumentPath), /<w:hyperlink[^>]*>\s*<w:hyperlink/);
});

test('insertHyperlink partial-cover inside existing hyperlink splits into sibling hyperlinks without nesting', () => {
  const doc = withBody(`<w:p xmlns:r="${OFFICE_REL_NS}"><w:hyperlink w:anchor="a"><w:r><w:t>linked text</w:t></w:r></w:hyperlink></w:p>`);
  const link = doc.insertHyperlink({ paragraph: 0, start: 0, end: 6 }, { url: 'https://example.com' });
  const links = doc.getHyperlinks();
  assert.equal(links.some((item) => item.url === 'https://example.com' && item.text === 'linked'), true);
  assert.equal(links.some((item) => item.anchor === 'a' && item.text === ' text'), true);
  assert.equal(links.some((item) => item.url === link.url && item.text === link.text), true);
  assert.doesNotMatch(doc.getPartXml(doc.mainDocumentPath), /<w:hyperlink[^>]*>\s*<w:hyperlink/);
});

test('insertHyperlink keeps insertion inside w:sdtContent for sdt-wrapped runs', () => {
  const doc = withBody('<w:p xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:sdt><w:sdtPr/><w:sdtContent><w:r><w:t>wrapped text</w:t></w:r></w:sdtContent></w:sdt></w:p>');
  const link = doc.insertHyperlink({ paragraph: 0, start: 0, end: 7 }, { url: 'https://example.com' });
  assert.equal(link.text, 'wrapped');
  const xml = doc.getPartXml(doc.mainDocumentPath);
  assert.match(xml, /<w:sdtContent><w:hyperlink [^>]*><w:r><w:rPr><w:rStyle w:val="Hyperlink"/);
  assert.doesNotMatch(xml, /<w:sdt><w:hyperlink /);
});

test('failed insertHyperlink does not change revision or leave orphan relationship', () => {
  const doc = DocxDocument.create();
  doc.setParagraphText(0, 'abc');
  const revision = doc.revision;
  assert.throws(() => doc.insertHyperlink({ paragraph: 0, start: 0, end: 99 }, { url: 'https://orphan.example' }), /out of bounds/);
  assert.equal(doc.revision, revision);
  assert.equal(doc.listParts().includes('word/_rels/document.xml.rels'), false);
});

test('insertHyperlink returns the inserted hyperlink, not last document hyperlink', () => {
  const doc = withBody(`<w:p xmlns:r="${OFFICE_REL_NS}"><w:r><w:t>hello world</w:t></w:r></w:p><w:p><w:hyperlink w:anchor="later"><w:r><w:t>later link</w:t></w:r></w:hyperlink></w:p>`);
  const link = doc.insertHyperlink({ paragraph: 0, start: 0, end: 5 }, { url: 'https://example.com' });
  assert.equal(link.paragraph, 0);
  assert.equal(link.text, 'hello');
  assert.equal(link.url, 'https://example.com');
});

test('applyOperations validates object-shaped hyperlink references', () => {
  const doc = DocxDocument.create();
  assert.throws(() => doc.applyOperations({ operations: [{ type: 'removeHyperlink', hyperlink: { paragraph: 0 } }] }), /hyperlink\.runs must be a non-empty array/);
});

test('body-level bookmarks are visible and block duplicate names', () => {
  const doc = withBody('<w:bookmarkStart w:id="5" w:name="tbl"/><w:p><w:r><w:t>x</w:t></w:r></w:p><w:bookmarkEnd w:id="5"/>');
  const bookmark = doc.getBookmarks({ includeInternal: true })[0];
  assert.equal(bookmark?.name, 'tbl');
  assert.equal(bookmark?.startParagraph, 0);
  assert.throws(() => doc.insertBookmark('tbl', { startParagraph: 0 }), /already exists/);
});

test('updateHyperlink writes fldSimple instruction with backslash-escaped quotes', () => {
  const doc = withBody('<w:p><w:fldSimple w:instr="HYPERLINK &quot;https://x.example&quot;"><w:r><w:t>x</w:t></w:r></w:fldSimple></w:p>');
  doc.updateHyperlink(0, { url: 'https://x.example/?q="a"' });
  assert.match(doc.getPartXml(doc.mainDocumentPath), /w:instr="HYPERLINK &quot;https:\/\/x\.example\/\?q=\\&quot;a\\&quot;&quot;"/);
});

test('snapshot includes hyperlinks and bookmarks', () => {
  const doc = withBody(`<w:bookmarkStart w:id="1" w:name="bm"/><w:p xmlns:r="${OFFICE_REL_NS}"><w:hyperlink w:anchor="bm"><w:r><w:t>x</w:t></w:r></w:hyperlink></w:p><w:bookmarkEnd w:id="1"/>`);
  const snapshot = doc.getSnapshot();
  assert.equal(snapshot.hyperlinks.length, 1);
  assert.equal(snapshot.bookmarks.length, 1);
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

test('note state fast path skips collection when main document has no note relationships', () => {
  const doc = withBody(Array.from({ length: 500 }, (_, index) => `<w:p><w:r><w:t>p${index}</w:t></w:r></w:p>`).join(''));
  const original = doc.collectNoteState;
  let calls = 0;
  doc.collectNoteState = function (...args) {
    calls++;
    return original.apply(this, args);
  };
  assert.equal(doc.getParagraphs().length, 500);
  assert.equal(doc.getBlocks().length, 500);
  const snapshot = doc.getSnapshot();
  assert.equal(snapshot.paragraphs.length, 500);
  assert.equal(snapshot.footnotes.length, 0);
  assert.equal(snapshot.endnotes.length, 0);
  assert.equal(calls, 0);
});

test('note state cache is reused across repeated reads until revision changes', () => {
  const doc = withBody(Array.from({ length: 20 }, (_, index) => `<w:p><w:r><w:t>p${index}</w:t></w:r></w:p>`).join(''));
  for (let index = 0; index < 20; index++) doc.insertFootnote(index, 1, `note ${index + 1}`);
  const original = doc.collectNoteState;
  let calls = 0;
  doc.collectNoteState = function (...args) {
    calls++;
    return original.apply(this, args);
  };
  assert.equal(doc.getParagraphs().length, 20);
  assert.equal(doc.getSnapshot().footnotes.length, 20);
  assert.equal(doc.getFootnotes().length, 20);
  assert.equal(calls, 1);
  doc.setNoteText('footnote', 1, 'updated');
  assert.equal(doc.getFootnotes()[0].blocks[0].paragraph.text.trim(), 'updated');
  assert.equal(calls, 2);
});

test('note state cache invalidates after note mutations and direct part rewrites', () => {
  const doc = DocxDocument.create();
  doc.setParagraphText(0, 'body');
  doc.getSnapshot();
  const footnote = doc.insertFootnote(0, 1, 'alpha');
  assert.equal(doc.getFootnotes()[0].id, footnote.id);
  doc.getSnapshot();
  const endnote = doc.insertEndnote(0, 2, 'omega');
  assert.equal(doc.getEndnotes()[0].id, endnote.id);

  doc.getSnapshot();
  doc.setNoteText('footnote', footnote.id, 'beta');
  assert.equal(doc.getFootnotes().find((item) => item.id === footnote.id)?.blocks[0]?.paragraph.text.trim(), 'beta');

  doc.getSnapshot();
  doc.setNoteSettings({ footnote: { numFmt: 'lowerRoman' } });
  assert.equal(doc.getFootnotes().find((item) => item.id === footnote.id)?.marker, 'i');

  doc.getSnapshot();
  doc.setPartXml('word/footnotes.xml', doc.getPartXml('word/footnotes.xml').replace('beta', 'gamma'));
  assert.equal(doc.getFootnotes().find((item) => item.id === footnote.id)?.blocks[0]?.paragraph.text.trim(), 'gamma');

  doc.getSnapshot();
  doc.updatePartXml('word/footnotes.xml', (document) => {
    document.getElementsByTagNameNS(WORD_NS, 't')[0].textContent = 'delta';
  });
  assert.equal(doc.getFootnotes().find((item) => item.id === footnote.id)?.blocks[0]?.paragraph.text.trim(), 'delta');

  doc.getSnapshot();
  doc.convertNote('footnote', footnote.id);
  assert.equal(doc.getFootnotes().length, 0);
  assert.equal(doc.getEndnotes().some((item) => item.id === footnote.id), true);

  doc.getSnapshot();
  doc.deleteNote('endnote', footnote.id);
  assert.equal(doc.getEndnotes().some((item) => item.id === footnote.id), false);
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

test('addComment creates comment parts and preserves visible paragraph text', () => {
  const doc = DocxDocument.create();
  doc.setParagraphText(0, 'body text');
  const id = doc.addComment({ paragraph: 0, start: 0, end: 4 }, { author: 'Alice', initials: 'AL', text: 'todo' });
  const comment = doc.getComments()[0];
  assert.equal(id, comment.id);
  assert.equal(doc.getParagraphs()[0].text, 'body text');
  assert.equal(comment.author, 'Alice');
  assert.equal(comment.anchor.paragraph, 0);
  assert.deepEqual(comment.anchor.runs, [0]);
  assert.match(doc.getPartXml('word/comments.xml'), /<w:rPr><w:rStyle w:val="CommentText"\/><\/w:rPr><w:t xml:space="preserve">todo<\/w:t>/);
  assert.match(doc.getPartXml('word/comments.xml'), /todo/);
  assert.match(doc.getPartXml('word/commentsExtended.xml'), /w15:commentEx/);
});

test('getComments reads replies and resolved state from commentsExtended.xml', () => {
  const doc = withCommentsDoc(
    '<w:p><w:commentRangeStart w:id="1"/><w:r><w:t>x</w:t></w:r><w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r></w:p>',
    `<w:comments xmlns:w="${WORD_NS}" xmlns:w14="${W14_NS}">
      <w:comment w:id="1" w:author="Alice"><w:p w14:paraId="0000000A"><w:r><w:annotationRef/></w:r><w:r><w:t>root</w:t></w:r></w:p></w:comment>
      <w:comment w:id="2" w:author="Bob"><w:p w14:paraId="0000000B"><w:r><w:annotationRef/></w:r><w:r><w:t>reply</w:t></w:r></w:p></w:comment>
    </w:comments>`,
    `<w15:commentsEx xmlns:w15="${W15_NS}">
      <w15:commentEx w15:paraId="0000000A" w15:done="1"/>
      <w15:commentEx w15:paraId="0000000B" w15:paraIdParent="0000000A" w15:done="0"/>
    </w15:commentsEx>`,
  );
  const comments = doc.getComments();
  assert.equal(comments.length, 2);
  assert.equal(comments[0].resolved, true);
  assert.equal(comments[1].parentId, 1);
  assert.equal(comments[1].isOrphan, false);
});

test('getComments reads header comment anchors with source part paths', () => {
  const doc = DocxDocument.create();
  doc.setPartXml(doc.mainDocumentPath, `<w:document xmlns:w="${WORD_NS}" xmlns:r="${OFFICE_REL_NS}"><w:body><w:p><w:r><w:t>body</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rId10"/></w:sectPr></w:body></w:document>`);
  doc.addPart('word/_rels/document.xml.rels', encoder.encode(`<Relationships xmlns="${REL_NS}">
    <Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="headerX.xml"/>
  </Relationships>`), RELS_TYPE);
  doc.addPart('word/headerX.xml', encoder.encode(`<w:hdr xmlns:w="${WORD_NS}"><w:p><w:commentRangeStart w:id="3"/><w:r><w:t>head</w:t></w:r><w:commentRangeEnd w:id="3"/><w:r><w:commentReference w:id="3"/></w:r></w:p></w:hdr>`), 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml');
  doc.addPart('word/_rels/headerX.xml.rels', encoder.encode(`<Relationships xmlns="${REL_NS}">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="hdr-comments.xml"/>
  </Relationships>`), RELS_TYPE);
  doc.addPart('word/hdr-comments.xml', encoder.encode(`<w:comments xmlns:w="${WORD_NS}" xmlns:w14="${W14_NS}"><w:comment w:id="3"><w:p w14:paraId="00000033"><w:r><w:annotationRef/></w:r><w:r><w:t>H</w:t></w:r></w:p></w:comment></w:comments>`), COMMENTS_TYPE);
  const comment = doc.getComments().find((item) => item.id === 3);
  assert.equal(comment?.anchor.sourcePartPath, 'word/headerX.xml');
  assert.equal(comment?.anchor.paragraph, 0);
});

test('getComments reads footnote comment anchors with source part paths', () => {
  const doc = DocxDocument.create();
  doc.setPartXml(doc.mainDocumentPath, `<w:document xmlns:w="${WORD_NS}"><w:body><w:p><w:r><w:t>x</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`);
  doc.addPart('word/_rels/document.xml.rels', encoder.encode(`<Relationships xmlns="${REL_NS}">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="fn.xml"/>
  </Relationships>`), RELS_TYPE);
  doc.addPart('word/fn.xml', encoder.encode(`<w:footnotes xmlns:w="${WORD_NS}">
    <w:footnote w:id="1"><w:p><w:commentRangeStart w:id="4"/><w:r><w:t>fn</w:t></w:r><w:commentRangeEnd w:id="4"/><w:r><w:commentReference w:id="4"/></w:r></w:p></w:footnote>
  </w:footnotes>`), 'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml');
  doc.addPart('word/_rels/fn.xml.rels', encoder.encode(`<Relationships xmlns="${REL_NS}">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="fn-comments.xml"/>
  </Relationships>`), RELS_TYPE);
  doc.addPart('word/fn-comments.xml', encoder.encode(`<w:comments xmlns:w="${WORD_NS}" xmlns:w14="${W14_NS}"><w:comment w:id="4"><w:p w14:paraId="00000044"><w:r><w:annotationRef/></w:r><w:r><w:t>FN</w:t></w:r></w:p></w:comment></w:comments>`), COMMENTS_TYPE);
  const comment = doc.getComments().find((item) => item.id === 4);
  assert.equal(comment?.anchor.sourcePartPath, 'word/fn.xml');
  assert.equal(comment?.anchor.paragraph, 0);
});

test('addComment across hyperlink keeps hyperlink intact', () => {
  const doc = withBody(`<w:p xmlns:r="${OFFICE_REL_NS}"><w:hyperlink w:anchor="x"><w:r><w:t>link</w:t></w:r></w:hyperlink><w:r><w:t>tail</w:t></w:r></w:p>`);
  doc.addComment({ start: { paragraph: 0, offset: 1 }, end: { paragraph: 0, offset: 6 } }, { text: 'cross' });
  assert.equal(doc.getHyperlinks().some((item) => item.anchor === 'x' && item.text.includes('link')), true);
});

test('addComment on a collapsed range keeps end marker and reference after the start marker', () => {
  const doc = DocxDocument.create();
  doc.setParagraphText(0, 'body');
  doc.addComment({ paragraph: 0, start: 2, end: 2 }, { text: 'caret' });
  const xml = doc.getPartXml(doc.mainDocumentPath);
  assert.match(xml, /<w:commentRangeStart w:id="1"\/><w:commentRangeEnd w:id="1"\/><w:r><w:rPr><w:rStyle w:val="CommentReference"\/><\/w:rPr><w:commentReference w:id="1"\/><\/w:r>/);
  assert.ok(xml.indexOf('<w:commentRangeStart w:id="1"/>') < xml.indexOf('<w:commentRangeEnd w:id="1"/>'));
  assert.ok(xml.indexOf('<w:commentRangeEnd w:id="1"/>') < xml.indexOf('<w:commentReference w:id="1"/>'));
});

test('comment mutations advance revision once per public call', () => {
  const doc = DocxDocument.create();
  doc.setParagraphText(0, 'abc');
  const start = doc.revision;
  const id = doc.addComment({ paragraph: 0, start: 0, end: 1 }, { text: 'one' });
  assert.equal(doc.revision, start + 1);
  doc.deleteComment(id);
  assert.equal(doc.revision, start + 2);
});

test('failed addComment is atomic', () => {
  const doc = DocxDocument.create();
  doc.setParagraphText(0, 'abc');
  const revision = doc.revision;
  const parts = doc.listParts();
  assert.throws(() => doc.addComment({ paragraph: 0, start: 0, end: 99 }, { text: 'bad' }), /out of bounds/);
  assert.equal(doc.revision, revision);
  assert.deepEqual(doc.listParts(), parts);
});

test('comment reads fast-path to empty without scanning anchors when markup is absent', () => {
  const doc = withBody('<w:p><w:r><w:t>x</w:t></w:r></w:p>');
  const original = doc.collectCommentLocations;
  let calls = 0;
  doc.collectCommentLocations = function (...args) {
    calls++;
    return original.apply(this, args);
  };
  assert.equal(doc.getComments().length, 0);
  assert.equal(doc.getSnapshot().comments.length, 0);
  assert.equal(calls, 0);
});

test('comment cache is reused until revision changes', () => {
  const doc = DocxDocument.create();
  doc.setParagraphText(0, 'abc');
  const id = doc.addComment({ paragraph: 0, start: 0, end: 1 }, { text: 'one' });
  const original = doc.collectCommentLocations;
  let calls = 0;
  doc.collectCommentLocations = function (...args) {
    calls++;
    return original.apply(this, args);
  };
  assert.equal(doc.getComments().length, 1);
  assert.equal(doc.getSnapshot().comments.length, 1);
  assert.equal(doc.getComments()[0].id, id);
  assert.equal(calls, 1);
  doc.setCommentText(id, 'two');
  assert.equal(doc.getComments()[0].text, 'two');
  assert.equal(calls, 2);
});

test('missing comment body for a reference is reported as orphan', () => {
  const doc = withBody('<w:p><w:commentRangeStart w:id="9"/><w:r><w:t>x</w:t></w:r><w:commentRangeEnd w:id="9"/><w:r><w:commentReference w:id="9"/></w:r></w:p>');
  const comment = doc.getComments()[0];
  assert.equal(comment.id, 9);
  assert.equal(comment.isOrphan, true);
  assert.equal(comment.text, '');
});

test('comment body without any anchor is reported as orphan', () => {
  const doc = withCommentsDoc(
    '<w:p><w:r><w:t>x</w:t></w:r></w:p>',
    `<w:comments xmlns:w="${WORD_NS}" xmlns:w14="${W14_NS}"><w:comment w:id="8"><w:p w14:paraId="00000088"><w:r><w:annotationRef/></w:r><w:r><w:t>body</w:t></w:r></w:p></w:comment></w:comments>`,
  );
  assert.equal(doc.getComments()[0].isOrphan, true);
});

test('invalid comment dates degrade to undefined', () => {
  const doc = withCommentsDoc(
    '<w:p><w:commentRangeStart w:id="1"/><w:r><w:t>x</w:t></w:r><w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r></w:p>',
    `<w:comments xmlns:w="${WORD_NS}" xmlns:w14="${W14_NS}"><w:comment w:id="1" w:date="not-a-date"><w:p w14:paraId="00000011"><w:r><w:annotationRef/></w:r></w:p></w:comment></w:comments>`,
  );
  assert.equal(doc.getComments()[0].date, undefined);
});

test('setCommentResolved creates commentsExtended when missing', () => {
  const doc = withCommentsDoc(
    '<w:p><w:commentRangeStart w:id="1"/><w:r><w:t>x</w:t></w:r><w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r></w:p>',
    `<w:comments xmlns:w="${WORD_NS}" xmlns:w14="${W14_NS}"><w:comment w:id="1"><w:p w14:paraId="00000021"><w:r><w:annotationRef/></w:r></w:p></w:comment></w:comments>`,
  );
  doc.setCommentResolved(1, true);
  assert.equal(doc.getComments()[0].resolved, true);
  assert.match(doc.getPartXml('word/commentsExtended.xml'), /w15:done="1"/);
});

test('setCommentText replaces comment text instead of appending', () => {
  const doc = DocxDocument.create();
  doc.setParagraphText(0, 'abc');
  const id = doc.addComment({ paragraph: 0, start: 0, end: 1 }, { text: 'old' });
  doc.setCommentText(id, 'new');
  const xml = doc.getPartXml('word/comments.xml');
  assert.match(xml, /new/);
  assert.doesNotMatch(xml, /old.*new/);
});

test('setCommentText updates existing header comment bodies', () => {
  const doc = DocxDocument.create();
  doc.setPartXml(doc.mainDocumentPath, `<w:document xmlns:w="${WORD_NS}" xmlns:r="${OFFICE_REL_NS}"><w:body><w:p><w:r><w:t>body</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rId10"/></w:sectPr></w:body></w:document>`);
  doc.addPart('word/_rels/document.xml.rels', encoder.encode(`<Relationships xmlns="${REL_NS}">
    <Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="headerX.xml"/>
  </Relationships>`), RELS_TYPE);
  doc.addPart('word/headerX.xml', encoder.encode(`<w:hdr xmlns:w="${WORD_NS}"><w:p><w:commentRangeStart w:id="3"/><w:r><w:t>head</w:t></w:r><w:commentRangeEnd w:id="3"/><w:r><w:commentReference w:id="3"/></w:r></w:p></w:hdr>`), 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml');
  doc.addPart('word/_rels/headerX.xml.rels', encoder.encode(`<Relationships xmlns="${REL_NS}">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="hdr-comments.xml"/>
  </Relationships>`), RELS_TYPE);
  doc.addPart('word/hdr-comments.xml', encoder.encode(`<w:comments xmlns:w="${WORD_NS}" xmlns:w14="${W14_NS}"><w:comment w:id="3"><w:p w14:paraId="00000033"><w:r><w:annotationRef/></w:r><w:r><w:t>old</w:t></w:r></w:p></w:comment></w:comments>`), COMMENTS_TYPE);
  doc.setCommentText(3, 'new');
  assert.match(doc.getPartXml('word/hdr-comments.xml'), /new/);
  assert.doesNotMatch(doc.getPartXml('word/hdr-comments.xml'), /old.*new/);
  assert.equal(doc.getComments().find((item) => item.id === 3)?.text, 'new');
});

test('deleteComment removes anchors and orphan relationships when last comment is removed', () => {
  const doc = DocxDocument.create();
  doc.setParagraphText(0, 'abc');
  const id = doc.addComment({ paragraph: 0, start: 0, end: 1 }, { text: 'gone' });
  doc.deleteComment(id);
  assert.equal(doc.getComments().length, 0);
  assert.equal(doc.listParts().includes('word/comments.xml'), false);
  assert.equal(doc.listParts().includes('word/commentsExtended.xml'), false);
  assert.doesNotMatch(doc.getPartXml(doc.mainDocumentPath), /commentRange(Start|End)|commentReference/);
});

test('deleteComment removes existing header comment anchors and parts', () => {
  const doc = DocxDocument.create();
  doc.setPartXml(doc.mainDocumentPath, `<w:document xmlns:w="${WORD_NS}" xmlns:r="${OFFICE_REL_NS}"><w:body><w:p><w:r><w:t>body</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rId10"/></w:sectPr></w:body></w:document>`);
  doc.addPart('word/_rels/document.xml.rels', encoder.encode(`<Relationships xmlns="${REL_NS}">
    <Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="headerX.xml"/>
  </Relationships>`), RELS_TYPE);
  doc.addPart('word/headerX.xml', encoder.encode(`<w:hdr xmlns:w="${WORD_NS}"><w:p><w:commentRangeStart w:id="3"/><w:r><w:t>head</w:t></w:r><w:commentRangeEnd w:id="3"/><w:r><w:commentReference w:id="3"/></w:r></w:p></w:hdr>`), 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml');
  doc.addPart('word/_rels/headerX.xml.rels', encoder.encode(`<Relationships xmlns="${REL_NS}">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="hdr-comments.xml"/>
  </Relationships>`), RELS_TYPE);
  doc.addPart('word/hdr-comments.xml', encoder.encode(`<w:comments xmlns:w="${WORD_NS}" xmlns:w14="${W14_NS}"><w:comment w:id="3"><w:p w14:paraId="00000033"><w:r><w:annotationRef/></w:r><w:r><w:t>old</w:t></w:r></w:p></w:comment></w:comments>`), COMMENTS_TYPE);
  doc.deleteComment(3);
  assert.equal(doc.getComments().length, 0);
  assert.equal(doc.listParts().includes('word/hdr-comments.xml'), false);
  assert.doesNotMatch(doc.getPartXml('word/headerX.xml'), /commentRange(Start|End)|commentReference/);
});

test('deleteComment cascades to replies by default', () => {
  const doc = DocxDocument.create();
  doc.setParagraphText(0, 'abc');
  const id = doc.addComment({ paragraph: 0, start: 0, end: 1 }, { text: 'root' });
  doc.replyComment(id, { text: 'child' });
  doc.deleteComment(id);
  assert.equal(doc.getComments().length, 0);
});

test('replyComment writes paraIdParent relation to commentsExtended', () => {
  const doc = DocxDocument.create();
  doc.setParagraphText(0, 'abc');
  const parentId = doc.addComment({ paragraph: 0, start: 0, end: 1 }, { text: 'root' });
  const childId = doc.replyComment(parentId, { text: 'child' });
  const child = doc.getComments().find((item) => item.id === childId);
  assert.equal(child?.parentId, parentId);
  assert.match(doc.getPartXml('word/commentsExtended.xml'), /paraIdParent/);
});

test('getComments supports author and resolved filters', () => {
  const doc = DocxDocument.create();
  doc.setParagraphText(0, 'abc');
  const first = doc.addComment({ paragraph: 0, start: 0, end: 1 }, { author: 'Alice', text: 'a' });
  doc.replyComment(first, { author: 'Bob', text: 'b' });
  doc.setCommentResolved(first, true);
  assert.equal(doc.getComments({ authors: ['Alice'] }).length, 1);
  assert.equal(doc.getComments({ resolved: true }).length, 1);
});

test('applyOperations supports comment operations and schema count stays aligned', () => {
  const doc = DocxDocument.create();
  doc.setParagraphText(0, 'abc');
  const snapshot = doc.applyOperations({
    operations: [{ type: 'addComment', range: { paragraph: 0, start: 0, end: 1 }, comment: { text: 'a' } }],
  });
  assert.equal(snapshot.comments.length, 1);
  assert.equal(AGENT_OPERATION_SCHEMA.properties.operations.items.oneOf.length, 62);
  assert.throws(() => doc.applyOperations({ operations: [{ type: 'replyComment', parentId: 0, comment: {} }] }), /comment\.text/);
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
  doc.addPart('word/settings.xml', encoder.encode(`<w:settings xmlns:w="${WORD_NS}"><w:defaultTabStop w:val="1440"/><w:evenAndOddHeaders/></w:settings>`), SETTINGS_TYPE);
  assert.deepEqual(doc.getSettings(), { defaultTabStop: 1440, evenAndOddHeaders: true, trackChanges: false });
});

test('getSettings reads explicit trackChanges off and on', () => {
  const disabled = withSettingsXml(`<w:settings xmlns:w="${WORD_NS}"><w:trackChanges w:val="0"/></w:settings>`);
  assert.equal(disabled.getSettings().trackChanges, false);
  const enabled = withSettingsXml(`<w:settings xmlns:w="${WORD_NS}"><w:trackChanges/></w:settings>`);
  assert.equal(enabled.getSettings().trackChanges, true);
});

test('setTrackChanges creates settings.xml and keeps CT_Settings order', () => {
  const doc = DocxDocument.create();
  doc.setTrackChanges(true);
  const settingsXml = doc.getPartXml('word/settings.xml');
  assert.match(settingsXml, /<w:trackChanges\/>/);
  assert.match(doc.getPartXml('word/_rels/document.xml.rels'), /relationships\/settings" Target="settings\.xml"/);
  const ordered = withSettingsXml(`<w:settings xmlns:w="${WORD_NS}"><w:revisionView/><w:doNotTrackMoves/></w:settings>`);
  ordered.setTrackChanges(true);
  assert.match(ordered.getPartXml('word/settings.xml'), /<w:revisionView\/><w:trackChanges\/><w:doNotTrackMoves\/>/);
});

test('setTrackChanges(false) writes explicit off value after being enabled', () => {
  const doc = DocxDocument.create();
  doc.setTrackChanges(true);
  doc.setTrackChanges(false);
  assert.equal(doc.getSettings().trackChanges, false);
  assert.match(doc.getPartXml('word/settings.xml'), /<w:trackChanges w:val="0"\/>/);
});

test('setTrackChanges(false) writes explicit off markup when settings were previously implicit', () => {
  const doc = DocxDocument.create();
  doc.setTrackChanges(false);
  assert.match(doc.getPartXml('word/settings.xml'), /<w:trackChanges w:val="0"\/>/);
});

test('setTrackChanges is a no-op when the explicit XML state is unchanged', () => {
  const disabled = withSettingsXml(`<w:settings xmlns:w="${WORD_NS}"><w:trackChanges w:val="0"/></w:settings>`);
  const disabledRevision = disabled.revision;
  disabled.setTrackChanges(false);
  assert.equal(disabled.revision, disabledRevision);
  const enabled = withSettingsXml(`<w:settings xmlns:w="${WORD_NS}"><w:trackChanges/></w:settings>`);
  const enabledRevision = enabled.revision;
  enabled.setTrackChanges(true);
  assert.equal(enabled.revision, enabledRevision);
});

test('getParagraphs keeps inserted text in body text and marks the run revision', () => {
  const doc = withBody('<w:p><w:ins w:id="7" w:author="Alice" w:date="2026-09-28T00:00:00Z"><w:r><w:t>inserted</w:t></w:r></w:ins><w:r><w:t>keep</w:t></w:r></w:p>');
  const paragraph = doc.getParagraphs()[0];
  assert.equal(paragraph.text, 'insertedkeep');
  assert.equal(paragraph.runs[0].text, 'inserted');
  assert.deepEqual(paragraph.runs[0].revisions, [{ id: 7, kind: 'insertion', author: 'Alice', date: '2026-09-28T00:00:00Z' }]);
});

test('deleted text stays out of paragraph text but keeps an empty run slot', () => {
  const doc = withBody('<w:p><w:del w:id="8"><w:r><w:delText>old</w:delText></w:r></w:del><w:r><w:t>keep</w:t></w:r></w:p>');
  const paragraph = doc.getParagraphs()[0];
  assert.equal(paragraph.text, 'keep');
  assert.equal(paragraph.runs.length, 2);
  assert.equal(paragraph.runs[0].text, '');
  assert.deepEqual(doc.getRevisions(), [{ id: 8, kind: 'deletion', paragraph: 0, run: 0, deletedText: 'old' }]);
});

test('nested ins/del markup degrades without throwing and preserves visible text rules', () => {
  const doc = withBody('<w:p><w:ins w:id="1"><w:del w:id="2"><w:r><w:delText>old</w:delText></w:r></w:del><w:r><w:t>new</w:t></w:r></w:ins></w:p>');
  assert.doesNotThrow(() => doc.getParagraphs());
  assert.equal(doc.getParagraphs()[0].text, 'new');
  assert.deepEqual(doc.getRevisions().map((revision) => revision.kind), ['insertion', 'deletion']);
});

test('getRevisions ignores malformed metadata but does not throw', () => {
  const doc = withBody('<w:p><w:ins w:author="" w:date="not-a-date"><w:r><w:t>bad</w:t></w:r></w:ins><w:del w:id="9" w:author="" w:date="bad"><w:r><w:delText>old</w:delText></w:r></w:del></w:p>');
  assert.doesNotThrow(() => doc.getRevisions());
  assert.deepEqual(doc.getRevisions(), [{ id: 9, kind: 'deletion', paragraph: 0, run: 1, deletedText: 'old' }]);
});

test('getRevisions filters by author and kind', () => {
  const doc = withBody('<w:p><w:ins w:id="1" w:author="Alice"><w:r><w:t>A</w:t></w:r></w:ins><w:del w:id="2" w:author="Bob"><w:r><w:delText>B</w:delText></w:r></w:del></w:p>');
  assert.deepEqual(doc.getRevisions({ authors: ['Alice'] }).map((revision) => revision.id), [1]);
  assert.deepEqual(doc.getRevisions({ kinds: ['deletion'] }).map((revision) => revision.id), [2]);
});

test('getRevisions reads run-property insertion and deletion markers', () => {
  const doc = withBody('<w:p><w:r><w:rPr><w:ins w:id="10"/><w:del w:id="11"/></w:rPr><w:t>A</w:t></w:r></w:p>');
  const revisions = doc.getRevisions();
  assert.deepEqual(revisions.map((revision) => ({ id: revision.id, kind: revision.kind, run: revision.run })), [
    { id: 10, kind: 'insertion', run: 0 },
    { id: 11, kind: 'deletion', run: 0 },
  ]);
});

test('getRevisions reads rPrChange previous format snapshots', () => {
  const doc = withBody('<w:p><w:r><w:rPr><w:rPrChange w:id="12" w:author="Alice"><w:rPr><w:b/><w:color w:val="FF0000"/></w:rPr></w:rPrChange></w:rPr><w:t>A</w:t></w:r></w:p>');
  const paragraph = doc.getParagraphs()[0];
  assert.deepEqual(paragraph.runs[0].revisions, [{ id: 12, kind: 'runFormatChange', author: 'Alice' }]);
  assert.deepEqual(doc.getRevisions()[0].previousFormat, { bold: true, color: 'FF0000' });
});

test('getParagraphs exposes paragraphRevision and getRevisions reads pPrChange snapshots', () => {
  const doc = withBody('<w:p><w:pPr><w:pPrChange w:id="13"><w:pPr><w:jc w:val="center"/></w:pPr></w:pPrChange></w:pPr><w:r><w:t>A</w:t></w:r></w:p>');
  assert.deepEqual(doc.getParagraphs()[0].paragraphRevision, { id: 13, kind: 'paragraphFormatChange' });
  assert.deepEqual(doc.getRevisions()[0].previousFormat, { alignment: 'center' });
});

test('getRevisions reads table, row, and cell format changes', () => {
  const doc = withBody('<w:tbl><w:tblPr><w:tblPrChange w:id="14"/></w:tblPr><w:tr><w:trPr><w:trPrChange w:id="15"/></w:trPr><w:tc><w:tcPr><w:tcPrChange w:id="16"/></w:tcPr><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc></w:tr></w:tbl>');
  assert.deepEqual(doc.getRevisions().map((revision) => ({ id: revision.id, kind: revision.kind, paragraph: revision.paragraph })), [
    { id: 14, kind: 'tableFormatChange', paragraph: 0 },
    { id: 15, kind: 'rowFormatChange', paragraph: 0 },
    { id: 16, kind: 'cellFormatChange', paragraph: 0 },
  ]);
});

test('getRevisions inside table cells uses the same paragraph index namespace as getParagraphs', () => {
  const doc = withBody('<w:p><w:r><w:t>before</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:p><w:ins w:id="21"><w:r><w:t>cell</w:t></w:r></w:ins></w:p></w:tc></w:tr></w:tbl><w:p><w:r><w:t>after</w:t></w:r></w:p>');
  const paragraphs = doc.getParagraphs();
  const cellParagraph = paragraphs.find((paragraph) => paragraph.text === 'cell');
  assert.ok(cellParagraph);
  assert.deepEqual(doc.getRevisions(), [{ id: 21, kind: 'insertion', paragraph: cellParagraph.index, run: 0 }]);
});

test('moveFrom and moveTo revisions are downgraded to deletion and insertion', () => {
  const doc = withBody('<w:p><w:moveFrom w:id="17"><w:r><w:delText>old</w:delText></w:r></w:moveFrom><w:moveTo w:id="18"><w:r><w:t>new</w:t></w:r></w:moveTo></w:p>');
  assert.deepEqual(doc.getRevisions().map((revision) => ({ id: revision.id, kind: revision.kind })), [
    { id: 17, kind: 'deletion' },
    { id: 18, kind: 'insertion' },
  ]);
  assert.equal(doc.getParagraphs()[0].text, 'new');
});

test('inserted images remain discoverable through transparent revision wrappers', () => {
  const doc = withImageDoc(
    '<w:p><w:ins w:id="19"><w:r><w:drawing><wp:inline><wp:extent cx="952500" cy="476250"/><wp:docPr id="1" name="img"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="rId1"/></pic:blipFill><pic:spPr/></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:ins></w:p>',
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>',
  );
  assert.equal(doc.getImages().length, 1);
  assert.equal(doc.getImages()[0].relationshipId, 'rId1');
});

test('getRevisions uses a per-revision cache and invalidates after mutation', () => {
  const doc = withBody('<w:p><w:ins w:id="20"><w:r><w:t>A</w:t></w:r></w:ins></w:p>');
  const first = doc.getRevisions();
  const second = doc.getRevisions();
  assert.equal(first, second);
  doc.setTrackChanges(true);
  const third = doc.getRevisions();
  assert.notEqual(third, first);
});

test('getReviewers aggregates revision and comment counts with date range', () => {
  const body = '<w:p><w:ins w:id="1" w:author="Alice" w:date="2026-01-02T03:04:05Z"><w:r><w:t>A</w:t></w:r></w:ins><w:del w:id="2" w:author="Alice" w:date="2026-01-04T03:04:05Z"><w:r><w:delText>B</w:delText></w:r></w:del><w:r><w:commentReference w:id="1"/></w:r><w:r><w:commentReference w:id="2"/></w:r></w:p>';
  const comments = `<w:comments xmlns:w="${WORD_NS}" xmlns:w14="${W14_NS}"><w:comment w:id="1" w:author="Alice" w:initials="AL" w:date="2026-01-01T00:00:00Z"><w:p w14:paraId="00000001"><w:r><w:t>one</w:t></w:r></w:p></w:comment><w:comment w:id="2" w:author="Alice" w:initials="A" w:date="2026-01-05T00:00:00Z"><w:p w14:paraId="00000002"><w:r><w:t>two</w:t></w:r></w:p></w:comment></w:comments>`;
  const commentsEx = `<w15:commentsEx xmlns:w15="${W15_NS}"><w15:commentEx w15:paraId="00000001" w15:done="1"/><w15:commentEx w15:paraId="00000002" w15:done="0"/></w15:commentsEx>`;
  const doc = withCommentsDoc(body, comments, commentsEx);
  const reviewer = doc.getReviewers().find((item) => item.author === 'Alice');
  assert.deepEqual(reviewer, {
    author: 'Alice',
    initials: 'AL',
    revisionCount: 2,
    commentCount: 2,
    unresolvedCommentCount: 1,
    firstDate: '2026-01-01T00:00:00Z',
    lastDate: '2026-01-05T00:00:00Z',
  });
});

test('getReviewers keeps missing, empty, and blank authors in explicit buckets', () => {
  const body = '<w:p><w:ins w:id="1"><w:r><w:t>A</w:t></w:r></w:ins><w:r><w:commentReference w:id="1"/></w:r><w:r><w:commentReference w:id="2"/></w:r><w:r><w:commentReference w:id="3"/></w:r></w:p>';
  const comments = `<w:comments xmlns:w="${WORD_NS}"><w:comment w:id="1"><w:p><w:r><w:t>missing</w:t></w:r></w:p></w:comment><w:comment w:id="2" w:author=""><w:p><w:r><w:t>empty</w:t></w:r></w:p></w:comment><w:comment w:id="3" w:author="   "><w:p><w:r><w:t>blank</w:t></w:r></w:p></w:comment></w:comments>`;
  const doc = withCommentsDoc(body, comments);
  const map = new Map(doc.getReviewers().map((item) => [item.author, item]));
  assert.equal(map.get('(unattributed)')?.revisionCount, 1);
  assert.equal(map.get('(unattributed)')?.commentCount, 1);
  assert.equal(map.get('(empty author)')?.commentCount, 1);
  assert.equal(map.get('(blank author)')?.commentCount, 1);
});

test('getReviewers picks the most frequent initials per author', () => {
  const body = '<w:p><w:r><w:commentReference w:id="1"/></w:r><w:r><w:commentReference w:id="2"/></w:r><w:r><w:commentReference w:id="3"/></w:r></w:p>';
  const comments = `<w:comments xmlns:w="${WORD_NS}"><w:comment w:id="1" w:author="Alice" w:initials="AA"><w:p><w:r><w:t>1</w:t></w:r></w:p></w:comment><w:comment w:id="2" w:author="Alice" w:initials="A"><w:p><w:r><w:t>2</w:t></w:r></w:p></w:comment><w:comment w:id="3" w:author="Alice" w:initials="AA"><w:p><w:r><w:t>3</w:t></w:r></w:p></w:comment></w:comments>`;
  const doc = withCommentsDoc(body, comments);
  assert.equal(doc.getReviewers().find((item) => item.author === 'Alice')?.initials, 'AA');
});

test('getReviewers keeps initials selection stable on frequency ties by first appearance', () => {
  const body = '<w:p><w:r><w:commentReference w:id="1"/></w:r><w:r><w:commentReference w:id="2"/></w:r></w:p>';
  const comments = `<w:comments xmlns:w="${WORD_NS}"><w:comment w:id="1" w:author="Alice" w:initials="ZX"><w:p><w:r><w:t>1</w:t></w:r></w:p></w:comment><w:comment w:id="2" w:author="Alice" w:initials="AB"><w:p><w:r><w:t>2</w:t></w:r></w:p></w:comment></w:comments>`;
  const doc = withCommentsDoc(body, comments);
  assert.equal(doc.getReviewers().find((item) => item.author === 'Alice')?.initials, 'ZX');
});

test('getReviewers ignores invalid dates while computing firstDate/lastDate', () => {
  const body = '<w:p><w:ins w:id="1" w:author="Alice" w:date="bad"><w:r><w:t>A</w:t></w:r></w:ins><w:r><w:commentReference w:id="1"/></w:r></w:p>';
  const comments = `<w:comments xmlns:w="${WORD_NS}"><w:comment w:id="1" w:author="Alice" w:date="2026-03-02T10:00:00Z"><w:p><w:r><w:t>ok</w:t></w:r></w:p></w:comment></w:comments>`;
  const doc = withCommentsDoc(body, comments);
  const reviewer = doc.getReviewers().find((item) => item.author === 'Alice');
  assert.equal(reviewer?.firstDate, '2026-03-02T10:00:00Z');
  assert.equal(reviewer?.lastDate, '2026-03-02T10:00:00Z');
});

test('getReviewers uses revision cache and invalidates after mutation', () => {
  const doc = withBody('<w:p><w:ins w:id="1" w:author="Alice"><w:r><w:t>A</w:t></w:r></w:ins></w:p>');
  const originalGetRevisions = doc.getRevisions.bind(doc);
  const originalGetComments = doc.getComments.bind(doc);
  let revisionCalls = 0;
  let commentCalls = 0;
  doc.getRevisions = (...args) => { revisionCalls++; return originalGetRevisions(...args); };
  doc.getComments = (...args) => { commentCalls++; return originalGetComments(...args); };
  doc.getReviewers();
  doc.getReviewers();
  assert.equal(revisionCalls, 1);
  assert.equal(commentCalls, 1);
  doc.setTrackChanges(true);
  doc.getReviewers();
  assert.equal(revisionCalls, 2);
  assert.equal(commentCalls, 2);
});

test('getReviewers fast-paths empty documents to an empty cached list', () => {
  const doc = DocxDocument.create();
  const first = doc.getReviewers();
  const second = doc.getReviewers();
  assert.deepEqual(first, []);
  assert.deepEqual(second, []);
});

test('getRevisions fast-paths empty documents to a cached empty list', () => {
  const doc = DocxDocument.create();
  const first = doc.getRevisions();
  const second = doc.getRevisions();
  assert.deepEqual(first, []);
  assert.equal(first, second);
});

test('setRevisionAuthor is used by tracked text edits', () => {
  const doc = trackedDoc('<w:p><w:r><w:t>Hello world</w:t></w:r></w:p>');
  doc.setRevisionAuthor('Alice');
  doc.setParagraphText(0, 'Hello brave world');
  assert.match(doc.getPartXml(doc.mainDocumentPath), /w:ins w:id="\d+" w:author="Alice"/);
});

test('tracked setParagraphText writes deleted text as w:delText and hides it from paragraph text', () => {
  const doc = trackedDoc('<w:p><w:r><w:t>Hello world</w:t></w:r></w:p>');
  doc.setParagraphText(0, 'Hello');
  assert.equal(doc.getParagraphs()[0].text, 'Hello');
  assert.match(doc.getPartXml(doc.mainDocumentPath), /<w:del w:id="\d+" w:author="docx-view"><w:r><w:delText[^>]*> world<\/w:delText><\/w:r><\/w:del>/);
  assert.deepEqual(doc.getRevisions().map((revision) => revision.deletedText), [' world']);
});

test('tracked setParagraphText writes inserted text inside w:ins', () => {
  const doc = trackedDoc('<w:p><w:r><w:t>Hello</w:t></w:r></w:p>');
  doc.setParagraphText(0, 'Hello world');
  assert.equal(doc.getParagraphs()[0].text, 'Hello world');
  assert.match(doc.getPartXml(doc.mainDocumentPath), /<w:ins w:id="\d+" w:author="docx-view"><w:r><w:t xml:space="preserve"> world<\/w:t><\/w:r><\/w:ins>/);
});

test('tracked setParagraphText keeps separate insertion wrappers across repeated edits', () => {
  const doc = trackedDoc('<w:p><w:r><w:t>ABCDEF</w:t></w:r></w:p>');
  doc.setParagraphText(0, 'XABCDEF');
  doc.setParagraphText(0, 'XABCYDEF');
  assert.equal(doc.getParagraphs()[0].text, 'XABCYDEF');
  assert.match(doc.getPartXml(doc.mainDocumentPath), /<w:p><w:ins w:id="\d+" w:author="docx-view"><w:r><w:t(?: xml:space="preserve")?>X<\/w:t><\/w:r><\/w:ins><w:r><w:t(?: xml:space="preserve")?>ABC<\/w:t><\/w:r><w:ins w:id="\d+" w:author="docx-view"><w:r><w:t(?: xml:space="preserve")?>Y<\/w:t><\/w:r><\/w:ins><w:r><w:t(?: xml:space="preserve")?>DEF<\/w:t><\/w:r><\/w:p>/);
  assert.equal((doc.getPartXml(doc.mainDocumentPath).match(/<w:ins\b/g) ?? []).length, 2);
});

test('tracked setParagraphText inserts after a hyperlink boundary instead of inside the hyperlink', () => {
  const doc = trackedDoc(`<w:p xmlns:r="${OFFICE_REL_NS}"><w:hyperlink w:anchor="a"><w:r><w:t>link</w:t></w:r></w:hyperlink></w:p>`);
  doc.setParagraphText(0, 'link!');
  assert.match(doc.getPartXml(doc.mainDocumentPath), /<\/w:hyperlink><w:ins w:id="\d+" w:author="docx-view"><w:r><w:t(?: xml:space="preserve")?>!<\/w:t><\/w:r><\/w:ins>/);
});

test('tracked setParagraphText keeps replacement markup inside a hyperlink container', () => {
  const doc = trackedDoc(`<w:p xmlns:r="${OFFICE_REL_NS}"><w:hyperlink w:anchor="a"><w:r><w:t>link</w:t></w:r></w:hyperlink><w:r><w:t> tail</w:t></w:r></w:p>`);
  doc.setParagraphText(0, 'LINK tail');
  const xml = doc.getPartXml(doc.mainDocumentPath);
  assert.match(xml, /<w:hyperlink w:anchor="a"><w:(?:del|ins)[\s\S]*<\/w:hyperlink>/);
  assert.doesNotMatch(xml, /<w:p><w:(?:del|ins)[^>]*><w:r><w:t[^>]*>LINK/);
});

test('tracked setParagraphText replacement creates both deletion and insertion markers in one revision step', () => {
  const doc = trackedDoc('<w:p><w:r><w:t>Hello world</w:t></w:r></w:p>');
  const before = doc.revision;
  doc.setParagraphText(0, 'Hello Earth');
  assert.equal(doc.revision, before + 1);
  const xml = doc.getPartXml(doc.mainDocumentPath);
  assert.match(xml, /<w:del /);
  assert.match(xml, /<w:ins /);
});

test('tracked formatRun writes rPrChange with the previous direct format snapshot', () => {
  const doc = trackedDoc('<w:p><w:r><w:rPr><w:b/><w:color w:val="FF0000"/></w:rPr><w:t>A</w:t></w:r></w:p>');
  doc.formatRun(0, 0, { italic: true });
  const xml = doc.getPartXml(doc.mainDocumentPath);
  assert.match(xml, /<w:rPr><w:b\/><w:i w:val="1"\/><w:color w:val="FF0000"\/><w:rPrChange w:id="\d+" w:author="docx-view"><w:rPr><w:b w:val="1"\/><w:color w:val="FF0000"\/><\/w:rPr><\/w:rPrChange><\/w:rPr>/);
});

test('tracked formatRun previousFormat can be written back through formatRun', () => {
  const doc = trackedDoc('<w:p><w:r><w:rPr><w:b/><w:color w:val="FF0000"/></w:rPr><w:t>A</w:t></w:r></w:p>');
  doc.formatRun(0, 0, { italic: true });
  const previousFormat = doc.getRevisions()[0].previousFormat;
  assert.doesNotThrow(() => doc.formatRun(0, 0, previousFormat));
});

test('tracked formatRun preserves the original previousFormat across repeated tracked changes', () => {
  const doc = trackedDoc('<w:p><w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:t>X</w:t></w:r></w:p>');
  doc.formatRun(0, 0, { bold: true });
  doc.formatRun(0, 0, { italic: true });
  const revision = doc.getRevisions().filter((entry) => entry.kind === 'runFormatChange').at(-1);
  assert.deepEqual(revision?.previousFormat, { fontSize: 10 });
  assert.doesNotMatch(doc.getPartXml(doc.mainDocumentPath), /<w:rPrChange[^>]*><w:rPr><w:b w:val="1"\/>/);
});

test('tracked formatRange writes rPrChange on the affected run slice', () => {
  const doc = trackedDoc('<w:p><w:r><w:rPr><w:color w:val="FF0000"/></w:rPr><w:t>abcd</w:t></w:r></w:p>');
  doc.formatRange({ paragraph: 0, start: 1, end: 3 }, { bold: true });
  const xml = doc.getPartXml(doc.mainDocumentPath);
  assert.match(xml, /<w:rPr><w:b w:val="1"\/><w:color w:val="FF0000"\/><w:rPrChange w:id="\d+" w:author="docx-view"><w:rPr><w:color w:val="FF0000"\/><\/w:rPr><\/w:rPrChange><\/w:rPr><w:t xml:space="preserve">bc<\/w:t>/);
});

test('tracked formatRange across multiple runs writes revision markup and only increments revision once', () => {
  const doc = trackedDoc('<w:p><w:r><w:rPr><w:color w:val="FF0000"/></w:rPr><w:t>ab</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>cd</w:t></w:r></w:p>');
  const before = doc.revision;
  doc.formatRange({ paragraph: 0, start: 1, end: 3 }, { bold: true });
  assert.equal(doc.revision, before + 1);
  const xml = doc.getPartXml(doc.mainDocumentPath);
  assert.match(xml, /<w:rPr><w:b w:val="1"\/><w:color w:val="FF0000"\/><w:rPrChange w:id="\d+" w:author="docx-view"><w:rPr><w:color w:val="FF0000"\/><\/w:rPr><\/w:rPrChange><\/w:rPr><w:t xml:space="preserve">b<\/w:t>/);
  assert.match(xml, /<w:rPr><w:b w:val="1"\/><w:i\/><w:rPrChange w:id="\d+" w:author="docx-view"><w:rPr><w:i w:val="1"\/><\/w:rPr><\/w:rPrChange><\/w:rPr><w:t xml:space="preserve">c<\/w:t>/);
});

test('tracked formatParagraph writes pPrChange with previous paragraph properties', () => {
  const doc = trackedDoc('<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>A</w:t></w:r></w:p>');
  doc.formatParagraph(0, { spacingAfter: 120 });
  const xml = doc.getPartXml(doc.mainDocumentPath);
  assert.match(xml, /<w:pPr><w:spacing w:after="120"\/><w:jc w:val="center"\/><w:pPrChange w:id="\d+" w:author="docx-view"><w:pPr><w:jc w:val="center"\/><\/w:pPr><\/w:pPrChange><\/w:pPr>/);
});

test('tracked insertParagraph wraps text in w:ins and marks the paragraph mark insertion', () => {
  const doc = trackedDoc();
  doc.insertParagraph('tail');
  const xml = doc.getPartXml(doc.mainDocumentPath);
  assert.match(xml, /<w:p><w:pPr><w:rPr><w:ins w:id="\d+" w:author="docx-view"\/><\/w:rPr><\/w:pPr><w:ins w:id="\d+" w:author="docx-view"><w:r><w:t xml:space="preserve">tail<\/w:t><\/w:r><\/w:ins><\/w:p>/);
});

test('tracked deleteParagraph keeps the paragraph node and marks deletion', () => {
  const doc = trackedDoc('<w:p><w:r><w:t>first</w:t></w:r></w:p><w:p><w:r><w:t>second</w:t></w:r></w:p>');
  doc.deleteParagraph(0);
  assert.equal(doc.getParagraphs().length, 2);
  assert.equal(doc.getParagraphs()[0].text, '');
  assert.match(doc.getPartXml(doc.mainDocumentPath), /<w:p><w:pPr><w:rPr><w:del w:id="\d+" w:author="docx-view"\/><\/w:rPr><\/w:pPr><w:del w:id="\d+" w:author="docx-view"><w:r><w:delText[^>]*>first<\/w:delText><\/w:r><\/w:del><\/w:p>/);
});

test('tracked deleteParagraph keeps deleted wrapper order across hyperlink boundaries', () => {
  const doc = trackedDoc(`<w:p xmlns:r="${OFFICE_REL_NS}"><w:r><w:t>AAA</w:t></w:r><w:hyperlink w:anchor="x"><w:r><w:t>BBB</w:t></w:r></w:hyperlink><w:r><w:t>CCC</w:t></w:r></w:p>`);
  doc.deleteParagraph(0);
  assert.deepEqual(doc.getRevisions().map((revision) => revision.deletedText).filter(Boolean), ['AAA', 'BBB', 'CCC']);
  assert.match(doc.getPartXml(doc.mainDocumentPath), /<w:del w:id="\d+" w:author="docx-view"><w:r><w:delText[^>]*>AAA<\/w:delText><\/w:r><\/w:del><w:hyperlink w:anchor="x"><w:del w:id="\d+" w:author="docx-view"><w:r><w:delText[^>]*>BBB<\/w:delText><\/w:r><\/w:del><\/w:hyperlink><w:del w:id="\d+" w:author="docx-view"><w:r><w:delText[^>]*>CCC<\/w:delText><\/w:r><\/w:del>/);
});

test('tracked setParagraphText keeps deleted wrapper positions across repeated deletions', () => {
  const doc = trackedDoc('<w:p><w:r><w:t>AAABBBCCCDDD</w:t></w:r></w:p>');
  doc.setParagraphText(0, 'BBBCCCDDD');
  doc.setParagraphText(0, 'BBBCCC');
  assert.equal(doc.getParagraphs()[0].text, 'BBBCCC');
  assert.match(doc.getPartXml(doc.mainDocumentPath), /<w:p><w:del w:id="\d+" w:author="docx-view"><w:r><w:delText[^>]*>AAA<\/w:delText><\/w:r><\/w:del><w:r><w:t(?: xml:space="preserve")?>BBBCCC<\/w:t><\/w:r><w:del w:id="\d+" w:author="docx-view"><w:r><w:delText[^>]*>DDD<\/w:delText><\/w:r><\/w:del><\/w:p>/);
});

test('tracked deleteParagraph clears non-run direct children while preserving deleted runs', () => {
  const doc = trackedDoc('<w:p><w:bookmarkStart w:id="1" w:name="b"/><w:r><w:t>first</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>');
  doc.deleteParagraph(0);
  const xml = doc.getPartXml(doc.mainDocumentPath);
  assert.match(xml, /bookmarkStart/);
  assert.match(xml, /bookmarkEnd/);
  assert.match(xml, /<w:del w:id="\d+" w:author="docx-view"><w:r><w:delText[^>]*>first<\/w:delText><\/w:r><\/w:del>/);
});

test('tracked deleteParagraph preserves bookmarks and comment anchors', () => {
  const doc = trackedDoc('<w:p><w:bookmarkStart w:id="1" w:name="bm"/><w:r><w:t>AAA</w:t></w:r><w:bookmarkEnd w:id="1"/><w:commentRangeStart w:id="7"/><w:r><w:t>BBB</w:t></w:r><w:commentRangeEnd w:id="7"/></w:p>');
  doc.deleteParagraph(0);
  const xml = doc.getPartXml(doc.mainDocumentPath);
  assert.match(xml, /bookmarkStart/);
  assert.match(xml, /bookmarkEnd/);
  assert.match(xml, /commentRangeStart/);
  assert.match(xml, /commentRangeEnd/);
});

test('tracked deleteParagraph places paragraph-mark deletion before existing para-rPr formatting', () => {
  const doc = trackedDoc('<w:p><w:pPr><w:rPr><w:b/><w:sz w:val="24"/></w:rPr></w:pPr><w:r><w:t>X</w:t></w:r></w:p><w:p><w:r><w:t>keep</w:t></w:r></w:p>');
  doc.deleteParagraph(0);
  assert.match(doc.getPartXml(doc.mainDocumentPath), /<w:pPr><w:rPr><w:del w:id="\d+" w:author="docx-view"\/><w:b\/><w:sz w:val="24"\/><\/w:rPr><\/w:pPr>/);
});

test('tracked deleteParagraph still rejects section-break paragraphs', () => {
  const doc = trackedDoc('<w:p><w:pPr><w:sectPr/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>');
  assert.throws(() => doc.deleteParagraph(0), /while track changes is enabled/);
});

test('tracked insertTableRow writes trPr ins metadata', () => {
  const doc = trackedDoc();
  doc.insertTable([['A']]);
  doc.insertTableRow(0, 1);
  assert.match(doc.getPartXml(doc.mainDocumentPath), /<w:tr><w:trPr><w:ins w:id="\d+" w:author="docx-view"\/><\/w:trPr><w:tc>/);
});

test('tracked deleteTableRow writes trPr del metadata without removing the row', () => {
  const doc = trackedDoc();
  doc.insertTable([['A'], ['B']]);
  doc.deleteTableRow(0, 0);
  assert.equal(doc.getTable(0).rows.length, 2);
  assert.match(doc.getPartXml(doc.mainDocumentPath), /<w:trPr><w:del w:id="\d+" w:author="docx-view"\/><\/w:trPr>/);
});

test('tracked insertImage wraps the image run in w:ins', () => {
  const doc = trackedDoc();
  doc.insertImage({ bytes: PNG_BYTES, contentType: 'image/png', alt: 'tracked' });
  assert.match(doc.getPartXml(doc.mainDocumentPath), /<w:ins w:id="\d+" w:author="docx-view"><w:r><w:drawing>/);
});

test('tracked deleteImage wraps the image run in w:del and keeps the media part', () => {
  const doc = trackedDoc();
  const image = doc.insertImage({ bytes: PNG_BYTES, contentType: 'image/png', alt: 'tracked' });
  doc.deleteImage(image);
  assert.equal(doc.listParts().includes('word/media/image1.png'), true);
  assert.match(doc.getPartXml(doc.mainDocumentPath), /<w:del w:id="\d+" w:author="docx-view"><w:r><w:drawing>/);
  assert.doesNotMatch(doc.getPartXml(doc.mainDocumentPath), /delText/);
});

test('tracked deleteImage isolates the target image from sibling run content', () => {
  const doc = withImageDoc(
    `<w:p><w:r><w:t>A</w:t><w:drawing><wp:inline><wp:extent cx="914400" cy="457200"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="1" name="one"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="${PIC_NS}"><pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="one"/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing><w:t>B</w:t><w:drawing><wp:inline><wp:extent cx="914400" cy="457200"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="2" name="two"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="${PIC_NS}"><pic:pic><pic:nvPicPr><pic:cNvPr id="2" name="two"/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`,
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image2.gif"/>`,
    [{ path: 'word/media/image1.png', bytes: PNG_BYTES, type: 'image/png' }, { path: 'word/media/image2.gif', bytes: GIF_BYTES, type: 'image/gif' }],
  );
  doc.setTrackChanges(true);
  const [first] = doc.getImages();
  doc.deleteImage(first);
  const xml = doc.getPartXml(doc.mainDocumentPath);
  assert.match(xml, /<w:r><w:t(?: xml:space="preserve")?>A<\/w:t><\/w:r><w:del w:id="\d+" w:author="docx-view"><w:r><w:drawing>/);
  assert.match(xml, /<w:r><w:t(?: xml:space="preserve")?>B<\/w:t><w:drawing>/);
  assert.doesNotMatch(xml, /<w:del[\s\S]*<w:delText/);
});

test('applyOperations accepts setTrackChanges and setRevisionAuthor', () => {
  const doc = DocxDocument.create();
  doc.applyOperations({
    operations: [
      { type: 'setTrackChanges', enabled: true },
      { type: 'setRevisionAuthor', author: 'Agent' },
      { type: 'setParagraphText', index: 0, text: 'body' },
    ],
  });
  assert.match(doc.getPartXml(doc.mainDocumentPath), /w:author="Agent"/);
  assert.equal(doc.getSettings().trackChanges, true);
});

test('tracked applyOperations keeps sequential run indices after earlier revision markup changes the DOM shape', () => {
  const doc = trackedDoc('<w:p><w:r><w:t>ab</w:t></w:r></w:p>');
  doc.applyOperations({
    operations: [
      { type: 'setParagraphText', index: 0, text: 'aXb' },
      { type: 'formatRun', paragraph: 0, run: 1, format: { bold: true } },
    ],
  });
  const run = doc.getParagraphs()[0].runs[1];
  assert.equal(run.text, 'X');
  assert.equal(run.bold, true);
});

test('tracked setParagraphText increments revision exactly once', () => {
  const doc = trackedDoc('<w:p><w:r><w:t>abc</w:t></w:r></w:p>');
  const before = doc.revision;
  doc.setParagraphText(0, 'abd');
  assert.equal(doc.revision, before + 1);
});

test('tracked formatRun increments revision exactly once', () => {
  const doc = trackedDoc();
  const before = doc.revision;
  doc.formatRun(0, 0, { bold: true });
  assert.equal(doc.revision, before + 1);
});

test('tracked formatParagraph increments revision exactly once', () => {
  const doc = trackedDoc();
  const before = doc.revision;
  doc.formatParagraph(0, { alignment: 'center' });
  assert.equal(doc.revision, before + 1);
});

test('tracked insertParagraph increments revision exactly once', () => {
  const doc = trackedDoc();
  const before = doc.revision;
  doc.insertParagraph('tail');
  assert.equal(doc.revision, before + 1);
});

test('tracked deleteParagraph increments revision exactly once', () => {
  const doc = trackedDoc('<w:p><w:r><w:t>first</w:t></w:r></w:p><w:p><w:r><w:t>second</w:t></w:r></w:p>');
  const before = doc.revision;
  doc.deleteParagraph(0);
  assert.equal(doc.revision, before + 1);
});

test('tracked insertTableRow increments revision exactly once', () => {
  const doc = trackedDoc();
  doc.insertTable([['A']]);
  const before = doc.revision;
  doc.insertTableRow(0, 1);
  assert.equal(doc.revision, before + 1);
});

test('tracked deleteTableRow increments revision exactly once', () => {
  const doc = trackedDoc();
  doc.insertTable([['A'], ['B']]);
  const before = doc.revision;
  doc.deleteTableRow(0, 0);
  assert.equal(doc.revision, before + 1);
});

test('tracked insertImage increments revision exactly once', () => {
  const doc = trackedDoc();
  const before = doc.revision;
  doc.insertImage({ bytes: PNG_BYTES, contentType: 'image/png' });
  assert.equal(doc.revision, before + 1);
});

test('tracked deleteImage increments revision exactly once', () => {
  const doc = trackedDoc();
  const image = doc.insertImage({ bytes: PNG_BYTES, contentType: 'image/png' });
  const before = doc.revision;
  doc.deleteImage(image);
  assert.equal(doc.revision, before + 1);
});

test('revision author survives undo and redo for later tracked edits', () => {
  const doc = trackedDoc('<w:p><w:r><w:t>ab</w:t></w:r></w:p>');
  doc.setRevisionAuthor('Alice');
  doc.setParagraphText(0, 'axb');
  doc.undo();
  doc.redo();
  doc.setParagraphText(0, 'axby');
  assert.match(doc.getPartXml(doc.mainDocumentPath), /w:author="Alice"/);
});

test('agent operation schema includes tracked-review settings operations', () => {
  const types = AGENT_OPERATION_SCHEMA.properties.operations.items.oneOf
    .map((entry) => entry.properties.type.const);
  assert.equal(AGENT_OPERATION_SCHEMA.properties.operations.items.oneOf.length, 62);
  assert.ok(types.includes('setTrackChanges'));
  assert.ok(types.includes('setRevisionAuthor'));
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

test('copyFormat mirrors getRangeFormat', () => {
  const doc = withBody('<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>abc</w:t></w:r></w:p>');
  assert.deepEqual(doc.copyFormat({ paragraph: 0, start: 0, end: 3 }), doc.getRangeFormat({ paragraph: 0, start: 0, end: 3 }));
});

test('applyFormat mirrors formatDocumentRange', () => {
  const doc = withBody('<w:p><w:r><w:t>abc</w:t></w:r></w:p>');
  const before = doc.revision;
  doc.applyFormat({ start: { paragraph: 0, offset: 1 }, end: { paragraph: 0, offset: 3 } }, { italic: true });
  assert.equal(doc.revision, before + 1);
  assert.equal(doc.getParagraphs()[0].runs.some((run) => run.italic), true);
});

test('copyClipboardFragment includes direct formatting, hyperlinks, and images', () => {
  const doc = DocxDocument.create();
  doc.setParagraphText(0, 'hello world');
  doc.formatRange({ paragraph: 0, start: 0, end: 5 }, { bold: true, color: 'FF0000' });
  doc.insertHyperlink({ paragraph: 0, start: 6, end: 11 }, { url: 'https://example.com' });
  doc.insertImage({ bytes: PNG_BYTES, contentType: 'image/png', paragraph: 0 });
  const fragment = doc.copyClipboardFragment({ start: { paragraph: 0, offset: 0 }, end: { paragraph: 0, offset: 11 } });
  const runs = fragment.paragraphs.flatMap((paragraph) => paragraph.runs);
  assert.equal(fragment.text, 'hello world');
  assert.equal(runs.some((run) => run.format?.bold), true);
  assert.equal(runs.some((run) => run.hyperlink?.url === 'https://example.com'), true);
  assert.equal(runs.some((run) => (run.images?.length ?? 0) > 0), true);
});

test('pasteClipboardFragment inserts rich runs and bumps revision once', () => {
  const doc = withBody('<w:p><w:r><w:t>base</w:t></w:r></w:p>');
  const fragment = {
    version: 1,
    text: 'A B',
    paragraphs: [{
      runs: [
        { text: 'A', format: { bold: true } },
        { text: ' ', format: {} },
        { text: 'B', hyperlink: { url: 'https://example.com' }, format: { underline: true } },
      ],
    }],
  };
  const before = doc.revision;
  assert.equal(doc.pasteClipboardFragment({ start: { paragraph: 0, offset: 0 }, end: { paragraph: 0, offset: 4 } }, fragment), true);
  assert.equal(doc.revision, before + 1);
  assert.equal(doc.getParagraphs()[0].text.includes('A B'), true);
  assert.equal(doc.getHyperlinks().some((item) => item.url === 'https://example.com'), true);
});

test('pasteClipboardFragment imports image bytes into a new media part in target docs', () => {
  const source = DocxDocument.create();
  source.setParagraphText(0, 'img');
  source.insertImage({ bytes: PNG_BYTES, contentType: 'image/png', paragraph: 0 });
  const fragment = source.copyClipboardFragment({ start: { paragraph: 0, offset: 0 }, end: { paragraph: 0, offset: 3 } });
  const target = DocxDocument.create();
  target.setParagraphText(0, 'X');
  const beforeSourceParts = source.listParts();
  assert.equal(target.pasteClipboardFragment({ start: { paragraph: 0, offset: 1 }, end: { paragraph: 0, offset: 1 } }, fragment), true);
  assert.equal(target.getImages().length > 0, true);
  assert.equal(target.listParts().some((path) => /^word\/media\/image\d+\./.test(path)), true);
  assert.deepEqual(source.listParts(), beforeSourceParts);
});

test('copy/paste rich image works when source relationship target uses a nonstandard media path', () => {
  const source = withImageDoc(
    '<w:p><w:r><w:t>x</w:t></w:r><w:r><w:drawing><wp:inline><wp:extent cx="19050" cy="19050"/><wp:docPr id="1" name="x"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="x"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId9"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="19050" cy="19050"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>',
    '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="assets/custom-image.png"/>',
    [{ path: 'word/assets/custom-image.png', bytes: PNG_BYTES, type: 'image/png' }],
  );
  const fragment = source.copyClipboardFragment({ start: { paragraph: 0, offset: 0 }, end: { paragraph: 0, offset: 1 } });
  const target = DocxDocument.create();
  target.setParagraphText(0, 'A');
  assert.equal(target.pasteClipboardFragment({ start: { paragraph: 0, offset: 1 }, end: { paragraph: 0, offset: 1 } }, fragment), true);
  assert.equal(target.getImages().length, 1);
  assert.equal(target.listParts().some((path) => /^word\/media\/image\d+\.png$/.test(path)), true);
});

test('copy/paste rich image degrades safely when source image target is malformed/encoded', () => {
  const source = withImageDoc(
    '<w:p><w:r><w:t>x</w:t></w:r><w:r><w:drawing><wp:inline><wp:extent cx="19050" cy="19050"/><wp:docPr id="1" name="x"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="x"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId9"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="19050" cy="19050"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>',
    '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="bad%ZZ.png"/>',
  );
  const target = DocxDocument.create();
  target.setParagraphText(0, 'A');
  assert.doesNotThrow(() => {
    const fragment = source.copyClipboardFragment({ start: { paragraph: 0, offset: 0 }, end: { paragraph: 0, offset: 1 } });
    target.pasteClipboardFragment({ start: { paragraph: 0, offset: 1 }, end: { paragraph: 0, offset: 1 } }, fragment);
  });
});

test('pasteClipboardFragment returns false for cross-paragraph targets', () => {
  const doc = withBody('<w:p><w:r><w:t>a</w:t></w:r></w:p><w:p><w:r><w:t>b</w:t></w:r></w:p>');
  const fragment = { version: 1, text: 'x', paragraphs: [{ runs: [{ text: 'x' }] }] };
  const revision = doc.revision;
  assert.equal(doc.pasteClipboardFragment({ start: { paragraph: 0, offset: 0 }, end: { paragraph: 1, offset: 1 } }, fragment), false);
  assert.equal(doc.revision, revision);
});

test('pasteClipboardFragment with empty fragment removes selected text', () => {
  const doc = withBody('<w:p><w:r><w:t>abcdef</w:t></w:r></w:p>');
  const before = doc.revision;
  assert.equal(doc.pasteClipboardFragment(
    { start: { paragraph: 0, offset: 2 }, end: { paragraph: 0, offset: 4 } },
    { version: 1, text: '', paragraphs: [] },
  ), true);
  assert.equal(doc.revision, before + 1);
  assert.equal(doc.getParagraphs()[0].text, 'abef');
});

test('pasteClipboardFragment applies 100-paragraph payload in one transaction', () => {
  const doc = withBody('<w:p><w:r><w:t>base</w:t></w:r></w:p>');
  const fragment = {
    version: 1,
    text: Array.from({ length: 100 }, (_, index) => `p${index}`).join('\n'),
    paragraphs: Array.from({ length: 100 }, (_, index) => ({ runs: [{ text: `p${index}` }] })),
  };
  const before = doc.revision;
  assert.equal(doc.pasteClipboardFragment({ start: { paragraph: 0, offset: 0 }, end: { paragraph: 0, offset: 4 } }, fragment), true);
  assert.equal(doc.revision, before + 1);
  const paragraphs = doc.getParagraphs();
  assert.equal(paragraphs.length >= 100, true);
  assert.equal(paragraphs[0].text, 'p0');
  assert.equal(paragraphs[99].text, 'p99');
});

test('pasteClipboardFragment sanitizes clipboard run text', () => {
  const doc = withBody('<w:p><w:r><w:t>x</w:t></w:r></w:p>');
  assert.equal(doc.pasteClipboardFragment(
    { start: { paragraph: 0, offset: 0 }, end: { paragraph: 0, offset: 1 } },
    { version: 1, text: 'a\u0000b', paragraphs: [{ runs: [{ text: 'a\u0000b' }] }] },
  ), true);
  assert.equal(doc.getParagraphs()[0].text, 'ab');
});

test('pasteClipboardFragment validates run format payload', () => {
  const doc = withBody('<w:p><w:r><w:t>x</w:t></w:r></w:p>');
  assert.throws(() => doc.pasteClipboardFragment(
    { start: { paragraph: 0, offset: 0 }, end: { paragraph: 0, offset: 1 } },
    { version: 1, text: 'x', paragraphs: [{ runs: [{ text: 'x', format: { color: 'BAD' } }] }] },
  ), /color must be six hexadecimal digits/);
});

test('pasteClipboardFragment rejects unsafe hyperlink urls', () => {
  const doc = withBody('<w:p><w:r><w:t>x</w:t></w:r></w:p>');
  assert.throws(() => doc.pasteClipboardFragment(
    { start: { paragraph: 0, offset: 0 }, end: { paragraph: 0, offset: 1 } },
    { version: 1, text: 'bad', paragraphs: [{ runs: [{ text: 'bad', hyperlink: { url: 'javascript:alert(1)' } }] }] },
  ), /link\.url must use http, https or mailto/);
});

test('pasteClipboardFragment enforces paragraph count limit', () => {
  const doc = withBody('<w:p><w:r><w:t>x</w:t></w:r></w:p>');
  const paragraphs = Array.from({ length: 1001 }, () => ({ runs: [{ text: 'x' }] }));
  assert.throws(() => doc.pasteClipboardFragment(
    { start: { paragraph: 0, offset: 0 }, end: { paragraph: 0, offset: 1 } },
    { version: 1, text: '', paragraphs },
  ), /clipboard paragraph count exceeds 1000/);
});

test('pasteClipboardFragment enforces run count limit', () => {
  const doc = withBody('<w:p><w:r><w:t>x</w:t></w:r></w:p>');
  const runs = Array.from({ length: 10001 }, () => ({ text: 'x' }));
  assert.throws(() => doc.pasteClipboardFragment(
    { start: { paragraph: 0, offset: 0 }, end: { paragraph: 0, offset: 1 } },
    { version: 1, text: '', paragraphs: [{ runs }] },
  ), /clipboard run count exceeds 10000/);
});

test('pasteClipboardFragment wraps inserted runs with revisions when track changes is enabled', () => {
  const doc = trackedDoc('<w:p><w:r><w:t>base</w:t></w:r></w:p>');
  doc.pasteClipboardFragment(
    { start: { paragraph: 0, offset: 4 }, end: { paragraph: 0, offset: 4 } },
    { version: 1, text: 'X', paragraphs: [{ runs: [{ text: 'X' }] }] },
  );
  const xml = doc.getPartXml(doc.mainDocumentPath);
  assert.match(xml, /<w:ins w:id="\d+" w:author="docx-view"><w:r><w:t(?: xml:space="preserve")?>X<\/w:t><\/w:r><\/w:ins>/);
});

test('clipboard copy/paste round-trip keeps richer direct run format fields', () => {
  const source = withBody('<w:p><w:r><w:t>format</w:t></w:r></w:p>');
  source.formatRange({ paragraph: 0, start: 0, end: 6 }, {
    strike: true,
    underline: true,
    underlineStyle: 'dotted',
    fontFamily: 'Arial',
    highlight: 'yellow',
  });
  const fragment = source.copyClipboardFragment({ start: { paragraph: 0, offset: 0 }, end: { paragraph: 0, offset: 6 } });
  const target = withBody('<w:p><w:r><w:t>xxxxxx</w:t></w:r></w:p>');
  target.pasteClipboardFragment({ start: { paragraph: 0, offset: 0 }, end: { paragraph: 0, offset: 6 } }, fragment);
  const run = target.getParagraphs()[0].runs.find((item) => item.text === 'format');
  assert.equal(run?.strike, true);
  assert.equal(run?.underline, true);
  assert.equal(run?.underlineStyle, 'dotted');
  assert.equal(run?.fontFamily, 'Arial');
  assert.equal(run?.highlight, 'yellow');
});

test('acceptRevision unwraps insertion wrappers', () => {
  const doc = withBody('<w:p><w:ins w:id="1"><w:r><w:t>A</w:t></w:r></w:ins><w:r><w:t>B</w:t></w:r></w:p>');
  doc.acceptRevision(1);
  assert.equal(doc.getParagraphs()[0].text, 'AB');
  assert.doesNotMatch(doc.getPartXml(doc.mainDocumentPath), /<w:ins\b/);
});

test('rejectRevision deletes insertion wrappers with their content', () => {
  const doc = withBody('<w:p><w:ins w:id="2"><w:r><w:t>A</w:t></w:r></w:ins><w:r><w:t>B</w:t></w:r></w:p>');
  doc.rejectRevision(2);
  assert.equal(doc.getParagraphs()[0].text, 'B');
  assert.doesNotMatch(doc.getPartXml(doc.mainDocumentPath), /<w:ins\b/);
});

test('acceptRevision deletes deletion wrappers with their content', () => {
  const doc = withBody('<w:p><w:del w:id="3"><w:r><w:delText>A</w:delText></w:r></w:del><w:r><w:t>B</w:t></w:r></w:p>');
  doc.acceptRevision(3);
  assert.equal(doc.getParagraphs()[0].text, 'B');
  assert.doesNotMatch(doc.getPartXml(doc.mainDocumentPath), /<w:del\b/);
});

test('rejectRevision unwraps deletion wrappers and converts delText back to text', () => {
  const doc = withBody('<w:p><w:del w:id="4"><w:r><w:delText>A</w:delText></w:r></w:del><w:r><w:t>B</w:t></w:r></w:p>');
  doc.rejectRevision(4);
  assert.equal(doc.getParagraphs()[0].text, 'AB');
  assert.match(doc.getPartXml(doc.mainDocumentPath), /<w:t(?: xml:space="preserve")?>A<\/w:t>/);
});

test('acceptRevision drops rPrChange while keeping current format', () => {
  const doc = withBody('<w:p><w:r><w:rPr><w:b/><w:rPrChange w:id="5"><w:rPr><w:i/></w:rPr></w:rPrChange></w:rPr><w:t>X</w:t></w:r></w:p>');
  doc.acceptRevision(5);
  const run = doc.getParagraphs()[0].runs[0];
  assert.equal(run.bold, true);
  assert.equal(run.italic, undefined);
  assert.doesNotMatch(doc.getPartXml(doc.mainDocumentPath), /rPrChange/);
});

test('rejectRevision restores previous format from rPrChange snapshot', () => {
  const doc = withBody('<w:p><w:r><w:rPr><w:b/><w:rPrChange w:id="6"><w:rPr><w:i/></w:rPr></w:rPrChange></w:rPr><w:t>X</w:t></w:r></w:p>');
  doc.rejectRevision(6);
  const run = doc.getParagraphs()[0].runs[0];
  assert.equal(run.bold, undefined);
  assert.equal(run.italic, true);
  assert.doesNotMatch(doc.getPartXml(doc.mainDocumentPath), /rPrChange/);
});

test('acceptRevision drops pPrChange while keeping current paragraph format', () => {
  const doc = withBody('<w:p><w:pPr><w:jc w:val="center"/><w:pPrChange w:id="7"><w:pPr><w:jc w:val="left"/></w:pPr></w:pPrChange></w:pPr><w:r><w:t>X</w:t></w:r></w:p>');
  doc.acceptRevision(7);
  assert.equal(doc.getParagraphs()[0].alignment, 'center');
  assert.doesNotMatch(doc.getPartXml(doc.mainDocumentPath), /pPrChange/);
});

test('rejectRevision restores previous paragraph format from pPrChange snapshot', () => {
  const doc = withBody('<w:p><w:pPr><w:jc w:val="center"/><w:pPrChange w:id="8"><w:pPr><w:jc w:val="left"/></w:pPr></w:pPrChange></w:pPr><w:r><w:t>X</w:t></w:r></w:p>');
  doc.rejectRevision(8);
  assert.equal(doc.getParagraphs()[0].alignment, 'left');
  assert.doesNotMatch(doc.getPartXml(doc.mainDocumentPath), /pPrChange/);
});

test('rejectRevision on paragraph insertion marker removes the paragraph and preserves body paragraph invariant', () => {
  const doc = withBody('<w:p><w:pPr><w:rPr><w:ins w:id="9"/></w:rPr></w:pPr><w:r><w:t>X</w:t></w:r></w:p>');
  doc.rejectRevision(9);
  assert.equal(doc.getParagraphs().length, 1);
  assert.equal(doc.getParagraphs()[0].text, '');
});

test('rejectRevision on inserted table-cell paragraph marker preserves required empty paragraph', () => {
  const doc = withBody('<w:tbl><w:tr><w:tc><w:p><w:pPr><w:rPr><w:ins w:id="10"/></w:rPr></w:pPr><w:r><w:t>X</w:t></w:r></w:p></w:tc></w:tr></w:tbl>');
  doc.rejectRevision(10);
  const table = doc.getTable(0);
  assert.equal(table.rows[0].cells[0].blocks.length, 1);
  assert.equal(table.rows[0].cells[0].blocks[0].type, 'paragraph');
  assert.equal(table.rows[0].cells[0].blocks[0].paragraph.text, '');
});

test('acceptRevision merges surviving content when only paragraph-mark deletion is tracked', () => {
  const doc = withBody('<w:p><w:pPr><w:pStyle w:val="Heading1"/><w:rPr><w:del w:id="101" w:author="Alice"/></w:rPr></w:pPr><w:r><w:t>上半</w:t></w:r></w:p><w:p><w:pPr><w:pStyle w:val="Quote"/></w:pPr><w:r><w:t>下半</w:t></w:r></w:p>');
  doc.acceptRevision(101);
  assert.deepEqual(doc.getParagraphs().map((paragraph) => [paragraph.text, paragraph.style]), [['上半下半', 'Quote']]);
  assert.equal(doc.getRevisions().length, 0);
});

test('acceptRevision keeps merge semantics when paragraph-mark deletion coexists with deleted runs', () => {
  const doc = withBody('<w:p><w:pPr><w:rPr><w:del w:id="102" w:author="Alice"/></w:rPr></w:pPr><w:r><w:t>上</w:t></w:r><w:del w:id="103" w:author="Bob"><w:r><w:delText>半</w:delText></w:r></w:del></w:p><w:p><w:r><w:t>下半</w:t></w:r></w:p>');
  doc.acceptRevision(102);
  assert.deepEqual(doc.getParagraphs().map((paragraph) => paragraph.text), ['上下半']);
  assert.deepEqual(doc.getRevisions().map((revision) => revision.id), [103]);
});

test('acceptRevision degrades to marker removal when paragraph-mark deletion is on the last paragraph', () => {
  const doc = withBody('<w:p><w:r><w:t>前</w:t></w:r></w:p><w:p><w:pPr><w:rPr><w:del w:id="104" w:author="Alice"/></w:rPr></w:pPr><w:r><w:t>后</w:t></w:r></w:p>');
  doc.acceptRevision(104);
  assert.deepEqual(doc.getParagraphs().map((paragraph) => paragraph.text), ['前', '后']);
  assert.equal(doc.getRevisions().length, 0);
});

test('acceptRevision degrades to marker removal when next paragraph is in another container', () => {
  const doc = withBody('<w:tbl><w:tr><w:tc><w:p><w:pPr><w:rPr><w:del w:id="105" w:author="Alice"/></w:rPr></w:pPr><w:r><w:t>单元格</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p><w:r><w:t>正文</w:t></w:r></w:p>');
  doc.acceptRevision(105);
  assert.deepEqual(doc.getParagraphs().map((paragraph) => paragraph.text), ['单元格', '正文']);
  assert.equal(doc.getRevisions().length, 0);
});

test('acceptRevision degrades to marker removal when paragraph-mark deletion paragraph has sectPr', () => {
  const doc = withBody('<w:p><w:pPr><w:rPr><w:del w:id="106" w:author="Alice"/></w:rPr><w:sectPr/></w:pPr><w:r><w:t>甲</w:t></w:r></w:p><w:p><w:r><w:t>乙</w:t></w:r></w:p>');
  doc.acceptRevision(106);
  assert.deepEqual(doc.getParagraphs().map((paragraph) => paragraph.text), ['甲', '乙']);
  assert.equal(doc.getRevisions().length, 0);
});

test('acceptAllRevisions merge path removes empty hyperlink and sdt wrapper shells', () => {
  const doc = withBody(
    `<w:p xmlns:r="${OFFICE_REL_NS}">
      <w:pPr><w:rPr><w:del w:id="107" w:author="Alice"/></w:rPr></w:pPr>
      <w:hyperlink w:anchor="x"><w:del w:id="108" w:author="Alice"><w:r><w:delText>link</w:delText></w:r></w:del></w:hyperlink>
      <w:sdt><w:sdtPr/><w:sdtContent><w:del w:id="109" w:author="Alice"><w:r><w:delText>tag</w:delText></w:r></w:del></w:sdtContent></w:sdt>
      <w:r><w:t>head</w:t></w:r>
    </w:p>
    <w:p><w:r><w:t>keep</w:t></w:r></w:p>`,
  );
  doc.acceptAllRevisions();
  assert.deepEqual(doc.getParagraphs().map((paragraph) => paragraph.text), ['headkeep']);
  const xml = doc.getPartXml(doc.mainDocumentPath);
  assert.doesNotMatch(xml, /<w:hyperlink[^>]*\/>/);
  assert.doesNotMatch(xml, /<w:sdt>(?:\s|<w:sdtPr\/>|<w:sdtContent\/>)*<\/w:sdt>/);
  assert.doesNotMatch(xml, /<w:pPr\/>/);
});

test('acceptRevision on row insertion marker keeps the row and removes marker', () => {
  const doc = withBody('<w:tbl><w:tr><w:trPr><w:ins w:id="11"/></w:trPr><w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc></w:tr></w:tbl>');
  doc.acceptRevision(11);
  assert.equal(doc.getTable(0).rows.length, 2);
  assert.doesNotMatch(doc.getPartXml(doc.mainDocumentPath), /<w:ins w:id="11"/);
});

test('rejectRevision on row insertion marker deletes the row', () => {
  const doc = withBody('<w:tbl><w:tr><w:trPr><w:ins w:id="12"/></w:trPr><w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc></w:tr></w:tbl>');
  doc.rejectRevision(12);
  assert.deepEqual(doc.getTable(0).rows.flatMap(row => row.cells[0].blocks[0].paragraph.text), ['B']);
});

test('acceptRevision on row deletion marker deletes the row', () => {
  const doc = withBody('<w:tbl><w:tr><w:trPr><w:del w:id="13"/></w:trPr><w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc></w:tr></w:tbl>');
  doc.acceptRevision(13);
  assert.deepEqual(doc.getTable(0).rows.flatMap(row => row.cells[0].blocks[0].paragraph.text), ['B']);
});

test('rejectRevision on row deletion marker keeps the row and removes marker', () => {
  const doc = withBody('<w:tbl><w:tr><w:trPr><w:del w:id="14"/></w:trPr><w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc></w:tr></w:tbl>');
  doc.rejectRevision(14);
  assert.equal(doc.getTable(0).rows.length, 2);
  assert.doesNotMatch(doc.getPartXml(doc.mainDocumentPath), /<w:del w:id="14"/);
});

test('rejectRevision on cellIns removes only the revised cell', () => {
  const doc = withBody('<w:tbl><w:tr><w:tc><w:tcPr><w:cellIns w:id="140"/></w:tcPr><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:tc><w:p><w:r><w:t>C</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>D</w:t></w:r></w:p></w:tc></w:tr></w:tbl>');
  doc.rejectRevision(140);
  const table = doc.getTable(0);
  assert.deepEqual(table.rows[0].cells.map((cell) => cell.blocks[0].paragraph.text), ['B']);
  assert.deepEqual(table.rows[1].cells.map((cell) => cell.blocks[0].paragraph.text), ['C', 'D']);
});

test('acceptRevision on cellDel removes only the revised cell', () => {
  const doc = withBody('<w:tbl><w:tr><w:tc><w:tcPr><w:cellDel w:id="141"/></w:tcPr><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:tc><w:p><w:r><w:t>C</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>D</w:t></w:r></w:p></w:tc></w:tr></w:tbl>');
  doc.acceptRevision(141);
  const table = doc.getTable(0);
  assert.deepEqual(table.rows[0].cells.map((cell) => cell.blocks[0].paragraph.text), ['B']);
  assert.deepEqual(table.rows[1].cells.map((cell) => cell.blocks[0].paragraph.text), ['C', 'D']);
});

test('acceptAllRevisions can filter by author', () => {
  const doc = withBody('<w:p><w:ins w:id="15" w:author="Alice"><w:r><w:t>A</w:t></w:r></w:ins><w:ins w:id="16" w:author="Bob"><w:r><w:t>B</w:t></w:r></w:ins><w:r><w:t>C</w:t></w:r></w:p>');
  doc.acceptAllRevisions({ authors: ['Alice'] });
  assert.equal(doc.getParagraphs()[0].text, 'ABC');
  assert.equal(doc.getRevisions().map((revision) => revision.id).includes(16), true);
  assert.equal(doc.getRevisions().map((revision) => revision.id).includes(15), false);
});

test('acceptAllRevisions processes 20+ revisions and increments revision once', () => {
  const wraps = Array.from({ length: 25 }, (_, index) => `<w:ins w:id="${200 + index}" w:author="A"><w:r><w:t>${index}</w:t></w:r></w:ins>`).join('');
  const doc = withBody(`<w:p>${wraps}</w:p>`);
  const before = doc.revision;
  doc.acceptAllRevisions();
  assert.equal(doc.revision, before + 1);
  assert.equal(doc.getRevisions().length, 0);
});

test('acceptAllRevisions removes a single-row deleted table', () => {
  const doc = withBody('<w:tbl><w:tr><w:trPr><w:del w:id="300"/></w:trPr><w:tc><w:p><w:r><w:t>only</w:t></w:r></w:p></w:tc></w:tr></w:tbl>');
  const beforeRevision = doc.revision;
  doc.acceptAllRevisions();
  assert.equal(doc.revision, beforeRevision + 1);
  assert.equal(doc.getBlocks().filter((block) => block.type === 'table').length, 0);
});

test('acceptAllRevisions is a no-op when author filter matches nothing', () => {
  const doc = withBody('<w:p><w:ins w:id="17" w:author="Alice"><w:r><w:t>A</w:t></w:r></w:ins></w:p>');
  const before = doc.revision;
  doc.acceptAllRevisions({ authors: ['Bob'] });
  assert.equal(doc.revision, before);
  assert.equal(doc.getRevisions().length, 1);
});

test('acceptRevision throws when revision id does not exist', () => {
  const doc = withBody('<w:p><w:r><w:t>A</w:t></w:r></w:p>');
  assert.throws(() => doc.acceptRevision(999), /does not exist/);
});

test('applyOperations keeps operation indices sequential after rejectRevision removes earlier runs', () => {
  const doc = withBody('<w:p><w:ins w:id="18"><w:r><w:t>X</w:t></w:r></w:ins><w:r><w:t>Y</w:t></w:r></w:p>');
  const snapshot = doc.applyOperations({
    operations: [
      { type: 'rejectRevision', id: 18 },
      { type: 'formatRun', paragraph: 0, run: 0, format: { bold: true } },
    ],
  });
  assert.equal(snapshot.paragraphs[0].text, 'Y');
  assert.equal(snapshot.paragraphs[0].runs[0].bold, true);
});

test('accept/reject revisions keep numbering sequence consistent', () => {
  const doc = withBody('<w:p><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>A</w:t></w:r></w:p><w:p><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr><w:rPr><w:ins w:id="19"/></w:rPr></w:pPr><w:r><w:t>B</w:t></w:r></w:p><w:p><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>C</w:t></w:r></w:p>');
  doc.addPart('word/numbering.xml', encoder.encode(`<w:numbering xmlns:w="${WORD_NS}"><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num></w:numbering>`), 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml');
  doc.addPart('word/_rels/document.xml.rels', encoder.encode(`<Relationships xmlns="${REL_NS}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>`), RELS_TYPE);
  assert.deepEqual(doc.getParagraphs().map((paragraph) => paragraph.numbering?.text), ['1.', '2.', '3.']);
  doc.rejectRevision(19);
  assert.deepEqual(doc.getParagraphs().map((paragraph) => paragraph.numbering?.text), ['1.', '2.']);
});

test('acceptRevision keeps hyperlinks and images discoverable after unwrapping', () => {
  const doc = withImageDoc(
    '<w:p><w:ins w:id="20"><w:hyperlink w:anchor="a"><w:r><w:t>L</w:t></w:r></w:hyperlink><w:r><w:drawing><wp:inline><wp:extent cx="952500" cy="476250"/><wp:docPr id="1" name="img"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="rId1"/></pic:blipFill><pic:spPr/></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:ins></w:p>',
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>',
  );
  doc.acceptRevision(20);
  assert.equal(doc.getHyperlinks().length, 1);
  assert.equal(doc.getImages().length, 1);
});

test('agent schema and runtime validation support revision accept/reject operations', () => {
  const opNames = AGENT_OPERATION_SCHEMA.properties.operations.items.oneOf
    .map((schema) => schema.properties?.type?.const)
    .filter(Boolean);
  assert.equal(opNames.includes('acceptRevision'), true);
  assert.equal(opNames.includes('rejectRevision'), true);
  assert.equal(opNames.includes('acceptAllRevisions'), true);
  assert.equal(opNames.includes('rejectAllRevisions'), true);
  const doc = withBody('<w:p><w:ins w:id="21"><w:r><w:t>A</w:t></w:r></w:ins></w:p>');
  const snapshot = doc.applyOperations({ operations: [{ type: 'acceptRevision', id: 21 }] });
  assert.equal(snapshot.paragraphs[0].text, 'A');
  assert.throws(() => doc.applyOperations({ operations: [{ type: 'acceptAllRevisions', filter: { authors: new Array(1001).fill('A') } }] }), /at most 1000 authors/);
});

test('compare roundtrip accept/reject restores revised/base text', async () => {
  const base = withBody('<w:p><w:r><w:t>Hello world</w:t></w:r></w:p><w:p><w:r><w:t>Tail</w:t></w:r></w:p>');
  const revised = withBody('<w:p><w:r><w:t>Hello brave world</w:t></w:r></w:p><w:p><w:r><w:t>Tail</w:t></w:r></w:p>');
  const { accepted, rejected } = await compareRoundTrip(base, revised, { author: 'Alice', date: '2026-09-28T00:00:00Z' });
  assert.deepEqual(paragraphTexts(accepted), paragraphTexts(revised));
  assert.deepEqual(paragraphTexts(rejected), paragraphTexts(base));
});

test('compare does not mutate input documents or their revisions', async () => {
  const base = withBody('<w:p><w:r><w:t>Base</w:t></w:r></w:p>');
  const revised = withBody('<w:p><w:r><w:t>Revised</w:t></w:r></w:p>');
  const beforeBaseBytes = await base.toUint8Array();
  const beforeRevisedBytes = await revised.toUint8Array();
  const beforeBaseRevision = base.revision;
  const beforeRevisedRevision = revised.revision;
  DocxDocument.compare(base, revised, { author: 'Alice' });
  assert.equal(base.revision, beforeBaseRevision);
  assert.equal(revised.revision, beforeRevisedRevision);
  assert.deepEqual(await base.toUint8Array(), beforeBaseBytes);
  assert.deepEqual(await revised.toUint8Array(), beforeRevisedBytes);
});

test('compare returns a new document whose revision starts at 0', () => {
  const compared = DocxDocument.compare(
    withBody('<w:p><w:r><w:t>A</w:t></w:r></w:p>'),
    withBody('<w:p><w:r><w:t>B</w:t></w:r></w:p>'),
  );
  assert.equal(compared.revision, 0);
});

test('compare keeps identical documents revision-free', () => {
  const compared = DocxDocument.compare(
    withBody('<w:p><w:r><w:t>Same</w:t></w:r></w:p>'),
    withBody('<w:p><w:r><w:t>Same</w:t></w:r></w:p>'),
  );
  assert.deepEqual(compared.getRevisions(), []);
});

test('compare tracks pure paragraph insertion', async () => {
  const base = withBody('<w:p><w:r><w:t>A</w:t></w:r></w:p>');
  const revised = withBody('<w:p><w:r><w:t>A</w:t></w:r></w:p><w:p><w:r><w:t>B</w:t></w:r></w:p>');
  const { compared, accepted, rejected } = await compareRoundTrip(base, revised, { author: 'Alice' });
  assert.equal(compared.getRevisions().some((revision) => revision.kind === 'insertion'), true);
  assert.deepEqual(paragraphTexts(accepted), ['A', 'B']);
  assert.deepEqual(paragraphTexts(rejected), ['A']);
});

test('compare tracks pure paragraph deletion', async () => {
  const base = withBody('<w:p><w:r><w:t>A</w:t></w:r></w:p><w:p><w:r><w:t>B</w:t></w:r></w:p>');
  const revised = withBody('<w:p><w:r><w:t>A</w:t></w:r></w:p>');
  const { compared, accepted, rejected } = await compareRoundTrip(base, revised, { author: 'Alice' });
  assert.equal(compared.getRevisions().some((revision) => revision.kind === 'deletion'), true);
  assert.deepEqual(paragraphTexts(accepted), ['A']);
  assert.deepEqual(paragraphTexts(rejected), ['A', 'B']);
});

test('compare tracks partial text replacement with insertion and deletion markers', () => {
  const compared = DocxDocument.compare(
    withBody('<w:p><w:r><w:t>Hello world</w:t></w:r></w:p>'),
    withBody('<w:p><w:r><w:t>Hello Earth</w:t></w:r></w:p>'),
    { author: 'Alice', date: '2026-09-28T00:00:00Z' },
  );
  const kinds = compared.getRevisions().map((revision) => revision.kind);
  assert.equal(kinds.includes('deletion'), true);
  assert.equal(kinds.includes('insertion'), true);
  assert.match(compared.getPartXml(compared.mainDocumentPath), /<w:del [^>]*w:date="2026-09-28T00:00:00Z"/);
});

test('compare handles paragraph reorder as delete plus insert', async () => {
  const base = withBody('<w:p><w:r><w:t>A</w:t></w:r></w:p><w:p><w:r><w:t>B</w:t></w:r></w:p><w:p><w:r><w:t>C</w:t></w:r></w:p>');
  const revised = withBody('<w:p><w:r><w:t>B</w:t></w:r></w:p><w:p><w:r><w:t>A</w:t></w:r></w:p><w:p><w:r><w:t>C</w:t></w:r></w:p>');
  const { accepted, rejected } = await compareRoundTrip(base, revised, { author: 'Alice' });
  assert.deepEqual(paragraphTexts(accepted), ['B', 'A', 'C']);
  assert.deepEqual(paragraphTexts(rejected), ['A', 'B', 'C']);
});

test('compare writes run format changes as rPrChange and reject restores base formatting', async () => {
  const base = withBody('<w:p><w:r><w:t>A</w:t></w:r></w:p>');
  const revised = withBody('<w:p><w:r><w:rPr><w:b/><w:sz w:val="28"/></w:rPr><w:t>A</w:t></w:r></w:p>');
  const { compared, accepted, rejected } = await compareRoundTrip(base, revised, { author: 'Alice' });
  assert.equal(compared.getRevisions().some((revision) => revision.kind === 'runFormatChange'), true);
  assert.match(compared.getPartXml(compared.mainDocumentPath), /<w:rPr><w:b w:val="1"\/><w:sz w:val="28"\/><w:szCs w:val="28"\/><w:rPrChange /);
  assert.equal(accepted.getParagraphs()[0].runs[0].bold, true);
  assert.equal(accepted.getParagraphs()[0].runs[0].fontSize, 14);
  assert.equal(rejected.getParagraphs()[0].runs[0].bold, undefined);
});

test('compare writes paragraph format changes as pPrChange and reject restores base paragraph format', async () => {
  const base = withBody('<w:p><w:r><w:t>A</w:t></w:r></w:p>');
  const revised = withBody('<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>A</w:t></w:r></w:p>');
  const { compared, accepted, rejected } = await compareRoundTrip(base, revised, { author: 'Alice' });
  assert.equal(compared.getRevisions().some((revision) => revision.kind === 'paragraphFormatChange'), true);
  assert.match(compared.getPartXml(compared.mainDocumentPath), /<w:pPr><w:jc w:val="center"\/><w:pPrChange /);
  assert.equal(accepted.getParagraphs()[0].alignment, 'center');
  assert.equal(rejected.getParagraphs()[0].alignment, undefined);
});

test('compare keeps revision child order for run and paragraph format changes', () => {
  const compared = DocxDocument.compare(
    withBody('<w:p><w:pPr><w:jc w:val="left"/></w:pPr><w:r><w:t>A</w:t></w:r></w:p>'),
    withBody('<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>A</w:t></w:r></w:p>'),
    { author: 'Alice' },
  );
  const xml = compared.getPartXml(compared.mainDocumentPath);
  assert.match(xml, /<w:pPr><w:jc w:val="center"\/><w:pPrChange /);
  assert.match(xml, /<w:rPr><w:b w:val="1"\/><w:rPrChange /);
});

test('compare treats table changes as coarse row deletion plus insertion', async () => {
  const base = withBody('<w:tbl><w:tr><w:tc><w:p><w:r><w:t>old</w:t></w:r></w:p></w:tc></w:tr></w:tbl>');
  const revised = withBody('<w:tbl><w:tr><w:tc><w:p><w:r><w:t>new</w:t></w:r></w:p></w:tc></w:tr></w:tbl>');
  const { compared, accepted, rejected } = await compareRoundTrip(base, revised, { author: 'Alice' });
  assert.deepEqual(compared.getRevisions().map((revision) => revision.kind), ['insertion', 'deletion']);
  assert.deepEqual(accepted.getTable(0).rows[0].cells[0].blocks[0].paragraph.text, 'new');
  assert.deepEqual(rejected.getTable(0).rows[0].cells[0].blocks[0].paragraph.text, 'old');
});

test('compare rejects oversized paragraph counts with a clear error', () => {
  const paragraphs = Array.from({ length: 1001 }, (_, index) => `<w:p><w:r><w:t>p${index}</w:t></w:r></w:p>`).join('');
  assert.throws(
    () => DocxDocument.compare(withBody(paragraphs), withBody('<w:p><w:r><w:t>x</w:t></w:r></w:p>')),
    /at most 1000 main-document paragraphs/,
  );
});

test('compare can stack on a document that already contains revisions', async () => {
  const base = DocxDocument.compare(
    withBody('<w:p><w:r><w:t>A</w:t></w:r></w:p>'),
    withBody('<w:p><w:r><w:t>AB</w:t></w:r></w:p>'),
    { author: 'Alice' },
  );
  const revised = withBody('<w:p><w:r><w:t>ABC</w:t></w:r></w:p>');
  const { accepted } = await compareRoundTrip(base, revised, { author: 'Bob' });
  assert.deepEqual(paragraphTexts(accepted), ['ABC']);
  assert.equal(DocxDocument.compare(base, revised, { author: 'Bob' }).getRevisions().length >= 2, true);
});

test('compare can apply insertion and run-format change in the same paragraph', async () => {
  const base = withBody('<w:p><w:r><w:t>AB</w:t></w:r></w:p>');
  const revised = withBody('<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>A</w:t></w:r><w:r><w:t>B</w:t></w:r><w:r><w:t>!</w:t></w:r></w:p>');
  const { compared, accepted, rejected } = await compareRoundTrip(base, revised, { author: 'Alice' });
  assert.equal(compared.getRevisions().some((revision) => revision.kind === 'runFormatChange'), true);
  assert.equal(compared.getRevisions().some((revision) => revision.kind === 'insertion'), true);
  assert.deepEqual(paragraphTexts(accepted), ['AB!']);
  assert.equal(rejected.getParagraphs()[0].runs[0].bold, undefined);
});

test('compare degrades zero-length formatted runs to coarse paragraph replacement', async () => {
  const base = withBody('<w:p><w:r><w:t>A</w:t></w:r></w:p>');
  const revised = withBody('<w:p><w:r><w:rPr><w:b/></w:rPr></w:r><w:r><w:t>A</w:t></w:r></w:p>');
  const { compared, accepted, rejected } = await compareRoundTrip(base, revised, { author: 'Alice' });
  assert.equal(compared.getRevisions().some((revision) => revision.kind === 'runFormatChange'), false);
  assert.equal(compared.getRevisions().some((revision) => revision.kind === 'insertion'), true);
  assert.deepEqual(paragraphTexts(accepted), ['A']);
  assert.deepEqual(paragraphTexts(rejected), ['A']);
});

test('compare preserves marker-only wrapper content when deleting a no-run paragraph', () => {
  const base = withBody('<w:p><w:sdt><w:sdtPr/><w:sdtContent><w:bookmarkStart w:id="1" w:name="keep"/><w:bookmarkEnd w:id="1"/></w:sdtContent></w:sdt></w:p><w:p><w:r><w:t>tail</w:t></w:r></w:p>');
  const revised = withBody('<w:p><w:r><w:t>tail</w:t></w:r></w:p>');
  const compared = DocxDocument.compare(base, revised, { author: 'Alice' });
  const xml = compared.getPartXml(compared.mainDocumentPath);
  assert.match(xml, /bookmarkStart/);
  assert.match(xml, /bookmarkEnd/);
});

test('compare validates compare options through assertText', () => {
  assert.throws(
    () => DocxDocument.compare(withBody('<w:p/>'), withBody('<w:p/>'), { author: 'bad\u0000' }),
    /author/,
  );
  assert.throws(
    () => DocxDocument.compare(withBody('<w:p/>'), withBody('<w:p/>'), { date: 'bad\u0000' }),
    /date/,
  );
});
