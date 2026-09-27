import { test } from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import { DocxDocument } from '../dist/document.js';
import { AGENT_OPERATION_SCHEMA } from '../dist/operations.js';
import { WORD_NS } from '../dist/xml.js';

function withBody(xml) {
  const doc = DocxDocument.create();
  doc.setPartXml(doc.mainDocumentPath, `<w:document xmlns:w="${WORD_NS}"><w:body>${xml}<w:sectPr/></w:body></w:document>`);
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
  assert.equal(emptyDoc.getParagraphs().length, 1);
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
  assert.equal(AGENT_OPERATION_SCHEMA.properties.operations.items.oneOf.length, 8);
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
