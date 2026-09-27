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
  assert.equal(AGENT_OPERATION_SCHEMA.properties.operations.items.oneOf.length, 13);
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

test('insertFootnote creates note part, reference run and visible marker data', () => {
  const doc = DocxDocument.create();
  doc.setParagraphText(0, '正文');
  const note = doc.insertFootnote(0, 1, '脚注内容');
  assert.equal(note.kind, 'footnote');
  assert.equal(note.number, 1);
  assert.equal(doc.getParagraphs()[0].text, '正文');
  assert.equal(doc.getParagraphs()[0].runs.at(-1).noteReference.marker, '1');
  assert.match(doc.getPartXml('word/footnotes.xml'), /脚注内容/);
  assert.match(doc.getPartXml('word/_rels\/document.xml.rels'), /footnotes/);
});

test('insertEndnote supports custom marker and keeps custom mark text', () => {
  const doc = DocxDocument.create();
  const note = doc.insertEndnote(0, 1, '尾注内容', { customMark: '*' });
  assert.equal(note.marker, '*');
  assert.match(doc.getPartXml('word/endnotes.xml'), /\*/);
  assert.match(doc.getPartXml(doc.mainDocumentPath), /customMarkFollows/);
});

test('footnote numbering follows reference order rather than id order', () => {
  const doc = DocxDocument.create();
  doc.setPartXml(doc.mainDocumentPath, `<w:document xmlns:w="${WORD_NS}"><w:body>
    <w:p><w:r><w:t>A</w:t></w:r><w:r><w:footnoteReference w:id="9"/></w:r></w:p>
    <w:p><w:r><w:t>B</w:t></w:r><w:r><w:footnoteReference w:id="3"/></w:r></w:p>
    <w:sectPr/></w:body></w:document>`);
  doc.addPart('word/footnotes.xml', new TextEncoder().encode(`<w:footnotes xmlns:w="${WORD_NS}">
    <w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>
    <w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>
    <w:footnote w:id="3"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t>三</w:t></w:r></w:p></w:footnote>
    <w:footnote w:id="9"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t>九</w:t></w:r></w:p></w:footnote>
  </w:footnotes>`), 'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml');
  doc.addPart('word/settings.xml', new TextEncoder().encode(`<w:settings xmlns:w="${WORD_NS}"><w:footnotePr><w:numFmt w:val="lowerRoman"/></w:footnotePr></w:settings>`), 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml');
  const notes = doc.getFootnotes();
  assert.deepEqual(notes.map(item => item.id), [9, 3]);
  assert.deepEqual(notes.map(item => item.marker), ['i', 'ii']);
});

test('footnote references inside table cells are recognized', () => {
  const doc = withBody('<w:tbl><w:tr><w:tc><w:p><w:r><w:t>x</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r></w:p></w:tc></w:tr></w:tbl>');
  doc.addPart('word/footnotes.xml', new TextEncoder().encode(`<w:footnotes xmlns:w="${WORD_NS}">
    <w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>
    <w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>
    <w:footnote w:id="1"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t>cell</w:t></w:r></w:p></w:footnote>
  </w:footnotes>`), 'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml');
  const note = doc.getFootnotes()[0];
  assert.equal(note.reference.paragraph, 0);
  assert.equal(note.blocks[0].paragraph.text.trim(), 'cell');
});

test('separator and continuationSeparator are hidden from getFootnotes results', () => {
  const doc = DocxDocument.create();
  doc.insertFootnote(0, 1, 'visible');
  const xml = doc.getPartXml('word/footnotes.xml');
  assert.match(xml, /w:type="separator"/);
  assert.match(xml, /w:type="continuationSeparator"/);
  assert.equal(doc.getFootnotes().length, 1);
});

test('deleteNote removes both document reference runs and note entries', () => {
  const doc = DocxDocument.create();
  const note = doc.insertFootnote(0, 1, 'to delete');
  doc.deleteNote('footnote', note.id);
  assert.equal(doc.getFootnotes().length, 0);
  assert.doesNotMatch(doc.getPartXml(doc.mainDocumentPath), /footnoteReference/);
  assert.doesNotMatch(doc.getPartXml('word/footnotes.xml'), /to delete/);
});

test('setNoteText rewrites note body and keeps reference marker run', () => {
  const doc = DocxDocument.create();
  const note = doc.insertFootnote(0, 1, 'old');
  doc.setNoteText('footnote', note.id, 'new text');
  assert.match(doc.getPartXml('word/footnotes.xml'), /new text/);
  assert.match(doc.getPartXml('word/footnotes.xml'), /footnoteRef/);
});

test('convertNote moves a footnote to endnotes and rewrites references', () => {
  const doc = DocxDocument.create();
  const note = doc.insertFootnote(0, 1, 'convert me');
  doc.insertEndnote(0, 1, 'existing');
  doc.convertNote('footnote', note.id);
  assert.equal(doc.getFootnotes().length, 0);
  assert.equal(doc.getEndnotes().length, 2);
  assert.doesNotMatch(doc.getPartXml(doc.mainDocumentPath), /footnoteReference/);
  assert.match(doc.getPartXml(doc.mainDocumentPath), /endnoteReference/);
  assert.match(doc.getPartXml('word/endnotes.xml'), /endnoteRef/);
  assert.doesNotMatch(doc.getPartXml('word/endnotes.xml'), /footnoteRef/);
  assert.match(doc.getPartXml(doc.mainDocumentPath), /w:id="2"/);
});

test('dangling note reference does not crash and yields empty blocks', () => {
  const doc = withBody('<w:p><w:r><w:footnoteReference w:id="99"/></w:r></w:p>');
  assert.doesNotThrow(() => doc.getFootnotes());
  const note = doc.getFootnotes()[0];
  assert.equal(note.id, 99);
  assert.deepEqual(note.blocks, []);
});

test('insertFootnote creates missing footnotes.xml automatically', () => {
  const doc = DocxDocument.create();
  assert.equal(doc.listParts().includes('word/footnotes.xml'), false);
  doc.insertFootnote(0, 1, 'created');
  assert.equal(doc.listParts().includes('word/footnotes.xml'), true);
});

test('note settings read and write through settings.xml', () => {
  const doc = DocxDocument.create();
  doc.setNoteSettings({
    footnote: { numFmt: 'lowerRoman', numStart: 3, numRestart: 'eachSect' },
    endnote: { numFmt: 'upperRoman', numStart: 5, numRestart: 'continuous' },
  });
  const settings = doc.getNoteSettings();
  assert.equal(settings.footnote.numFmt, 'lowerRoman');
  assert.equal(settings.footnote.numStart, 3);
  assert.equal(settings.footnote.numRestart, 'eachSect');
  assert.equal(settings.endnote.numFmt, 'upperRoman');
  assert.equal(settings.endnote.numStart, 5);
  assert.equal(settings.endnote.numRestart, 'continuous');
  assert.match(doc.getPartXml('word/settings.xml'), /footnotePr/);
  assert.match(doc.getPartXml('word/settings.xml'), /endnotePr/);
  assert.match(doc.getPartXml('word/_rels/document.xml.rels'), /relationships\/settings/);
});

test('agent note operations validate and run in transactions', () => {
  const doc = DocxDocument.create();
  const result = doc.applyOperations({ operations: [
    { type: 'insertFootnote', paragraph: 0, run: 1, text: 'x' },
    { type: 'setNoteText', kind: 'footnote', id: 1, text: 'y' },
    { type: 'convertNote', kind: 'footnote', id: 1 },
  ] });
  assert.equal(result.endnotes[0].marker, '1');
  assert.throws(() => doc.applyOperations({ operations: [{ type: 'deleteNote', kind: 'bad', id: 1 }] }), /Invalid note kind/);
});
