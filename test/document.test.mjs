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

test('reads external hyperlink relationship and run metadata', () => {
  const doc = withBody('<w:p><w:hyperlink r:id="rId9" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:r><w:t>官网</w:t></w:r></w:hyperlink></w:p>');
  doc.addPart('word/_rels/document.xml.rels', new TextEncoder().encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/></Relationships>'), 'application/vnd.openxmlformats-package.relationships+xml');
  const links = doc.getHyperlinks();
  assert.equal(links.length, 1);
  assert.equal(links[0].url, 'https://example.com');
  assert.equal(doc.getParagraphs()[0].runs[0].hyperlink?.url, 'https://example.com');
});

test('marks javascript and data hyperlinks as unsafe', () => {
  const doc = withBody('<w:p><w:hyperlink r:id="rId1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:r><w:t>bad</w:t></w:r></w:hyperlink></w:p>');
  doc.addPart('word/_rels/document.xml.rels', new TextEncoder().encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="javascript:alert(1)" TargetMode="External"/></Relationships>'), 'application/vnd.openxmlformats-package.relationships+xml');
  assert.equal(doc.getHyperlinks()[0].unsafe, true);
  assert.equal(doc.getParagraphs()[0].runs[0].hyperlink?.unsafe, true);
});

test('marks external javascript hyperlink unsafe even when anchor is present', () => {
  const doc = withBody('<w:p><w:hyperlink r:id="rId7" w:anchor="top" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:r><w:t>x</w:t></w:r></w:hyperlink></w:p>');
  doc.addPart('word/_rels/document.xml.rels', new TextEncoder().encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="javascript:alert(1)" TargetMode="External"/></Relationships>'), 'application/vnd.openxmlformats-package.relationships+xml');
  const link = doc.getHyperlinks()[0];
  assert.equal(link.anchor, 'top');
  assert.equal(link.unsafe, true);
});

test('reads anchor hyperlinks and tooltip', () => {
  const doc = withBody('<w:p><w:hyperlink w:anchor="chapter1" w:tooltip="跳转" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:t>目录</w:t></w:r></w:hyperlink></w:p>');
  const link = doc.getHyperlinks()[0];
  assert.equal(link.anchor, 'chapter1');
  assert.equal(link.tooltip, '跳转');
  assert.equal(link.isExternal, false);
});

test('parses fldSimple HYPERLINK field URLs', () => {
  const doc = withBody('<w:p><w:fldSimple w:instr=" HYPERLINK &quot;https://legacy.example&quot; "><w:r><w:t>legacy</w:t></w:r></w:fldSimple></w:p>');
  const link = doc.getHyperlinks()[0];
  assert.equal(link.url, 'https://legacy.example');
  assert.equal(link.text, 'legacy');
});

test('bookmarks are listed and internal entries hidden by default', () => {
  const doc = withBody('<w:p><w:bookmarkStart w:id="1" w:name="user"/><w:r><w:t>a</w:t></w:r></w:p><w:p><w:r><w:t>b</w:t></w:r><w:bookmarkEnd w:id="1"/><w:bookmarkStart w:id="2" w:name="_GoBack"/><w:bookmarkEnd w:id="2"/></w:p>');
  const visible = doc.getBookmarks();
  const all = doc.getBookmarks({ includeInternal: true });
  assert.equal(visible.length, 1);
  assert.equal(visible[0].name, 'user');
  assert.equal(visible[0].endParagraph, 1);
  assert.equal(all.length, 2);
});

test('getBookmarks includes body-level bookmarks and duplicate check rejects their names', () => {
  const doc = withBody('<w:bookmarkStart w:id="5" w:name="tbl"/><w:p><w:r><w:t>row</w:t></w:r></w:p><w:bookmarkEnd w:id="5"/>');
  const all = doc.getBookmarks({ includeInternal: true });
  assert.equal(all[0].name, 'tbl');
  assert.equal(all[0].startParagraph, 0);
  assert.equal(all[0].endParagraph, 0);
  assert.throws(() => doc.insertBookmark('tbl', { startParagraph: 0 }), /already exists/);
});

test('insertBookmark and deleteBookmark maintain ids and boundaries', () => {
  const doc = DocxDocument.create();
  doc.insertParagraph('next');
  const mark = doc.insertBookmark('range', { startParagraph: 0, endParagraph: 1 });
  assert.equal(mark.startParagraph, 0);
  assert.equal(mark.endParagraph, 1);
  assert.equal(doc.getBookmarks().length, 1);
  doc.deleteBookmark('range');
  assert.equal(doc.getBookmarks().length, 0);
});

test('insertBookmark rejects duplicate names', () => {
  const doc = DocxDocument.create();
  doc.insertBookmark('dup', { startParagraph: 0 });
  assert.throws(() => doc.insertBookmark('dup', { startParagraph: 0 }), /already exists/);
});

test('insertHyperlink wraps selected text and creates external relationship', () => {
  const doc = DocxDocument.create();
  doc.setParagraphText(0, 'hello world');
  const link = doc.insertHyperlink({ paragraph: 0, start: 6, end: 11 }, { url: 'https://example.com', tooltip: 'site' });
  assert.equal(link.text, 'world');
  assert.equal(link.url, 'https://example.com');
  assert.match(doc.getPartXml(doc.mainDocumentPath), /w:hyperlink/);
  assert.match(doc.getPartXml('word/_rels/document.xml.rels'), /Target="https:\/\/example.com"/);
});

test('insertHyperlink handles runs inside wrappers and returns inserted hyperlink', () => {
  const doc = withBody('<w:p><w:hyperlink w:anchor="later"><w:r><w:t>later link</w:t></w:r></w:hyperlink></w:p><w:p><w:ins><w:r><w:t>linked text</w:t></w:r></w:ins></w:p>');
  const link = doc.insertHyperlink({ paragraph: 1, start: 0, end: 6 }, { url: 'https://example.com' });
  assert.equal(link.paragraph, 1);
  assert.equal(link.text, 'linked');
  assert.equal(link.url, 'https://example.com');
});

test('insertHyperlink failures do not create orphan rels or increment revision', () => {
  const doc = DocxDocument.create();
  doc.setParagraphText(0, 'abc');
  const revision = doc.revision;
  assert.throws(() => doc.insertHyperlink({ paragraph: 0, start: 0, end: 99 }, { url: 'https://orphan.example' }), /out of bounds/);
  assert.equal(doc.revision, revision);
  assert.throws(() => doc.getPartXml('word/_rels/document.xml.rels'), /Package part not found/);
});

test('insertHyperlink rejects non-whitelisted schemes', () => {
  const doc = DocxDocument.create();
  doc.setParagraphText(0, 'x');
  assert.throws(() => doc.insertHyperlink({ paragraph: 0, start: 0, end: 1 }, { url: 'javascript:alert(1)' }), /http, https or mailto/);
  assert.throws(() => doc.insertHyperlink({ paragraph: 0, start: 0, end: 1 }, { url: 'file:///tmp/a' }), /http, https or mailto/);
});

test('insertHyperlink supports mailto and anchors', () => {
  const doc = DocxDocument.create();
  doc.setParagraphText(0, 'contact');
  const external = doc.insertHyperlink({ paragraph: 0, start: 0, end: 7 }, { url: 'mailto:test@example.com' });
  assert.equal(external.unsafe, false);
  doc.removeHyperlink(0);
  doc.insertHyperlink({ paragraph: 0, start: 0, end: 7 }, { anchor: 'dest' });
  assert.equal(doc.getHyperlinks()[0].anchor, 'dest');
});

test('removeHyperlink keeps text by default and cleans dangling relationship', () => {
  const doc = DocxDocument.create();
  doc.setParagraphText(0, 'demo');
  doc.insertHyperlink({ paragraph: 0, start: 0, end: 4 }, { url: 'https://example.com' });
  const id = doc.getHyperlinks()[0].relationshipId;
  doc.removeHyperlink(0);
  assert.equal(doc.getParagraphs()[0].text, 'demo');
  assert.equal(doc.getHyperlinks().length, 0);
  assert.doesNotMatch(doc.getPartXml('word/_rels/document.xml.rels'), new RegExp(`Id="${id}"`));
});

test('removeHyperlink with keepText false removes linked text', () => {
  const doc = DocxDocument.create();
  doc.setParagraphText(0, 'demo');
  doc.insertHyperlink({ paragraph: 0, start: 0, end: 4 }, { url: 'https://example.com' });
  doc.removeHyperlink(0, { keepText: false });
  assert.equal(doc.getParagraphs()[0].text, '');
});

test('updateHyperlink rewrites attributes and relationship target', () => {
  const doc = DocxDocument.create();
  doc.setParagraphText(0, 'abc');
  doc.insertHyperlink({ paragraph: 0, start: 0, end: 3 }, { url: 'https://a.example' });
  doc.updateHyperlink(0, { url: 'https://b.example', tooltip: 'B' });
  assert.equal(doc.getHyperlinks()[0].url, 'https://b.example');
  assert.equal(doc.getHyperlinks()[0].tooltip, 'B');
  assert.match(doc.getPartXml('word/_rels/document.xml.rels'), /https:\/\/b.example/);
});

test('updateHyperlink supports fldSimple links', () => {
  const doc = withBody('<w:p><w:fldSimple w:instr="HYPERLINK &quot;https://legacy.example&quot;"><w:r><w:t>legacy</w:t></w:r></w:fldSimple></w:p>');
  doc.updateHyperlink(0, { url: 'https://new.example?q=""', anchor: 'a"b', tooltip: 'tip' });
  assert.equal(doc.getHyperlinks()[0].url, 'https://new.example?q=""');
  assert.equal(doc.getHyperlinks()[0].anchor, 'a"b');
  assert.match(doc.getPartXml(doc.mainDocumentPath), /w:tooltip="tip"/);
  const fld = doc.getPartDocument(doc.mainDocumentPath).getElementsByTagNameNS(WORD_NS, 'fldSimple')[0];
  const instruction = fld?.getAttributeNS(WORD_NS, 'instr') ?? '';
  assert.match(instruction, /\?q=\\\"\\\"/);
  assert.match(instruction, /\\l "a\\\"b"/);
});

test('updateHyperlink does not retarget other links sharing a relationship id', () => {
  const doc = withBody('<w:p><w:hyperlink r:id="rId1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:r><w:t>one</w:t></w:r></w:hyperlink><w:r><w:t> </w:t></w:r><w:hyperlink r:id="rId1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:r><w:t>two</w:t></w:r></w:hyperlink></w:p>');
  doc.addPart('word/_rels/document.xml.rels', new TextEncoder().encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://old.example" TargetMode="External"/></Relationships>'), 'application/vnd.openxmlformats-package.relationships+xml');
  doc.updateHyperlink(0, { url: 'https://new.example' });
  const links = doc.getHyperlinks();
  assert.equal(links[0].url, 'https://new.example');
  assert.equal(links[1].url, 'https://old.example');
  assert.match(doc.getPartXml('word/_rels/document.xml.rels'), /https:\/\/new.example/);
  assert.match(doc.getPartXml('word/_rels/document.xml.rels'), /https:\/\/old.example/);
});

test('updateHyperlink can convert external link to anchor-only link', () => {
  const doc = DocxDocument.create();
  doc.setParagraphText(0, 'abc');
  doc.insertHyperlink({ paragraph: 0, start: 0, end: 3 }, { url: 'https://a.example' });
  const id = doc.getHyperlinks()[0].relationshipId;
  doc.updateHyperlink(0, { anchor: 'dest' });
  assert.equal(doc.getHyperlinks()[0].url, undefined);
  assert.equal(doc.getHyperlinks()[0].anchor, 'dest');
  assert.doesNotMatch(doc.getPartXml(doc.mainDocumentPath), /r:id=/);
  assert.doesNotMatch(doc.getPartXml('word/_rels/document.xml.rels'), new RegExp(`Id="${id}"`));
});

test('agent operations validate hyperlink and bookmark commands', () => {
  const doc = DocxDocument.create();
  assert.throws(() => doc.applyOperations({ operations: [{ type: 'insertHyperlink', target: { paragraph: 0, start: 0, end: 0 }, link: { url: 'javascript:1' } }] }), /http, https or mailto/);
  assert.throws(() => doc.applyOperations({ operations: [{
    type: 'insertHyperlink',
    target: { paragraph: 0, start: 0, end: 0 },
    link: { url: 'https://ok.example', tooltip: 'a\u0000b' },
  }] }), /valid XML/);
  assert.throws(() => doc.applyOperations({ operations: [{ type: 'removeHyperlink', hyperlink: { paragraph: 0 } }] }), /hyperlink\.(text|runs)/);
  const snapshot = doc.applyOperations({ operations: [{ type: 'insertBookmark', name: 'b1', range: { startParagraph: 0 } }] });
  assert.equal(snapshot.bookmarks[0].name, 'b1');
  doc.applyOperations({ operations: [{ type: 'deleteBookmark', name: 'b1' }] });
  assert.equal(doc.getBookmarks().length, 0);
  assert.throws(() => doc.insertBookmark('reverse', { startParagraph: 1, endParagraph: 0 }), />=/);
});
