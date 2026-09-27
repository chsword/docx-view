import { test } from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import { DocxDocument } from '../dist/document.js';
import { AGENT_OPERATION_SCHEMA } from '../dist/operations.js';
import { WORD_NS } from '../dist/xml.js';

const STYLES_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml';
const THEME_TYPE = 'application/vnd.openxmlformats-officedocument.theme+xml';

function withBody(xml) {
  const doc = DocxDocument.create();
  doc.setPartXml(doc.mainDocumentPath, `<w:document xmlns:w="${WORD_NS}"><w:body>${xml}<w:sectPr/></w:body></w:document>`);
  return doc;
}

function withStyles(bodyXml, stylesXml, themeXml) {
  const doc = withBody(bodyXml);
  doc.addPart('word/styles.xml', new TextEncoder().encode(stylesXml), STYLES_TYPE);
  if (themeXml) doc.addPart('word/theme/theme1.xml', new TextEncoder().encode(themeXml), THEME_TYPE);
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

test('table style conditions apply first-row and horizontal band run formatting', () => {
  const doc = withStyles(
    `<w:tbl>
      <w:tblPr><w:tblStyle w:val="FancyTable"/><w:tblLook w:firstRow="1" w:noHBand="0"/></w:tblPr>
      <w:tr><w:tc><w:p><w:r><w:t>H1</w:t></w:r></w:p></w:tc></w:tr>
      <w:tr><w:tc><w:p><w:r><w:t>R2</w:t></w:r></w:p></w:tc></w:tr>
    </w:tbl>`,
    `<w:styles xmlns:w="${WORD_NS}">
      <w:style w:type="table" w:styleId="FancyTable"><w:name w:val="Fancy Table"/>
        <w:tblStylePr w:type="firstRow"><w:rPr><w:b/></w:rPr></w:tblStylePr>
        <w:tblStylePr w:type="band2Horz"><w:rPr><w:color w:val="AA5500"/></w:rPr></w:tblStylePr>
      </w:style>
    </w:styles>`,
  );
  const paragraphs = doc.getParagraphs();
  assert.equal(paragraphs[0].runs[0].effective.bold, true);
  assert.equal(paragraphs[1].runs[0].effective.color, 'AA5500');
});
