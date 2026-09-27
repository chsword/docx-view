import { test } from 'node:test';
import assert from 'node:assert/strict';
import { DocxDocument } from '../dist/document.js';
import { REL_NS, WORD_NS } from '../dist/xml.js';
import { AGENT_OPERATION_SCHEMA } from '../dist/operations.js';

const NUMBERING_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml';
const RELS_TYPE = 'application/vnd.openxmlformats-package.relationships+xml';
const encoder = new TextEncoder();

function withBody(xml) {
  const doc = DocxDocument.create();
  doc.setPartXml(doc.mainDocumentPath, `<w:document xmlns:w="${WORD_NS}"><w:body>${xml}<w:sectPr/></w:body></w:document>`);
  return doc;
}

function attachNumbering(doc, numberingXml, stylesXml) {
  doc.addPart('word/numbering.xml', encoder.encode(numberingXml), NUMBERING_TYPE);
  doc.addPart('word/_rels/document.xml.rels', encoder.encode(`<Relationships xmlns="${REL_NS}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>`), RELS_TYPE);
  if (stylesXml) doc.addPart('word/styles.xml', encoder.encode(stylesXml), 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml');
}

test('paragraphs with numPr but no numbering.xml degrade without throwing', () => {
  const doc = withBody('<w:p><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>A</w:t></w:r></w:p>');
  assert.equal(doc.getParagraphs()[0].numbering, undefined);
});

test('dangling numId degrades without throwing', () => {
  const doc = withBody('<w:p><w:pPr><w:numPr><w:numId w:val="99"/></w:numPr></w:pPr><w:r><w:t>A</w:t></w:r></w:p>');
  attachNumbering(doc, `<w:numbering xmlns:w="${WORD_NS}"></w:numbering>`);
  assert.equal(doc.getParagraphs()[0].numbering, undefined);
});

test('bullet numbering parses marker font and visible bullet text', () => {
  const doc = withBody('<w:p><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Bullet</w:t></w:r></w:p>');
  attachNumbering(doc, `<w:numbering xmlns:w="${WORD_NS}"><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val=""/><w:suff w:val="space"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol"/></w:rPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num></w:numbering>`);
  const numbering = doc.getParagraphs()[0].numbering;
  assert.equal(numbering.text, '•');
  assert.equal(numbering.isBullet, true);
  assert.equal(numbering.suffix, 'space');
  assert.equal(numbering.runFormat.fontFamily, 'Symbol');
});

test('numbering counts main-body paragraphs including table cells', () => {
  const doc = withBody('<w:p><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>A</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:p><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>B</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>C</w:t></w:r></w:p>');
  attachNumbering(doc, `<w:numbering xmlns:w="${WORD_NS}"><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num></w:numbering>`);
  assert.deepEqual(doc.getParagraphs().map(paragraph => paragraph.numbering?.text), ['1.', '2.', '3.']);
});

test('multilevel placeholders use current counters', () => {
  const doc = withBody('<w:p><w:pPr><w:numPr><w:numId w:val="1"/><w:ilvl w:val="0"/></w:numPr></w:pPr><w:r><w:t>A</w:t></w:r></w:p><w:p><w:pPr><w:numPr><w:numId w:val="1"/><w:ilvl w:val="1"/></w:numPr></w:pPr><w:r><w:t>B</w:t></w:r></w:p><w:p><w:pPr><w:numPr><w:numId w:val="1"/><w:ilvl w:val="2"/></w:numPr></w:pPr><w:r><w:t>C</w:t></w:r></w:p>');
  attachNumbering(doc, `<w:numbering xmlns:w="${WORD_NS}"><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl><w:lvl w:ilvl="1"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2."/></w:lvl><w:lvl w:ilvl="2"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2.%3."/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num></w:numbering>`);
  assert.deepEqual(doc.getParagraphs().map(paragraph => paragraph.numbering?.text), ['1.', '1.1.', '1.1.1.']);
});

test('startOverride changes the initial value for an instance level', () => {
  const doc = withBody('<w:p><w:pPr><w:numPr><w:numId w:val="7"/></w:numPr></w:pPr><w:r><w:t>A</w:t></w:r></w:p><w:p><w:pPr><w:numPr><w:numId w:val="7"/></w:numPr></w:pPr><w:r><w:t>B</w:t></w:r></w:p>');
  attachNumbering(doc, `<w:numbering xmlns:w="${WORD_NS}"><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="7"><w:abstractNumId w:val="1"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="4"/></w:lvlOverride></w:num></w:numbering>`);
  assert.deepEqual(doc.getParagraphs().map(paragraph => paragraph.numbering?.text), ['4.', '5.']);
});

test('legal numbering forces placeholder output to decimal', () => {
  const doc = withBody('<w:p><w:pPr><w:numPr><w:numId w:val="3"/><w:ilvl w:val="0"/></w:numPr></w:pPr><w:r><w:t>A</w:t></w:r></w:p><w:p><w:pPr><w:numPr><w:numId w:val="3"/><w:ilvl w:val="1"/></w:numPr></w:pPr><w:r><w:t>B</w:t></w:r></w:p>');
  attachNumbering(doc, `<w:numbering xmlns:w="${WORD_NS}"><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:numFmt w:val="lowerRoman"/><w:lvlText w:val="%1."/></w:lvl><w:lvl w:ilvl="1"><w:numFmt w:val="lowerLetter"/><w:isLgl/><w:lvlText w:val="%1.%2."/></w:lvl></w:abstractNum><w:num w:numId="3"><w:abstractNumId w:val="1"/></w:num></w:numbering>`);
  assert.deepEqual(doc.getParagraphs().map(paragraph => paragraph.numbering?.text), ['i.', '1.1.']);
});

test('paragraph styles with numPr contribute numbering', () => {
  const doc = withBody('<w:p><w:pPr><w:pStyle w:val="ListStyle"/></w:pPr><w:r><w:t>A</w:t></w:r></w:p>');
  attachNumbering(doc, `<w:numbering xmlns:w="${WORD_NS}"><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="5"><w:abstractNumId w:val="1"/></w:num></w:numbering>`, `<w:styles xmlns:w="${WORD_NS}"><w:style w:type="paragraph" w:styleId="ListStyle"><w:pPr><w:numPr><w:numId w:val="5"/></w:numPr></w:pPr></w:style></w:styles>`);
  assert.equal(doc.getParagraphs()[0].numbering.text, '1.');
});

test('unknown numFmt falls back to decimal instead of throwing', () => {
  const doc = withBody('<w:p><w:pPr><w:numPr><w:numId w:val="5"/></w:numPr></w:pPr><w:r><w:t>A</w:t></w:r></w:p>');
  attachNumbering(doc, `<w:numbering xmlns:w="${WORD_NS}"><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:numFmt w:val="totallyUnknown"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="5"><w:abstractNumId w:val="1"/></w:num></w:numbering>`);
  assert.equal(doc.getParagraphs()[0].numbering.text, '1.');
});

test('createNumbering creates package parts and survives export roundtrip', async () => {
  const doc = DocxDocument.create();
  const numId = doc.createNumbering('bullet');
  doc.setParagraphNumbering(0, numId);
  const xml = doc.getPartXml('word/numbering.xml');
  assert.match(xml, /<w:abstractNum[\s\S]*<w:num /);
  assert.ok(doc.listParts().includes('word/numbering.xml'));
  assert.ok(doc.listParts().includes('word/_rels/document.xml.rels'));
  const reopened = await DocxDocument.load(await doc.toUint8Array());
  assert.equal(reopened.getParagraphs()[0].numbering.isBullet, true);
});

test('setParagraphLevel clamps between 0 and 8 and clearParagraphNumbering removes direct numPr', () => {
  const doc = DocxDocument.create();
  const numId = doc.createNumbering('multilevel');
  doc.setParagraphNumbering(0, numId, 0);
  doc.setParagraphLevel(0, 99);
  assert.equal(doc.getParagraphs()[0].numbering.level, 8);
  doc.clearParagraphNumbering(0);
  assert.equal(doc.getParagraphs()[0].numbering, undefined);
});

test('agent operations support numbering mutations atomically', () => {
  const doc = DocxDocument.create();
  const numId = doc.createNumbering('decimal');
  const snapshot = doc.applyOperations({ expectedRevision: doc.revision, operations: [
    { type: 'setParagraphNumbering', index: 0, numId },
    { type: 'setParagraphLevel', index: 0, delta: 1 },
    { type: 'clearParagraphNumbering', index: 0 },
  ] });
  assert.equal(snapshot.paragraphs[0].numbering, undefined);
  assert.equal(AGENT_OPERATION_SCHEMA.properties.operations.items.oneOf.length, 11);
});

test('cyclic numStyleLink does not throw', () => {
  const doc = withBody('<w:p><w:pPr><w:numPr><w:numId w:val="2"/></w:numPr></w:pPr><w:r><w:t>A</w:t></w:r></w:p>');
  attachNumbering(doc, `<w:numbering xmlns:w="${WORD_NS}"><w:abstractNum w:abstractNumId="1"><w:numStyleLink w:val="LoopStyle"/></w:abstractNum><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num></w:numbering>`, `<w:styles xmlns:w="${WORD_NS}"><w:style w:type="paragraph" w:styleId="LoopStyle"><w:pPr><w:numPr><w:numId w:val="2"/></w:numPr></w:pPr></w:style></w:styles>`);
  assert.equal(doc.getParagraphs()[0].numbering, undefined);
});
