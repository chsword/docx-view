import { test } from 'node:test';
import assert from 'node:assert/strict';
import { DocxDocument } from '../dist/document.js';
import { findReusableNumberingId } from '../dist/numbering.js';
import { parseXml, REL_NS, WORD_NS } from '../dist/xml.js';
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

function numberedDocument(levels, numberingXml = `<w:numbering xmlns:w="${WORD_NS}"><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num></w:numbering>`) {
  const paragraphs = levels.map((level, index) =>
    `<w:p><w:pPr><w:numPr><w:ilvl w:val="${level}"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>${index}</w:t></w:r></w:p>`);
  const doc = withBody(paragraphs.join(''));
  attachNumbering(doc, numberingXml);
  return doc;
}

function elementNames(element) {
  return [...element.childNodes].filter((node) => node.nodeType === 1).map((node) => node.localName);
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

test('malformed numbering relationship targets degrade without throwing', () => {
  for (const target of ['..', 'sub\numbering.xml', 'num%bering.xml']) {
    const doc = withBody('<w:p><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>A</w:t></w:r></w:p>');
    doc.addPart('word/_rels/document.xml.rels', encoder.encode(`<Relationships xmlns="${REL_NS}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="${target}"/></Relationships>`), RELS_TYPE);
    assert.doesNotThrow(() => doc.getParagraphs());
    assert.equal(doc.getParagraphs()[0].numbering, undefined);
    assert.doesNotThrow(() => doc.getBlocks());
    assert.doesNotThrow(() => doc.getSnapshot());
  }
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

test('picture bullets resolve their own image relationship and survive package roundtrip', async () => {
  const doc = withBody('<w:p><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Item</w:t></w:r></w:p>');
  const numberingXml = `<w:numbering xmlns:w="${WORD_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:v="urn:schemas-microsoft-com:vml"><w:numPicBullet w:numPicBulletId="5"><w:pict><v:shape><v:imagedata r:id="rBullet"/></v:shape></w:pict></w:numPicBullet><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="▪"/><w:lvlPicBulletId w:val="5"/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num></w:numbering>`;
  attachNumbering(doc, numberingXml);
  doc.addPart('word/media/bullet.png', Uint8Array.of(1, 2, 3), 'image/png');
  doc.addPart('word/_rels/numbering.xml.rels', encoder.encode(`<Relationships xmlns="${REL_NS}"><Relationship Id="rBullet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/bullet.png"/></Relationships>`), RELS_TYPE);
  const image = doc.getParagraphs()[0].numbering.image;
  assert.equal(image.relationshipId, 'rBullet');
  assert.equal(image.partPath, 'word/media/bullet.png');
  assert.equal(image.sourcePartPath, 'word/numbering.xml');
  assert.equal(image.isExternal, false);
  assert.equal(doc.getImageDataUrl(image), 'data:image/png;base64,AQID');
  const xml = doc.getPartXml('word/numbering.xml');
  const reloaded = await DocxDocument.load(await doc.toUint8Array());
  assert.equal(reloaded.getPartXml('word/numbering.xml'), xml);
  assert.match(reloaded.getPartXml('word/numbering.xml'), /w:numPicBulletId="5"/);
  assert.match(reloaded.getPartXml('word/numbering.xml'), /w:lvlPicBulletId w:val="5"/);
});

test('external picture bullets are placeholders and missing pictures fall back to marker text', () => {
  const doc = withBody('<w:p><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Item</w:t></w:r></w:p>');
  const numberingXml = `<w:numbering xmlns:w="${WORD_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:v="urn:schemas-microsoft-com:vml"><w:numPicBullet w:numPicBulletId="5"><w:pict><v:shape><v:imagedata r:id="rBullet"/></v:shape></w:pict></w:numPicBullet><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="▪"/><w:lvlPicBulletId w:val="5"/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num></w:numbering>`;
  attachNumbering(doc, numberingXml);
  doc.addPart('word/_rels/numbering.xml.rels', encoder.encode(`<Relationships xmlns="${REL_NS}"><Relationship Id="rBullet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="https://example.invalid/bullet.png" TargetMode="External"/></Relationships>`), RELS_TYPE);
  const image = doc.getParagraphs()[0].numbering.image;
  assert.equal(image.isExternal, true);
  assert.match(doc.getImageDataUrl(image), /^data:image\/svg\+xml;base64,/);

  const missing = withBody('<w:p><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Item</w:t></w:r></w:p>');
  attachNumbering(missing, numberingXml);
  assert.equal(missing.getParagraphs()[0].numbering.text, '▪');
  assert.equal(missing.getParagraphs()[0].numbering.image, undefined);
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

test('restartNumbering restarts this and following same-list items without changing earlier references', () => {
  const doc = numberedDocument([0, 0, 0]);
  const oldRevision = doc.revision;
  const newNumId = doc.restartNumbering(1);
  const paragraphs = doc.getParagraphs();
  assert.deepEqual(paragraphs.map((paragraph) => paragraph.numbering?.text), ['1.', '1.', '2.']);
  assert.deepEqual(paragraphs.map((paragraph) => paragraph.numbering?.numId), [1, newNumId, newNumId]);
  assert.equal(doc.revision, oldRevision + 1);
});

test('restartNumbering honors a positive custom start', () => {
  const doc = numberedDocument([0, 0, 0]);
  doc.restartNumbering(1, { start: 5 });
  assert.deepEqual(doc.getParagraphs().map((paragraph) => paragraph.numbering?.text), ['1.', '5.', '6.']);
});

test('restartNumbering of a nested item leaves parent numbering unchanged', () => {
  const doc = numberedDocument([0, 1, 1, 0], `<w:numbering xmlns:w="${WORD_NS}"><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl><w:lvl w:ilvl="1"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2."/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num></w:numbering>`);
  const before = doc.getParagraphs().map((paragraph) => paragraph.numbering?.text);
  doc.restartNumbering(2);
  const paragraphs = doc.getParagraphs();
  assert.deepEqual(paragraphs.map((paragraph) => paragraph.numbering?.text), ['1.', '1.1.', '1.1.', '2.']);
  assert.equal(paragraphs[0].numbering.numId, 1);
  assert.equal(paragraphs[3].numbering.numId, 1);
  assert.equal(paragraphs[0].numbering.text, before[0]);
  assert.equal(paragraphs[3].numbering.text, before[3]);
});

test('restartNumbering keeps deeper child items in the restarted segment', () => {
  const doc = numberedDocument([0, 1, 2, 2, 1], `<w:numbering xmlns:w="${WORD_NS}"><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl><w:lvl w:ilvl="1"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2."/></w:lvl><w:lvl w:ilvl="2"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2.%3."/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num></w:numbering>`);
  const newNumId = doc.restartNumbering(1);
  const paragraphs = doc.getParagraphs();
  assert.deepEqual(paragraphs.map((paragraph) => paragraph.numbering?.text), ['1.', '1.1.', '1.1.1.', '1.1.2.', '1.2.']);
  assert.deepEqual(paragraphs.map((paragraph) => paragraph.numbering?.numId), [1, newNumId, newNumId, newNumId, newNumId]);
});

test('restartNumbering stops before a different numbering instance', () => {
  const xml = `<w:numbering xmlns:w="${WORD_NS}"><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num></w:numbering>`;
  const doc = numberedDocument([0, 0], xml);
  doc.setPartXml(doc.mainDocumentPath, `<w:document xmlns:w="${WORD_NS}"><w:body><w:p><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr></w:p><w:p><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr></w:p><w:p><w:pPr><w:numPr><w:numId w:val="2"/></w:numPr></w:pPr></w:p><w:sectPr/></w:body></w:document>`);
  const newNumId = doc.restartNumbering(1);
  assert.deepEqual(doc.getParagraphs().map((paragraph) => paragraph.numbering?.numId), [1, newNumId, 2]);
  assert.deepEqual(doc.getParagraphs().map((paragraph) => paragraph.numbering?.text), ['1.', '1.', '1.']);
});

test('restartNumbering stops when a shallower list level begins', () => {
  const doc = numberedDocument([0, 1, 1, 0, 1], `<w:numbering xmlns:w="${WORD_NS}"><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl><w:lvl w:ilvl="1"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2."/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num></w:numbering>`);
  const newNumId = doc.restartNumbering(2);
  assert.deepEqual(doc.getParagraphs().map((paragraph) => paragraph.numbering?.numId), [1, 1, newNumId, 1, 1]);
});

test('a repeated restart creates a fresh numbering instance and replaces its start value', () => {
  const doc = numberedDocument([0, 0, 0]);
  const firstNumId = doc.restartNumbering(1, { start: 4 });
  const secondNumId = doc.restartNumbering(1, { start: 7 });
  assert.notEqual(secondNumId, firstNumId);
  assert.deepEqual(doc.getParagraphs().map((paragraph) => paragraph.numbering?.numId), [1, secondNumId, secondNumId]);
  assert.deepEqual(doc.getParagraphs().map((paragraph) => paragraph.numbering?.text), ['1.', '7.', '8.']);
});

test('continueNumbering restores the previous instance and its continuous sequence', () => {
  const doc = numberedDocument([0, 0, 0]);
  doc.restartNumbering(1, { start: 5 });
  doc.continueNumbering(1);
  assert.deepEqual(doc.getParagraphs().map((paragraph) => paragraph.numbering?.numId), [1, 1, 1]);
  assert.deepEqual(doc.getParagraphs().map((paragraph) => paragraph.numbering?.text), ['1.', '2.', '3.']);
});

test('continueNumbering without a preceding same-level item is a true no-op', () => {
  const doc = numberedDocument([0, 0]);
  const revision = doc.revision;
  doc.continueNumbering(0);
  assert.equal(doc.revision, revision);
});

test('continueNumbering without a restart does not mutate an already continuous list', () => {
  const doc = numberedDocument([0, 0, 0]);
  const revision = doc.revision;
  doc.continueNumbering(1);
  assert.equal(doc.revision, revision);
});

test('numbering mutations reject paragraphs outside a list and invalid indices', () => {
  const doc = withBody('<w:p><w:r><w:t>Plain</w:t></w:r></w:p>');
  assert.throws(() => doc.restartNumbering(0), /does not have numbering/);
  assert.throws(() => doc.continueNumbering(0), /does not have numbering/);
  assert.throws(() => doc.restartNumbering(1), /does not exist/);
  assert.throws(() => doc.continueNumbering(1), /does not exist/);
  assert.throws(() => doc.restartNumbering(-1), /non-negative safe integer/);
});

test('restartNumbering validates options.start before changing the document', () => {
  const doc = numberedDocument([0, 0]);
  const revision = doc.revision;
  for (const start of [0, -1, 1.5, Number.MAX_SAFE_INTEGER + 1, '5', null]) {
    assert.throws(() => doc.restartNumbering(1, { start }), /start|safe integer/i);
    assert.equal(doc.revision, revision);
  }
  assert.throws(() => doc.restartNumbering(1, { start: 2, extra: true }), /options/);
});

test('restartNumbering preserves resolved numbering part paths and repairs package links', () => {
  const doc = numberedDocument([0, 0]);
  const relationshipsPath = 'word/_rels/document.xml.rels';
  const relationships = doc.getPartXml(relationshipsPath).replace(/<Relationship[^>]*numbering[^>]*\/>/, '');
  doc.setPartXml(relationshipsPath, relationships);
  const contentTypes = doc.getPartXml('[Content_Types].xml').replace(/<Override[^>]*numbering\.xml[^>]*\/>/, '');
  doc.setPartXml('[Content_Types].xml', contentTypes);
  doc.restartNumbering(1);
  assert.match(doc.getPartXml(relationshipsPath), /relationships\/numbering/);
  assert.match(doc.getPartXml('[Content_Types].xml'), /word\/numbering\.xml/);
  assert.deepEqual(doc.getParagraphs().map((paragraph) => paragraph.numbering?.text), ['1.', '1.']);
});

test('restartNumbering writes CT_Num and CT_NumLvl children in schema order', () => {
  const xml = `<w:numbering xmlns:w="${WORD_NS}"><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="1"/><w:lvlOverride w:ilvl="0"><w:lvl w:ilvl="0"><w:start w:val="2"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:lvlOverride></w:num></w:numbering>`;
  const doc = numberedDocument([0, 0], xml);
  const numId = doc.restartNumbering(1);
  const numbering = parseXml(doc.getPartXml('word/numbering.xml')).documentElement;
  const num = [...numbering.getElementsByTagNameNS(WORD_NS, 'num')]
    .find((item) => Number(item.getAttributeNS(WORD_NS, 'numId')) === numId);
  assert.deepEqual(elementNames(num), ['abstractNumId', 'lvlOverride']);
  const override = [...num.getElementsByTagNameNS(WORD_NS, 'lvlOverride')].find((item) =>
    Number(item.getAttributeNS(WORD_NS, 'ilvl')) === 0);
  const names = elementNames(override);
  assert.deepEqual(names, ['startOverride', 'lvl']);
  assert.ok(names.indexOf('startOverride') < names.indexOf('lvl'));
});

test('restartNumbering and continueNumbering are supported by atomic Agent operations', () => {
  const doc = numberedDocument([0, 0, 0]);
  const operationTypes = AGENT_OPERATION_SCHEMA.properties.operations.items.oneOf
    .map((entry) => entry.properties.type.const);
  assert.ok(operationTypes.includes('restartNumbering'));
  assert.ok(operationTypes.includes('continueNumbering'));
  const revision = doc.revision;
  doc.applyOperations({ operations: [
    { type: 'restartNumbering', index: 1, options: { start: 4 } },
    { type: 'continueNumbering', index: 1 },
  ] });
  assert.equal(doc.revision, revision + 1);
  assert.deepEqual(doc.getParagraphs().map((paragraph) => paragraph.numbering?.text), ['1.', '2.', '3.']);
  assert.throws(() => doc.applyOperations({ operations: [
    { type: 'restartNumbering', index: 1, options: { start: 0 } },
  ] }), /positive integer/);
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

test('numbered paragraphs keep style-derived effective formatting', () => {
  const doc = withBody('<w:p><w:pPr><w:pStyle w:val="HeadingList"/></w:pPr><w:r><w:t>A</w:t></w:r></w:p>');
  attachNumbering(
    doc,
    `<w:numbering xmlns:w="${WORD_NS}"><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="5"><w:abstractNumId w:val="1"/></w:num></w:numbering>`,
    `<w:styles xmlns:w="${WORD_NS}"><w:style w:type="paragraph" w:styleId="HeadingList"><w:pPr><w:numPr><w:numId w:val="5"/></w:numPr><w:spacing w:after="240"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style></w:styles>`,
  );
  const paragraph = doc.getParagraphs()[0];
  assert.equal(paragraph.numbering.text, '1.');
  assert.equal(paragraph.effective.spacingAfter, 240);
  assert.equal(paragraph.runs[0].effective.bold, true);
  assert.equal(paragraph.runs[0].effective.fontSize, 16);
});

test('unknown numFmt falls back to decimal instead of throwing', () => {
  const doc = withBody('<w:p><w:pPr><w:numPr><w:numId w:val="5"/></w:numPr></w:pPr><w:r><w:t>A</w:t></w:r></w:p>');
  attachNumbering(doc, `<w:numbering xmlns:w="${WORD_NS}"><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:numFmt w:val="totallyUnknown"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="5"><w:abstractNumId w:val="1"/></w:num></w:numbering>`);
  assert.equal(doc.getParagraphs()[0].numbering.text, '1.');
});

test('deeper first item initializes missing ancestor counters', () => {
  const doc = withBody('<w:p><w:pPr><w:numPr><w:numId w:val="1"/><w:ilvl w:val="2"/></w:numPr></w:pPr><w:r><w:t>A</w:t></w:r></w:p>');
  attachNumbering(doc, `<w:numbering xmlns:w="${WORD_NS}"><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl><w:lvl w:ilvl="1"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2."/></w:lvl><w:lvl w:ilvl="2"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2.%3."/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num></w:numbering>`);
  assert.equal(doc.getParagraphs()[0].numbering.text, '1.1.1.');
});

test('seeded ancestor counters do not consume their start values', () => {
  const doc = withBody('<w:p><w:pPr><w:numPr><w:numId w:val="1"/><w:ilvl w:val="1"/></w:numPr></w:pPr><w:r><w:t>A</w:t></w:r></w:p><w:p><w:pPr><w:numPr><w:numId w:val="1"/><w:ilvl w:val="0"/></w:numPr></w:pPr><w:r><w:t>B</w:t></w:r></w:p><w:p><w:pPr><w:numPr><w:numId w:val="1"/><w:ilvl w:val="1"/></w:numPr></w:pPr><w:r><w:t>C</w:t></w:r></w:p>');
  attachNumbering(doc, `<w:numbering xmlns:w="${WORD_NS}"><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl><w:lvl w:ilvl="1"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2."/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num></w:numbering>`);
  assert.deepEqual(doc.getParagraphs().map(paragraph => paragraph.numbering?.text), ['1.1.', '1.', '1.2.']);
});

test('additional numbering formats render supported text', () => {
  const doc = withBody('<w:p><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>A</w:t></w:r></w:p><w:p><w:pPr><w:numPr><w:numId w:val="2"/></w:numPr></w:pPr><w:r><w:t>B</w:t></w:r></w:p><w:p><w:pPr><w:numPr><w:numId w:val="3"/></w:numPr></w:pPr><w:r><w:t>C</w:t></w:r></w:p>');
  attachNumbering(doc, `<w:numbering xmlns:w="${WORD_NS}"><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:numFmt w:val="decimalZero"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:abstractNum w:abstractNumId="2"><w:lvl w:ilvl="0"><w:numFmt w:val="ordinal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:abstractNum w:abstractNumId="3"><w:lvl w:ilvl="0"><w:numFmt w:val="decimalEnclosedCircle"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="2"/></w:num><w:num w:numId="3"><w:abstractNumId w:val="3"/></w:num></w:numbering>`);
  assert.deepEqual(doc.getParagraphs().map(paragraph => paragraph.numbering?.text), ['01.', '1st.', '①.']);
});

test('createNumbering creates package parts and survives export roundtrip', async () => {
  const doc = DocxDocument.create();
  doc.setParagraphText(0, 'Item');
  const numId = doc.createNumbering('bullet');
  doc.setParagraphNumbering(0, numId, 1);
  const xml = doc.getPartXml('word/numbering.xml');
  assert.match(xml, /<w:abstractNum[\s\S]*<w:num /);
  assert.ok(doc.listParts().includes('word/numbering.xml'));
  assert.ok(doc.listParts().includes('word/_rels/document.xml.rels'));
  const reopened = await DocxDocument.load(await doc.toUint8Array());
  assert.equal(reopened.getParagraphs()[0].numbering.isBullet, true);
  assert.equal(reopened.getParagraphs()[0].numbering.level, 1);
  assert.equal(reopened.getParagraphs()[0].text, 'Item');
});

test('createNumbering works when the main document is stored at a custom package path', async () => {
  const base = await DocxDocument.create().toUint8Array();
  const JSZip = (await import('jszip')).default;
  const zip = await JSZip.loadAsync(base);
  const main = await zip.file('word/document.xml').async('uint8array');
  zip.file('custom/main.xml', main);
  zip.remove('word/document.xml');
  zip.file('_rels/.rels', (await zip.file('_rels/.rels').async('string')).replace('word/document.xml', 'custom/main.xml'));
  zip.file('[Content_Types].xml', (await zip.file('[Content_Types].xml').async('string')).replace('/word/document.xml', '/custom/main.xml'));
  const doc = await DocxDocument.load(await zip.generateAsync({ type: 'uint8array' }));
  const numId = doc.createNumbering('decimal');
  doc.setParagraphNumbering(0, numId);
  const reopened = await DocxDocument.load(await doc.toUint8Array());
  assert.equal(reopened.mainDocumentPath, 'custom/main.xml');
  assert.equal(reopened.getParagraphs()[0].numbering.text, '1.');
  assert.ok(reopened.listParts().includes('custom/numbering.xml'));
  assert.match(reopened.getPartXml('custom/_rels/main.xml.rels'), /Target="numbering.xml"/);
});

test('style-based numbering resolves through custom main-document relationships', async () => {
  const source = withBody('<w:p><w:pPr><w:pStyle w:val="ListStyle"/></w:pPr><w:r><w:t>A</w:t></w:r></w:p>');
  attachNumbering(source, `<w:numbering xmlns:w="${WORD_NS}"><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="5"><w:abstractNumId w:val="1"/></w:num></w:numbering>`, `<w:styles xmlns:w="${WORD_NS}"><w:style w:type="paragraph" w:styleId="ListStyle"><w:pPr><w:numPr><w:numId w:val="5"/></w:numPr></w:pPr></w:style></w:styles>`);
  const JSZip = (await import('jszip')).default;
  const zip = await JSZip.loadAsync(await source.toUint8Array());
  zip.file('custom/main.xml', await zip.file('word/document.xml').async('uint8array'));
  zip.remove('word/document.xml');
  zip.file('customAssets/styles.xml', await zip.file('word/styles.xml').async('uint8array'));
  zip.remove('word/styles.xml');
  zip.remove('word/_rels/document.xml.rels');
  zip.file('custom/_rels/main.xml.rels', `<Relationships xmlns="${REL_NS}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="../word/numbering.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="../customAssets/styles.xml"/></Relationships>`);
  zip.file('_rels/.rels', (await zip.file('_rels/.rels').async('string')).replace('word/document.xml', 'custom/main.xml'));
  zip.file('[Content_Types].xml', (await zip.file('[Content_Types].xml').async('string'))
    .replace('/word/document.xml', '/custom/main.xml')
    .replace('/word/styles.xml', '/customAssets/styles.xml'));
  const doc = await DocxDocument.load(await zip.generateAsync({ type: 'uint8array' }));
  assert.equal(doc.mainDocumentPath, 'custom/main.xml');
  assert.equal(doc.getParagraphs()[0].numbering.text, '1.');
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

test('clearParagraphNumbering suppresses style-derived numbering and invalid numId is rejected', () => {
  const doc = withBody('<w:p><w:pPr><w:pStyle w:val="ListStyle"/></w:pPr><w:r><w:t>A</w:t></w:r></w:p>');
  attachNumbering(doc, `<w:numbering xmlns:w="${WORD_NS}"><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="5"><w:abstractNumId w:val="1"/></w:num></w:numbering>`, `<w:styles xmlns:w="${WORD_NS}"><w:style w:type="paragraph" w:styleId="ListStyle"><w:pPr><w:numPr><w:numId w:val="5"/></w:numPr></w:pPr></w:style></w:styles>`);
  assert.throws(() => doc.setParagraphNumbering(0, 0), /numId/);
  doc.clearParagraphNumbering(0);
  assert.equal(doc.getParagraphs()[0].numbering, undefined);
});

test('numStyleLink resolves numbering styles and styleLink does not override local levels', () => {
  const doc = withBody('<w:p><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>A</w:t></w:r></w:p><w:p><w:pPr><w:numPr><w:numId w:val="2"/></w:numPr></w:pPr><w:r><w:t>B</w:t></w:r></w:p>');
  attachNumbering(doc, `<w:numbering xmlns:w="${WORD_NS}"><w:abstractNum w:abstractNumId="1"><w:numStyleLink w:val="MyList"/></w:abstractNum><w:abstractNum w:abstractNumId="2"><w:styleLink w:val="MyList"/><w:lvl w:ilvl="0"><w:numFmt w:val="upperLetter"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:abstractNum w:abstractNumId="3"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num><w:num w:numId="10"><w:abstractNumId w:val="3"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="2"/></w:num></w:numbering>`, `<w:styles xmlns:w="${WORD_NS}"><w:style w:type="numbering" w:styleId="MyList"><w:pPr><w:numPr><w:numId w:val="10"/></w:numPr></w:pPr></w:style></w:styles>`);
  assert.equal(doc.getParagraphs()[0].numbering.text, '1.');
  assert.equal(doc.getParagraphs()[1].numbering.text, 'A.');
});

test('findReusableNumberingId reuses adjacent or remembered list definitions', () => {
  const doc = DocxDocument.create();
  doc.insertParagraph('B');
  doc.insertParagraph('C');
  const bullet = doc.createNumbering('bullet');
  doc.setParagraphNumbering(0, bullet);
  assert.equal(findReusableNumberingId(doc.getParagraphs(), doc.getNumberingDefinitions(), 'bullet', 1), bullet);
  assert.equal(findReusableNumberingId(doc.getParagraphs(), doc.getNumberingDefinitions(), 'bullet', 2, bullet), bullet);
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
  const operationTypes = AGENT_OPERATION_SCHEMA.properties.operations.items.oneOf
    .map((entry) => entry.properties.type.const);
  assert.ok(operationTypes.includes('setParagraphNumbering'));
  assert.ok(operationTypes.includes('setParagraphLevel'));
  assert.ok(operationTypes.includes('clearParagraphNumbering'));
  assert.throws(() => doc.applyOperations({ operations: [{ type: 'setParagraphNumbering', index: 0, numId: 0 }] }), /numId/);
});

test('cyclic numStyleLink does not throw', () => {
  const doc = withBody('<w:p><w:pPr><w:numPr><w:numId w:val="2"/></w:numPr></w:pPr><w:r><w:t>A</w:t></w:r></w:p>');
  attachNumbering(doc, `<w:numbering xmlns:w="${WORD_NS}"><w:abstractNum w:abstractNumId="1"><w:numStyleLink w:val="LoopStyle"/></w:abstractNum><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num></w:numbering>`, `<w:styles xmlns:w="${WORD_NS}"><w:style w:type="paragraph" w:styleId="LoopStyle"><w:pPr><w:numPr><w:numId w:val="2"/></w:numPr></w:pPr></w:style></w:styles>`);
  assert.equal(doc.getParagraphs()[0].numbering, undefined);
});
