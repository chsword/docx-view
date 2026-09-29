import { test } from 'node:test';
import assert from 'node:assert/strict';
import { DocxDocument } from '../dist/document.js';
import { AGENT_OPERATION_SCHEMA } from '../dist/operations.js';
import { OFFICE_REL_NS, WORD_NS } from '../dist/xml.js';

const W14_NS = 'http://schemas.microsoft.com/office/word/2010/wordml';

function withBody(xml) {
  const doc = DocxDocument.create();
  doc.setPartXml(doc.mainDocumentPath, `<w:document xmlns:w="${WORD_NS}" xmlns:r="${OFFICE_REL_NS}"><w:body>${xml}<w:sectPr/></w:body></w:document>`);
  return doc;
}

test('reads metadata, locks, binding, placeholders, and list items for controls', () => {
  const doc = withBody(`<w:sdt><w:sdtPr><w:alias w:val="City"/><w:tag w:val="city-tag"/><w:id w:val="5"/><w:lock w:val="sdtLocked"/><w:placeholder><w:docPart w:val="DefaultPlaceholder"/></w:placeholder><w:showingPlcHdr/><w:dataBinding w:prefixMappings="xmlns:ns='urn:test'" w:xpath="/ns:root/ns:city" w:storeItemID="{abc}"/><w:dropDownList><w:listItem w:displayText="Beijing" w:value="BJ"/><w:listItem w:displayText="Shanghai" w:value="SH"/></w:dropDownList></w:sdtPr><w:sdtContent><w:p><w:r><w:t>Choose city</w:t></w:r></w:p></w:sdtContent></w:sdt>`);
  assert.deepEqual(doc.getContentControls()[0], {
    id: 5, kind: 'dropDownList', alias: 'City', tag: 'city-tag', lock: 'sdtLocked',
    showingPlaceholder: true, placeholderDocPart: 'DefaultPlaceholder',
    items: [{ displayText: 'Beijing', value: 'BJ' }, { displayText: 'Shanghai', value: 'SH' }],
    dataBinding: { prefixMappings: "xmlns:ns='urn:test'", xpath: '/ns:root/ns:city', storeItemId: '{abc}' },
    paragraphs: [0], nested: false, text: 'Choose city',
  });
});

test('reads block and inline controls with body paragraph indexes', () => {
  const doc = withBody(`<w:sdt><w:sdtPr><w:id w:val="1"/><w:text/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>Block</w:t></w:r></w:p></w:sdtContent></w:sdt><w:p><w:r><w:t>pre</w:t></w:r><w:sdt><w:sdtPr><w:id w:val="2"/><w:text/></w:sdtPr><w:sdtContent><w:r><w:t>inline</w:t></w:r></w:sdtContent></w:sdt><w:r><w:t>post</w:t></w:r></w:p>`);
  assert.deepEqual(doc.getContentControls().map(({ paragraphs, text }) => ({ paragraphs, text })), [
    { paragraphs: [0], text: 'Block' }, { paragraphs: [1], text: 'inline' },
  ]);
  assert.deepEqual(doc.getParagraphs().map((paragraph) => paragraph.text), ['Block', 'preinlinepost']);
});

test('returns nested controls independently and marks the inner control', () => {
  const doc = withBody(`<w:sdt><w:sdtPr><w:id w:val="11"/><w:richText/></w:sdtPr><w:sdtContent><w:p><w:sdt><w:sdtPr><w:id w:val="12"/><w:text/></w:sdtPr><w:sdtContent><w:r><w:t>Inner</w:t></w:r></w:sdtContent></w:sdt></w:p></w:sdtContent></w:sdt>`);
  assert.deepEqual(doc.getContentControls().map(({ id, nested, paragraphs }) => ({ id, nested, paragraphs })), [
    { id: 11, nested: false, paragraphs: [0] }, { id: 12, nested: true, paragraphs: [0] },
  ]);
});

test('unknown and incomplete controls degrade without throwing', () => {
  const doc = withBody(`<w:sdt><w:sdtPr><w:alias w:val="unknown"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>Safe</w:t></w:r></w:p></w:sdtContent></w:sdt><w:sdt><w:sdtPr/><w:sdtContent/></w:sdt>`);
  assert.deepEqual(doc.getContentControls().map(({ id, kind, paragraphs, text }) => ({ id, kind, paragraphs, text })), [
    { id: undefined, kind: 'unknown', paragraphs: [0], text: 'Safe' },
    { id: undefined, kind: 'unknown', paragraphs: [], text: '' },
  ]);
});

test('reads date format and checkbox state from their respective namespaces', () => {
  const doc = withBody(`<w:sdt><w:sdtPr><w:id w:val="20"/><w:date><w:dateFormat w:val="yyyy-MM-dd"/></w:date></w:sdtPr><w:sdtContent><w:p><w:r><w:t>2026-09-29</w:t></w:r></w:p></w:sdtContent></w:sdt><w:sdt><w:sdtPr><w:id w:val="21"/><w14:checkbox xmlns:w14="${W14_NS}"><w14:checked w14:val="1"/></w14:checkbox></w:sdtPr><w:sdtContent><w:p><w:r><w:t>☒</w:t></w:r></w:p></w:sdtContent></w:sdt>`);
  const [date, checkbox] = doc.getContentControls();
  assert.equal(date.kind, 'date');
  assert.equal(date.dateFormat, 'yyyy-MM-dd');
  assert.equal(checkbox.kind, 'checkbox');
  assert.equal(checkbox.checked, true);
});

test('date and checkbox values can be absent', () => {
  const doc = withBody(`<w:sdt><w:sdtPr><w:id w:val="30"/><w14:checkbox xmlns:w14="${W14_NS}"/></w:sdtPr><w:sdtContent/></w:sdt>`);
  assert.equal(doc.getContentControls()[0].checked, undefined);
});

test('setContentControlText clears showing placeholder and keeps sdt properties', () => {
  const doc = withBody(`<w:sdt><w:sdtPr><w:id w:val="31"/><w:showingPlcHdr/><w:placeholder><w:docPart w:val="Default"/></w:placeholder><w:text/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>Placeholder</w:t></w:r></w:p></w:sdtContent></w:sdt>`);
  doc.setContentControlText(31, 'Entered');
  assert.equal(doc.getContentControls()[0].text, 'Entered');
  assert.equal(doc.getContentControls()[0].showingPlaceholder, false);
  assert.match(doc.getPartXml(doc.mainDocumentPath), /<w:placeholder>/);
  assert.match(doc.getPartXml(doc.mainDocumentPath), /<w:docPart w:val="Default"\/>/);
});

test('setContentControlText writes inline control content without changing adjacent text', () => {
  const doc = withBody(`<w:p><w:r><w:t>before</w:t></w:r><w:sdt><w:sdtPr><w:id w:val="32"/><w:text/></w:sdtPr><w:sdtContent><w:r><w:t>old</w:t></w:r></w:sdtContent></w:sdt><w:r><w:t>after</w:t></w:r></w:p>`);
  doc.setContentControlText(32, 'new');
  assert.equal(doc.getParagraphs()[0].text, 'beforenewafter');
});

test('setContentControlText preserves paragraph properties and unknown surrounding XML', () => {
  const doc = withBody(`<w:sdt><w:sdtPr><w:id w:val="45"/><w:text/></w:sdtPr><w:sdtContent><w:p><w:pPr><w:keepNext/></w:pPr><w:bookmarkStart w:id="1" w:name="before"/><w:r><w:rPr><w:b/></w:rPr><w:t>old</w:t></w:r><w:bookmarkEnd w:id="1"/><w:customXml w:uri="urn:custom"><w:customXmlPr/></w:customXml></w:p></w:sdtContent></w:sdt>`);
  doc.setContentControlText(45, 'new');
  const xml = doc.getPartXml(doc.mainDocumentPath);
  assert.match(xml, /<w:pPr><w:keepNext\/><\/w:pPr>/);
  assert.match(xml, /<w:bookmarkStart w:id="1" w:name="before"\/>/);
  assert.match(xml, /<w:bookmarkEnd w:id="1"\/>/);
  assert.match(xml, /<w:customXml w:uri="urn:custom"><w:customXmlPr\/><\/w:customXml>/);
  assert.match(xml, /<w:rPr><w:b\/><\/w:rPr><w:t(?: xml:space="preserve")?>new<\/w:t>/);
});

test('setContentControlText rejects content-locked controls while paragraph editing remains available', () => {
  const doc = withBody(`<w:sdt><w:sdtPr><w:id w:val="33"/><w:lock w:val="contentLocked"/><w:text/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>old</w:t></w:r></w:p></w:sdtContent></w:sdt>`);
  assert.throws(() => doc.setContentControlText(33, 'new'), /locked/);
  doc.setParagraphText(0, 'paragraph edit');
  assert.equal(doc.getParagraphs()[0].text, 'paragraph edit');
});

test('setContentControlText permits structural-only locks', () => {
  const doc = withBody(`<w:sdt><w:sdtPr><w:id w:val="34"/><w:lock w:val="sdtLocked"/><w:text/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>old</w:t></w:r></w:p></w:sdtContent></w:sdt>`);
  doc.setContentControlText(34, 'new');
  assert.equal(doc.getContentControls()[0].text, 'new');
});

test('checkbox setter updates w14 state and can be repeated without revision changes', () => {
  const doc = withBody(`<w:sdt><w:sdtPr><w:id w:val="35"/><w14:checkbox xmlns:w14="${W14_NS}"><w14:checked w14:val="0"/></w14:checkbox></w:sdtPr><w:sdtContent><w:p/></w:sdtContent></w:sdt>`);
  doc.setContentControlChecked(35, true);
  assert.equal(doc.getContentControls()[0].checked, true);
  const revision = doc.revision;
  doc.setContentControlChecked(35, true);
  assert.equal(doc.revision, revision);
});

test('properties patch preserves unmodified fields and unknown sdtPr elements', () => {
  const doc = withBody(`<w:sdt><w:sdtPr><w:alias w:val="old"/><w:id w:val="36"/><w:tag w:val="keep"/><w:customXml w:uri="urn:custom"><w:customXmlPr/></w:customXml><w:text/></w:sdtPr><w:sdtContent><w:p/></w:sdtContent></w:sdt>`);
  doc.setContentControlProperties(36, { alias: 'new' });
  const xml = doc.getPartXml(doc.mainDocumentPath);
  assert.match(xml, /<w:alias w:val="new"\/>/);
  assert.match(xml, /<w:tag w:val="keep"\/>/);
  assert.match(xml, /<w:customXml w:uri="urn:custom"><w:customXmlPr\/><\/w:customXml>/);
});

test('properties setter can clear alias, update tag, and lock can round-trip unlocked', () => {
  const doc = withBody(`<w:sdt><w:sdtPr><w:alias w:val="title"/><w:tag w:val="old"/><w:id w:val="37"/><w:lock w:val="sdtLocked"/></w:sdtPr><w:sdtContent><w:p/></w:sdtContent></w:sdt>`);
  doc.setContentControlProperties(37, { alias: null, tag: 'new', lock: 'unlocked' });
  assert.equal(doc.getContentControls()[0].alias, undefined);
  assert.equal(doc.getContentControls()[0].tag, 'new');
  assert.equal(doc.getContentControls()[0].lock, 'unlocked');
});

test('properties with no effective changes do not increment revision', () => {
  const doc = withBody(`<w:sdt><w:sdtPr><w:alias w:val="same"/><w:tag w:val="tag"/><w:id w:val="38"/></w:sdtPr><w:sdtContent><w:p/></w:sdtContent></w:sdt>`);
  const control = doc.getContentControls()[0];
  const revision = doc.revision;
  doc.setContentControlProperties(control.id, { alias: control.alias, tag: control.tag, lock: control.lock });
  assert.equal(doc.revision, revision);
});

test('text no-op preserves revision and checked setter rejects non-checkbox controls', () => {
  const doc = withBody(`<w:sdt><w:sdtPr><w:id w:val="44"/><w:text/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>same</w:t></w:r></w:p></w:sdtContent></w:sdt>`);
  const revision = doc.revision;
  doc.setContentControlText(44, doc.getContentControls()[0].text);
  assert.equal(doc.revision, revision);
  assert.throws(() => doc.setContentControlChecked(44, true), /not a checkbox/);
});

test('removeContentControl keeps content by default and preserves paragraph order', () => {
  const doc = withBody(`<w:p><w:r><w:t>before</w:t></w:r></w:p><w:sdt><w:sdtPr><w:id w:val="39"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>middle</w:t></w:r></w:p></w:sdtContent></w:sdt><w:p><w:r><w:t>after</w:t></w:r></w:p>`);
  doc.removeContentControl(39);
  assert.deepEqual(doc.getParagraphs().map((paragraph) => paragraph.text), ['before', 'middle', 'after']);
  assert.equal(doc.getContentControls().length, 0);
});

test('removeContentControl with keepContent false keeps body structurally nonempty', () => {
  const doc = withBody(`<w:sdt><w:sdtPr><w:id w:val="40"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>remove</w:t></w:r></w:p></w:sdtContent></w:sdt>`);
  doc.removeContentControl(40, { keepContent: false });
  assert.equal(doc.getParagraphs().length, 1);
  assert.equal(doc.getParagraphs()[0].text, '');
});

test('removeContentControl keeps a table cell paragraph when removing the only content', () => {
  const doc = withBody(`<w:tbl><w:tr><w:tc><w:tcPr/><w:sdt><w:sdtPr><w:id w:val="41"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>remove</w:t></w:r></w:p></w:sdtContent></w:sdt></w:tc></w:tr></w:tbl>`);
  doc.removeContentControl(41, { keepContent: false });
  const xml = doc.getPartXml(doc.mainDocumentPath);
  assert.match(xml, /<w:tc><w:tcPr\/><w:p\/><\/w:tc>/);
  assert.deepEqual(doc.getParagraphs().map((paragraph) => paragraph.text), ['']);
});

test('content control methods validate ids and text', () => {
  const doc = withBody('<w:p/>');
  assert.throws(() => doc.setContentControlText(-1, 'x'), /Index/);
  assert.throws(() => doc.setContentControlText(1, 'a\u0000b'), /valid XML text/);
  assert.throws(() => doc.setContentControlChecked(1, 1), /checked/);
});

test('Agent operations support all content control mutations atomically', () => {
  const doc = withBody(`<w:sdt><w:sdtPr><w:id w:val="42"/><w:text/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>old</w:t></w:r></w:p></w:sdtContent></w:sdt>`);
  const result = doc.applyOperations({ operations: [
    { type: 'setContentControlText', id: 42, text: 'updated' },
    { type: 'setContentControlProperties', id: 42, patch: { alias: 'Title' } },
  ] });
  assert.equal(result.paragraphs[0].text, 'updated');
  assert.equal(result.revision, doc.revision);
  assert.equal(doc.getContentControls()[0].alias, 'Title');
  assert.equal(AGENT_OPERATION_SCHEMA.properties.operations.items.oneOf.length, 66);
  const revision = doc.revision;
  assert.throws(() => doc.applyOperations({ operations: [
    { type: 'setContentControlText', id: 42, text: 'partial' },
    { type: 'setContentControlText', id: 99, text: 'missing' },
  ] }), /does not exist/);
  assert.equal(doc.revision, revision);
  assert.equal(doc.getContentControls()[0].text, 'updated');
});

test('Agent schema accepts the declared content control operation shapes', () => {
  const doc = withBody(`<w:sdt><w:sdtPr><w:id w:val="43"/><w14:checkbox xmlns:w14="${W14_NS}"/></w:sdtPr><w:sdtContent><w:p/></w:sdtContent></w:sdt>`);
  const result = doc.applyOperations({ operations: [
    { type: 'setContentControlChecked', id: 43, checked: true },
    { type: 'setContentControlProperties', id: 43, patch: { tag: null, lock: 'contentLocked' } },
  ] });
  assert.equal(result.paragraphs.length, 1);
  assert.equal(doc.getContentControls()[0].checked, true);
  assert.equal(doc.getContentControls()[0].lock, 'contentLocked');
});
