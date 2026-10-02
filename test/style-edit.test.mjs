import { test } from 'node:test';
import assert from 'node:assert/strict';
import { DocxDocument } from '../dist/document.js';
import { AGENT_OPERATION_SCHEMA, STYLE_PATCH_KEYS, validateStylePatch } from '../dist/operations.js';
import { WORD_NS } from '../dist/xml.js';

const STYLES_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml';
const encoder = new TextEncoder();

function withStyles(bodyXml, stylesXml) {
  const doc = DocxDocument.create();
  doc.setPartXml(doc.mainDocumentPath, `<w:document xmlns:w="${WORD_NS}"><w:body>${bodyXml}<w:sectPr/></w:body></w:document>`);
  doc.addPart('word/styles.xml', encoder.encode(`<w:styles xmlns:w="${WORD_NS}">${stylesXml}</w:styles>`), STYLES_TYPE);
  return doc;
}

const STYLES = `
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/>
    <w:rPr><w:sz w:val="22"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Base"><w:name w:val="Base"/><w:basedOn w:val="Normal"/>
    <w:rPr><w:color w:val="112233"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Heading"><w:name w:val="Heading"/><w:basedOn w:val="Base"/>
    <w:next w:val="Base"/><w:link w:val="HeadingChar"/><w:uiPriority w:val="9"/><w:qFormat/>
    <w:pPr><w:keepNext/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="4"/></w:numPr><w:spacing w:before="240" w:after="60"/><w:jc w:val="center"/></w:pPr>
    <w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi"/><w:b/><w:sz w:val="32"/></w:rPr></w:style>
  <w:style w:type="character" w:styleId="HeadingChar"><w:name w:val="Heading Char"/><w:link w:val="Heading"/>
    <w:rPr><w:b/></w:rPr></w:style>
  <w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/></w:style>`;

const BODY = `
  <w:p><w:pPr><w:pStyle w:val="Heading"/></w:pPr><w:r><w:t>Title</w:t></w:r></w:p>
  <w:p><w:pPr><w:pStyle w:val="Base"/></w:pPr><w:r><w:rPr><w:rStyle w:val="HeadingChar"/></w:rPr><w:t>body</w:t></w:r></w:p>`;

function styleXml(doc, id) {
  const xml = new TextDecoder().decode(doc.getPartBytes('word/styles.xml'));
  return xml.match(new RegExp(`<w:style [^>]*w:styleId="${id}"[^>]*>[\\s\\S]*?</w:style>`))?.[0];
}

test('defineStyle is undoable and does not undo the edit before it', () => {
  const doc = DocxDocument.create();
  doc.setParagraphText(0, 'kept');
  doc.defineStyle({ id: 'Fresh', name: 'Fresh', type: 'paragraph', run: { bold: true } });
  assert.ok(doc.getStyle('Fresh'));
  doc.undo();
  // 原先 defineStyle 不进撤销历史：这一步撤掉的是 setParagraphText，样式还在。
  assert.equal(doc.getStyle('Fresh'), undefined);
  assert.equal(doc.getParagraphs()[0].text, 'kept');
  doc.redo();
  assert.equal(doc.getStyle('Fresh')?.run?.bold, true);
});

test('updateStyle patches fields and keeps what the read model does not know', () => {
  const doc = withStyles(BODY, STYLES);
  const before = doc.getParagraphs()[0].effective;
  assert.equal(before.alignment, 'center');

  doc.updateStyle('Heading', {
    name: 'Big Heading',
    paragraph: { alignment: 'left', spacingAfter: 120 },
    run: { color: 'FF0000', bold: null },
  });
  const style = doc.getStyle('Heading');
  assert.equal(style.name, 'Big Heading');
  assert.equal(style.paragraph.alignment, 'left');
  assert.equal(style.paragraph.spacingBefore, 240, '补丁里没提的字段不动');
  assert.equal(style.paragraph.spacingAfter, 120);
  assert.equal(style.paragraph.keepNext, true);
  assert.equal(style.run.color, 'FF0000');
  assert.equal(style.run.bold, undefined, 'null 从样式里删掉该属性');
  assert.equal(style.basedOn, 'Base');
  assert.equal(style.next, 'Base');
  assert.equal(style.uiPriority, 9);

  // defineStyle 会整个重写 pPr / rPr，这两样就没了；补丁不该碰它们。
  const xml = styleXml(doc, 'Heading');
  assert.match(xml, /<w:numPr><w:ilvl w:val="0"\/><w:numId w:val="4"\/><\/w:numPr>/);
  assert.match(xml, /<w:rFonts w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi"\/>/);

  // 读者立刻看到新值（样式缓存按 revision 失效）。
  const after = doc.getParagraphs()[0];
  assert.equal(after.effective.alignment, 'left');
  assert.equal(after.runs[0].effective.color, 'FF0000');
});

test('updateStyle writes CT_Style children in schema order', () => {
  const doc = withStyles(BODY, STYLES);
  // HeadingChar 原本没有 basedOn / uiPriority / qFormat；新加的元素要插到正确位置，不能追加在末尾。
  doc.updateStyle('HeadingChar', { basedOn: 'DefaultParagraphFont', uiPriority: 3, quickFormat: true, aliases: ['HC', 'Hc2'] });
  const names = [...styleXml(doc, 'HeadingChar').matchAll(/<w:(\w+)[ />]/g)].map((match) => match[1]).slice(1);
  assert.deepEqual(names, ['name', 'aliases', 'basedOn', 'link', 'uiPriority', 'qFormat', 'rPr', 'b']);
  assert.deepEqual(doc.getStyle('HeadingChar').aliases, ['HC', 'Hc2']);
});

test('updateStyle twice keeps stacking on the first patch, and clears with null', () => {
  const doc = withStyles(BODY, STYLES);
  doc.updateStyle('Base', { run: { italic: true } });
  doc.updateStyle('Base', { run: { fontSize: 14 } });
  assert.deepEqual(
    { color: doc.getStyle('Base').run.color, italic: doc.getStyle('Base').run.italic, fontSize: doc.getStyle('Base').run.fontSize },
    { color: '112233', italic: true, fontSize: 14 },
  );
  doc.updateStyle('Heading', { next: null, link: null, uiPriority: null, quickFormat: false, basedOn: null });
  const style = doc.getStyle('Heading');
  assert.equal(style.next, undefined);
  assert.equal(style.link, undefined);
  assert.equal(style.uiPriority, undefined);
  assert.equal(style.quickFormat, undefined);
  assert.equal(style.basedOn, undefined);
  // 不再继承 Base 之后，Base 的颜色不该再出现在 Heading 段落上。
  assert.equal(doc.getParagraphs()[0].runs[0].effective.color, undefined);
});

test('updateStyle is one undo step, and a no-op patch does not bump the revision', () => {
  const doc = withStyles(BODY, STYLES);
  const revision = doc.revision;
  const undoDepth = doc.getHistory().undo.length;
  doc.updateStyle('Base', { run: { color: '112233' } });
  assert.equal(doc.revision, revision);
  assert.equal(doc.getHistory().undo.length, undoDepth);
  doc.updateStyle('Base', { run: { color: 'ABCDEF', bold: true }, paragraph: { alignment: 'right' } });
  assert.equal(doc.revision, revision + 1);
  doc.undo();
  assert.equal(doc.getStyle('Base').run.color, '112233');
  assert.equal(doc.getStyle('Base').run.bold, undefined);
});

test('updateStyle rejects patches Word could not open, and leaves the document untouched', () => {
  const doc = withStyles(BODY, STYLES);
  const revision = doc.revision;
  for (const [id, patch, pattern] of [
    ['Missing', { name: 'x' }, /not found/],
    ['Normal', { basedOn: 'Heading' }, /inherit from itself/],
    ['Base', { basedOn: 'Base' }, /inherit from itself/],
    ['Base', { basedOn: 'HeadingChar' }, /must be a paragraph style/],
    ['Base', { basedOn: 'Nope' }, /not found/],
    ['Base', { name: 'Heading' }, /already named/],
    ['Base', { name: '  ' }, /must not be empty/],
    ['Base', { next: 'HeadingChar' }, /existing paragraph style/],
    ['HeadingChar', { next: 'Base' }, /only applies to paragraph/],
    ['Base', { link: 'Heading' }, /existing character style/],
    ['HeadingChar', { paragraph: { alignment: 'left' } }, /no paragraph properties/],
    ['Base', { paragraph: { style: 'Heading' } }, /basedOn/],
    ['Base', { run: { color: 'red' } }, /hexadecimal/],
    ['Base', { aliases: ['a,b'] }, /comma/],
    ['Base', { colour: 'x' }, /colour/],
  ]) {
    assert.throws(() => doc.updateStyle(id, patch), pattern, `${id} ${JSON.stringify(patch)}`);
  }
  assert.equal(doc.revision, revision);
});

test('deleteStyle rewires inheritance and strips references from every story part', () => {
  const doc = withStyles(BODY, STYLES);
  doc.setHeaderText(0, 'head');
  const headerPath = doc.listParts().find((path) => /header\d*\.xml$/.test(path));
  doc.updatePartXml(headerPath, (header) => {
    const paragraph = header.getElementsByTagNameNS(WORD_NS, 'p')[0];
    const pPr = header.createElementNS(WORD_NS, 'w:pPr');
    const pStyle = header.createElementNS(WORD_NS, 'w:pStyle');
    pStyle.setAttributeNS(WORD_NS, 'w:val', 'Base');
    pPr.appendChild(pStyle);
    paragraph.insertBefore(pPr, paragraph.firstChild);
  });
  const revision = doc.revision;

  doc.deleteStyle('Base');
  assert.equal(doc.revision, revision + 1, '跨多个部件也只提交一次');
  assert.equal(doc.getStyle('Base'), undefined);
  // Heading 原来继承 Base，Base 继承 Normal：删掉中间一层后 Heading 接到 Normal 上。
  assert.equal(doc.getStyle('Heading').basedOn, 'Normal');
  assert.equal(doc.getStyle('Heading').next, undefined, 'next 指向被删样式的那条删掉');
  const paragraphs = doc.getParagraphs();
  assert.equal(paragraphs[1].style, undefined, '正文里的 pStyle 删掉，回落到默认样式');
  assert.equal(paragraphs[1].runs[0].style, 'HeadingChar', '别的样式引用不受影响');
  const header = new TextDecoder().decode(doc.getPartBytes(headerPath));
  assert.doesNotMatch(header, /pStyle/);
  assert.doesNotMatch(header, /<w:pPr\/>/, '删空的 pPr 一并删掉');

  doc.undo();
  assert.equal(doc.getStyle('Base')?.name, 'Base');
  assert.equal(doc.getParagraphs()[1].style, 'Base');
  assert.match(new TextDecoder().decode(doc.getPartBytes(headerPath)), /<w:pStyle w:val="Base"\/>/);
});

test('deleteStyle removes the other half of a linked pair and character references', () => {
  const doc = withStyles(BODY, STYLES);
  doc.deleteStyle('HeadingChar');
  assert.equal(doc.getStyle('Heading').link, undefined);
  assert.equal(doc.getParagraphs()[1].runs[0].style, undefined);
  assert.doesNotMatch(new TextDecoder().decode(doc.getPartBytes('word/document.xml')), /<w:rPr\/>/);
});

test('deleteStyle refuses default styles and unknown ids', () => {
  const doc = withStyles(BODY, STYLES);
  assert.throws(() => doc.deleteStyle('Normal'), /default paragraph style/);
  assert.throws(() => doc.deleteStyle('Nope'), /not found/);
});

test('style operations run in agent batches and roll back together', () => {
  const doc = withStyles(BODY, STYLES);
  doc.applyOperations({ operations: [
    { type: 'defineStyle', style: { id: 'Note', name: 'Note', type: 'paragraph', basedOn: 'Normal', run: { italic: true } } },
    { type: 'updateStyle', id: 'Note', patch: { run: { color: '00FF00' } } },
    { type: 'applyParagraphStyle', index: 1, styleId: 'Note' },
    { type: 'deleteStyle', id: 'HeadingChar' },
  ] });
  assert.deepEqual([doc.getStyle('Note').run.italic, doc.getStyle('Note').run.color], [true, '00FF00']);
  assert.equal(doc.getParagraphs()[1].style, 'Note');
  assert.equal(doc.getStyle('HeadingChar'), undefined);

  const revision = doc.revision;
  assert.throws(() => doc.applyOperations({ operations: [
    { type: 'updateStyle', id: 'Note', patch: { name: 'Renamed' } },
    { type: 'deleteStyle', id: 'Normal' },
  ] }), /default/);
  assert.equal(doc.getStyle('Note').name, 'Note');
  assert.equal(doc.revision, revision);
});

test('style operation schemas match runtime validation', () => {
  const operations = AGENT_OPERATION_SCHEMA.properties.operations.items.oneOf;
  const find = (name) => operations.find((operation) => operation.properties.type.const === name);
  assert.deepEqual(Object.keys(find('updateStyle').properties.patch.properties).sort(), [...STYLE_PATCH_KEYS].sort());
  // defineStyle 是整体定义，运行时拒绝 null 字段，所以 schema 里不能是 nullable。
  const defined = find('defineStyle').properties.style.properties;
  for (const [key, schema] of Object.entries(defined.paragraph.properties)) {
    assert.ok(!schema.anyOf?.some((entry) => entry.type === 'null'), `defineStyle.paragraph.${key}`);
  }
  for (const schema of Object.values(defined.run.properties)) assert.ok(!schema.anyOf?.some((entry) => entry.type === 'null'));
  // patch 里不暴露 paragraph.style / run.style：继承只能走 basedOn。
  assert.ok(!('style' in find('updateStyle').properties.patch.properties.paragraph.properties));
  assert.ok(!('style' in find('updateStyle').properties.patch.properties.run.properties));
  assert.doesNotThrow(() => validateStylePatch({ name: 'x', basedOn: null, aliases: ['a'], uiPriority: 0, quickFormat: true, paragraph: { alignment: null }, run: { bold: null } }));
});
