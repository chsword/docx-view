import { test } from 'node:test';
import assert from 'node:assert/strict';
import { DocxDocument } from '../dist/document.js';
import { WORD_NS } from '../dist/xml.js';

const STYLES_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml';

function docWith(body, styles) {
  const doc = DocxDocument.create();
  doc.setPartXml(doc.mainDocumentPath, `<w:document xmlns:w="${WORD_NS}"><w:body>${body}<w:sectPr/></w:body></w:document>`);
  if (styles !== undefined) doc.addPart('word/styles.xml', new TextEncoder().encode(`<w:styles xmlns:w="${WORD_NS}">${styles}</w:styles>`), STYLES_TYPE);
  return doc;
}

const SOURCE_STYLES = `
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:rPr><w:sz w:val="20"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Base"><w:name w:val="Base"/><w:basedOn w:val="Normal"/><w:rPr><w:color w:val="112233"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Heading"><w:name w:val="Heading"/><w:basedOn w:val="Base"/><w:next w:val="Body"/><w:link w:val="HeadingChar"/>
    <w:pPr><w:jc w:val="center"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>
  <w:style w:type="character" w:styleId="HeadingChar"><w:name w:val="Heading Char"/><w:link w:val="Heading"/><w:rPr><w:b/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Body"><w:name w:val="Body"/></w:style>
  <w:style w:type="character" w:styleId="Strong"><w:name w:val="Strong"/><w:rPr><w:i/></w:rPr></w:style>`;

const SOURCE_BODY = `<w:p><w:pPr><w:pStyle w:val="Heading"/></w:pPr><w:r><w:t xml:space="preserve">Title </w:t></w:r><w:r><w:rPr><w:rStyle w:val="Strong"/></w:rPr><w:t>here</w:t></w:r></w:p>`;

test('copying carries the referenced styles with their basedOn chain and linked pair', () => {
  const source = docWith(SOURCE_BODY, SOURCE_STYLES);
  const fragment = source.copyClipboardFragment({ start: { paragraph: 0, offset: 0 }, end: { paragraph: 0, offset: 10 } });
  const byId = new Map(fragment.styles.map((style) => [style.id, style]));
  assert.deepEqual([...byId.keys()].sort(), ['Base', 'Heading', 'HeadingChar', 'Normal', 'Strong']);
  assert.equal(byId.get('Normal').isDefault, undefined, '默认样式标记不带：目标文档会出现两个默认样式');
  assert.equal(byId.get('Heading').next, undefined, 'next 指向集合外的 Body，去掉');
  assert.equal(byId.get('Heading').link, 'HeadingChar');
  assert.deepEqual(byId.get('Heading').paragraph, { alignment: 'center' });
  // 片段要能 JSON 往返（编辑器把它放进自定义 MIME），所以不能有 undefined 键之外的怪东西。
  assert.deepEqual(JSON.parse(JSON.stringify(fragment.styles)), fragment.styles);
});

test('pasting into another document defines only the styles it lacks, in one undo step', () => {
  const source = docWith(SOURCE_BODY, SOURCE_STYLES);
  const fragment = JSON.parse(JSON.stringify(source.copyClipboardFragment({ start: { paragraph: 0, offset: 0 }, end: { paragraph: 0, offset: 10 } })));
  // 目标文档自己的 Normal 是 14 磅、红色；同 ID 的样式以目标为准。
  const target = docWith('<w:p><w:r><w:t>x</w:t></w:r></w:p>',
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:rPr><w:sz w:val="28"/><w:color w:val="FF0000"/></w:rPr></w:style>');
  const revision = target.revision;
  assert.equal(target.pasteClipboardFragment({ start: { paragraph: 0, offset: 1 }, end: { paragraph: 0, offset: 1 } }, fragment), true);
  assert.equal(target.revision, revision + 1, '定义样式和粘贴是同一次提交');
  assert.deepEqual(target.getStyles().map((style) => style.id).sort(), ['Base', 'Heading', 'HeadingChar', 'Normal', 'Strong']);
  assert.equal(target.getStyle('Normal').run.fontSize, 14, '目标文档的 Normal 不被覆盖');
  assert.equal(target.getStyles().filter((style) => style.isDefault && style.type === 'paragraph').length, 1);
  const pasted = target.getParagraphs()[0];
  assert.equal(pasted.style, 'Heading');
  assert.equal(pasted.effective.alignment, 'center', '样式定义跟过来了，不再悬空回落');
  assert.equal(pasted.runs.find((run) => run.text === 'here').effective.italic, true);
  // Base 继承的是**目标**文档的 Normal：颜色来自 Base 自己，字号来自 Heading。
  assert.equal(pasted.runs[0].effective.color, '112233');

  target.undo();
  assert.equal(target.getStyle('Heading'), undefined, '撤销一步样式也一起回去');
});

test('clipboard style definitions are untrusted input', () => {
  const target = docWith('<w:p><w:r><w:t>x</w:t></w:r></w:p>', '');
  const at = { start: { paragraph: 0, offset: 1 }, end: { paragraph: 0, offset: 1 } };
  const paragraphs = [{ runs: [{ text: 'y' }] }];
  const revision = target.revision;
  for (const styles of [
    [{ id: 'X', type: 'paragraph', run: { color: 'red' } }],
    [{ id: 'X', type: 'nonsense' }],
    [{ id: 'X', type: 'paragraph', paragraph: { alignment: null } }],
    [{ id: 'X', type: 'paragraph', onload: 'alert(1)' }],
    Array.from({ length: 201 }, (_, index) => ({ id: `S${index}`, type: 'paragraph' })),
    'not an array',
  ]) {
    assert.throws(() => target.pasteClipboardFragment(at, { version: 1, text: 'y', paragraphs, styles }));
  }
  assert.equal(target.revision, revision);
  assert.deepEqual(target.getStyles(), []);
});
