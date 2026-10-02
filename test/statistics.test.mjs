import { test } from 'node:test';
import assert from 'node:assert/strict';
import { DocxDocument } from '../dist/document.js';
import { WORD_NS } from '../dist/xml.js';
import { countTextStatistics } from '../dist/docprops.js';

function withBody(xml) {
  const doc = DocxDocument.create();
  doc.setPartXml(doc.mainDocumentPath, `<w:document xmlns:w="${WORD_NS}"><w:body>${xml}<w:sectPr/></w:body></w:document>`);
  return doc;
}

test('word and character counts follow Word for Latin, CJK and mixed text', () => {
  assert.deepEqual(countTextStatistics('Hello world'), { words: 2, characters: 10, charactersWithSpaces: 11 });
  // 每个汉字算一个词；中文标点不算词，但算字符。
  assert.deepEqual(countTextStatistics('中文测试，好。'), { words: 5, characters: 7, charactersWithSpaces: 7 });
  // 「Word」「2026」各一个词，「年度报告」四个。
  assert.equal(countTextStatistics('Word 2026 年度报告').words, 6);
  assert.equal(countTextStatistics('e-mail, co-op').words, 2);
  // 字符按码点计：emoji 是一个字符，不是两个 UTF-16 单元。
  assert.deepEqual(countTextStatistics('a😀b'), { words: 1, characters: 3, charactersWithSpaces: 3 });
  assert.deepEqual(countTextStatistics('かなカナ한국'), { words: 6, characters: 6, charactersWithSpaces: 6 });
  assert.deepEqual(countTextStatistics('  \t '), { words: 0, characters: 0, charactersWithSpaces: 4 });
});

test('document statistics count the body, tables included, without hidden or deleted text', () => {
  const doc = withBody(`
    <w:p><w:r><w:t xml:space="preserve">Hello world </w:t></w:r><w:r><w:rPr><w:vanish/></w:rPr><w:t>secret words</w:t></w:r></w:p>
    <w:p/>
    <w:p><w:r><w:t>中文</w:t></w:r><w:del w:id="1" w:author="A"><w:r><w:delText>删掉</w:delText></w:r></w:del></w:p>
    <w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>cell text</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`);
  // 「Hello world 」10 / 12，「中文」2 / 2，「cell text」8 / 9；隐藏的 secret words 与删除的「删掉」不算，空段落不算段落。
  assert.deepEqual(doc.getDocumentStatistics(), { words: 6, characters: 20, charactersWithSpaces: 23, paragraphs: 3 });
});

test('updateDocumentStatistics writes app.xml in schema order, once, and only writes pages with pagination', () => {
  const doc = withBody('<w:p><w:r><w:t>one two three</w:t></w:r></w:p>');
  doc.setDocumentProperties({ company: 'Acme' });
  const revision = doc.revision;
  const written = doc.updateDocumentStatistics();
  assert.deepEqual(written, { words: 3, characters: 11, charactersWithSpaces: 13, paragraphs: 1 });
  assert.equal(doc.revision, revision + 1);
  const app = () => new TextDecoder().decode(doc.getPartBytes('docProps/app.xml'));
  const order = () => [...app().matchAll(/<(\w+)>/g)].map((match) => match[1]).filter((name) => name !== 'Properties');
  // CT_Properties 是 sequence：Company 在 Pages 之前，CharactersWithSpaces 在 Paragraphs 之后。
  assert.deepEqual(order(), ['Company', 'Words', 'Characters', 'Paragraphs', 'CharactersWithSpaces']);
  assert.doesNotMatch(app(), /<Pages>/, '没有分页结果就不写页数');

  // 值没变：不推进 revision、不进撤销历史。
  doc.updateDocumentStatistics();
  assert.equal(doc.revision, revision + 1);

  const paged = doc.updateDocumentStatistics({ pagination: { pageCount: 2, lineCount: 37, pageOfParagraph: () => 0, numberOfPage: (index) => index + 1 } });
  assert.equal(paged.pages, 2);
  assert.deepEqual(order(), ['Company', 'Pages', 'Words', 'Characters', 'Lines', 'Paragraphs', 'CharactersWithSpaces']);
  assert.match(app(), /<Lines>37<\/Lines>/);
  assert.equal(doc.getDocumentProperties().company, 'Acme', '已有的属性不动');

  doc.undo();
  assert.doesNotMatch(app(), /<Pages>/);
  assert.throws(() => doc.updateDocumentStatistics({ pagination: { pageCount: -1 } }), /pageCount/);
});

test('updateDocumentStatistics creates app.xml with its relationship when the package has none', async () => {
  const doc = withBody('<w:p><w:r><w:t>x</w:t></w:r></w:p>');
  assert.ok(!doc.listParts().includes('docProps/app.xml'));
  doc.updateDocumentStatistics();
  const reopened = await DocxDocument.load(await doc.toUint8Array());
  assert.match(new TextDecoder().decode(reopened.getPartBytes('_rels/.rels')), /extended-properties/);
  assert.match(new TextDecoder().decode(reopened.getPartBytes('[Content_Types].xml')), /\/docProps\/app\.xml/);
  assert.match(new TextDecoder().decode(reopened.getPartBytes('docProps/app.xml')), /<Words>1<\/Words>/);
});
