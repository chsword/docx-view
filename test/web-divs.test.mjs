import { test } from 'node:test';
import assert from 'node:assert/strict';
import { DocxDocument } from '../dist/document.js';
import { WORD_NS } from '../dist/xml.js';
import { webDivIndents } from '../dist/web-divs.js';

const WEB_SETTINGS_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml';
const WEB_SETTINGS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings';

// Outlook 回复邮件另存为 docx 的形状：body div 里套一层 blockquote，blockquote 带左边距与左框线。
const div = (id, inner, { left = 0, right = 0, blockQuote = false, bodyDiv = false, border = '' } = {}) => `
  <w:div w:id="${id}">${blockQuote ? '<w:blockQuote w:val="1"/>' : ''}${bodyDiv ? '<w:bodyDiv w:val="1"/>' : ''}
    <w:marLeft w:val="${left}"/><w:marRight w:val="${right}"/><w:marTop w:val="0"/><w:marBottom w:val="0"/>
    <w:divBdr><w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>${border}
      <w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/><w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/></w:divBdr>
    ${inner ? `<w:divsChild>${inner}</w:divsChild>` : ''}
  </w:div>`;
const WEB_SETTINGS = `<w:webSettings xmlns:w="${WORD_NS}"><w:divs>${div(1450904384,
  div(733114029, div(98765, '', { left: 360 }), { left: 720, right: 120, blockQuote: true,
    border: '<w:left w:val="single" w:sz="12" w:space="4" w:color="1F4E79"/>' }),
  { bodyDiv: true })}</w:divs><w:optimizeForBrowser/></w:webSettings>`;

function withWebSettings(body) {
  const doc = DocxDocument.create();
  doc.setPartXml(doc.mainDocumentPath, `<w:document xmlns:w="${WORD_NS}"><w:body>${body}<w:sectPr/></w:body></w:document>`);
  // 部件名不用惯例的 webSettings.xml：路径必须从关系解析（第 1 条）。
  doc.addPart('word/web/ws.xml', new TextEncoder().encode(WEB_SETTINGS), WEB_SETTINGS_TYPE);
  doc.addPart('word/_rels/document.xml.rels', new TextEncoder().encode(
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`
    + `<Relationship Id="rIdWeb" Type="${WEB_SETTINGS_REL}" Target="web/ws.xml"/></Relationships>`),
  'application/vnd.openxmlformats-package.relationships+xml');
  return doc;
}

const BODY = `
  <w:p><w:pPr><w:divId w:val="733114029"/><w:ind w:left="200"/></w:pPr><w:r><w:t>quoted</w:t></w:r></w:p>
  <w:p><w:pPr><w:divId w:val="98765"/></w:pPr><w:r><w:t>nested</w:t></w:r></w:p>
  <w:p><w:pPr><w:divId w:val="424242"/></w:pPr><w:r><w:t>dangling</w:t></w:r></w:p>
  <w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>
    <w:tr><w:trPr><w:divId w:val="733114029"/></w:trPr><w:tc><w:p/></w:tc></w:tr></w:tbl>`;

test('getWebDivs flattens the w:divs tree from webSettings.xml', () => {
  const divs = withWebSettings(BODY).getWebDivs();
  assert.deepEqual(divs.map((entry) => [entry.id, entry.parentId, entry.bodyDiv, entry.blockQuote, entry.marginLeft, entry.marginRight]), [
    [1450904384, undefined, true, false, 0, 0],
    [733114029, 1450904384, false, true, 720, 120],
    [98765, 733114029, false, false, 360, 0],
  ]);
  assert.deepEqual(divs[1].borders.left, { style: 'single', none: false, size: 12, space: 4, color: '1F4E79' });
  assert.equal(divs[1].borders.top.none, true);
  assert.deepEqual(DocxDocument.create().getWebDivs(), [], '没有 webSettings 部件时为空');
});

test('div margins accumulate down the nesting chain', () => {
  const indents = webDivIndents(withWebSettings(BODY).getWebDivs());
  assert.deepEqual(indents.get(733114029), { left: 720, right: 120 });
  assert.deepEqual(indents.get(98765), { left: 1080, right: 120 }, '自己的 360 + blockquote 的 720');
  // 构造出来的成环 parentId 不能死循环。
  const looped = webDivIndents([{ id: 1, parentId: 2, marginLeft: 10, marginRight: 0 }, { id: 2, parentId: 1, marginLeft: 20, marginRight: 0 }]);
  assert.deepEqual(looped.get(1), { left: 30, right: 0 });
});

test('divId reads and writes on paragraphs and rows, and stays out of the clipboard', () => {
  const doc = withWebSettings(BODY);
  const paragraphs = doc.getParagraphs();
  assert.equal(paragraphs[0].divId, 733114029);
  assert.equal(doc.getBlocks().find((block) => block.type === 'table').rows[0].format.divId, 733114029);

  doc.formatParagraph(2, { divId: 98765 });
  assert.equal(doc.getParagraphs()[2].divId, 98765);
  doc.formatParagraph(2, { divId: null });
  assert.equal(doc.getParagraphs()[2].divId, undefined);
  doc.formatTableRow(0, 0, { divId: null });
  assert.equal(doc.getBlocks().find((block) => block.type === 'table').rows[0].format?.divId, undefined);
  assert.throws(() => doc.formatParagraph(0, { divId: -1 }), /divId/);
  assert.throws(() => doc.formatTableRow(0, 0, { divId: 1.5 }), /divId/);

  // div id 只在本文档的 webSettings 里有意义，剪贴板不带。
  const fragment = doc.copyClipboardFragment({ start: { paragraph: 0, offset: 0 }, end: { paragraph: 0, offset: 6 } });
  assert.equal(fragment.paragraphs[0].format.divId, undefined);
  assert.equal(fragment.paragraphs[0].format.indentLeft, 200);
});
