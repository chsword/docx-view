import { test } from 'node:test';
import assert from 'node:assert/strict';
import { DocxDocument } from '../dist/document.js';
import { WORD_NS } from '../dist/xml.js';

function withBody(xml) {
  const doc = DocxDocument.create();
  doc.setPartXml(doc.mainDocumentPath, `<w:document xmlns:w="${WORD_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>${xml}</w:body></w:document>`);
  return doc;
}

const xmlOf = (doc) => new TextDecoder().decode(doc.getPartBytes(doc.mainDocumentPath));

// Word 的形状：sectPrChange 是 sectPr 的最后一个子元素，里面的 sectPr 只有版面属性（CT_SectPrBase），
// 没有页眉页脚引用。
const BODY = `
  <w:p><w:pPr><w:sectPr>
    <w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
    <w:sectPrChange w:id="5" w:author="Alice" w:date="2026-09-30T08:00:00Z"><w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr></w:sectPrChange>
  </w:sectPr></w:pPr><w:r><w:t>first section</w:t></w:r></w:p>
  <w:p><w:r><w:t>second section</w:t></w:r></w:p>
  <w:p><w:r><w:t>last</w:t></w:r></w:p>
  <w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/>
    <w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="720" w:footer="720" w:gutter="0"/>
    <w:cols w:num="2" w:space="425"/>
    <w:sectPrChange w:id="6" w:author="Bob"><w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr></w:sectPrChange>
  </w:sectPr>`;

test('section property changes read as revisions anchored at the paragraph that ends the section', () => {
  const doc = withBody(BODY);
  const revisions = doc.getRevisions({ kinds: ['sectionFormatChange'] });
  assert.deepEqual(revisions.map((revision) => [revision.id, revision.paragraph, revision.author, revision.run]), [
    [5, 0, 'Alice', undefined],
    // 正文末尾的 sectPr 结束的是最后一节，锚在最后一段；通用锚点查找会落到第一段。
    [6, 2, 'Bob', undefined],
  ]);
  assert.equal(revisions[0].previousSection.orientation, 'portrait');
  assert.equal(revisions[0].previousSection.pageWidth, 11906);
  assert.equal(revisions[1].previousSection.margins.top, 1440);
  assert.equal(revisions[1].previousSection.columns.count, 1, '快照里没有 cols，就是单栏');
  assert.equal(revisions[1].date, undefined);
});

test('rejecting a section change restores the layout and keeps header references; accepting keeps the new layout', () => {
  const rejected = withBody(BODY);
  rejected.rejectRevision(6);
  const last = rejected.getSections().at(-1);
  assert.equal(last.margins.top, 1440);
  assert.equal(last.columns.count, 1);
  assert.match(xmlOf(rejected), /<w:headerReference w:type="default" r:id="rIdHeader"\/>/, '快照里没有页眉引用，拒绝时不能把它删掉');
  assert.equal(rejected.getRevisions({ kinds: ['sectionFormatChange'] }).length, 1);
  // 子元素顺序：headerReference 仍在最前，pgSz 在 pgMar 前。
  const tail = xmlOf(rejected).match(/<w:sectPr><w:headerReference[\s\S]*?<\/w:sectPr>\s*<\/w:body>/)[0];
  assert.deepEqual([...tail.matchAll(/<w:(\w+)[ />]/g)].map((match) => match[1]), ['sectPr', 'headerReference', 'pgSz', 'pgMar']);

  const accepted = withBody(BODY);
  accepted.acceptRevision(5);
  assert.equal(accepted.getSections()[0].orientation, 'landscape');
  assert.doesNotMatch(xmlOf(accepted).split('first section')[0], /sectPrChange/);
  accepted.acceptAllRevisions();
  assert.deepEqual(accepted.getRevisions(), []);
  assert.equal(accepted.getSections().at(-1).columns.count, 2);
});

test('tracked page setup records the original layout once, and a change back removes the revision', () => {
  const doc = withBody('<w:p><w:r><w:t>x</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>');
  doc.setTrackChanges(true);
  doc.setRevisionAuthor('Carol');

  // no-op 不留修订，也不推进 revision。
  const revision = doc.revision;
  doc.setPageSetup(0, { pageWidth: 11906 });
  assert.equal(doc.revision, revision);
  assert.deepEqual(doc.getRevisions(), []);

  doc.setPageSetup(0, { orientation: 'landscape' });
  doc.setPageSetup(0, { margins: { top: 720 } });
  const [change] = doc.getRevisions();
  assert.equal(change.kind, 'sectionFormatChange');
  assert.equal(change.author, 'Carol');
  // 第 17 条：快照始终是最初的未修订属性，不是第一次修改后的中间态。
  assert.equal(change.previousSection.orientation, 'portrait');
  assert.equal(change.previousSection.margins.top, 1440);
  assert.equal(doc.getSections()[0].margins.top, 720);
  assert.equal((xmlOf(doc).match(/sectPrChange/g) ?? []).length, 2, '只有一条（开闭两个标签）');

  doc.rejectRevision(change.id);
  assert.equal(doc.getSections()[0].orientation, 'portrait');
  assert.equal(doc.getSections()[0].margins.top, 1440);

  doc.setPageSetup(0, { margins: { top: 900 } });
  assert.equal(doc.getRevisions().length, 1);
  doc.setPageSetup(0, { margins: { top: 1440 } });
  assert.deepEqual(doc.getRevisions(), [], '改回最初的值之后修订没有内容，删掉');
  assert.doesNotMatch(xmlOf(doc), /sectPrChange/);
});

test('section revision ids stay unique across the document', () => {
  const doc = withBody(`<w:p><w:ins w:id="40" w:author="A"><w:r><w:t>x</w:t></w:r></w:ins></w:p>
    <w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>`);
  doc.setTrackChanges(true);
  doc.setPageSetup(0, { orientation: 'landscape' });
  const ids = doc.getRevisions().map((revision) => revision.id);
  assert.equal(new Set(ids).size, ids.length);
  assert.ok(ids.includes(41));
});
