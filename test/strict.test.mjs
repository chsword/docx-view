import { test } from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import { DocxDocument } from '../dist/document.js';

const S = 'http://purl.oclc.org/ooxml';

// 照 Word「Strict Open XML 文档」的输出拼：命名空间、关系类型都在 purl.oclc.org 下，根上有
// w:conformance="strict"，表格宽度写成百分比字符串，内容类型与 Transitional 相同。
async function strictBytes() {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
    + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    + '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
    + '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>');
  zip.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`
    + `<Relationship Id="rId1" Type="${S}/officeDocument/relationships/officeDocument" Target="word/document.xml"/>`
    + `<Relationship Id="rId2" Type="${S}/officeDocument/relationships/extendedProperties" Target="docProps/app.xml"/></Relationships>`);
  zip.file('docProps/app.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="${S}/officeDocument/extendedProperties" xmlns:vt="${S}/officeDocument/docPropsVTypes"><Company>Strict Co</Company></Properties>`);
  zip.file('word/_rels/document.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${S}/officeDocument/relationships/styles" Target="styles.xml"/></Relationships>`);
  zip.file('word/styles.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="${S}/wordprocessingml/main"><w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:rPr><w:b w:val="true"/><w:sz w:val="40"/></w:rPr></w:style></w:styles>`);
  zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${S}/wordprocessingml/main" xmlns:r="${S}/officeDocument/relationships" xmlns:m="${S}/officeDocument/math" w:conformance="strict"><w:body>`
    + '<w:p><w:pPr><w:pStyle w:val="Title"/></w:pPr><w:r><w:t>Strict title</w:t></w:r></w:p>'
    + '<w:p><w:r><w:rPr><w:w w:val="150%"/></w:rPr><w:t>wide</w:t></w:r><m:oMath><m:r><m:t>x</m:t></m:r></m:oMath></w:p>'
    + '<w:tbl><w:tblPr><w:tblW w:w="50%" w:type="pct"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
    + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>');
  return zip.generateAsync({ type: 'uint8array' });
}

test('Strict OOXML opens: namespaces, relationship types and percentage values convert to Transitional', async () => {
  const doc = await DocxDocument.load(await strictBytes());
  assert.equal(doc.getPackageKind().convertedFromStrict, true);
  const [title, wide] = doc.getParagraphs();
  assert.equal(title.text, 'Strict title');
  assert.equal(title.style, 'Title');
  assert.equal(title.runs[0].effective.bold, true, '样式关系按 Strict 的关系类型解析到了');
  assert.equal(title.runs[0].effective.fontSize, 20);
  assert.equal(wide.runs[0].characterScale, 150, '150% → 150');
  assert.equal(doc.getMath()[0].linear, 'x');
  assert.deepEqual(doc.getBlocks().find((block) => block.type === 'table').format.width, { type: 'pct', value: 2500 }, '50% → 2500（五十分之一个百分点）');
  assert.equal(doc.getDocumentProperties().company, 'Strict Co', 'extendedProperties 关系类型名字不同，单独映射');

  // 存出来是 Transitional：再打开时没有任何 Strict 的痕迹。
  doc.setParagraphText(1, 'edited');
  const saved = await doc.toUint8Array();
  const reopened = await DocxDocument.load(saved);
  assert.equal(reopened.getPackageKind().convertedFromStrict, false);
  const zip = await JSZip.loadAsync(saved);
  for (const name of Object.keys(zip.files)) {
    if (zip.files[name].dir) continue;
    assert.doesNotMatch(await zip.file(name).async('string'), /purl\.oclc\.org|conformance/, name);
  }
  assert.equal(reopened.getParagraphs()[1].text, 'edited');
});

test('Transitional packages are left byte-for-byte alone, even when their text quotes a Strict namespace', async () => {
  // 一份讲 OOXML 的普通文档，正文里引用了 Strict 的命名空间 URI：不是 Strict 包，正文不能被改。
  const original = DocxDocument.create();
  original.setParagraphText(0, 'Strict 的正文命名空间是 http://purl.oclc.org/ooxml/wordprocessingml/main');
  const bytes = await original.toUint8Array();
  const reopened = await DocxDocument.load(bytes);
  assert.equal(reopened.getPackageKind().convertedFromStrict, false);
  assert.deepEqual(reopened.getPartBytes('word/document.xml'), original.getPartBytes('word/document.xml'));
  assert.match(reopened.getParagraphs()[0].text, /purl\.oclc\.org\/ooxml\/wordprocessingml\/main/);
});
