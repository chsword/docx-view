import { test } from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import { DocxDocument } from '../dist/document.js';

const VBA_BYTES = Uint8Array.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1, 1, 2, 3, 4]);

// 照 Word 存 .docm 的形状拼包：主文档内容类型是 macroEnabled，Default 里有 bin → vbaProject，
// 主文档关系指向 vbaProject.bin，它自己的关系再指向 vbaData.xml。
async function docmBytes({ template = false, defaultTyped = false } = {}) {
  const zip = new JSZip();
  const mainType = template ? 'application/vnd.ms-word.template.macroEnabledTemplate.main+xml' : 'application/vnd.ms-word.document.macroEnabled.main+xml';
  zip.file('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    + '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    + `<Default Extension="xml" ContentType="${defaultTyped ? mainType : 'application/xml'}"/>`
    + (defaultTyped ? '' : `<Override PartName="/word/document.xml" ContentType="${mainType}"/>`)
    + '<Override PartName="/word/vbaData.xml" ContentType="application/vnd.ms-word.vbaData+xml"/></Types>');
  zip.file('_rels/.rels', '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file('word/document.xml', '<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Macro host</w:t></w:r></w:p><w:sectPr/></w:body></w:document>');
  zip.file('word/_rels/document.xml.rels', '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId9" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/></Relationships>');
  zip.file('word/vbaProject.bin', VBA_BYTES);
  zip.file('word/_rels/vbaProject.bin.rels', '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2006/relationships/wordVbaData" Target="vbaData.xml"/></Relationships>');
  zip.file('word/vbaData.xml', '<?xml version="1.0" encoding="UTF-8"?><wne:vbaSuppData xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"/>');
  return zip.generateAsync({ type: 'uint8array' });
}

test('macro-enabled documents open, edit and save with the VBA project untouched', async () => {
  const doc = await DocxDocument.load(await docmBytes());
  assert.deepEqual(doc.getPackageKind(), {
    kind: 'macroEnabledDocument', extension: 'docm', mimeType: 'application/vnd.ms-word.document.macroEnabled.12', hasMacros: true,
  });
  doc.setParagraphText(0, 'Edited, macros carried along');
  const reopened = await DocxDocument.load(await doc.toUint8Array());
  assert.equal(reopened.getParagraphs()[0].text, 'Edited, macros carried along');
  assert.deepEqual([...reopened.getPartBytes('word/vbaProject.bin')], [...VBA_BYTES], 'VBA 工程逐字节保留');
  assert.equal(reopened.getPackageKind().kind, 'macroEnabledDocument');
  // Blob 按规范把 MIME 转成小写。
  assert.equal((await reopened.toBlob()).type, 'application/vnd.ms-word.document.macroenabled.12');
});

test('removeMacros strips the VBA project and turns .docm into .docx, as one undo step', async () => {
  const doc = await DocxDocument.load(await docmBytes());
  const revision = doc.revision;
  doc.removeMacros();
  assert.equal(doc.revision, revision + 1);
  assert.deepEqual(doc.getPackageKind(), {
    kind: 'document', extension: 'docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', hasMacros: false,
  });
  for (const path of ['word/vbaProject.bin', 'word/_rels/vbaProject.bin.rels', 'word/vbaData.xml']) assert.ok(!doc.listParts().includes(path), path);
  const types = new TextDecoder().decode(doc.getPartBytes('[Content_Types].xml'));
  assert.doesNotMatch(types, /vbaData/);
  assert.doesNotMatch(new TextDecoder().decode(doc.getPartBytes('word/_rels/document.xml.rels')), /vbaProject/);
  // 存出来再打开，是一份普通 .docx。
  const reopened = await DocxDocument.load(await doc.toUint8Array());
  assert.equal(reopened.getPackageKind().kind, 'document');
  doc.removeMacros();
  assert.equal(doc.revision, revision + 1, '没有宏了再调是空操作');
  doc.undo();
  assert.equal(doc.getPackageKind().hasMacros, true);
});

test('templates keep their kind, and a main type that comes from a Default entry is handled', async () => {
  const template = await DocxDocument.load(await docmBytes({ template: true }));
  assert.equal(template.getPackageKind().extension, 'dotm');
  template.removeMacros();
  assert.equal(template.getPackageKind().extension, 'dotx', '.dotm 去掉宏是 .dotx，不是 .docx');
  const defaultTyped = await DocxDocument.load(await docmBytes({ defaultTyped: true }));
  assert.equal(defaultTyped.getPackageKind().kind, 'macroEnabledDocument');
  defaultTyped.removeMacros();
  assert.equal(defaultTyped.getPackageKind().kind, 'document');
  assert.equal(DocxDocument.create().getPackageKind().kind, 'document');
});
