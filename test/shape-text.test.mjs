import { test } from 'node:test';
import assert from 'node:assert/strict';
import { DocxDocument } from '../dist/document.js';
import { WORD_NS, REL_NS, OFFICE_REL_NS } from '../dist/xml.js';

const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const WPS = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';
const M = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
const DSP = 'http://schemas.microsoft.com/office/drawing/2008/diagram';
const encoder = new TextEncoder();

test('SmartArt node text keeps run formatting, alignment and line breaks', () => {
  const doc = DocxDocument.create();
  doc.setPartXml(doc.mainDocumentPath, `<w:document xmlns:w="${WORD_NS}" xmlns:r="${OFFICE_REL_NS}" xmlns:wp="${WP}" xmlns:a="${A}" xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"><w:body><w:p><w:r><w:drawing><wp:inline><wp:extent cx="5486400" cy="3200400"/><wp:docPr id="9"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:relIds r:dm="rIdData" r:lo="rIdLayout" r:qs="rIdStyle" r:cs="rIdColors"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p><w:sectPr/></w:body></w:document>`);
  doc.addPart('word/_rels/document.xml.rels', encoder.encode(`<Relationships xmlns="${REL_NS}"><Relationship Id="rIdDrawing" Type="http://schemas.microsoft.com/office/word/2008/relationships/diagramDrawing" Target="diagrams/drawing1.xml"/><Relationship Id="rIdData" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="diagrams/data1.xml"/></Relationships>`), 'application/vnd.openxmlformats-package.relationships+xml');
  // Word 预渲染的 dsp:txBody：bodyPr、lstStyle、带 pPr 的段落，run 有 lang / sz / b 与实心填充。
  doc.addPart('word/diagrams/drawing1.xml', encoder.encode(`<dsp:drawing xmlns:dsp="${DSP}" xmlns:a="${A}"><dsp:spTree><dsp:nvGrpSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvGrpSpPr/></dsp:nvGrpSpPr><dsp:grpSpPr/>
    <dsp:sp modelId="{1}"><dsp:nvSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvSpPr/></dsp:nvSpPr>
      <dsp:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="1097280"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></dsp:spPr>
      <dsp:txBody><a:bodyPr spcFirstLastPara="0" vert="horz" wrap="square" lIns="64008" tIns="64008" rIns="64008" bIns="64008" numCol="1" anchor="ctr" anchorCtr="0"><a:noAutofit/></a:bodyPr><a:lstStyle/>
        <a:p><a:pPr marL="0" lvl="0" indent="0" algn="l" defTabSz="1866900"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc></a:pPr>
          <a:r><a:rPr lang="zh-CN" altLang="en-US" sz="1800" b="1" kern="1200"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:latin typeface="Microsoft YaHei"/></a:rPr><a:t>标题</a:t></a:r>
          <a:r><a:rPr lang="zh-CN" sz="1800" i="1" u="sng" kern="1200"/><a:t>斜</a:t></a:r>
          <a:br><a:rPr lang="zh-CN" sz="1800"/></a:br>
          <a:r><a:rPr lang="en-US" sz="1200" kern="1200"><a:latin typeface="+mn-lt"/></a:rPr><a:t>second line</a:t></a:r></a:p>
        <a:p><a:pPr algn="r"/><a:r><a:rPr lang="en-US" sz="1000" strike="sngStrike"/><a:t>right</a:t></a:r></a:p>
        <a:p><a:pPr algn="ctr"/><a:endParaRPr lang="en-US"/></a:p>
      </dsp:txBody></dsp:sp></dsp:spTree></dsp:drawing>`), 'application/vnd.ms-office.drawingml.diagramDrawing+xml');
  doc.addPart('word/diagrams/data1.xml', encoder.encode('<diagram/>'), 'application/xml');
  const child = doc.getShapes()[0].children[0];
  assert.equal(child.text, '标题斜second lineright', '旧的纯文本字段保留');
  assert.deepEqual(child.paragraphs, [
    { alignment: 'left', lines: [
      [{ text: '标题', bold: true, fontSize: 18, fontFamily: 'Microsoft YaHei', color: '#FFFFFF' }, { text: '斜', italic: true, underline: true, fontSize: 18 }],
      // +mn-lt 是主题字体占位，不是字体名。
      [{ text: 'second line', fontSize: 12 }],
    ] },
    { alignment: 'right', lines: [[{ text: 'right', strike: true, fontSize: 10 }]] },
  ], '空段落不留');
});

test('formulas inside a text box come back from getShapeParagraphs', () => {
  const doc = DocxDocument.create();
  doc.setPartXml(doc.mainDocumentPath, `<w:document xmlns:w="${WORD_NS}" xmlns:wp="${WP}" xmlns:a="${A}" xmlns:wps="${WPS}" xmlns:m="${M}"><w:body><w:p><w:r><w:t>before</w:t></w:r><w:r><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="1828800" cy="914400"/><wp:wrapSquare wrapText="bothSides"/><wp:docPr id="3" name="Text Box 3"/><a:graphic><a:graphicData uri="${WPS}"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr>
    <wps:txbx><w:txbxContent><w:p><w:r><w:t xml:space="preserve">Area: </w:t></w:r><m:oMath><m:sSup><m:e><m:r><m:t>r</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup></m:oMath></w:p></w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p><w:sectPr/></w:body></w:document>`);
  const shape = doc.getShapes()[0];
  const [paragraph] = doc.getShapeParagraphs(shape.id);
  assert.equal(paragraph.math.length, 1);
  assert.equal(paragraph.math[0].linear, 'r^2');
  // 文本框里的公式不属于锚定它的正文段落（第 10 条）。原先按后代查找把它算进了正文第 0 段：
  // getMath() 报在第 0 段、编辑器在正文里再画一遍、deleteMath(0, 0) 会删到文本框里去。
  assert.deepEqual(doc.getMath(), []);
  assert.deepEqual(doc.getParagraphs()[0].math ?? [], []);
  assert.throws(() => doc.deleteMath(0, 0), /does not exist/);
  assert.match(new TextDecoder().decode(doc.getPartBytes(doc.mainDocumentPath)), /<m:oMath>/);
});
