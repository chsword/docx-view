import { test } from 'node:test';
import assert from 'node:assert/strict';
import { DocxDocument } from '../dist/document.js';
import { WORD_NS, parseXml } from '../dist/xml.js';
import { resolveDrawingColor } from '../dist/styles.js';
import { presetGeometryIsOpen, presetGeometryPath, supportedPresetGeometries } from '../dist/geometry.js';

const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const STYLES_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml';
const THEME_TYPE = 'application/vnd.openxmlformats-officedocument.theme+xml';
const encoder = new TextEncoder();
const OFFICE_THEME = { colors: { accent1: '4472C4', accent2: 'ED7D31', dk1: '000000', lt1: 'FFFFFF' }, fonts: {} };

const color = (xml) => resolveDrawingColor(OFFICE_THEME, parseXml(`<r xmlns:a="${A_NS}">${xml}</r>`).documentElement.firstChild);

test('DrawingML colour transforms apply in document order, like Word theme tints and shades', () => {
  // Word「主题颜色」面板里的淡色 / 深色就是 lumMod + lumOff；原先这两个都被忽略，整格落成原色。
  // 对照 Office 色板（差 1 以内：Office 自己的几处色板之间也差 1）。
  const near = (actual, expected) => {
    const channels = (hex) => hex.match(/../g).map((part) => parseInt(part, 16));
    channels(actual).forEach((value, index) => assert.ok(Math.abs(value - channels(expected)[index]) <= 1, `${actual} vs ${expected}`));
  };
  near(color('<a:schemeClr val="accent1"><a:lumMod val="20000"/><a:lumOff val="80000"/></a:schemeClr>'), 'D9E2F3');
  near(color('<a:schemeClr val="accent1"><a:lumMod val="60000"/><a:lumOff val="40000"/></a:schemeClr>'), '8EAADB');
  near(color('<a:schemeClr val="accent1"><a:lumMod val="75000"/></a:schemeClr>'), '2F5496');
  near(color('<a:schemeClr val="accent2"><a:lumMod val="40000"/><a:lumOff val="60000"/></a:schemeClr>'), 'F8CBAD');
  near(color('<a:schemeClr val="accent2"><a:lumMod val="50000"/></a:schemeClr>'), '843C0C');
  // 顺序有意义：先乘再加，和先加再乘结果不同。
  assert.notEqual(color('<a:srgbClr val="808080"><a:lumMod val="50000"/><a:lumOff val="50000"/></a:srgbClr>'),
    color('<a:srgbClr val="808080"><a:lumOff val="50000"/><a:lumMod val="50000"/></a:srgbClr>'));
  // srgbClr 上的变换也要套（原先十六进制颜色直接返回，变换全丢）。
  assert.equal(color('<a:srgbClr val="FF0000"><a:lumMod val="50000"/></a:srgbClr>'), '800000');
  assert.equal(color('<a:srgbClr val="FF0000"><a:inv/></a:srgbClr>'), '00FFFF');
  assert.equal(color('<a:srgbClr val="FF0000"><a:comp/></a:srgbClr>'), '00FFFF');
  assert.equal(color('<a:srgbClr val="FF0000"><a:gray/></a:srgbClr>'), '4D4D4D');
});

test('every DrawingML colour source resolves', () => {
  assert.equal(color('<a:sysClr val="windowText" lastClr="112233"/>'), '112233', 'lastClr 是 Word 存下的实际颜色');
  assert.equal(color('<a:sysClr val="window"/>'), 'FFFFFF');
  assert.equal(color('<a:prstClr val="red"/>'), 'FF0000');
  assert.equal(color('<a:prstClr val="nonsense"/>'), undefined);
  assert.equal(color('<a:scrgbClr r="100000" g="0" b="0"/>'), 'FF0000');
  assert.equal(color('<a:scrgbClr r="50000" g="50000" b="50000"/>'), 'BCBCBC', 'scRGB 是线性的，50% 不是 80');
  assert.equal(color('<a:hslClr hue="7200000" sat="100000" lum="50000"/>'), '00FF00');
  // 占位色 phClr 没有引用处的颜色时解析不出来，而不是瞎猜一个。
  assert.equal(color('<a:schemeClr val="phClr"/>'), undefined);
});

function themedShapes(spPrs, styles) {
  const theme = `<a:theme xmlns:a="${A_NS}"><a:themeElements>
    <a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1>
      <a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4>
      <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink>
      <a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme>
    <a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/></a:minorFont></a:fontScheme>
    <a:fmtScheme name="Office">
      <a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="110000"/><a:satMod val="105000"/><a:tint val="67000"/></a:schemeClr></a:gs>
          <a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="109000"/><a:tint val="81000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>
        <a:solidFill><a:schemeClr val="phClr"><a:shade val="50000"/></a:schemeClr></a:solidFill></a:fillStyleLst>
      <a:lnStyleLst><a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>
        <a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="sysDash"/></a:ln>
        <a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>
      <a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>
      <a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/></a:schemeClr></a:solidFill></a:bgFillStyleLst>
    </a:fmtScheme></a:themeElements></a:theme>`;
  const shape = (id, spPr, style) => `<w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="457200"/><wp:docPr id="${id}"/><a:graphic>`
    + '<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp>'
    + `<wps:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${spPr}</wps:spPr><wps:style>${style}</wps:style></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
  const doc = DocxDocument.create();
  doc.addPart('word/styles.xml', encoder.encode(`<w:styles xmlns:w="${WORD_NS}"/>`), STYLES_TYPE);
  doc.addPart('word/theme/theme1.xml', encoder.encode(theme), THEME_TYPE);
  doc.setPartXml(doc.mainDocumentPath, `<w:document xmlns:w="${WORD_NS}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="${A_NS}" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><w:body><w:p>${spPrs.map((spPr, index) => shape(index + 1, spPr, styles[index])).join('')}</w:p><w:sectPr/></w:body></w:document>`);
  return doc.getShapes();
}

// Word 插入「矩形」时写的样式：线条取主题第 2 条、填充取第 1 条，颜色是 accent1 的不同深浅。
const WORD_DEFAULT_STYLE = '<a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="50000"/></a:schemeClr></a:lnRef>'
  + '<a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>'
  + '<a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef>';

test('fillRef and lnRef resolve through the theme format scheme, substituting the placeholder colour', () => {
  const [plain, gradient, darkened, background, none, widened, noLine] = themedShapes([
    '', '', '', '', '', '<a:ln w="38100"/>', '<a:ln><a:noFill/></a:ln>',
  ], [
    WORD_DEFAULT_STYLE,
    WORD_DEFAULT_STYLE.replace('fillRef idx="1"', 'fillRef idx="2"'),
    WORD_DEFAULT_STYLE.replace('fillRef idx="1"', 'fillRef idx="3"'),
    WORD_DEFAULT_STYLE.replace('fillRef idx="1"', 'fillRef idx="1001"'),
    WORD_DEFAULT_STYLE.replace('fillRef idx="1"', 'fillRef idx="0"').replace('lnRef idx="2"', 'lnRef idx="0"'),
    WORD_DEFAULT_STYLE,
    WORD_DEFAULT_STYLE,
  ]);
  assert.deepEqual(plain.fill, { type: 'solid', color: '#4472C4' });
  // lnStyleLst 第 2 条：12700 EMU = 1pt = 4/3 px，sysDash；颜色是占位色（accent1 shade 50%）。
  assert.ok(Math.abs(plain.line.widthPx - 4 / 3) < 0.001);
  assert.equal(plain.line.dash, '4 2');
  assert.equal(plain.line.color, color('<a:schemeClr val="accent1"><a:shade val="50000"/></a:schemeClr>').replace(/^/, '#'));

  assert.equal(gradient.fill.type, 'gradient', 'fillStyleLst 第 2 条是渐变');
  assert.equal(gradient.fill.stops.length, 2);
  assert.notEqual(gradient.fill.stops[0].color, gradient.fill.stops[1].color, '两个停止点各自的变换都套上了');
  assert.equal(gradient.fill.angle, 90);
  // 表里的变换叠在占位色之上：accent1 再 shade 50%。
  assert.equal(darkened.fill.color, `#${color('<a:schemeClr val="accent1"><a:shade val="50000"/></a:schemeClr>')}`);
  assert.equal(background.fill.color, `#${color('<a:schemeClr val="accent1"><a:tint val="95000"/></a:schemeClr>')}`, '1001 起指向背景填充表');
  assert.deepEqual(none.fill, { type: 'none' }, 'fillRef idx 0 是没有填充');
  assert.equal(none.line, undefined, 'lnRef idx 0 是没有主题线条');
  // spPr 里的 a:ln 只写了宽度：宽度用它的，颜色和线型仍来自主题线条。
  assert.ok(Math.abs(widened.line.widthPx - 4) < 0.001);
  assert.equal(widened.line.dash, '4 2');
  assert.equal(widened.line.color, plain.line.color);
  // 显式 noFill 的线条就是没有线，主题线条不能透上来。
  assert.deepEqual(noLine.line, {});
});

test('every supported preset geometry yields a finite path, and open ones render unfilled', () => {
  const names = supportedPresetGeometries();
  assert.ok(names.length >= 100, `${names.length} presets`);
  for (const name of names) {
    for (const [w, h] of [[200, 100], [50, 300], [1, 1]]) {
      const path = presetGeometryPath(name, w, h);
      assert.ok(path && !/NaN|undefined|Infinity/.test(path), `${name} ${w}x${h}: ${path}`);
    }
  }
  for (const name of ['line', 'arc', 'leftBracket', 'bracePair', 'bentConnector3']) assert.equal(presetGeometryIsOpen(name), true, name);
  for (const name of ['rect', 'chord', 'pie', 'blockArc', 'wave']) assert.equal(presetGeometryIsOpen(name), false, name);
});

test('the README preset count is the code\'s count', async () => {
  const { readFileSync } = await import('node:fs');
  const readme = readFileSync(new URL('../README.md', import.meta.url), 'utf8');
  const count = Number(readme.match(/形状 SVG 渲染支持 (\d+) 种 DrawingML 预设几何/)?.[1]);
  assert.equal(count, supportedPresetGeometries().length);
});
