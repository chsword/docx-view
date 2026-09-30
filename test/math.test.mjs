import { test } from 'node:test';
import assert from 'node:assert/strict';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import {
  linearToMathMl, mathMlToOmml, ommlToLinearText, ommlToLinearTextWithInfo, ommlToMathMl, ommlToMathMlWithInfo,
} from '../dist/math.js';
import { readFileSync } from 'node:fs';

const ns = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
const parse = (xml) => new DOMParser().parseFromString(
  `<m:oMath xmlns:m="${ns}">${xml}</m:oMath>`, 'text/xml').documentElement;

test('converts common OMML structures to MathML data', () => {
  const formula = parse(`
    <m:f><m:num><m:r><m:t>x+1</m:t></m:r></m:num><m:den><m:r><m:t>2</m:t></m:r></m:den></m:f>
    <m:sSup><m:e><m:r><m:t>a</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup>
    <m:rad><m:e><m:r><m:t>y</m:t></m:r></m:e></m:rad>`);
  const math = ommlToMathMl(formula);
  assert.equal(math.tag, 'math');
  assert.equal(math.children[0].tag, 'mfrac');
  assert.equal(math.children[1].tag, 'msup');
  assert.equal(math.children[2].tag, 'msqrt');
  assert.equal(ommlToLinearText(formula), '(x+1)/2a^2√(y)');
});

test('supports delimiters and unknown elements without dropping text', () => {
  const formula = parse('<m:d><m:e><m:r><m:t>a</m:t></m:r></m:e><m:e><m:r><m:t>b</m:t></m:r></m:e></m:d><m:unknown><m:r><m:t>z</m:t></m:r></m:unknown>');
  assert.match(ommlToLinearText(formula), /^\(a,b\)z$/);
});

test('tokenizes mixed math text and recognizes Unicode operators', () => {
  const formula = parse('<m:r><m:t>x+y≠z∈∞</m:t></m:r>');
  const nodes = ommlToMathMl(formula).children;
  assert.deepEqual(nodes.map(({ tag, text }) => [tag, text]), [
    ['mi', 'x'], ['mo', '+'], ['mi', 'y'], ['mo', '≠'], ['mi', 'z'], ['mo', '∈'], ['mo', '∞'],
  ]);
});

test('preserves the n-ary body after limits', () => {
  const formula = parse('<m:nary><m:naryPr><m:limLoc m:val="undOvr"/></m:naryPr><m:sub><m:r><m:t>i=1</m:t></m:r></m:sub><m:sup><m:r><m:t>n</m:t></m:r></m:sup><m:e><m:r><m:t>a_i</m:t></m:r></m:e></m:nary>');
  const node = ommlToMathMl(formula).children[0];
  assert.equal(node.tag, 'mrow');
  assert.equal(node.children[0].tag, 'munderover');
  assert.equal(node.children[1].text, 'a_i');
});

test('renders sPre with preceding scripts and preserves empty placeholders', () => {
  const formula = parse('<m:sPre><m:e><m:r><m:t>U</m:t></m:r></m:e><m:sub><m:r><m:t>92</m:t></m:r></m:sub><m:sup><m:r><m:t>238</m:t></m:r></m:sup></m:sPre>');
  const node = ommlToMathMl(formula).children[0];
  assert.equal(node.tag, 'mmultiscripts');
  assert.deepEqual(node.children.map((child) => child.tag), ['mi', 'mprescripts', 'mn', 'mn']);
  assert.equal(ommlToLinearText(formula), '_92^238U');

  const missing = ommlToMathMl(parse('<m:sPre><m:e><m:r><m:t>U</m:t></m:r></m:e></m:sPre>')).children[0];
  assert.deepEqual(missing.children.map((child) => child.tag), ['mi', 'mprescripts', 'mrow', 'mrow']);
  assert.equal(ommlToLinearText(parse('<m:sPre><m:e><m:r><m:t>X</m:t></m:r></m:e><m:sub><m:r><m:t>a</m:t></m:r></m:sub></m:sPre>')), '_aX');
  assert.equal(ommlToLinearText(parse('<m:sPre><m:e><m:r><m:t>X</m:t></m:r></m:e><m:sup><m:r><m:t>b</m:t></m:r></m:sup></m:sPre>')), '^bX');
});

test('converts block, scripts, functions, limits, accents, tables, and boxes', () => {
  const formula = parse(`
    <m:oMathPara><m:jc m:val="centerGroup"/><m:r><m:t>x</m:t></m:r></m:oMathPara>
    <m:sSub><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sub><m:r><m:t>i</m:t></m:r></m:sub></m:sSub>
    <m:sSubSup><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sub><m:r><m:t>i</m:t></m:r></m:sub><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSubSup>
    <m:func><m:fName><m:r><m:t>sin</m:t></m:r></m:fName><m:e><m:r><m:t>x</m:t></m:r></m:e></m:func>
    <m:limLow><m:e><m:r><m:t>f</m:t></m:r></m:e><m:lim><m:r><m:t>0</m:t></m:r></m:lim></m:limLow>
    <m:limUpp><m:e><m:r><m:t>f</m:t></m:r></m:e><m:lim><m:r><m:t>∞</m:t></m:r></m:lim></m:limUpp>
    <m:acc><m:accPr><m:chr m:val="^"/></m:accPr><m:e><m:r><m:t>x</m:t></m:r></m:e></m:acc>
    <m:bar><m:e><m:r><m:t>x</m:t></m:r></m:e></m:bar>
    <m:groupChr><m:groupChrPr><m:pos m:val="bot"/></m:groupChrPr><m:e><m:r><m:t>x</m:t></m:r></m:e></m:groupChr>
    <m:m><m:mr><m:e><m:r><m:t>a</m:t></m:r></m:e></m:mr></m:m>
    <m:eqArr><m:e><m:r><m:t>b</m:t></m:r></m:e></m:eqArr>
    <m:box><m:e><m:r><m:t>c</m:t></m:r></m:e></m:box>
    <m:borderBox><m:e><m:r><m:t>d</m:t></m:r></m:e></m:borderBox>
    <m:phant><m:e><m:r><m:t>e</m:t></m:r></m:e></m:phant>`);
  const children = ommlToMathMl(formula).children;
  assert.equal(children[0].attrs.style, 'text-align:center');
  const tags = [];
  const visit = (node) => { tags.push(node.tag); for (const child of node.children ?? []) visit(child); };
  children.forEach(visit);
  for (const tag of ['msub', 'msubsup', 'mrow', 'munder', 'mover', 'mtable', 'menclose', 'mphantom']) {
    assert.ok(tags.includes(tag), tag);
  }
});

test('limits recursive conversion depth', () => {
  let nested = '<m:r><m:t>x</m:t></m:r>';
  for (let i = 0; i < 100; i++) nested = `<m:f><m:num>${nested}</m:num><m:den><m:r><m:t>2</m:t></m:r></m:den></m:f>`;
  const formula = parse(nested);
  assert.doesNotThrow(() => ommlToMathMl(formula));
  assert.equal(ommlToMathMlWithInfo(formula).truncated, true);
  assert.equal(ommlToLinearTextWithInfo(formula).truncated, true);
  assert.equal(ommlToLinearText(parse('<m:r><m:t>a…b</m:t></m:r>')), 'a…b');
});

test('math converter has no browser DOM dependency', () => {
  const source = readFileSync(new URL('../src/math.ts', import.meta.url), 'utf8');
  assert.doesNotMatch(source, /\b(?:document|window)\b/);
  assert.deepEqual([...source.matchAll(/from ['"]([^'"]+)['"]/g)].map((match) => match[1]),
    ['@xmldom/xmldom', './types.js']);
});

test('MathML-to-OMML round trips every structure emitted by the reader', () => {
  const cases = [
    ['token', '<m:r><m:t>x</m:t></m:r>'],
    ['row', '<m:box><m:r><m:t>x</m:t></m:r><m:r><m:t>+</m:t></m:r><m:r><m:t>y</m:t></m:r></m:box>'],
    ['fraction', '<m:f><m:num><m:r><m:t>x</m:t></m:r></m:num><m:den><m:r><m:t>2</m:t></m:r></m:den></m:f>'],
    ['no-bar fraction', '<m:f><m:fPr><m:type m:val="noBar"/></m:fPr><m:num><m:r><m:t>x</m:t></m:r></m:num><m:den><m:r><m:t>2</m:t></m:r></m:den></m:f>'],
    ['superscript', '<m:sSup><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup>'],
    ['subscript', '<m:sSub><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sub><m:r><m:t>i</m:t></m:r></m:sub></m:sSub>'],
    ['subscript and superscript', '<m:sSubSup><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sub><m:r><m:t>i</m:t></m:r></m:sub><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSubSup>'],
    ['square root', '<m:rad><m:e><m:r><m:t>x</m:t></m:r></m:e></m:rad>'],
    ['indexed root', '<m:rad><m:deg><m:r><m:t>3</m:t></m:r></m:deg><m:e><m:r><m:t>x</m:t></m:r></m:e></m:rad>'],
    ['n-ary limits', '<m:nary><m:sub><m:r><m:t>i</m:t></m:r></m:sub><m:sup><m:r><m:t>n</m:t></m:r></m:sup><m:e><m:r><m:t>x</m:t></m:r></m:e></m:nary>'],
    ['delimiters', '<m:d><m:e><m:r><m:t>a</m:t></m:r></m:e><m:e><m:r><m:t>b</m:t></m:r></m:e></m:d>'],
    ['lower limit', '<m:limLow><m:e><m:r><m:t>x</m:t></m:r></m:e><m:lim><m:r><m:t>i</m:t></m:r></m:lim></m:limLow>'],
    ['upper limit', '<m:limUpp><m:e><m:r><m:t>x</m:t></m:r></m:e><m:lim><m:r><m:t>i</m:t></m:r></m:lim></m:limUpp>'],
    ['accent', '<m:acc><m:accPr><m:chr m:val="^"/></m:accPr><m:e><m:r><m:t>x</m:t></m:r></m:e></m:acc>'],
    ['matrix', '<m:m><m:mr><m:e><m:r><m:t>a</m:t></m:r></m:e><m:e><m:r><m:t>b</m:t></m:r></m:e></m:mr></m:m>'],
    ['enclosure', '<m:borderBox><m:e><m:r><m:t>x</m:t></m:r></m:e></m:borderBox>'],
    ['phantom', '<m:phant><m:e><m:r><m:t>x</m:t></m:r></m:e></m:phant>'],
    ['prescripts', '<m:sPre><m:e><m:r><m:t>U</m:t></m:r></m:e><m:sub><m:r><m:t>92</m:t></m:r></m:sub><m:sup><m:r><m:t>238</m:t></m:r></m:sup></m:sPre>'],
  ];
  const owner = new DOMParser().parseFromString('<root/>', 'text/xml');
  for (const [name, xml] of cases) {
    const original = ommlToMathMl(parse(xml));
    const written = mathMlToOmml(original, owner);
    assert.deepEqual(ommlToMathMl(written), original, name);
  }
  const blockMath = {
    tag: 'math', attrs: { display: 'block' },
    children: [{ tag: 'mi', attrs: { mathvariant: 'italic' }, text: 'x' }],
  };
  assert.deepEqual(ommlToMathMl(mathMlToOmml(blockMath, owner)), blockMath);
});

test('linearToMathMl parses only the documented expression subset', () => {
  const cases = [
    ['a/b', 'mfrac', 'a/b'],
    ['a^b', 'msup', 'a^b'],
    ['a_b', 'msub', 'a_b'],
    ['a_b^c', 'msubsup', 'a_b^c'],
    ['√(a)', 'msqrt', '√(a)'],
    ['sqrt(a)', 'msqrt', '√(a)'],
    ['(a+b)', 'mrow', '(a+b)'],
    ['∑_(a)^(b) c', 'mrow', '∑_(a)^(b)c'],
    ['∫_(a)^(b) c', 'mrow', '∫_(a)^(b)c'],
  ];
  const owner = new DOMParser().parseFromString('<root/>', 'text/xml');
  for (const [source, tag, expectedLinear] of cases) {
    const tree = linearToMathMl(source);
    assert.equal(tree.children[0].tag, tag, source);
    assert.equal(ommlToLinearText(mathMlToOmml(tree, owner)), expectedLinear, source);
  }
  assert.throws(() => linearToMathMl('x \\\\alpha'), error =>
    error.message.includes('a/b') && error.message.includes('∑_(a)^(b) c'));
});

test('MathML writes reject unknown tags, attributes, invalid text, and excessive depth', () => {
  const owner = new DOMParser().parseFromString('<root/>', 'text/xml');
  assert.throws(() => mathMlToOmml({ tag: 'script', text: 'x' }, owner), /Unsupported MathML tag/);
  assert.throws(() => mathMlToOmml({ tag: 'mi', attrs: { onclick: 'alert(1)' }, text: 'x' }, owner),
    /Unsupported MathML attribute/);
  assert.throws(() => mathMlToOmml({ tag: 'math', text: 'ignored' }, owner), /cannot contain text/);
  assert.throws(() => mathMlToOmml({ tag: 'mi', text: '\u0000' }, owner), /valid XML text/);
  let nested = { tag: 'mi', text: 'x' };
  for (let index = 0; index < 66; index++) nested = { tag: 'mrow', children: [nested] };
  assert.throws(() => mathMlToOmml(nested, owner), /maximum depth/);
  const injected = mathMlToOmml({ tag: 'mi', text: '<w:evil/>&' }, owner);
  const xml = new XMLSerializer().serializeToString(injected);
  assert.match(xml, /&lt;w:evil\/&gt;&amp;/);
  assert.equal(injected.getElementsByTagNameNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'evil').length, 0);
  const text = mathMlToOmml({ tag: 'mtext', text: '<w:evil/>&' }, owner);
  assert.match(new XMLSerializer().serializeToString(text), /&lt;w:evil\/&gt;&amp;/);
});

test('generated OMML follows required child order and grouping containers', () => {
  const owner = new DOMParser().parseFromString('<root/>', 'text/xml');
  const childNames = (element) => Array.from(element.childNodes)
    .filter((child) => child.nodeType === 1)
    .map((child) => child.localName);
  const row = mathMlToOmml({ tag: 'mrow', children: [{ tag: 'mi', text: 'x' }] }, owner);
  assert.deepEqual(childNames(row.firstChild), ['e']);
  const root = mathMlToOmml({ tag: 'mroot', children: [
    { tag: 'mi', text: 'x' }, { tag: 'mn', text: '3' },
  ] }, owner);
  assert.deepEqual(childNames(root.firstChild), ['radPr', 'deg', 'e']);
  const phantom = mathMlToOmml({ tag: 'mphantom', children: [{ tag: 'mi', text: 'x' }] }, owner);
  assert.deepEqual(childNames(phantom.firstChild), ['e']);
});
