import { test } from 'node:test';
import assert from 'node:assert/strict';
import { DOMParser } from '@xmldom/xmldom';
import { ommlToLinearText, ommlToLinearTextWithInfo, ommlToMathMl, ommlToMathMlWithInfo } from '../dist/math.js';
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
