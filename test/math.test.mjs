import { test } from 'node:test';
import assert from 'node:assert/strict';
import { DOMParser } from '@xmldom/xmldom';
import { ommlToLinearText, ommlToMathMl } from '../dist/math.js';
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

test('limits recursive conversion depth', () => {
  let nested = '<m:r><m:t>x</m:t></m:r>';
  for (let i = 0; i < 100; i++) nested = `<m:f><m:num>${nested}</m:num><m:den><m:r><m:t>2</m:t></m:r></m:den></m:f>`;
  const formula = parse(nested);
  assert.doesNotThrow(() => ommlToMathMl(formula));
  assert.match(ommlToLinearText(formula), /…/);
});

test('math converter has no browser DOM dependency', () => {
  const source = readFileSync(new URL('../src/math.ts', import.meta.url), 'utf8');
  assert.doesNotMatch(source, /\b(?:document|window)\b/);
});
