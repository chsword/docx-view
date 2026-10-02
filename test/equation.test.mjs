import { test } from 'node:test';
import assert from 'node:assert/strict';
import { equationToMathMl, parseEquationInstruction } from '../dist/fields.js';

const mathMl = (instruction) => equationToMathMl(parseEquationInstruction(instruction));
const tags = (node) => node.children ? [node.tag, node.children.map(tags)] : node.text !== undefined ? `${node.tag}:${node.text}` : node.tag;

test('EQ parser keeps numeric values per switch and bracket characters', () => {
  assert.deepEqual(parseEquationInstruction(' EQ \\a\\al\\co2\\hs3(1,2,3,4) ').values, { co: 2, hs: 3 });
  assert.deepEqual(parseEquationInstruction(' EQ \\b\\lc\\{(x) ').characters, { lc: '{' });
  assert.deepEqual(parseEquationInstruction(' EQ \\b \\bc\\[ (x) ').characters, { bc: '[' });
  assert.deepEqual(parseEquationInstruction(' EQ \\i\\fc\\∮(a,b,f) ').characters, { fc: '∮' });
  // 转义的逗号是文字本身，不是分隔符。
  assert.deepEqual(parseEquationInstruction(' EQ \\l(a\\,b,c) ').parts.map((part) => part.text), ['a,b', 'c']);
  // 一个参数里几个元素并排：原先只读第一个，后面的一声不响地丢掉。
  assert.deepEqual(tags(mathMl(' EQ x\\s\\up8(2) ')), ['mrow', ['mtext:x', ['mpadded', ['mtext:2']]]]);
  assert.deepEqual(tags(mathMl(' EQ \\f(a\\r(2),3) ')), ['mfrac', [['mrow', ['mtext:a', ['msqrt', ['mtext:2']]]], 'mtext:3']]);
  // 原有的 \s\up / \s\do 位移量不受影响。
  assert.equal(parseEquationInstruction(' EQ \\s\\do5(x) ').raisePoints, -5);
});

test('EQ switches map to the MathML a browser lays out', () => {
  assert.deepEqual(tags(mathMl(' EQ \\f(1,2) ')), ['mfrac', ['mtext:1', 'mtext:2']]);
  assert.deepEqual(tags(mathMl(' EQ \\r(x) ')), ['msqrt', ['mtext:x']]);
  assert.deepEqual(tags(mathMl(' EQ \\r(3,x) ')), ['mroot', ['mtext:x', 'mtext:3']], '第一个参数是根指数');
  assert.deepEqual(tags(mathMl(' EQ \\i(0,1,x) ')), ['mrow', [['msubsup', ['mo:∫', 'mtext:0', 'mtext:1']], 'mtext:x']]);
  assert.deepEqual(tags(mathMl(' EQ \\i\\su(i=1,n,i) ')), ['mrow', [['munderover', ['mo:∑', 'mtext:i=1', 'mtext:n']], 'mtext:i']]);
  assert.deepEqual(tags(mathMl(' EQ \\i\\su\\in(i,n,i) ')).at(1)[0][0], 'msubsup', '\\in 把上下限放成行内');
  assert.deepEqual(tags(mathMl(' EQ \\i\\pr(a,b,c) ')).at(1)[0][1][0], 'mo:∏');
  assert.deepEqual(tags(mathMl(' EQ \\b(x) ')), ['mrow', ['mo:(', 'mtext:x', 'mo:)']]);
  assert.deepEqual(tags(mathMl(' EQ \\b\\bc\\{(x) ')), ['mrow', ['mo:{', 'mtext:x', 'mo:}']], '\\bc 两侧配对');
  assert.deepEqual(tags(mathMl(' EQ \\b\\lc\\[(x) ')), ['mrow', ['mo:[', 'mtext:x']], '只给左边时右边为空');
  assert.deepEqual(tags(mathMl(' EQ \\l(a,b,c) ')), ['mrow', ['mtext:a', 'mo:,', 'mtext:b', 'mo:,', 'mtext:c']]);

  const array = mathMl(' EQ \\a\\ar\\co2(1,2,3,4,5) ');
  assert.equal(array.tag, 'mtable');
  assert.equal(array.attrs.columnalign, 'right');
  assert.deepEqual(array.children.map((row) => row.children.length), [2, 2, 1]);

  assert.deepEqual(mathMl(' EQ \\s\\up8(2) ').attrs, { voffset: '8pt' });
  assert.deepEqual(mathMl(' EQ \\s\\do4(2) ').attrs, { voffset: '-4pt' });
  assert.match(mathMl(' EQ \\x(x) ').attrs.style, /border-top:1px solid;border-bottom:1px solid;border-left:1px solid;border-right:1px solid/);
  assert.match(mathMl(' EQ \\x\\bo(x) ').attrs.style, /^border-bottom:1px solid;padding/);
  assert.deepEqual(mathMl(' EQ \\d\\fo10() '), { tag: 'mspace', attrs: { width: '10pt' } });
  // 嵌套：分数里套根号。
  assert.deepEqual(tags(mathMl(' EQ \\f(\\r(2),2) ')), ['mfrac', [['msqrt', ['mtext:2']], 'mtext:2']]);
  // 不认识的开关退化为参数并排，内容不丢。
  assert.deepEqual(tags(mathMl(' EQ \\zz(a,b) ')), ['mrow', ['mtext:a', 'mtext:b']]);
});
