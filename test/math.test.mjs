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
  // CT_Rad 里 radPr 是可选的,degHide 的默认值又正是 0,所以有 m:deg 时不必写出 radPr。
  assert.deepEqual(childNames(root.firstChild), ['deg', 'e']);
  const phantom = mathMlToOmml({ tag: 'mphantom', children: [{ tag: 'mi', text: 'x' }] }, owner);
  assert.deepEqual(childNames(phantom.firstChild), ['e']);
});

test('writing back what the reader produced keeps the OMML element structure', () => {
  // MathML 的一个标签对应多个 OMML 元素(mover 可能来自 bar / acc / groupChr / limUpp,
  // mrow 可能来自 d / func / box / nary,mtable 可能来自 m / eqArr),所以读取侧把原元素名
  // 记在 MathMlNode.source 上,写入侧照它还原。少了 source 就只能按标签猜,下面这些结构会
  // 被静静换成另一种。
  //
  // 比较的是 OMML 元素序列,其中相邻的文本被并起来看:读取侧按 MathML 语义把 "i=1" 拆成
  // mi / mo / mn 三个节点,写回就是三个 m:r。这不丢信息、渲染一致,但 run 的粒度确实会变粗,
  // 所以这条用例只钉元素结构,不钉 run 的切分。
  const run = (text) => `<m:r><m:t>${text}</m:t></m:r>`;
  const sequence = (element, out = []) => {
    for (const child of Array.from(element.childNodes)) {
      if (child.nodeType !== 1) continue;
      if (child.localName === 't') {
        if (typeof out.at(-1) === 'object') out.at(-1).text += child.textContent;
        else out.push({ text: child.textContent });
        continue;
      }
      if (child.localName === 'r') { sequence(child, out); continue; }
      out.push(child.localName);
      sequence(child, out);
    }
    return out.map((entry) => (typeof entry === 'object' ? `t:${entry.text}` : entry)).join('>');
  };
  const cases = [
    ['fraction', `<m:f><m:num>${run('x')}</m:num><m:den>${run('2')}</m:den></m:f>`],
    ['radical', `<m:rad><m:e>${run('y')}</m:e></m:rad>`],
    ['indexed radical', `<m:rad><m:deg>${run('3')}</m:deg><m:e>${run('y')}</m:e></m:rad>`],
    ['superscript', `<m:sSup><m:e>${run('a')}</m:e><m:sup>${run('2')}</m:sup></m:sSup>`],
    ['prescripts', `<m:sPre><m:sub>${run('92')}</m:sub><m:sup>${run('238')}</m:sup><m:e>${run('U')}</m:e></m:sPre>`],
    ['n-ary', `<m:nary><m:naryPr><m:chr m:val="∑"/></m:naryPr><m:sub>${run('i')}</m:sub><m:sup>${run('n')}</m:sup><m:e>${run('a')}</m:e></m:nary>`],
    // 容器(m:sub)里装多个节点:读取侧按语义把 "i=1" 拆成 mi/mo/mn 再凑成一个 mrow,
    // 写回时这层 mrow 必须摊开,不能凭空多出一层 m:box。
    ['n-ary with a compound limit', `<m:nary><m:naryPr><m:chr m:val="∑"/></m:naryPr><m:sub>${run('i=1')}</m:sub><m:sup>${run('n')}</m:sup><m:e>${run('a')}</m:e></m:nary>`],
    ['n-ary without an upper limit', `<m:nary><m:naryPr><m:chr m:val="∑"/></m:naryPr><m:sub>${run('i')}</m:sub><m:e>${run('a')}</m:e></m:nary>`],
    ['n-ary with inline limits', `<m:nary><m:naryPr><m:chr m:val="∫"/><m:limLoc m:val="subSup"/></m:naryPr><m:sub>${run('0')}</m:sub><m:sup>${run('1')}</m:sup><m:e>${run('f')}</m:e></m:nary>`],
    ['delimiters', `<m:d><m:e>${run('a')}</m:e><m:e>${run('b')}</m:e></m:d>`],
    ['delimiters with custom characters', `<m:d><m:dPr><m:begChr m:val="["/><m:endChr m:val="]"/></m:dPr><m:e>${run('a')}</m:e></m:d>`],
    ['function', `<m:func><m:fName>${run('sin')}</m:fName><m:e>${run('x')}</m:e></m:func>`],
    ['lower limit', `<m:limLow><m:e>${run('lim')}</m:e><m:lim>${run('n')}</m:lim></m:limLow>`],
    ['upper limit', `<m:limUpp><m:e>${run('max')}</m:e><m:lim>${run('k')}</m:lim></m:limUpp>`],
    ['accent', `<m:acc><m:e>${run('x')}</m:e></m:acc>`],
    ['accent with a character', `<m:acc><m:accPr><m:chr m:val="~"/></m:accPr><m:e>${run('x')}</m:e></m:acc>`],
    ['bar', `<m:bar><m:e>${run('x')}</m:e></m:bar>`],
    ['bar below', `<m:bar><m:barPr><m:pos m:val="bot"/></m:barPr><m:e>${run('x')}</m:e></m:bar>`],
    ['group character', `<m:groupChr><m:e>${run('x')}</m:e></m:groupChr>`],
    ['group character below', `<m:groupChr><m:groupChrPr><m:chr m:val="⏟"/><m:pos m:val="bot"/></m:groupChrPr><m:e>${run('x')}</m:e></m:groupChr>`],
    ['matrix', `<m:m><m:mr><m:e>${run('a')}</m:e><m:e>${run('b')}</m:e></m:mr><m:mr><m:e>${run('c')}</m:e><m:e>${run('d')}</m:e></m:mr></m:m>`],
    ['equation array', `<m:eqArr><m:e>${run('a')}</m:e><m:e>${run('b')}</m:e></m:eqArr>`],
    ['enclosure', `<m:borderBox><m:e>${run('x')}</m:e></m:borderBox>`],
    ['phantom', `<m:phant><m:e>${run('x')}</m:e></m:phant>`],
    ['box', `<m:box><m:e>${run('x')}</m:e></m:box>`],
    ['box around a bar', `<m:box><m:e><m:bar><m:e>${run('x')}</m:e></m:bar></m:e></m:box>`],
    ['box around several runs', `<m:box><m:e>${run('a')}${run('b')}</m:e></m:box>`],
  ];
  const owner = new DOMParser().parseFromString('<root/>', 'text/xml');
  for (const [name, xml] of cases) {
    const original = parse(xml);
    const before = sequence(original);
    const written = mathMlToOmml(ommlToMathMl(original), owner);
    assert.equal(sequence(written), before, name);
  }
});

test('reading a matrix gives one cell per m:e, not one cell per row', () => {
  // 一行里的 m:e 并成一个 mtd 的话,2x2 矩阵会读成 2x1,而且写回时丢掉的列再也补不回来。
  const run = (text) => `<m:r><m:t>${text}</m:t></m:r>`;
  const table = ommlToMathMl(parse(
    `<m:m><m:mr><m:e>${run('a')}</m:e><m:e>${run('b')}</m:e></m:mr>`
    + `<m:mr><m:e>${run('c')}</m:e><m:e>${run('d')}</m:e></m:mr></m:m>`)).children[0];
  assert.equal(table.tag, 'mtable');
  assert.deepEqual(table.children.map((row) => row.children.length), [2, 2]);
  // m:eqArr 的每个 m:e 本身就是一整行,所以那边一行一个单元格才是对的。
  const array = ommlToMathMl(parse(`<m:eqArr><m:e>${run('a')}</m:e><m:e>${run('b')}</m:e></m:eqArr>`)).children[0];
  assert.deepEqual(array.children.map((row) => row.children.length), [1, 1]);
});

test('a hand-built row keeps the shapes that only the heuristics can recognise', () => {
  // 不带 source 的 mrow 平时会被摊进外层容器(省掉一层 m:box),但 m:d 和 m:nary 只能靠形状
  // 认出来:首尾可伸缩的 mo,以及 munderover 后面跟着底数。摊开就把这两种结构变没了,
  // 而手工构造节点(setMath / insertMath 收的就是 MathML)正是唯一能走到这两条启发式的路。
  const owner = new DOMParser().parseFromString('<root/>', 'text/xml');
  const names = (element) => Array.from(element.childNodes)
    .filter((child) => child.nodeType === 1).map((child) => child.localName);
  const stretchy = (text) => ({ tag: 'mo', attrs: { stretchy: 'true' }, text });
  const fraction = mathMlToOmml({ tag: 'mfrac', children: [
    { tag: 'mrow', children: [stretchy('('), { tag: 'mi', text: 'a' }, stretchy(')')] },
    { tag: 'mn', text: '2' },
  ] }, owner);
  const numerator = fraction.getElementsByTagNameNS(ns, 'num')[0];
  assert.deepEqual(names(numerator), ['d'], 'numerator keeps the delimiter group');
  const summed = mathMlToOmml({ tag: 'mfrac', children: [
    { tag: 'mrow', children: [
      { tag: 'munderover', children: [{ tag: 'mo', text: '∑' }, { tag: 'mi', text: 'i' }, { tag: 'mi', text: 'n' }] },
      { tag: 'mi', text: 'a' },
    ] },
    { tag: 'mn', text: '2' },
  ] }, owner);
  // 摊开会让 munderover 单独走 nary 分支、body 空着,底数 a 掉成一个兄弟 run——
  // 和 #122 丢掉求和操作数是同一类。所以钉的是 m:nary 的 m:e 里确实装着底数。
  const operators = summed.getElementsByTagNameNS(ns, 'nary');
  assert.equal(operators.length, 1, 'numerator keeps exactly one n-ary operator');
  const body = Array.from(operators[0].childNodes).filter((child) => child.localName === 'e');
  assert.equal(body.length, 1);
  assert.equal(body[0].textContent, 'a', 'the n-ary operator keeps its operand');
});

test('linear math round-trips: what getMath() writes, linearToMathMl() reads back to the same text', async () => {
  const { linearToMathMl, ommlToLinearText, ommlToMathMl } = await import('../dist/math.js');
  const { parseXml } = await import('../dist/xml.js');
  const M = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
  const r = (text) => `<m:r><m:t>${text}</m:t></m:r>`;
  const omml = (inner) => parseXml(`<m:oMath xmlns:m="${M}">${inner}</m:oMath>`).documentElement;
  for (const [inner, expected] of [
    // 分母是多个记号：原先写成 (a+b)/c−d，读回来是 (a+b)/c 再减 d。
    [`<m:f><m:num>${r('a+b')}</m:num><m:den>${r('c−d')}</m:den></m:f>`, '(a+b)/(c−d)'],
    // 上标是多个记号：原先写成 x^2n，读回来是 x² 再乘 n。
    [`<m:sSup><m:e>${r('x')}</m:e><m:sup>${r('2n')}</m:sup></m:sSup>`, 'x^(2n)'],
    [`<m:sSup><m:e><m:d><m:e>${r('a+b')}</m:e></m:d></m:e><m:sup>${r('2')}</m:sup></m:sSup>`, '(a+b)^2'],
    [`<m:f><m:num>${r('1')}</m:num><m:den>${r('2')}</m:den></m:f>`, '1/2'],
  ]) {
    const text = ommlToLinearText(omml(inner));
    assert.equal(text, expected);
    // 读回来再写一次，文字不变：对话框里不改直接保存是空操作。
    const reread = linearToMathMl(text);
    const again = ommlToLinearText(omml(inner));
    assert.equal(again, text);
    assert.equal(reread.children[0].tag, ommlToMathMl(omml(inner)).children[0].tag, `${text} 的结构`);
  }
});

test('linear math accepts grouped fraction operands, brackets and multi-token scripts', async () => {
  const { linearToMathMl } = await import('../dist/math.js');
  const fraction = linearToMathMl('(a+b)/(c-d)').children[0];
  assert.equal(fraction.tag, 'mfrac');
  // 分子分母外面那层括号只是分组，不画出来（UnicodeMath 的约定）。
  assert.ok(!JSON.stringify(fraction).includes('"("'));
  const power = linearToMathMl('(a+b)^2').children[0];
  assert.equal(power.tag, 'msup');
  assert.ok(JSON.stringify(power.children[0]).includes('"("'), '作为上下标的底，括号是要画的');
  assert.equal(linearToMathMl('x^(2n)').children[0].children[1].tag, 'mrow');
  assert.doesNotThrow(() => linearToMathMl('x=(-b±√(b^2-4ac))/(2a)'));
  assert.deepEqual(linearToMathMl('[a,b]').children.map((node) => node.text), ['[', 'a', ',', 'b', ']']);
  assert.throws(() => linearToMathMl('\\frac{a}{b}'), /Unsupported linear math/);
});
