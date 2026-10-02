import { test } from 'node:test';
import assert from 'node:assert/strict';
import { styleFormValues, stylePatchFromForm, uniqueStyleId, withoutNulls } from '../examples/style-editor.ts';

const HEADING = {
  id: 'Heading1', name: 'heading 1', type: 'paragraph', basedOn: 'Normal', next: 'Normal', quickFormat: true,
  paragraph: { alignment: 'center', spacingBefore: 240, spacingAfter: 60, lineSpacing: 360, lineSpacingRule: 'auto', keepNext: true },
  run: { bold: true, fontSize: 16, color: '2F5496', fontFamily: 'Calibri Light' },
};

test('the style form shows the style\'s own values, in points and multiples', () => {
  const values = styleFormValues(HEADING);
  assert.deepEqual(values, {
    name: 'heading 1', basedOn: 'Normal', next: 'Normal', fontFamily: 'Calibri Light', fontSize: '16',
    bold: 'on', italic: '', underline: '', color: '2F5496', alignment: 'center',
    spacingBefore: '12', spacingAfter: '3', lineSpacing: '1.5', indentLeft: '', indentFirstLine: '',
    quickFormat: true,
  });
  // 固定行距不能用「倍数」表达：显示为空，且没动它就不会出现在补丁里。
  assert.equal(styleFormValues({ ...HEADING, paragraph: { lineSpacing: 400, lineSpacingRule: 'exact' } }).lineSpacing, '');
});

test('an untouched form yields an empty patch, so unknown style properties survive', () => {
  const values = styleFormValues(HEADING);
  assert.deepEqual(stylePatchFromForm(values, { ...values }, 'paragraph'), {});
});

test('changed fields become patch values, and cleared fields become null (inherit again)', () => {
  const initial = styleFormValues(HEADING);
  const patch = stylePatchFromForm(initial, {
    ...initial, name: ' Big ', bold: '', italic: 'off', color: '#ff0000', fontSize: '13.3',
    spacingBefore: '', lineSpacing: '2', indentFirstLine: '21', basedOn: '', quickFormat: false,
  }, 'paragraph');
  assert.deepEqual(patch, {
    name: 'Big',
    basedOn: null,
    quickFormat: false,
    run: { fontSize: 13.5, bold: null, italic: false, color: 'FF0000' },
    paragraph: { spacingBefore: null, indentFirstLine: 420, lineSpacing: 480, lineSpacingRule: 'auto' },
  });
  const cleared = stylePatchFromForm(initial, { ...initial, lineSpacing: '' }, 'paragraph');
  assert.deepEqual(cleared.paragraph, { lineSpacing: null, lineSpacingRule: null });
});

test('character styles never get paragraph properties or a next style', () => {
  const initial = styleFormValues({ id: 'Strong', name: 'Strong', type: 'character', run: { bold: true } });
  const patch = stylePatchFromForm(initial, { ...initial, alignment: 'right', next: 'Normal', spacingAfter: '6', bold: 'off' }, 'character');
  assert.deepEqual(patch, { run: { bold: false } });
});

test('invalid form input is reported before anything is written', () => {
  const initial = styleFormValues(HEADING);
  for (const [field, value, pattern] of [
    ['name', '   ', /不能为空/],
    ['fontSize', 'big', /字号/],
    ['fontSize', '500', /字号/],
    ['color', 'red', /十六进制/],
    ['spacingAfter', '-1', /段后/],
    ['lineSpacing', '0.1', /行距/],
  ]) {
    assert.throws(() => stylePatchFromForm(initial, { ...initial, [field]: value }, 'paragraph'), pattern, `${field}=${value}`);
  }
});

test('new style ids come from the name and never collide', () => {
  assert.equal(uniqueStyleId('My Quote', ['Normal']), 'MyQuote');
  assert.equal(uniqueStyleId('My Quote', ['myquote']), 'MyQuote1', 'Word 的样式 ID 比较不分大小写');
  assert.equal(uniqueStyleId('重点', ['Normal']), 'Style1', '没有 ASCII 字符时退回 StyleN');
  assert.equal(uniqueStyleId('提示', ['Style1', 'style2']), 'Style3');
});

test('withoutNulls drops cleared fields for defineStyle', () => {
  assert.deepEqual(withoutNulls({ bold: true, color: null }), { bold: true });
  assert.equal(withoutNulls({ color: null }), undefined);
  assert.equal(withoutNulls(undefined), undefined);
});
