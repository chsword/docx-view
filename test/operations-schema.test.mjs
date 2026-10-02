import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import {
  AGENT_OPERATION_SCHEMA, CELL_FORMAT_KEYS, EAST_ASIAN_LAYOUT_KEYS, PARAGRAPH_FORMAT_KEYS, PARAGRAPH_FRAME_KEYS,
  ROW_FORMAT_KEYS, RUN_FORMAT_FIELDS, TABLE_FORMAT_KEYS,
  validateCellFormat, validateEastAsianLayout, validateParagraphFormat, validateParagraphFrame, validateRowFormat,
  validateRunFormat, validateTableFormat,
} from '../dist/operations.js';

test('operation schema and README list stay aligned', () => {
  const operationTypes = AGENT_OPERATION_SCHEMA.properties.operations.items.oneOf
    .map((operation) => operation.properties.type.const);
  assert.equal(operationTypes.length, 72);
  assert.equal(new Set(operationTypes).size, operationTypes.length);

  const readme = readFileSync(new URL('../README.md', import.meta.url), 'utf8');
  const line = readme.match(/^支持的操作类型（按 `AGENT_OPERATION_SCHEMA` 顺序）：(.+)$/m)?.[1];
  assert.ok(line);
  const documentedTypes = [...line.matchAll(/`([^`]+)`/g)].map((match) => match[1]);
  assert.deepEqual(documentedTypes, operationTypes);
});

test('run text-effect fields have matching schema and runtime validation', () => {
  const fields = [
    'position', 'characterScale', 'kerning', 'fitTextWidth', 'textEffect',
    'textOutline', 'textShadow', 'emboss', 'imprint',
  ];
  const operationSchemas = AGENT_OPERATION_SCHEMA.properties.operations.items.oneOf
    .filter((operation) => ['formatRun', 'formatRange', 'formatDocumentRange'].includes(operation.properties.type.const));
  assert.equal(operationSchemas.length, 3);
  for (const operation of operationSchemas) {
    const formatSchema = operation.properties.format.properties;
    for (const field of fields) assert.ok(field in formatSchema, `${operation.properties.type.const}.${field}`);
    assert.deepEqual(formatSchema.position.anyOf[0], { type: 'integer', minimum: -32768, maximum: 32767 });
    assert.deepEqual(formatSchema.characterScale.anyOf[0], { type: 'integer', minimum: 1, maximum: 600 });
    assert.deepEqual(formatSchema.kerning.anyOf[0], { type: 'integer', minimum: 0, maximum: 32767 });
    assert.deepEqual(formatSchema.fitTextWidth.anyOf[0], { type: 'integer', minimum: 0, maximum: 31680 });
    for (const field of ['textOutline', 'textShadow', 'emboss', 'imprint']) {
      assert.deepEqual(formatSchema[field].anyOf[0], { type: 'boolean' });
    }
    assert.equal(formatSchema.textEffect.anyOf[0].type, 'string');
    assert.equal(formatSchema.textEffect.anyOf[0].maxLength, 1_000_000);
  }
  assert.doesNotThrow(() => validateRunFormat({
    position: -32768, characterScale: 600, kerning: 32767, fitTextWidth: 31680, textEffect: 'sparkle',
    textOutline: false, textShadow: true, emboss: false, imprint: true,
  }));
  for (const format of [
    { position: 32768 }, { characterScale: 601 }, { kerning: 32768 }, { fitTextWidth: 31681 },
    { textOutline: 'yes' }, { textEffect: 1 },
  ]) assert.throws(() => validateRunFormat(format));
});

test('every format field the validators accept is declared in the agent schema', () => {
  // 栽过三次：floatingPosition、行的 grid*、段落的 frame 都被运行时校验收下、却没进 schema。
  // schema 比校验窄，等于这些字段对 agent 不存在，而两边都是绿的——所以这条用例按名字那一份
  // 真相（operations.ts 导出的四张名单）核对 schema 的属性集，两个方向都查。
  const operations = AGENT_OPERATION_SCHEMA.properties.operations.items.oneOf;
  const formatProperties = (name) => operations.find((operation) => operation.properties.type.const === name)
    .properties.format.properties;
  for (const [operation, names, validate] of [
    // run 格式有三个入口，schema 原先各抄一份；eastAsianLayout 就是只进了校验、三份 schema 都没有。
    ['formatRun', RUN_FORMAT_FIELDS, validateRunFormat],
    ['formatRange', RUN_FORMAT_FIELDS, validateRunFormat],
    ['formatDocumentRange', RUN_FORMAT_FIELDS, validateRunFormat],
    ['formatParagraph', PARAGRAPH_FORMAT_KEYS, validateParagraphFormat],
    ['formatTable', TABLE_FORMAT_KEYS, validateTableFormat],
    ['formatTableRow', ROW_FORMAT_KEYS, validateRowFormat],
    ['formatCell', CELL_FORMAT_KEYS, validateCellFormat],
  ]) {
    const declared = Object.keys(formatProperties(operation)).sort();
    assert.deepEqual(declared, [...names].sort(), `${operation} schema does not match its key list`);
    // 名单本身也得是校验真认的名字，不能只是一张对不上实现的字符串表。这里只问「名字认不认」，
    // 所以不需要给每个字段都造一个合法取值——值不合法照样会抛，但抛的不是「不认识这个属性」。
    for (const name of names) {
      try {
        validate({ [name]: null });
      } catch (error) {
        assert.ok(!/Unknown operation or format property/.test(String(error.message)),
          `${operation} does not know the field ${name}`);
      }
    }
    assert.throws(() => validate({ definitelyNotAField: 1 }), /definitelyNotAField/,
      `${operation} must name the unknown property it rejects`);
  }
});

test('the framePr schema is derived from the frame validator, not copied', () => {
  const operations = AGENT_OPERATION_SCHEMA.properties.operations.items.oneOf;
  const frame = operations.find((operation) => operation.properties.type.const === 'formatParagraph')
    .properties.format.properties.frame.anyOf[0];
  // 属性集按名字那一份真相核对。只抽查几个字段是不够的：把派生换成手抄的名单、少抄两项
  // （horizontalSpaceTwips / verticalSpaceTwips），抽查式的断言一条都不会红。
  assert.deepEqual(Object.keys(frame.properties).sort(), [...PARAGRAPH_FRAME_KEYS].sort());
  // 首字下沉走的就是 framePr，agent 拿不到这个属性就做不了首字下沉。
  assert.deepEqual(frame.properties.dropCap.enum, ['none', 'drop', 'margin']);
  assert.equal(frame.properties.lines.minimum, 0);
  // 坐标可负（挪到页边距外），尺寸不可负 —— 与 validateParagraphFrame 同一条判断。
  assert.equal(frame.properties.xTwips.minimum, -31680);
  assert.equal(frame.properties.widthTwips.minimum, 0);
  assert.equal(frame.additionalProperties, false);
  // 校验与 schema 共用 PARAGRAPH_FRAME_KEYS，所以收的名字也得对得上：schema 声明
  // additionalProperties: false，校验就不能悄悄多收一个。
  assert.throws(() => validateParagraphFrame({ notAFrameField: 1 }), /notAFrameField/);
  for (const name of PARAGRAPH_FRAME_KEYS) assert.ok(name in frame.properties, name);
});

test('the eastAsianLayout schema is derived from its validator, not copied', () => {
  const operations = AGENT_OPERATION_SCHEMA.properties.operations.items.oneOf;
  const layout = operations.find((operation) => operation.properties.type.const === 'formatRun')
    .properties.format.properties.eastAsianLayout.anyOf[0];
  assert.deepEqual(Object.keys(layout.properties).sort(), [...EAST_ASIAN_LAYOUT_KEYS].sort());
  assert.deepEqual(layout.properties.combineBrackets.enum, ['none', 'round', 'square', 'angle', 'curly']);
  assert.equal(layout.properties.id.maximum, 0xffff);
  assert.equal(layout.additionalProperties, false);
  assert.throws(() => validateEastAsianLayout({ notALayoutField: 1 }), /notALayoutField/);
});
