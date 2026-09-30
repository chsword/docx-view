import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { AGENT_OPERATION_SCHEMA, validateRunFormat } from '../dist/operations.js';

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
