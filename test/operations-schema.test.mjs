import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { AGENT_OPERATION_SCHEMA } from '../dist/operations.js';

test('operation schema and README list stay aligned', () => {
  const operationTypes = AGENT_OPERATION_SCHEMA.properties.operations.items.oneOf
    .map((operation) => operation.properties.type.const);
  assert.equal(operationTypes.length, 70);
  assert.equal(new Set(operationTypes).size, operationTypes.length);

  const readme = readFileSync(new URL('../README.md', import.meta.url), 'utf8');
  const line = readme.match(/^支持的操作类型（按 `AGENT_OPERATION_SCHEMA` 顺序）：(.+)$/m)?.[1];
  assert.ok(line);
  const documentedTypes = [...line.matchAll(/`([^`]+)`/g)].map((match) => match[1]);
  assert.deepEqual(documentedTypes, operationTypes);
});
