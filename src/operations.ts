import type { AgentRequest, ParagraphFormat, RunFormat } from './types.js';
import { assertText } from './xml.js';

function object(value: unknown): asserts value is Record<string, unknown> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new Error('Expected a JSON object.');
  }
}

function keys(value: Record<string, unknown>, allowed: string[]): void {
  if (Object.keys(value).some(key => !allowed.includes(key))) {
    throw new Error('Unknown operation or format property.');
  }
}

export function assertIndex(value: unknown): asserts value is number {
  if (!Number.isSafeInteger(value) || (value as number) < 0) {
    throw new Error('Index/revision must be a non-negative safe integer.');
  }
}

export function assertInteger(value: unknown, name = 'value'): asserts value is number {
  if (!Number.isSafeInteger(value)) throw new Error(`${name} must be a safe integer.`);
}

export function validateRunFormat(value: unknown): asserts value is RunFormat {
  object(value);
  keys(value, ['bold', 'italic', 'underline', 'fontSize', 'fontFamily', 'color']);
  for (const key of ['bold', 'italic', 'underline']) {
    if (key in value && typeof value[key] !== 'boolean') throw new Error(`${key} must be boolean.`);
  }
  if ('fontSize' in value && (typeof value.fontSize !== 'number' ||
      !Number.isFinite(value.fontSize) || value.fontSize < 1 || value.fontSize > 400 ||
      !Number.isInteger(value.fontSize * 2))) {
    throw new Error('fontSize must be 1–400 points in half-point increments.');
  }
  if ('fontFamily' in value) assertText(value.fontFamily, 'fontFamily');
  if ('color' in value && (typeof value.color !== 'string' || !/^[a-f\d]{6}$/i.test(value.color))) {
    throw new Error('color must be six hexadecimal digits without #.');
  }
}

export function validateParagraphFormat(value: unknown): asserts value is ParagraphFormat {
  object(value);
  keys(value, ['alignment', 'style']);
  if ('alignment' in value && !['left', 'center', 'right', 'both'].includes(String(value.alignment))) {
    throw new Error('Invalid paragraph alignment.');
  }
  if ('style' in value) assertText(value.style, 'style');
}

export function validateRows(rows: unknown): asserts rows is string[][] {
  if (!Array.isArray(rows) || rows.length === 0 || rows.length > 1000 ||
      rows.some(row => !Array.isArray(row) || row.length === 0 || row.length > 100) ||
      rows.reduce((total, row) => total + row.length, 0) > 10_000) {
    throw new Error('Table must contain 1–1000 rows, 1–100 cells per row, and at most 10,000 cells.');
  }
  rows.forEach(row => row.forEach((text: unknown) => assertText(text)));
}

export function validateRequest(value: unknown): asserts value is AgentRequest {
  object(value);
  keys(value, ['expectedRevision', 'operations']);
  if ('expectedRevision' in value) assertIndex(value.expectedRevision);
  if (!Array.isArray(value.operations) || value.operations.length > 1000) {
    throw new Error('operations must be an array of at most 1000 operations.');
  }
  for (const op of value.operations) {
    object(op);
    switch (op.type) {
      case 'setParagraphText':
        keys(op, ['type', 'index', 'text']); assertIndex(op.index); assertText(op.text); break;
      case 'insertParagraph':
        keys(op, ['type', 'text', 'before']); assertText(op.text);
        if ('before' in op) assertIndex(op.before);
        break;
      case 'deleteParagraph':
        keys(op, ['type', 'index']); assertIndex(op.index); break;
      case 'formatParagraph':
        keys(op, ['type', 'index', 'format']); assertIndex(op.index); validateParagraphFormat(op.format); break;
      case 'setParagraphNumbering':
        keys(op, ['type', 'index', 'numId', 'level']);
        assertIndex(op.index); assertIndex(op.numId);
        if (op.numId < 1) throw new Error('numId must be at least 1.');
        if ('level' in op) {
          assertIndex(op.level);
          if (op.level > 8) throw new Error('level must be between 0 and 8.');
        }
        break;
      case 'clearParagraphNumbering':
        keys(op, ['type', 'index']); assertIndex(op.index); break;
      case 'setParagraphLevel':
        keys(op, ['type', 'index', 'delta']); assertIndex(op.index); assertInteger(op.delta, 'delta'); break;
      case 'formatRun':
        keys(op, ['type', 'paragraph', 'run', 'format']);
        assertIndex(op.paragraph); assertIndex(op.run); validateRunFormat(op.format); break;
      case 'replaceText':
        keys(op, ['type', 'search', 'replacement']); assertText(op.search); assertText(op.replacement);
        if (!op.search) throw new Error('search must not be empty.');
        break;
      case 'insertTable':
        keys(op, ['type', 'rows']); validateRows(op.rows); break;
      case 'setPartXml':
        keys(op, ['type', 'path', 'xml']); assertText(op.path, 'path'); assertText(op.xml, 'xml'); break;
      default: throw new Error(`Unknown operation type: ${String(op.type)}`);
    }
  }
}

const text = { type: 'string', maxLength: 1_000_000 };
const index = { type: 'integer', minimum: 0, maximum: Number.MAX_SAFE_INTEGER };
const integer = { type: 'integer', minimum: -Number.MAX_SAFE_INTEGER, maximum: Number.MAX_SAFE_INTEGER };
const shape = (properties: Record<string, unknown>, required = Object.keys(properties)) => ({
  type: 'object', properties, required, additionalProperties: false,
});
const operation = (type: string, properties: Record<string, unknown>, required = Object.keys(properties)) =>
  shape({ type: { const: type }, ...properties }, ['type', ...required]);

/** JSON Schema for tool/function calling; requests are also validated at runtime. */
export const AGENT_OPERATION_SCHEMA = {
  $schema: 'http://json-schema.org/draft-07/schema#',
  title: 'DocxViewOperations',
  description: 'Atomic DOCX edits. Paragraph/run indices refer to the current state before each operation.',
  ...shape({
    expectedRevision: index,
    operations: {
      type: 'array', maxItems: 1000,
      items: {
        oneOf: [
          operation('setParagraphText', { index, text }),
          operation('insertParagraph', { text, before: index }, ['text']),
          operation('deleteParagraph', { index }),
          operation('formatParagraph', { index, format: shape({
            alignment: { enum: ['left', 'center', 'right', 'both'] }, style: text,
          }, []) }),
          operation('setParagraphNumbering', { index, numId: { ...index, minimum: 1 }, level: { ...index, maximum: 8 } }, ['index', 'numId']),
          operation('clearParagraphNumbering', { index }),
          operation('setParagraphLevel', { index, delta: integer }),
          operation('formatRun', { paragraph: index, run: index, format: shape({
            bold: { type: 'boolean' }, italic: { type: 'boolean' }, underline: { type: 'boolean' },
            fontSize: { type: 'number', minimum: 1, maximum: 400, multipleOf: 0.5 },
            fontFamily: text, color: { type: 'string', pattern: '^[a-fA-F0-9]{6}$' },
          }, []) }),
          operation('replaceText', { search: { ...text, minLength: 1 }, replacement: text }),
          operation('insertTable', { rows: {
            type: 'array', minItems: 1, maxItems: 1000,
            items: { type: 'array', minItems: 1, maxItems: 100, items: text },
          } }),
          operation('setPartXml', { path: text, xml: text }),
        ],
      },
    },
  }, ['operations']),
} as const;
