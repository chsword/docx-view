import type { AgentRequest, ParagraphFormat, RunFormat } from './types.js';
import { assertText } from './xml.js';

function object(value: unknown): asserts value is Record<string, unknown> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new Error('Expected a JSON object.');
  }
}

function keys(value: Record<string, unknown>, allowed: string[]): void {
  if (Object.keys(value).some((key) => !allowed.includes(key))) {
    throw new Error('Unknown operation or format property.');
  }
}

export function assertIndex(value: unknown): asserts value is number {
  if (!Number.isSafeInteger(value) || (value as number) < 0) {
    throw new Error('Index/revision must be a non-negative safe integer.');
  }
}

export function validateRunFormat(value: unknown): asserts value is RunFormat {
  object(value);
  keys(value, [
    'style', 'bold', 'italic', 'underline', 'underlineStyle', 'underlineColor', 'fontSize', 'fontFamily',
    'fontFamilyEastAsia', 'color', 'strike', 'doubleStrike', 'verticalAlign', 'smallCaps', 'allCaps',
    'highlight', 'characterSpacing',
  ]);
  for (const key of ['bold', 'italic', 'underline', 'strike', 'doubleStrike', 'smallCaps', 'allCaps']) {
    if (key in value && typeof value[key] !== 'boolean') throw new Error(`${key} must be boolean.`);
  }
  for (const key of ['style', 'fontFamily', 'fontFamilyEastAsia', 'underlineStyle', 'highlight']) {
    if (key in value) assertText(value[key], key);
  }
  for (const key of ['color', 'underlineColor']) {
    if (key in value && (typeof value[key] !== 'string' || !/^[a-f\d]{6}$/i.test(value[key] as string))) {
      throw new Error(`${key} must be six hexadecimal digits without #.`);
    }
  }
  if ('fontSize' in value && (typeof value.fontSize !== 'number' ||
      !Number.isFinite(value.fontSize) || value.fontSize < 1 || value.fontSize > 400 ||
      !Number.isInteger(value.fontSize * 2))) {
    throw new Error('fontSize must be 1–400 points in half-point increments.');
  }
  if ('verticalAlign' in value && !['baseline', 'subscript', 'superscript'].includes(String(value.verticalAlign))) {
    throw new Error('verticalAlign must be baseline, subscript or superscript.');
  }
  if ('characterSpacing' in value && (!Number.isSafeInteger(value.characterSpacing as number) || Math.abs(value.characterSpacing as number) > 31680)) {
    throw new Error('characterSpacing must be a safe integer within OOXML spacing bounds.');
  }
}

export function validateParagraphFormat(value: unknown): asserts value is ParagraphFormat {
  object(value);
  keys(value, [
    'alignment', 'style', 'indentLeft', 'indentRight', 'indentFirstLine', 'indentHanging', 'spacingBefore',
    'spacingAfter', 'lineSpacing', 'lineSpacingRule', 'keepNext', 'keepLines', 'pageBreakBefore',
    'widowControl', 'outlineLevel',
  ]);
  if ('alignment' in value && !['left', 'center', 'right', 'both', 'distribute'].includes(String(value.alignment))) {
    throw new Error('Invalid paragraph alignment.');
  }
  if ('style' in value) assertText(value.style, 'style');
  for (const key of ['indentLeft', 'indentRight', 'indentFirstLine', 'indentHanging', 'spacingBefore', 'spacingAfter', 'lineSpacing', 'outlineLevel']) {
    if (key in value && (!Number.isSafeInteger(value[key]) || Math.abs(value[key] as number) > 31680)) {
      throw new Error(`${key} must be a safe integer within OOXML bounds.`);
    }
  }
  if ('lineSpacingRule' in value && !['auto', 'atLeast', 'exact'].includes(String(value.lineSpacingRule))) {
    throw new Error('Invalid lineSpacingRule.');
  }
  for (const key of ['keepNext', 'keepLines', 'pageBreakBefore', 'widowControl']) {
    if (key in value && typeof value[key] !== 'boolean') throw new Error(`${key} must be boolean.`);
  }
}

export function validateRows(rows: unknown): asserts rows is string[][] {
  if (!Array.isArray(rows) || rows.length === 0 || rows.length > 1000 ||
      rows.some((row) => !Array.isArray(row) || row.length === 0 || row.length > 100) ||
      rows.reduce((total, row) => total + row.length, 0) > 10_000) {
    throw new Error('Table must contain 1–1000 rows, 1–100 cells per row, and at most 10,000 cells.');
  }
  rows.forEach((row) => row.forEach((text: unknown) => assertText(text)));
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
      case 'formatRun':
        keys(op, ['type', 'paragraph', 'run', 'format']);
        assertIndex(op.paragraph); assertIndex(op.run); validateRunFormat(op.format); break;
      case 'replaceText':
        keys(op, ['type', 'search', 'replacement']); assertText(op.search, 'search'); assertText(op.replacement, 'replacement');
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
const signedInteger = { type: 'integer', minimum: -31680, maximum: 31680 };
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
            alignment: { enum: ['left', 'center', 'right', 'both', 'distribute'] },
            style: text,
            indentLeft: signedInteger,
            indentRight: signedInteger,
            indentFirstLine: signedInteger,
            indentHanging: signedInteger,
            spacingBefore: signedInteger,
            spacingAfter: signedInteger,
            lineSpacing: signedInteger,
            lineSpacingRule: { enum: ['auto', 'atLeast', 'exact'] },
            keepNext: { type: 'boolean' },
            keepLines: { type: 'boolean' },
            pageBreakBefore: { type: 'boolean' },
            widowControl: { type: 'boolean' },
            outlineLevel: signedInteger,
          }, []) }),
          operation('formatRun', { paragraph: index, run: index, format: shape({
            style: text,
            bold: { type: 'boolean' },
            italic: { type: 'boolean' },
            underline: { type: 'boolean' },
            underlineStyle: text,
            underlineColor: { type: 'string', pattern: '^[a-fA-F0-9]{6}$' },
            fontSize: { type: 'number', minimum: 1, maximum: 400, multipleOf: 0.5 },
            fontFamily: text,
            fontFamilyEastAsia: text,
            color: { type: 'string', pattern: '^[a-fA-F0-9]{6}$' },
            strike: { type: 'boolean' },
            doubleStrike: { type: 'boolean' },
            verticalAlign: { enum: ['baseline', 'subscript', 'superscript'] },
            smallCaps: { type: 'boolean' },
            allCaps: { type: 'boolean' },
            highlight: text,
            characterSpacing: signedInteger,
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
