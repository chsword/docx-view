import type { AgentRequest, BorderSide, ParagraphFormat, RunFormat, Shading, TabStop } from './types.js';
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

export function validateRunFormat(value: unknown): asserts value is RunFormat {
  object(value);
  keys(value, ['bold', 'italic', 'underline', 'fontSize', 'fontFamily', 'color', 'border', 'shading']);
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
  if ('border' in value && value.border !== null) validateBorderSide(value.border);
  if ('shading' in value && value.shading !== null) validateShading(value.shading);
}

export function validateParagraphFormat(value: unknown): asserts value is ParagraphFormat {
  object(value);
  keys(value, ['alignment', 'style', 'tabs', 'borders', 'shading',
    'keepNext', 'keepLines', 'pageBreakBefore', 'widowControl', 'suppressLineNumbers', 'suppressAutoHyphens']);
  if ('alignment' in value && !['left', 'center', 'right', 'both'].includes(String(value.alignment))) {
    throw new Error('Invalid paragraph alignment.');
  }
  if ('style' in value) assertText(value.style, 'style');
  if ('tabs' in value && value.tabs !== null) validateTabs(value.tabs);
  if ('borders' in value && value.borders !== null) validateBorders(value.borders);
  if ('shading' in value && value.shading !== null) validateShading(value.shading);
  for (const key of ['keepNext', 'keepLines', 'pageBreakBefore', 'widowControl', 'suppressLineNumbers', 'suppressAutoHyphens']) {
    if (key in value && typeof value[key] !== 'boolean') throw new Error(`${key} must be boolean.`);
  }
}

export function validateTabStop(value: unknown): asserts value is TabStop {
  object(value);
  keys(value, ['position', 'alignment', 'leader']);
  if (typeof value.position !== 'number' || !Number.isFinite(value.position)) throw new Error('tab.position must be a finite number.');
  if (!['left', 'center', 'right', 'decimal', 'bar', 'clear', 'num'].includes(String(value.alignment))) {
    throw new Error('tab.alignment is invalid.');
  }
  if ('leader' in value && (typeof value.leader !== 'string' || !['none', 'dot', 'hyphen', 'underscore', 'heavy', 'middleDot'].includes(value.leader))) {
    throw new Error('tab.leader is invalid.');
  }
}

export function validateTabs(value: unknown): asserts value is TabStop[] {
  if (!Array.isArray(value)) throw new Error('tabs must be an array.');
  if (value.length > 200) throw new Error('tabs must contain at most 200 items.');
  value.forEach(validateTabStop);
}

export function validateBorderSide(value: unknown): asserts value is BorderSide {
  object(value);
  keys(value, ['style', 'size', 'space', 'color', 'shadow']);
  assertText(value.style, 'border.style');
  if (typeof value.size !== 'number' || !Number.isFinite(value.size) || value.size < 0) throw new Error('border.size must be a non-negative number.');
  if (typeof value.space !== 'number' || !Number.isFinite(value.space) || value.space < 0) throw new Error('border.space must be a non-negative number.');
  if (typeof value.color !== 'string' || !/^(auto|[a-f\d]{6})$/i.test(value.color)) throw new Error('border.color must be auto or six hexadecimal digits.');
  if ('shadow' in value && typeof value.shadow !== 'boolean') throw new Error('border.shadow must be boolean.');
}

export function validateBorders(value: unknown): asserts value is ParagraphFormat['borders'] {
  object(value);
  keys(value, ['top', 'left', 'bottom', 'right', 'between', 'bar']);
  for (const key of ['top', 'left', 'bottom', 'right', 'between', 'bar'] as const) {
    if (key in value && value[key]) validateBorderSide(value[key]);
  }
}

export function validateShading(value: unknown): asserts value is Shading {
  object(value);
  keys(value, ['pattern', 'fill', 'color']);
  assertText(value.pattern, 'shading.pattern');
  if (typeof value.fill !== 'string' || !/^(auto|[a-f\d]{6})$/i.test(value.fill)) throw new Error('shading.fill must be auto or six hexadecimal digits.');
  if ('color' in value && (typeof value.color !== 'string' || !/^(auto|[a-f\d]{6})$/i.test(value.color))) {
    throw new Error('shading.color must be auto or six hexadecimal digits.');
  }
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
      case 'formatRun':
        keys(op, ['type', 'paragraph', 'run', 'format']);
        assertIndex(op.paragraph); assertIndex(op.run); validateRunFormat(op.format); break;
      case 'replaceText':
        keys(op, ['type', 'search', 'replacement']); assertText(op.search); assertText(op.replacement);
        if (!op.search) throw new Error('search must not be empty.');
        break;
      case 'setParagraphTabs':
        keys(op, ['type', 'index', 'tabs']); assertIndex(op.index); validateTabs(op.tabs); break;
      case 'setParagraphBorders':
        keys(op, ['type', 'index', 'borders']); assertIndex(op.index); validateBorders(op.borders); break;
      case 'setParagraphShading':
        keys(op, ['type', 'index', 'shading']); assertIndex(op.index); validateShading(op.shading); break;
      case 'insertBreak':
        keys(op, ['type', 'paragraph', 'run', 'breakType']);
        assertIndex(op.paragraph); assertIndex(op.run);
        if (!['textWrapping', 'page', 'column'].includes(String(op.breakType))) throw new Error('Invalid break type.');
        break;
      case 'insertSymbol':
        keys(op, ['type', 'paragraph', 'run', 'font', 'charCode']);
        assertIndex(op.paragraph); assertIndex(op.run); assertText(op.font, 'font');
        if (typeof op.charCode !== 'number' || !Number.isSafeInteger(op.charCode) || op.charCode < 0 || op.charCode > 0xffff) {
          throw new Error('charCode must be an integer in [0, 65535].');
        }
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
            tabs: {
              type: 'array', maxItems: 200, items: shape({
                position: { type: 'number' },
                alignment: { enum: ['left', 'center', 'right', 'decimal', 'bar', 'clear', 'num'] },
                leader: { enum: ['none', 'dot', 'hyphen', 'underscore', 'heavy', 'middleDot'] },
              }, ['position', 'alignment']),
            },
            borders: shape({
              top: shape({ style: text, size: { type: 'number', minimum: 0 }, space: { type: 'number', minimum: 0 }, color: { type: 'string', pattern: '^(auto|[a-fA-F0-9]{6})$' }, shadow: { type: 'boolean' } }, []),
              left: shape({ style: text, size: { type: 'number', minimum: 0 }, space: { type: 'number', minimum: 0 }, color: { type: 'string', pattern: '^(auto|[a-fA-F0-9]{6})$' }, shadow: { type: 'boolean' } }, []),
              bottom: shape({ style: text, size: { type: 'number', minimum: 0 }, space: { type: 'number', minimum: 0 }, color: { type: 'string', pattern: '^(auto|[a-fA-F0-9]{6})$' }, shadow: { type: 'boolean' } }, []),
              right: shape({ style: text, size: { type: 'number', minimum: 0 }, space: { type: 'number', minimum: 0 }, color: { type: 'string', pattern: '^(auto|[a-fA-F0-9]{6})$' }, shadow: { type: 'boolean' } }, []),
              between: shape({ style: text, size: { type: 'number', minimum: 0 }, space: { type: 'number', minimum: 0 }, color: { type: 'string', pattern: '^(auto|[a-fA-F0-9]{6})$' }, shadow: { type: 'boolean' } }, []),
              bar: shape({ style: text, size: { type: 'number', minimum: 0 }, space: { type: 'number', minimum: 0 }, color: { type: 'string', pattern: '^(auto|[a-fA-F0-9]{6})$' }, shadow: { type: 'boolean' } }, []),
            }, []),
            shading: shape({
              pattern: text, fill: { type: 'string', pattern: '^(auto|[a-fA-F0-9]{6})$' }, color: { type: 'string', pattern: '^(auto|[a-fA-F0-9]{6})$' },
            }, ['pattern', 'fill']),
            keepNext: { type: 'boolean' }, keepLines: { type: 'boolean' }, pageBreakBefore: { type: 'boolean' },
            widowControl: { type: 'boolean' }, suppressLineNumbers: { type: 'boolean' }, suppressAutoHyphens: { type: 'boolean' },
          }, []) }),
          operation('formatRun', { paragraph: index, run: index, format: shape({
            bold: { type: 'boolean' }, italic: { type: 'boolean' }, underline: { type: 'boolean' },
            fontSize: { type: 'number', minimum: 1, maximum: 400, multipleOf: 0.5 },
            fontFamily: text, color: { type: 'string', pattern: '^[a-fA-F0-9]{6}$' },
            border: shape({
              style: text, size: { type: 'number', minimum: 0 }, space: { type: 'number', minimum: 0 },
              color: { type: 'string', pattern: '^(auto|[a-fA-F0-9]{6})$' }, shadow: { type: 'boolean' },
            }),
            shading: shape({
              pattern: text, fill: { type: 'string', pattern: '^(auto|[a-fA-F0-9]{6})$' }, color: { type: 'string', pattern: '^(auto|[a-fA-F0-9]{6})$' },
            }, ['pattern', 'fill']),
          }, []) }),
          operation('setParagraphTabs', { index, tabs: {
            type: 'array', maxItems: 200, items: shape({
              position: { type: 'number' },
              alignment: { enum: ['left', 'center', 'right', 'decimal', 'bar', 'clear', 'num'] },
              leader: { enum: ['none', 'dot', 'hyphen', 'underscore', 'heavy', 'middleDot'] },
            }, ['position', 'alignment']),
          } }),
          operation('setParagraphBorders', { index, borders: shape({
            top: shape({ style: text, size: { type: 'number', minimum: 0 }, space: { type: 'number', minimum: 0 }, color: { type: 'string', pattern: '^(auto|[a-fA-F0-9]{6})$' }, shadow: { type: 'boolean' } }, []),
            left: shape({ style: text, size: { type: 'number', minimum: 0 }, space: { type: 'number', minimum: 0 }, color: { type: 'string', pattern: '^(auto|[a-fA-F0-9]{6})$' }, shadow: { type: 'boolean' } }, []),
            bottom: shape({ style: text, size: { type: 'number', minimum: 0 }, space: { type: 'number', minimum: 0 }, color: { type: 'string', pattern: '^(auto|[a-fA-F0-9]{6})$' }, shadow: { type: 'boolean' } }, []),
            right: shape({ style: text, size: { type: 'number', minimum: 0 }, space: { type: 'number', minimum: 0 }, color: { type: 'string', pattern: '^(auto|[a-fA-F0-9]{6})$' }, shadow: { type: 'boolean' } }, []),
            between: shape({ style: text, size: { type: 'number', minimum: 0 }, space: { type: 'number', minimum: 0 }, color: { type: 'string', pattern: '^(auto|[a-fA-F0-9]{6})$' }, shadow: { type: 'boolean' } }, []),
            bar: shape({ style: text, size: { type: 'number', minimum: 0 }, space: { type: 'number', minimum: 0 }, color: { type: 'string', pattern: '^(auto|[a-fA-F0-9]{6})$' }, shadow: { type: 'boolean' } }, []),
          }, []) }),
          operation('setParagraphShading', { index, shading: shape({
            pattern: text, fill: { type: 'string', pattern: '^(auto|[a-fA-F0-9]{6})$' }, color: { type: 'string', pattern: '^(auto|[a-fA-F0-9]{6})$' },
          }, ['pattern', 'fill']) }),
          operation('insertBreak', { paragraph: index, run: index, breakType: { enum: ['textWrapping', 'page', 'column'] } }),
          operation('insertSymbol', { paragraph: index, run: index, font: text, charCode: { type: 'integer', minimum: 0, maximum: 65535 } }),
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
