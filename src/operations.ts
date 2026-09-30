import type {
  AgentRequest, BorderSide, CellFormat, EditableRegionEditorGroup, ParagraphFormat, RowFormat, RunFormat, Shading,
  TableFormat, TabStop,
} from './types.js';
import { assertBase64 } from './drawing.js';
import { assertText, isValidXmlCharCode } from './xml.js';
import { assertHyperlinkInput } from './hyperlink.js';
import { fieldKindFromInstruction, NEVER_EVALUATE } from './fields.js';

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

function maybeNull<T>(value: T | null | undefined, validate: (value: T) => void): void {
  if (value !== null && value !== undefined) validate(value);
}

const RUN_FORMAT_FIELDS = [
  'style', 'bold', 'italic', 'hidden', 'webHidden', 'underline', 'underlineStyle', 'underlineColor', 'fontSize', 'fontFamily',
  'fontFamilyEastAsia', 'color', 'strike', 'doubleStrike', 'verticalAlign', 'smallCaps', 'allCaps',
  'rtl', 'complexScript', 'highlight', 'characterSpacing', 'border', 'shading',
] as const;

function validateTextRange(value: unknown): asserts value is { paragraph: number; start: number; end: number } {
  object(value);
  keys(value, ['paragraph', 'start', 'end']);
  assertIndex(value.paragraph);
  assertIndex(value.start);
  assertIndex(value.end);
  if (value.end < value.start) throw new Error('range.end must be greater than or equal to range.start.');
}

function validateDocumentRange(value: unknown): asserts value is {
  start: { paragraph: number; offset: number };
  end: { paragraph: number; offset: number };
} {
  object(value);
  keys(value, ['start', 'end']);
  object(value.start);
  keys(value.start, ['paragraph', 'offset']);
  object(value.end);
  keys(value.end, ['paragraph', 'offset']);
  assertIndex(value.start.paragraph);
  assertIndex(value.start.offset);
  assertIndex(value.end.paragraph);
  assertIndex(value.end.offset);
  if (value.start.paragraph > value.end.paragraph ||
      (value.start.paragraph === value.end.paragraph && value.start.offset > value.end.offset)) {
    throw new Error('range.start must not be after range.end.');
  }
}

function validateStyleApplyOptions(value: unknown): asserts value is { clearDirectFormat?: boolean } {
  object(value);
  keys(value, ['clearDirectFormat']);
  if ('clearDirectFormat' in value && value.clearDirectFormat !== undefined && typeof value.clearDirectFormat !== 'boolean') {
    throw new Error('options.clearDirectFormat must be boolean.');
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
  keys(value, [
    'style', 'bold', 'italic', 'underline', 'underlineStyle', 'underlineColor', 'fontSize', 'fontFamily',
    'fontFamilyEastAsia', 'color', 'strike', 'doubleStrike', 'verticalAlign', 'smallCaps', 'allCaps',
    'rtl', 'complexScript', 'highlight', 'characterSpacing', 'border', 'shading',
  ]);
  for (const key of ['bold', 'italic', 'hidden', 'webHidden', 'underline', 'strike', 'doubleStrike', 'smallCaps', 'allCaps', 'rtl', 'complexScript']) {
    if (key in value && value[key] !== null && typeof value[key] !== 'boolean') throw new Error(`${key} must be boolean.`);
  }
  for (const key of ['style', 'fontFamily', 'fontFamilyEastAsia', 'underlineStyle', 'highlight']) {
    if (key in value) maybeNull(value[key] as string | null | undefined, (entry) => assertText(entry, key));
  }
  for (const key of ['color', 'underlineColor']) {
    if (key in value && value[key] !== null && (typeof value[key] !== 'string' || !/^[a-f\d]{6}$/i.test(value[key] as string))) {
      throw new Error(`${key} must be six hexadecimal digits without #.`);
    }
  }
  if ('fontSize' in value && value.fontSize !== null && (typeof value.fontSize !== 'number' ||
      !Number.isFinite(value.fontSize) || value.fontSize < 1 || value.fontSize > 400 ||
      !Number.isInteger(value.fontSize * 2))) {
    throw new Error('fontSize must be 1–400 points in half-point increments.');
  }
  if ('verticalAlign' in value && value.verticalAlign !== null && !['baseline', 'subscript', 'superscript'].includes(String(value.verticalAlign))) {
    throw new Error('verticalAlign must be baseline, subscript or superscript.');
  }
  if ('characterSpacing' in value && value.characterSpacing !== null &&
      (!Number.isSafeInteger(value.characterSpacing as number) || Math.abs(value.characterSpacing as number) > 31680)) {
    throw new Error('characterSpacing must be a safe integer within OOXML spacing bounds.');
  }
  if ('border' in value) maybeNull(value.border as BorderSide | null | undefined, (entry) => validateBorderSide(entry));
  if ('shading' in value) maybeNull(value.shading as Shading | null | undefined, (entry) => validateDocShading(entry));
}

export function validateParagraphFormat(value: unknown): asserts value is ParagraphFormat {
  object(value);
  keys(value, [
    'alignment', 'style', 'indentLeft', 'indentRight', 'indentFirstLine', 'indentHanging', 'spacingBefore',
    'spacingAfter', 'lineSpacing', 'lineSpacingRule', 'keepNext', 'keepLines', 'pageBreakBefore',
    'widowControl', 'outlineLevel', 'tabs', 'borders', 'shading', 'suppressLineNumbers', 'suppressAutoHyphens',
    'kinsoku', 'wordWrap', 'overflowPunct', 'topLinePunct', 'autoSpaceDE', 'autoSpaceDN', 'bidi', 'textDirection',
  ]);
  if ('alignment' in value && value.alignment !== null && !['left', 'center', 'right', 'both', 'distribute'].includes(String(value.alignment))) {
    throw new Error('Invalid paragraph alignment.');
  }
  if ('style' in value) maybeNull(value.style as string | null | undefined, (entry) => assertText(entry, 'style'));
  for (const key of ['indentLeft', 'indentRight', 'lineSpacing']) {
    if (key in value && value[key] !== null && (!Number.isSafeInteger(value[key]) || Math.abs(value[key] as number) > 31680)) {
      throw new Error(`${key} must be a safe integer within OOXML bounds.`);
    }
  }
  for (const key of ['indentFirstLine', 'indentHanging', 'spacingBefore', 'spacingAfter']) {
    if (key in value && value[key] !== null && (!Number.isSafeInteger(value[key]) || (value[key] as number) < 0 || (value[key] as number) > 31680)) {
      throw new Error(`${key} must be an unsigned twips value within OOXML bounds.`);
    }
  }
  if ('outlineLevel' in value && value.outlineLevel !== null &&
      (!Number.isSafeInteger(value.outlineLevel as number) || (value.outlineLevel as number) < 0 || (value.outlineLevel as number) > 9)) {
    throw new Error('outlineLevel must be an integer from 0 to 9.');
  }
  if ('lineSpacingRule' in value && value.lineSpacingRule !== null && !['auto', 'atLeast', 'exact'].includes(String(value.lineSpacingRule))) {
    throw new Error('Invalid lineSpacingRule.');
  }
  for (const key of [
    'keepNext', 'keepLines', 'pageBreakBefore', 'widowControl', 'suppressLineNumbers', 'suppressAutoHyphens',
    'kinsoku', 'wordWrap', 'overflowPunct', 'topLinePunct', 'autoSpaceDE', 'autoSpaceDN', 'bidi',
  ]) {
    if (key in value && value[key] !== null && typeof value[key] !== 'boolean') throw new Error(`${key} must be boolean.`);
  }
  if ('textDirection' in value) maybeNull(value.textDirection as string | null | undefined, (entry) => assertText(entry, 'textDirection'));
  if ('tabs' in value) maybeNull(value.tabs as TabStop[] | null | undefined, (entry) => validateTabs(entry));
  if ('borders' in value) maybeNull(value.borders as ParagraphFormat['borders'] | null | undefined, (entry) => validateParagraphBorders(entry));
  if ('shading' in value) maybeNull(value.shading as Shading | null | undefined, (entry) => validateDocShading(entry, 'shading'));
}

export function validateTabStop(value: unknown): asserts value is TabStop {
  object(value);
  keys(value, ['position', 'alignment', 'leader']);
  if (typeof value.position !== 'number' || !Number.isFinite(value.position)) throw new Error('tab.position must be a finite number.');
  if (!['left', 'center', 'right', 'decimal', 'bar', 'clear', 'num'].includes(String(value.alignment))) {
    throw new Error('tab.alignment is invalid.');
  }
  if ('leader' in value && value.leader !== undefined &&
      (typeof value.leader !== 'string' || !['none', 'dot', 'hyphen', 'underscore', 'heavy', 'middleDot'].includes(value.leader))) {
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
  if ('shadow' in value && value.shadow !== undefined && typeof value.shadow !== 'boolean') throw new Error('border.shadow must be boolean.');
}

export function validateParagraphBorders(value: unknown): asserts value is ParagraphFormat['borders'] {
  object(value as Record<string, unknown>);
  keys(value as Record<string, unknown>, ['top', 'left', 'bottom', 'right', 'between', 'bar']);
  for (const key of ['top', 'left', 'bottom', 'right', 'between', 'bar'] as const) {
    if (key in (value as Record<string, unknown>) && (value as Record<string, unknown>)[key] !== undefined) {
      validateBorderSide((value as Record<string, unknown>)[key]);
    }
  }
}

export function validateDocShading(value: unknown, name = 'shading'): asserts value is Shading {
  object(value as Record<string, unknown>);
  const shading = value as Record<string, unknown>;
  keys(shading, ['pattern', 'fill', 'color']);
  assertText(shading.pattern, `${name}.pattern`);
  if (typeof shading.fill !== 'string' || !/^(auto|[a-f\d]{6})$/i.test(shading.fill)) {
    throw new Error(`${name}.fill must be auto or six hexadecimal digits.`);
  }
  if ('color' in shading && shading.color !== undefined && (typeof shading.color !== 'string' || !/^(auto|[a-f\d]{6})$/i.test(shading.color))) {
    throw new Error(`${name}.color must be auto or six hexadecimal digits.`);
  }
}

function validateWidth(value: unknown, name: string): void {
  object(value as Record<string, unknown>);
  const width = value as Record<string, unknown>;
  keys(width, ['type', 'value']);
  if (!['auto', 'dxa', 'pct'].includes(String(width.type))) throw new Error(`${name}.type must be auto, dxa or pct.`);
  if (!Number.isFinite(width.value) || typeof width.value !== 'number' || width.value < 0) {
    throw new Error(`${name}.value must be a non-negative number.`);
  }
}

function validateBorder(value: unknown, name: string): void {
  object(value as Record<string, unknown>);
  const border = value as Record<string, unknown>;
  keys(border, ['style', 'size', 'space', 'color', 'none']);
  if ('style' in border && typeof border.style !== 'string') throw new Error(`${name}.style must be a string.`);
  for (const key of ['size', 'space']) {
    if (key in border && (typeof border[key] !== 'number' || !Number.isFinite(border[key]) || border[key] < 0)) {
      throw new Error(`${name}.${key} must be a non-negative number.`);
    }
  }
  if ('color' in border && (typeof border.color !== 'string' || !/^(auto|[a-f\d]{6})$/i.test(border.color))) {
    throw new Error(`${name}.color must be auto or six hexadecimal digits.`);
  }
  if ('none' in border && typeof border.none !== 'boolean') throw new Error(`${name}.none must be boolean.`);
}

function validateBorders(value: unknown, name: string): void {
  object(value as Record<string, unknown>);
  const borders = value as Record<string, unknown>;
  keys(borders, ['top', 'right', 'bottom', 'left', 'insideH', 'insideV']);
  for (const side of ['top', 'right', 'bottom', 'left', 'insideH', 'insideV']) {
    if (side in borders) validateBorder(borders[side], `${name}.${side}`);
  }
}

function validateShading(value: unknown, name: string): void {
  object(value as Record<string, unknown>);
  const shading = value as Record<string, unknown>;
  keys(shading, ['fill', 'color', 'value']);
  for (const key of ['fill', 'color']) {
    if (key in shading && (typeof shading[key] !== 'string' || !/^(auto|[a-f\d]{6})$/i.test(String(shading[key])))) {
      throw new Error(`${name}.${key} must be auto or six hexadecimal digits.`);
    }
  }
  if ('value' in shading && typeof shading.value !== 'string') throw new Error(`${name}.value must be a string.`);
}

function validateMargins(value: unknown, name: string): void {
  object(value as Record<string, unknown>);
  const margin = value as Record<string, unknown>;
  keys(margin, ['top', 'right', 'bottom', 'left']);
  for (const side of ['top', 'right', 'bottom', 'left']) if (side in margin) validateWidth(margin[side], `${name}.${side}`);
}

export function validateTableFormat(value: unknown): asserts value is TableFormat {
  object(value);
  keys(value, ['width', 'alignment', 'indent', 'borders', 'shading', 'cellMargin', 'layout', 'style', 'look', 'caption', 'description', 'bidiVisual']);
  if ('width' in value) validateWidth(value.width, 'width');
  if ('alignment' in value && !['left', 'center', 'right'].includes(String(value.alignment))) throw new Error('Invalid table alignment.');
  if ('indent' in value && (typeof value.indent !== 'number' || !Number.isFinite(value.indent) || value.indent < 0)) throw new Error('indent must be a non-negative number.');
  if ('borders' in value) validateBorders(value.borders, 'borders');
  if ('shading' in value) validateShading(value.shading, 'shading');
  if ('cellMargin' in value) validateMargins(value.cellMargin, 'cellMargin');
  if ('layout' in value && !['fixed', 'autofit'].includes(String(value.layout))) throw new Error('Invalid table layout.');
  if ('bidiVisual' in value && value.bidiVisual !== null && typeof value.bidiVisual !== 'boolean') throw new Error('bidiVisual must be boolean.');
  for (const key of ['style', 'look', 'caption', 'description'] as const) if (key in value) assertText(value[key], key);
}

export function validateRowFormat(value: unknown): asserts value is RowFormat {
  object(value);
  keys(value, ['height', 'cantSplit', 'header', 'alignment', 'deleted', 'inserted', 'revision']);
  if ('height' in value) {
    object(value.height);
    const height = value.height as Record<string, unknown>;
    keys(height, ['value', 'rule']);
    if (typeof height.value !== 'number' || !Number.isFinite(height.value) || height.value < 0) throw new Error('height.value must be a non-negative number.');
    if ('rule' in height && !['atLeast', 'exact'].includes(String(height.rule))) throw new Error('height.rule must be atLeast or exact.');
  }
  for (const key of ['cantSplit', 'header', 'deleted', 'inserted'] as const) {
    if (key in value && typeof value[key] !== 'boolean') throw new Error(`${key} must be boolean.`);
  }
  if ('alignment' in value && !['left', 'center', 'right'].includes(String(value.alignment))) throw new Error('Invalid row alignment.');
  if ('revision' in value) {
    object(value.revision);
    const revision = value.revision as Record<string, unknown>;
    keys(revision, ['author', 'date']);
    if ('author' in revision) assertText(revision.author, 'revision.author');
    if ('date' in revision) assertText(revision.date, 'revision.date');
  }
}

export function validateCellFormat(value: unknown): asserts value is CellFormat {
  object(value);
  keys(value, ['width', 'borders', 'shading', 'margin', 'verticalAlign', 'textDirection', 'noWrap', 'hideMark', 'hMerge', 'vMerge']);
  if ('width' in value) validateWidth(value.width, 'width');
  if ('borders' in value) validateBorders(value.borders, 'borders');
  if ('shading' in value) validateShading(value.shading, 'shading');
  if ('margin' in value) validateMargins(value.margin, 'margin');
  if ('verticalAlign' in value && !['top', 'center', 'bottom'].includes(String(value.verticalAlign))) throw new Error('Invalid verticalAlign.');
  if ('textDirection' in value) assertText(value.textDirection, 'textDirection');
  for (const key of ['noWrap', 'hideMark'] as const) {
    if (key in value && typeof value[key] !== 'boolean') throw new Error(`${key} must be boolean.`);
  }
  for (const key of ['hMerge', 'vMerge'] as const) {
    if (key in value && !['restart', 'continue'].includes(String(value[key]))) throw new Error(`${key} must be restart or continue.`);
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

function validateHyperlinkReference(value: unknown): void {
  if (typeof value === 'number') {
    assertIndex(value);
    return;
  }
  object(value);
  keys(value, ['paragraph', 'runs', 'text']);
  assertIndex(value.paragraph);
  if (!Array.isArray(value.runs) || value.runs.length === 0) throw new Error('hyperlink.runs must be a non-empty array.');
  for (const run of value.runs) assertIndex(run);
  assertText(value.text, 'hyperlink.text');
}

function validateCommentInput(value: unknown, name = 'comment'): void {
  object(value);
  keys(value, ['author', 'initials', 'text']);
  if ('author' in value && value.author !== undefined) assertText(value.author, `${name}.author`);
  if ('initials' in value && value.initials !== undefined) assertText(value.initials, `${name}.initials`);
  assertText(value.text, `${name}.text`);
}

function validateEditableRegionOptions(value: unknown): asserts value is {
  editorGroup?: EditableRegionEditorGroup;
  editorId?: string;
} {
  object(value);
  keys(value, ['editorGroup', 'editorId']);
  if (value.editorGroup !== undefined &&
      !['none', 'everyone', 'administrators', 'contributors', 'editors', 'owners', 'current'].includes(String(value.editorGroup))) {
    throw new Error('options.editorGroup must be a supported editor group.');
  }
  if (value.editorId !== undefined) assertText(value.editorId, 'options.editorId');
  if ((value.editorGroup !== undefined) === (value.editorId !== undefined)) {
    throw new Error('Specify exactly one of options.editorGroup or options.editorId.');
  }
}

function validateRevisionFilter(value: unknown): void {
  object(value);
  keys(value, ['authors']);
  if ('authors' in value) {
    if (!Array.isArray(value.authors) || value.authors.length > 1000) {
      throw new Error('filter.authors must be an array of at most 1000 authors.');
    }
    value.authors.forEach((author, index) => assertText(author, `filter.authors[${index}]`));
  }
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
      case 'setTrackChanges':
        keys(op, ['type', 'enabled']);
        if (typeof op.enabled !== 'boolean') throw new Error('enabled must be boolean.');
        break;
      case 'setRevisionAuthor':
        keys(op, ['type', 'author']);
        assertText(op.author, 'author');
        break;
      case 'acceptRevision':
      case 'rejectRevision':
        keys(op, ['type', 'id']);
        assertIndex(op.id);
        break;
      case 'acceptAllRevisions':
      case 'rejectAllRevisions':
        keys(op, ['type', 'filter']);
        if ('filter' in op && op.filter !== undefined) validateRevisionFilter(op.filter);
        break;
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
      case 'applyParagraphStyle':
        keys(op, ['type', 'index', 'styleId', 'options']);
        assertIndex(op.index);
        assertText(op.styleId, 'styleId');
        if ('options' in op && op.options !== undefined) validateStyleApplyOptions(op.options);
        break;
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
      case 'restartNumbering':
        keys(op, ['type', 'index', 'options']);
        assertIndex(op.index);
        if ('options' in op && op.options !== undefined) {
          object(op.options);
          keys(op.options, ['start']);
          if ('start' in op.options && op.options.start !== undefined) {
            assertIndex(op.options.start);
            if (op.options.start < 1) throw new Error('options.start must be a positive integer.');
          }
        }
        break;
      case 'continueNumbering':
        keys(op, ['type', 'index']); assertIndex(op.index); break;
      case 'formatRun':
        keys(op, ['type', 'paragraph', 'run', 'format']);
        assertIndex(op.paragraph); assertIndex(op.run); validateRunFormat(op.format); break;
      case 'formatRange':
        keys(op, ['type', 'range', 'format']);
        validateTextRange(op.range); validateRunFormat(op.format); break;
      case 'applyCharacterStyle':
        keys(op, ['type', 'range', 'styleId', 'options']);
        validateTextRange(op.range);
        assertText(op.styleId, 'styleId');
        if ('options' in op && op.options !== undefined) validateStyleApplyOptions(op.options);
        break;
      case 'clearRangeFormat':
        keys(op, ['type', 'range', 'fields']);
        validateTextRange(op.range);
        if ('fields' in op) {
          if (!Array.isArray(op.fields)) throw new Error('fields must be an array.');
          for (const field of op.fields) {
            if (!RUN_FORMAT_FIELDS.includes(field as typeof RUN_FORMAT_FIELDS[number])) {
              throw new Error(`Unsupported run format field: ${String(field)}`);
            }
          }
        }
        break;
      case 'formatDocumentRange':
        keys(op, ['type', 'range', 'format']);
        validateDocumentRange(op.range); validateRunFormat(op.format); break;
      case 'setOutlineLevel':
        keys(op, ['type', 'index', 'level']);
        assertIndex(op.index);
        if (op.level !== null) {
          assertIndex(op.level);
          if (op.level > 8) throw new Error('level must be between 0 and 8 or null.');
        }
        break;
      case 'moveOutlineSection':
        keys(op, ['type', 'from', 'to']);
        assertIndex(op.from);
        assertIndex(op.to);
        break;
      case 'setParagraphTabs':
        keys(op, ['type', 'index', 'tabs']); assertIndex(op.index); validateTabs(op.tabs); break;
      case 'setParagraphBorders':
        keys(op, ['type', 'index', 'borders']); assertIndex(op.index); validateParagraphBorders(op.borders); break;
      case 'setParagraphShading':
        keys(op, ['type', 'index', 'shading']); assertIndex(op.index); validateDocShading(op.shading); break;
      case 'insertBreak':
        keys(op, ['type', 'paragraph', 'run', 'breakType']);
        assertIndex(op.paragraph); assertIndex(op.run);
        if (!['textWrapping', 'page', 'column'].includes(String(op.breakType))) throw new Error('Invalid break type.');
        break;
      case 'insertSymbol':
        keys(op, ['type', 'paragraph', 'run', 'font', 'charCode']);
        assertIndex(op.paragraph); assertIndex(op.run); assertText(op.font, 'font');
        if (typeof op.charCode !== 'number' || !Number.isInteger(op.charCode) || op.charCode < 0 || op.charCode > 0xFFFF ||
            !isValidXmlCharCode(op.charCode)) throw new Error('charCode must be an XML-valid BMP code point.');
        break;
      case 'replaceText':
        keys(op, ['type', 'search', 'replacement']); assertText(op.search, 'search'); assertText(op.replacement, 'replacement');
        if (!op.search) throw new Error('search must not be empty.');
        break;
      case 'insertTable':
        keys(op, ['type', 'rows']); validateRows(op.rows); break;
      case 'insertTableAt':
        keys(op, ['type', 'rows', 'cols', 'before', 'format']);
        assertIndex(op.rows); assertIndex(op.cols);
        if (op.rows < 1 || op.cols < 1) throw new Error('rows and cols must be positive.');
        if ('before' in op) assertIndex(op.before);
        if ('format' in op) validateTableFormat(op.format);
        break;
      case 'insertTableRow':
      case 'deleteTableRow':
      case 'insertTableColumn':
      case 'deleteTableColumn':
        keys(op, ['type', 'table', 'at']); assertIndex(op.table); assertIndex(op.at); break;
      case 'mergeCells':
        keys(op, ['type', 'table', 'range']); assertIndex(op.table); object(op.range);
        {
          const range = op.range as Record<string, unknown>;
          keys(range, ['row', 'col', 'rowSpan', 'colSpan']);
          for (const key of ['row', 'col', 'rowSpan', 'colSpan'] as const) assertIndex(range[key]);
          if ((range.rowSpan as number) < 1 || (range.colSpan as number) < 1) throw new Error('merge range must be positive.');
        }
        break;
      case 'splitCell':
        keys(op, ['type', 'table', 'row', 'col', 'rows', 'cols']);
        for (const key of ['table', 'row', 'col', 'rows', 'cols'] as const) assertIndex(op[key]);
        if ((op.rows as number) < 1 || (op.cols as number) < 1) throw new Error('split rows/cols must be positive.');
        break;
      case 'formatTable':
        keys(op, ['type', 'table', 'format']); assertIndex(op.table); validateTableFormat(op.format); break;
      case 'formatTableRow':
        keys(op, ['type', 'table', 'row', 'format']); assertIndex(op.table); assertIndex(op.row); validateRowFormat(op.format); break;
      case 'formatCell':
        keys(op, ['type', 'table', 'row', 'col', 'format']);
        assertIndex(op.table); assertIndex(op.row); assertIndex(op.col); validateCellFormat(op.format); break;
      case 'setCellText':
        keys(op, ['type', 'table', 'row', 'col', 'text']);
        assertIndex(op.table); assertIndex(op.row); assertIndex(op.col); assertText(op.text);
        break;
      case 'insertHyperlink':
        keys(op, ['type', 'target', 'link']);
        object(op.target);
        keys(op.target, ['paragraph', 'start', 'end']);
        assertIndex(op.target.paragraph);
        assertIndex(op.target.start);
        assertIndex(op.target.end);
        object(op.link);
        assertHyperlinkInput(op.link as { url?: string; anchor?: string; tooltip?: string });
        break;
      case 'updateHyperlink':
        keys(op, ['type', 'hyperlink', 'link']);
        validateHyperlinkReference(op.hyperlink);
        object(op.link);
        assertHyperlinkInput(op.link as { url?: string; anchor?: string; tooltip?: string });
        break;
      case 'removeHyperlink':
        keys(op, ['type', 'hyperlink', 'options']);
        validateHyperlinkReference(op.hyperlink);
        if ('options' in op) {
          object(op.options);
          keys(op.options, ['keepText']);
          if ('keepText' in op.options && typeof op.options.keepText !== 'boolean') {
            throw new Error('options.keepText must be boolean.');
          }
        }
        break;
      case 'insertBookmark':
        keys(op, ['type', 'name', 'range']);
        assertText(op.name, 'name');
        object(op.range);
        keys(op.range, ['startParagraph', 'endParagraph']);
        assertIndex(op.range.startParagraph);
        if ('endParagraph' in op.range) assertIndex(op.range.endParagraph);
        break;
      case 'updateFields':
        keys(op, ['type', 'kinds', 'now', 'filename']);
        if ('kinds' in op && op.kinds !== undefined && (!Array.isArray(op.kinds) || op.kinds.some(kind => typeof kind !== 'string'))) {
          throw new Error('kinds must be an array.');
        }
        if ('now' in op && op.now !== undefined) assertText(op.now, 'now');
        if ('filename' in op && op.filename !== undefined) assertText(op.filename, 'filename');
        break;
      case 'insertField':
        keys(op, ['type', 'paragraph', 'instruction', 'result']);
        assertIndex(op.paragraph); assertText(op.instruction, 'instruction');
        if (op.instruction.length > 4096) throw new Error('Field instruction exceeds 4096 characters.');
        if ('result' in op && op.result !== undefined) assertText(op.result, 'result');
        if (NEVER_EVALUATE.has(fieldKindFromInstruction(op.instruction))) throw new Error('This field type is not allowed.');
        break;
      case 'deleteBookmark':
        keys(op, ['type', 'name']);
        assertText(op.name, 'name');
        break;
      case 'setContentControlText':
        keys(op, ['type', 'id', 'text']);
        assertIndex(op.id); assertText(op.text);
        break;
      case 'setContentControlChecked':
        keys(op, ['type', 'id', 'checked']);
        assertIndex(op.id);
        if (typeof op.checked !== 'boolean') throw new Error('checked must be boolean.');
        break;
      case 'setContentControlProperties':
        keys(op, ['type', 'id', 'patch']);
        assertIndex(op.id);
        object(op.patch);
        keys(op.patch, ['alias', 'tag', 'lock']);
        for (const key of ['alias', 'tag'] as const) {
          if (key in op.patch) maybeNull(op.patch[key] as string | null | undefined, (entry) => assertText(entry, key));
        }
        if ('lock' in op.patch && op.patch.lock !== undefined &&
            !['sdtLocked', 'contentLocked', 'sdtContentLocked', 'unlocked'].includes(String(op.patch.lock))) {
          throw new Error('Invalid content control lock.');
        }
        break;
      case 'removeContentControl':
        keys(op, ['type', 'id', 'options']);
        assertIndex(op.id);
        if ('options' in op && op.options !== undefined) {
          object(op.options);
          keys(op.options, ['keepContent']);
          if ('keepContent' in op.options && typeof op.options.keepContent !== 'boolean') {
            throw new Error('options.keepContent must be boolean.');
          }
        }
        break;
      case 'addEditableRegion':
        keys(op, ['type', 'range', 'options']);
        validateDocumentRange(op.range);
        validateEditableRegionOptions(op.options);
        break;
      case 'removeEditableRegion':
        keys(op, ['type', 'id']);
        assertIndex(op.id);
        break;
      case 'insertImage':
        keys(op, ['type', 'bytes', 'contentType', 'paragraph', 'run', 'widthEmu', 'heightEmu', 'alt', 'placement']);
        assertText(op.bytes, 'bytes');
        assertBase64(op.bytes);
        assertText(op.contentType, 'contentType');
        if ('paragraph' in op) assertIndex(op.paragraph);
        if ('run' in op) assertIndex(op.run);
        if ('widthEmu' in op && (typeof op.widthEmu !== 'number' || !Number.isFinite(op.widthEmu) || op.widthEmu <= 0)) throw new Error('widthEmu must be a positive number.');
        if ('heightEmu' in op && (typeof op.heightEmu !== 'number' || !Number.isFinite(op.heightEmu) || op.heightEmu <= 0)) throw new Error('heightEmu must be a positive number.');
        if ('alt' in op) assertText(op.alt, 'alt');
        if ('placement' in op && !['inline', 'floating'].includes(String(op.placement))) throw new Error('Invalid image placement.');
        break;
      case 'resizeImage':
        keys(op, ['type', 'image', 'size']); assertText(op.image, 'image'); object(op.size);
        keys(op.size, ['widthEmu', 'heightEmu', 'keepAspect']);
        if (!('widthEmu' in op.size) && !('heightEmu' in op.size)) throw new Error('resizeImage requires widthEmu and/or heightEmu.');
        if ('widthEmu' in op.size && (typeof op.size.widthEmu !== 'number' || !Number.isFinite(op.size.widthEmu) || op.size.widthEmu <= 0)) throw new Error('widthEmu must be a positive number.');
        if ('heightEmu' in op.size && (typeof op.size.heightEmu !== 'number' || !Number.isFinite(op.size.heightEmu) || op.size.heightEmu <= 0)) throw new Error('heightEmu must be a positive number.');
        if ('keepAspect' in op.size && typeof op.size.keepAspect !== 'boolean') throw new Error('keepAspect must be boolean.');
        break;
      case 'replaceImageBytes':
        keys(op, ['type', 'image', 'bytes', 'contentType']);
        assertText(op.image, 'image');
        assertText(op.bytes, 'bytes');
        assertBase64(op.bytes);
        if ('contentType' in op) assertText(op.contentType, 'contentType');
        break;
      case 'setImageAlt':
        keys(op, ['type', 'image', 'alt', 'title']); assertText(op.image, 'image'); assertText(op.alt, 'alt');
        if ('title' in op) assertText(op.title, 'title');
        break;
      case 'deleteImage':
        keys(op, ['type', 'image']); assertText(op.image, 'image'); break;
      case 'setPartXml':
        keys(op, ['type', 'path', 'xml']); assertText(op.path, 'path'); assertText(op.xml, 'xml'); break;
      case 'insertFootnote':
      case 'insertEndnote':
        keys(op, ['type', 'paragraph', 'run', 'text', 'customMark']);
        assertIndex(op.paragraph); assertIndex(op.run); assertText(op.text);
        if ('customMark' in op && op.customMark !== undefined) assertText(op.customMark, 'customMark');
        break;
      case 'setNoteText':
        keys(op, ['type', 'kind', 'id', 'text']);
        if (!['footnote', 'endnote'].includes(String(op.kind))) throw new Error('Invalid note kind.');
        assertIndex(op.id); assertText(op.text);
        break;
      case 'deleteNote':
        keys(op, ['type', 'kind', 'id']);
        if (!['footnote', 'endnote'].includes(String(op.kind))) throw new Error('Invalid note kind.');
        assertIndex(op.id);
        break;
      case 'convertNote':
        keys(op, ['type', 'kind', 'id']);
        if (!['footnote', 'endnote'].includes(String(op.kind))) throw new Error('Invalid note kind.');
        assertIndex(op.id);
        break;
      case 'addComment':
        keys(op, ['type', 'range', 'comment']);
        object(op.range);
        if ('paragraph' in op.range) validateTextRange(op.range);
        else validateDocumentRange(op.range);
        validateCommentInput(op.comment);
        break;
      case 'replyComment':
        keys(op, ['type', 'parentId', 'comment']);
        assertIndex(op.parentId);
        validateCommentInput(op.comment);
        break;
      case 'setCommentResolved':
        keys(op, ['type', 'id', 'resolved']);
        assertIndex(op.id);
        if (typeof op.resolved !== 'boolean') throw new Error('resolved must be boolean.');
        break;
      case 'setCommentText':
        keys(op, ['type', 'id', 'text']);
        assertIndex(op.id);
        assertText(op.text, 'text');
        break;
      case 'deleteComment':
        keys(op, ['type', 'id', 'options']);
        assertIndex(op.id);
        if ('options' in op) {
          object(op.options);
          keys(op.options, ['withReplies']);
          if ('withReplies' in op.options && typeof op.options.withReplies !== 'boolean') {
            throw new Error('options.withReplies must be boolean.');
          }
        }
      case 'undo':
      case 'redo':
        keys(op, ['type']);
        break;
      default: throw new Error(`Unknown operation type: ${String(op.type)}`);
    }
  }
}

const text = { type: 'string', maxLength: 1_000_000 };
const index = { type: 'integer', minimum: 0, maximum: Number.MAX_SAFE_INTEGER };
const integer = { type: 'integer', minimum: -Number.MAX_SAFE_INTEGER, maximum: Number.MAX_SAFE_INTEGER };
const signedInteger = { type: 'integer', minimum: -31680, maximum: 31680 };
const unsignedTwips = { type: 'integer', minimum: 0, maximum: 31680 };
const outlineLevel = { type: 'integer', minimum: 0, maximum: 9 };
const nullable = <T extends Record<string, unknown>>(schema: T) => ({ anyOf: [schema, { type: 'null' }] });
const shape = (properties: Record<string, unknown>, required = Object.keys(properties)) => ({
  type: 'object', properties, required, additionalProperties: false,
});
const imageSizeShape = {
  ...shape({
    widthEmu: { type: 'number', exclusiveMinimum: 0 },
    heightEmu: { type: 'number', exclusiveMinimum: 0 },
    keepAspect: { type: 'boolean' },
  }, []),
  anyOf: [{ required: ['widthEmu'] }, { required: ['heightEmu'] }],
};
const operation = (type: string, properties: Record<string, unknown>, required = Object.keys(properties)) =>
  shape({ type: { const: type }, ...properties }, ['type', ...required]);
const hyperlinkTarget = shape({ paragraph: index, start: index, end: index });
const textRange = shape({ paragraph: index, start: index, end: index });
const documentRange = shape({
  start: shape({ paragraph: index, offset: index }),
  end: shape({ paragraph: index, offset: index }),
}, ['start', 'end']);
const styleApplyOptions = shape({
  clearDirectFormat: { type: 'boolean' },
}, []);
const revisionFilter = shape({
  authors: { type: 'array', maxItems: 1000, items: text },
}, []);
const runFormatField = {
  enum: [...RUN_FORMAT_FIELDS],
};
const hyperlinkLink = shape({ url: text, anchor: text, tooltip: text }, []);
const commentInput = shape({ author: text, initials: text, text }, ['text']);
const hyperlinkRef = {
  anyOf: [
    index,
    shape({
      paragraph: index,
      runs: { type: 'array', minItems: 1, items: index },
      text,
    }),
  ],
};
const editableRegionOptions = {
  ...shape({
    editorGroup: { enum: ['none', 'everyone', 'administrators', 'contributors', 'editors', 'owners', 'current'] },
    editorId: text,
  }, []),
  oneOf: [{ required: ['editorGroup'] }, { required: ['editorId'] }],
};
const width = shape({ type: { enum: ['auto', 'dxa', 'pct'] }, value: { type: 'number', minimum: 0 } });
const border = shape({
  style: text,
  size: { type: 'number', minimum: 0 },
  space: { type: 'number', minimum: 0 },
  color: { type: 'string', pattern: '^(auto|[a-fA-F0-9]{6})$' },
  none: { type: 'boolean' },
}, []);
const docBorderSide = shape({
  style: text,
  size: { type: 'number', minimum: 0 },
  space: { type: 'number', minimum: 0 },
  color: { type: 'string', pattern: '^(auto|[a-fA-F0-9]{6})$' },
  shadow: { type: 'boolean' },
}, ['style', 'size', 'space', 'color']);
const docShading = shape({
  pattern: text,
  fill: { type: 'string', pattern: '^(auto|[a-fA-F0-9]{6})$' },
  color: { type: 'string', pattern: '^(auto|[a-fA-F0-9]{6})$' },
}, ['pattern', 'fill']);
const docTabs = {
  type: 'array',
  maxItems: 200,
  items: shape({
    position: { type: 'number' },
    alignment: { enum: ['left', 'center', 'right', 'decimal', 'bar', 'clear', 'num'] },
    leader: { enum: ['none', 'dot', 'hyphen', 'underscore', 'heavy', 'middleDot'] },
  }, ['position', 'alignment']),
};
const borders = shape({
  top: border, right: border, bottom: border, left: border, insideH: border, insideV: border,
}, []);
const shading = shape({
  fill: { type: 'string', pattern: '^(auto|[a-fA-F0-9]{6})$' },
  color: { type: 'string', pattern: '^(auto|[a-fA-F0-9]{6})$' },
  value: text,
}, []);
const margins = shape({ top: width, right: width, bottom: width, left: width }, []);
const tableFormat = shape({
  width, alignment: { enum: ['left', 'center', 'right'] }, indent: { type: 'number', minimum: 0 },
  borders, shading, cellMargin: margins, layout: { enum: ['fixed', 'autofit'] },
  style: text, look: text, caption: text, description: text, bidiVisual: nullable({ type: 'boolean' }),
}, []);
const rowFormat = shape({
  height: shape({ value: { type: 'number', minimum: 0 }, rule: { enum: ['atLeast', 'exact'] } }, ['value']),
  cantSplit: { type: 'boolean' }, header: { type: 'boolean' }, alignment: { enum: ['left', 'center', 'right'] },
  deleted: { type: 'boolean' }, inserted: { type: 'boolean' },
  revision: shape({ author: { type: 'string' }, date: { type: 'string' } }),
}, []);
const cellFormat = shape({
  width, borders, shading, margin: margins, verticalAlign: { enum: ['top', 'center', 'bottom'] },
  textDirection: text, noWrap: { type: 'boolean' }, hideMark: { type: 'boolean' },
  hMerge: { enum: ['restart', 'continue'] }, vMerge: { enum: ['restart', 'continue'] },
}, []);

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
          operation('setTrackChanges', { enabled: { type: 'boolean' } }),
          operation('setRevisionAuthor', { author: text }),
          operation('acceptRevision', { id: index }),
          operation('rejectRevision', { id: index }),
          operation('acceptAllRevisions', { filter: revisionFilter }, []),
          operation('rejectAllRevisions', { filter: revisionFilter }, []),
          operation('setParagraphText', { index, text }),
          operation('insertParagraph', { text, before: index }, ['text']),
          operation('deleteParagraph', { index }),
          operation('formatParagraph', { index, format: shape({
            alignment: nullable({ enum: ['left', 'center', 'right', 'both', 'distribute'] }),
            style: nullable(text),
            indentLeft: nullable(signedInteger),
            indentRight: nullable(signedInteger),
            indentFirstLine: nullable(unsignedTwips),
            indentHanging: nullable(unsignedTwips),
            spacingBefore: nullable(unsignedTwips),
            spacingAfter: nullable(unsignedTwips),
            lineSpacing: nullable(signedInteger),
            lineSpacingRule: nullable({ enum: ['auto', 'atLeast', 'exact'] }),
            keepNext: nullable({ type: 'boolean' }),
            keepLines: nullable({ type: 'boolean' }),
            pageBreakBefore: nullable({ type: 'boolean' }),
            widowControl: nullable({ type: 'boolean' }),
            suppressLineNumbers: nullable({ type: 'boolean' }),
            suppressAutoHyphens: nullable({ type: 'boolean' }),
            kinsoku: nullable({ type: 'boolean' }),
            wordWrap: nullable({ type: 'boolean' }),
            overflowPunct: nullable({ type: 'boolean' }),
            topLinePunct: nullable({ type: 'boolean' }),
            autoSpaceDE: nullable({ type: 'boolean' }),
            autoSpaceDN: nullable({ type: 'boolean' }),
            bidi: nullable({ type: 'boolean' }),
            textDirection: nullable(text),
            outlineLevel: nullable(outlineLevel),
            tabs: nullable(docTabs),
            borders: nullable(shape({
              top: docBorderSide,
              left: docBorderSide,
              bottom: docBorderSide,
              right: docBorderSide,
              between: docBorderSide,
              bar: docBorderSide,
            }, [])),
            shading: nullable(docShading),
          }, []) }),
          operation('applyParagraphStyle', { index, styleId: text, options: styleApplyOptions }, ['index', 'styleId']),
          operation('setParagraphNumbering', { index, numId: { ...index, minimum: 1 }, level: { ...index, maximum: 8 } }, ['index', 'numId']),
          operation('clearParagraphNumbering', { index }),
          operation('setParagraphLevel', { index, delta: integer }),
          operation('restartNumbering', { index, options: shape({ start: { ...index, minimum: 1 } }, []) }, ['index']),
          operation('continueNumbering', { index }),
          operation('formatRun', { paragraph: index, run: index, format: shape({
            style: nullable(text),
            bold: nullable({ type: 'boolean' }),
            italic: nullable({ type: 'boolean' }),
            underline: nullable({ type: 'boolean' }),
            underlineStyle: nullable(text),
            underlineColor: nullable({ type: 'string', pattern: '^[a-fA-F0-9]{6}$' }),
            fontSize: nullable({ type: 'number', minimum: 1, maximum: 400, multipleOf: 0.5 }),
            fontFamily: nullable(text),
            fontFamilyEastAsia: nullable(text),
            color: nullable({ type: 'string', pattern: '^[a-fA-F0-9]{6}$' }),
            strike: nullable({ type: 'boolean' }),
            doubleStrike: nullable({ type: 'boolean' }),
            rtl: nullable({ type: 'boolean' }),
            complexScript: nullable({ type: 'boolean' }),
            verticalAlign: nullable({ enum: ['baseline', 'subscript', 'superscript'] }),
            smallCaps: nullable({ type: 'boolean' }),
            allCaps: nullable({ type: 'boolean' }),
            highlight: nullable(text),
            characterSpacing: nullable(signedInteger),
            border: nullable(docBorderSide),
            shading: nullable(docShading),
          }, []) }),
          operation('formatRange', {
            range: textRange,
            format: shape({
              style: nullable(text),
              bold: nullable({ type: 'boolean' }),
              italic: nullable({ type: 'boolean' }),
              underline: nullable({ type: 'boolean' }),
              underlineStyle: nullable(text),
              underlineColor: nullable({ type: 'string', pattern: '^[a-fA-F0-9]{6}$' }),
              fontSize: nullable({ type: 'number', minimum: 1, maximum: 400, multipleOf: 0.5 }),
              fontFamily: nullable(text),
              fontFamilyEastAsia: nullable(text),
              color: nullable({ type: 'string', pattern: '^[a-fA-F0-9]{6}$' }),
              strike: nullable({ type: 'boolean' }),
              doubleStrike: nullable({ type: 'boolean' }),
              rtl: nullable({ type: 'boolean' }),
              complexScript: nullable({ type: 'boolean' }),
              verticalAlign: nullable({ enum: ['baseline', 'subscript', 'superscript'] }),
              smallCaps: nullable({ type: 'boolean' }),
              allCaps: nullable({ type: 'boolean' }),
              highlight: nullable(text),
              characterSpacing: nullable(signedInteger),
              border: nullable(docBorderSide),
              shading: nullable(docShading),
            }, []),
          }),
          operation('applyCharacterStyle', { range: textRange, styleId: text, options: styleApplyOptions }, ['range', 'styleId']),
          operation('clearRangeFormat', {
            range: textRange,
            fields: { type: 'array', items: runFormatField },
          }, ['range']),
          operation('formatDocumentRange', {
            range: documentRange,
            format: shape({
              style: nullable(text),
              bold: nullable({ type: 'boolean' }),
              italic: nullable({ type: 'boolean' }),
              underline: nullable({ type: 'boolean' }),
              underlineStyle: nullable(text),
              underlineColor: nullable({ type: 'string', pattern: '^[a-fA-F0-9]{6}$' }),
              fontSize: nullable({ type: 'number', minimum: 1, maximum: 400, multipleOf: 0.5 }),
              fontFamily: nullable(text),
              fontFamilyEastAsia: nullable(text),
              color: nullable({ type: 'string', pattern: '^[a-fA-F0-9]{6}$' }),
              strike: nullable({ type: 'boolean' }),
              doubleStrike: nullable({ type: 'boolean' }),
              rtl: nullable({ type: 'boolean' }),
              complexScript: nullable({ type: 'boolean' }),
              verticalAlign: nullable({ enum: ['baseline', 'subscript', 'superscript'] }),
              smallCaps: nullable({ type: 'boolean' }),
              allCaps: nullable({ type: 'boolean' }),
              highlight: nullable(text),
              characterSpacing: nullable(signedInteger),
              border: nullable(docBorderSide),
              shading: nullable(docShading),
            }, []),
          }),
          operation('setOutlineLevel', { index, level: nullable({ ...outlineLevel, maximum: 8 }) }),
          operation('moveOutlineSection', { from: index, to: index }),
          operation('setParagraphTabs', { index, tabs: docTabs }),
          operation('setParagraphBorders', { index, borders: shape({
            top: docBorderSide,
            left: docBorderSide,
            bottom: docBorderSide,
            right: docBorderSide,
            between: docBorderSide,
            bar: docBorderSide,
          }, []) }),
          operation('setParagraphShading', { index, shading: docShading }),
          operation('insertBreak', { paragraph: index, run: index, breakType: { enum: ['textWrapping', 'page', 'column'] } }),
          operation('insertSymbol', { paragraph: index, run: index, font: text, charCode: { type: 'integer', minimum: 0, maximum: 65535 } }),
          operation('replaceText', { search: { ...text, minLength: 1 }, replacement: text }),
          operation('insertTable', { rows: {
            type: 'array', minItems: 1, maxItems: 1000,
            items: { type: 'array', minItems: 1, maxItems: 100, items: text },
          } }),
          operation('insertTableAt', { rows: { ...index, minimum: 1 }, cols: { ...index, minimum: 1 }, before: index, format: tableFormat }, ['rows', 'cols']),
          operation('insertTableRow', { table: index, at: index }),
          operation('deleteTableRow', { table: index, at: index }),
          operation('insertTableColumn', { table: index, at: index }),
          operation('deleteTableColumn', { table: index, at: index }),
          operation('mergeCells', { table: index, range: shape({ row: index, col: index, rowSpan: { ...index, minimum: 1 }, colSpan: { ...index, minimum: 1 } }) }),
          operation('splitCell', { table: index, row: index, col: index, rows: { ...index, minimum: 1 }, cols: { ...index, minimum: 1 } }),
          operation('formatTable', { table: index, format: tableFormat }),
          operation('formatTableRow', { table: index, row: index, format: rowFormat }),
          operation('formatCell', { table: index, row: index, col: index, format: cellFormat }),
          operation('setCellText', { table: index, row: index, col: index, text }),
          operation('insertHyperlink', { target: hyperlinkTarget, link: hyperlinkLink }),
          operation('updateHyperlink', { hyperlink: hyperlinkRef, link: hyperlinkLink }),
          operation('removeHyperlink', { hyperlink: hyperlinkRef, options: shape({ keepText: { type: 'boolean' } }, []) }, ['hyperlink']),
          operation('insertBookmark', { name: text, range: shape({ startParagraph: index, endParagraph: index }, ['startParagraph']) }),
          operation('deleteBookmark', { name: text }),
          operation('updateFields', {
            kinds: { type: 'array', items: text },
            now: text,
            filename: text,
          }),
          operation('insertField', { paragraph: index, instruction: text, result: text }, ['paragraph', 'instruction']),
          operation('setContentControlText', { id: index, text }),
          operation('setContentControlChecked', { id: index, checked: { type: 'boolean' } }),
          operation('setContentControlProperties', {
            id: index,
            patch: shape({
              alias: nullable(text),
              tag: nullable(text),
              lock: { enum: ['sdtLocked', 'contentLocked', 'sdtContentLocked', 'unlocked'] },
            }, []),
          }),
          operation('removeContentControl', {
            id: index,
            options: shape({ keepContent: { type: 'boolean' } }, []),
          }, ['id']),
          operation('addEditableRegion', { range: documentRange, options: editableRegionOptions }),
          operation('removeEditableRegion', { id: index }),
          operation('insertImage', {
            bytes: { type: 'string', maxLength: 22_500_000 },
            contentType: text,
            paragraph: index,
            run: index,
            widthEmu: { type: 'number', exclusiveMinimum: 0 },
            heightEmu: { type: 'number', exclusiveMinimum: 0 },
            alt: text,
            placement: { enum: ['inline', 'floating'] },
          }, ['bytes', 'contentType']),
          operation('replaceImageBytes', {
            image: text,
            bytes: { type: 'string', maxLength: 22_500_000 },
            contentType: text,
          }, ['image', 'bytes']),
          operation('resizeImage', {
            image: text,
            size: imageSizeShape,
          }),
          operation('setImageAlt', { image: text, alt: text, title: text }, ['image', 'alt']),
          operation('deleteImage', { image: text }),
          operation('setPartXml', { path: text, xml: text }),
          operation('insertFootnote', { paragraph: index, run: index, text, customMark: text }, ['paragraph', 'run', 'text']),
          operation('insertEndnote', { paragraph: index, run: index, text, customMark: text }, ['paragraph', 'run', 'text']),
          operation('setNoteText', { kind: { enum: ['footnote', 'endnote'] }, id: index, text }),
          operation('deleteNote', { kind: { enum: ['footnote', 'endnote'] }, id: index }),
          operation('convertNote', { kind: { enum: ['footnote', 'endnote'] }, id: index }),
          operation('addComment', { range: { anyOf: [textRange, documentRange] }, comment: commentInput }),
          operation('replyComment', { parentId: index, comment: commentInput }),
          operation('setCommentResolved', { id: index, resolved: { type: 'boolean' } }),
          operation('setCommentText', { id: index, text }),
          operation('deleteComment', { id: index, options: shape({ withReplies: { type: 'boolean' } }, []) }, ['id']),
          operation('undo', {}),
          operation('redo', {}),
        ],
      },
    },
  }, ['operations']),
} as const;
