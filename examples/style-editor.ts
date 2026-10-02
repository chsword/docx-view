import type { ParagraphFormat, RunFormat, StyleInfo, StylePatch } from '../src/index.js';

/**
 * 「修改样式」对话框的表单值。全部是字符串 / 三态，**空串表示「不在这个样式里设置，继续继承」**
 * ——这正是样式的语义：样式只记它自己写了什么，其余从 basedOn 链上来。所以表单显示的是样式
 * **自己的**值，不是有效值；把继承来的值填进表单，保存时就会被当成新写入的值固化进样式。
 */
export interface StyleFormValues {
  name: string;
  basedOn: string;
  next: string;
  fontFamily: string;
  /** 磅。 */
  fontSize: string;
  bold: TriState;
  italic: TriState;
  underline: TriState;
  /** `RRGGBB`，不带 #。 */
  color: string;
  alignment: '' | NonNullable<ParagraphFormat['alignment']>;
  /** 磅。 */
  spacingBefore: string;
  spacingAfter: string;
  /** 倍数（单倍 = 1）。只表达 `lineSpacingRule: auto`；固定 / 最小行距的样式这里显示为空且不改动。 */
  lineSpacing: string;
  /** 磅。 */
  indentLeft: string;
  indentFirstLine: string;
  quickFormat: boolean;
}

export type TriState = '' | 'on' | 'off';

const TWIPS_PER_POINT = 20;
const AUTO_LINE_UNIT = 240;

function points(twips: number | undefined): string {
  return twips === undefined ? '' : String(twips / TWIPS_PER_POINT);
}

function triState(value: boolean | undefined): TriState {
  return value === undefined ? '' : value ? 'on' : 'off';
}

export function styleFormValues(style: StyleInfo): StyleFormValues {
  const paragraph = style.paragraph ?? {};
  const run = style.run ?? {};
  const autoLine = paragraph.lineSpacing !== undefined && paragraph.lineSpacing !== null &&
    (paragraph.lineSpacingRule === undefined || paragraph.lineSpacingRule === null || paragraph.lineSpacingRule === 'auto');
  return {
    name: style.name,
    basedOn: style.basedOn ?? '',
    next: style.next ?? '',
    fontFamily: run.fontFamily ?? '',
    fontSize: run.fontSize === undefined || run.fontSize === null ? '' : String(run.fontSize),
    bold: triState(run.bold ?? undefined),
    italic: triState(run.italic ?? undefined),
    underline: triState(run.underline ?? undefined),
    color: run.color ?? '',
    alignment: paragraph.alignment ?? '',
    spacingBefore: points(paragraph.spacingBefore ?? undefined),
    spacingAfter: points(paragraph.spacingAfter ?? undefined),
    lineSpacing: autoLine ? String(Number(((paragraph.lineSpacing as number) / AUTO_LINE_UNIT).toFixed(2))) : '',
    indentLeft: points(paragraph.indentLeft ?? undefined),
    indentFirstLine: points(paragraph.indentFirstLine ?? undefined),
    quickFormat: style.quickFormat === true,
  };
}

function number(value: string, label: string, min: number, max: number): number {
  const parsed = Number(value.trim());
  if (!Number.isFinite(parsed) || parsed < min || parsed > max) throw new Error(`${label}必须是 ${min}–${max} 之间的数字。`);
  return parsed;
}

function twips(value: string, label: string, min = 0): number {
  return Math.round(number(value, label, min / TWIPS_PER_POINT, 1584) * TWIPS_PER_POINT);
}

/**
 * 只比较改过的字段：没动的字段不进补丁，于是样式里表单不认识的属性、以及没打开过的字段都原样
 * 保留（`updateStyle()` 是补丁语义）。改成空串的字段映射为 `null`，即从样式里删掉、改回继承。
 */
export function stylePatchFromForm(initial: StyleFormValues, current: StyleFormValues, type: StyleInfo['type']): StylePatch {
  const patch: StylePatch = {};
  const run: RunFormat = {};
  const paragraph: ParagraphFormat = {};
  const changed = (key: keyof StyleFormValues) => initial[key] !== current[key];
  const trimmed = (key: keyof StyleFormValues) => String(current[key]).trim();

  if (changed('name')) {
    if (!trimmed('name')) throw new Error('样式名称不能为空。');
    patch.name = trimmed('name');
  }
  if (changed('basedOn')) patch.basedOn = current.basedOn || null;
  if (type === 'paragraph' && changed('next')) patch.next = current.next || null;
  if (changed('quickFormat')) patch.quickFormat = current.quickFormat;

  if (changed('fontFamily')) run.fontFamily = trimmed('fontFamily') || null;
  if (changed('fontSize')) {
    if (!trimmed('fontSize')) run.fontSize = null;
    else {
      // Word 的字号以半磅为单位存储。
      run.fontSize = Math.round(number(current.fontSize, '字号', 1, 400) * 2) / 2;
    }
  }
  for (const key of ['bold', 'italic', 'underline'] as const) {
    if (changed(key)) run[key] = current[key] === '' ? null : current[key] === 'on';
  }
  if (changed('color')) {
    const color = trimmed('color').replace(/^#/, '');
    if (color && !/^[0-9a-f]{6}$/i.test(color)) throw new Error('颜色必须是六位十六进制，例如 3567D6。');
    run.color = color ? color.toUpperCase() : null;
  }

  if (type === 'paragraph') {
    if (changed('alignment')) paragraph.alignment = current.alignment || null;
    if (changed('spacingBefore')) paragraph.spacingBefore = trimmed('spacingBefore') ? twips(current.spacingBefore, '段前') : null;
    if (changed('spacingAfter')) paragraph.spacingAfter = trimmed('spacingAfter') ? twips(current.spacingAfter, '段后') : null;
    if (changed('indentLeft')) paragraph.indentLeft = trimmed('indentLeft') ? twips(current.indentLeft, '左缩进', -31680) : null;
    if (changed('indentFirstLine')) paragraph.indentFirstLine = trimmed('indentFirstLine') ? twips(current.indentFirstLine, '首行缩进') : null;
    if (changed('lineSpacing')) {
      if (!trimmed('lineSpacing')) {
        paragraph.lineSpacing = null;
        paragraph.lineSpacingRule = null;
      } else {
        paragraph.lineSpacing = Math.round(number(current.lineSpacing, '行距', 0.5, 10) * AUTO_LINE_UNIT);
        paragraph.lineSpacingRule = 'auto';
      }
    }
  }
  if (Object.keys(run).length) patch.run = run;
  if (Object.keys(paragraph).length) patch.paragraph = paragraph;
  return patch;
}

/** 补丁里的 `null` 在新建样式时没有意义（没有可清的东西），`defineStyle()` 也拒绝它。 */
export function withoutNulls<T extends object>(format: T | undefined): T | undefined {
  if (!format) return undefined;
  const entries = Object.entries(format).filter(([, value]) => value !== null && value !== undefined);
  return entries.length ? Object.fromEntries(entries) as T : undefined;
}

/**
 * 从显示名派生样式 ID。Word 的做法是去掉空白、保留字母数字；中文名没有可用的 ASCII 字符时
 * 退回 `Style1`、`Style2`……ID 只用于引用，用户看到的始终是名称。
 */
export function uniqueStyleId(name: string, existing: Iterable<string>): string {
  const taken = new Set([...existing].map((id) => id.toLowerCase()));
  const base = name.replace(/[^A-Za-z0-9]/g, '') || 'Style';
  if (!taken.has(base.toLowerCase()) && base !== 'Style') return base;
  for (let suffix = 1; ; suffix++) {
    const candidate = `${base}${suffix}`;
    if (!taken.has(candidate.toLowerCase())) return candidate;
  }
}
