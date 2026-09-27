import type { Document, Element } from '@xmldom/xmldom';
import type { NoteSettings, NoteSettingsValue } from './types.js';
import { WORD_NS, children, descendants, setWordValue, wordElement, wordValue } from './xml.js';

export type NoteKind = 'footnote' | 'endnote';

const REL_TYPE = {
  footnote: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes',
  endnote: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes',
} as const;

const CONTENT_TYPE = {
  footnote: 'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml',
  endnote: 'application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml',
} as const;

export function notePartPath(kind: NoteKind): string {
  return `word/${kind}s.xml`;
}

const SETTINGS_ORDER = [
  'writeProtection', 'view', 'zoom', 'removePersonalInformation', 'removeDateAndTime', 'doNotDisplayPageBoundaries',
  'displayBackgroundShape', 'printPostScriptOverText', 'printFractionalCharacterWidth', 'printFormsData',
  'embedTrueTypeFonts', 'embedSystemFonts', 'saveSubsetFonts', 'saveFormsData', 'mirrorMargins', 'alignBordersAndEdges',
  'bordersDoNotSurroundHeader', 'bordersDoNotSurroundFooter', 'gutterAtTop', 'hideSpellingErrors',
  'hideGrammaticalErrors', 'activeWritingStyle', 'proofState', 'formsDesign', 'attachedTemplate', 'linkStyles',
  'stylePaneFormatFilter', 'stylePaneSortMethod', 'documentType', 'mailMerge', 'revisionView', 'trackRevisions',
  'doNotTrackMoves', 'doNotTrackFormatting', 'documentProtection', 'autoFormatOverride', 'styleLockTheme',
  'styleLockQFSet', 'defaultTabStop', 'autoHyphenation', 'consecutiveHyphenLimit', 'hyphenationZone',
  'doNotHyphenateCaps', 'showEnvelope', 'summaryLength', 'clickAndTypeStyle', 'defaultTableStyle', 'evenAndOddHeaders',
  'bookFoldRevPrinting', 'bookFoldPrinting', 'bookFoldPrintingSheets', 'drawingGridHorizontalSpacing',
  'drawingGridVerticalSpacing', 'displayHorizontalDrawingGridEvery', 'displayVerticalDrawingGridEvery',
  'doNotUseMarginsForDrawingGridOrigin', 'drawingGridHorizontalOrigin', 'drawingGridVerticalOrigin', 'doNotShadeFormData',
  'noPunctuationKerning', 'characterSpacingControl', 'printTwoOnOne', 'strictFirstAndLastChars', 'noLineBreaksAfter',
  'noLineBreaksBefore', 'savePreviewPicture', 'doNotValidateAgainstSchema', 'saveInvalidXml', 'ignoreMixedContent',
  'alwaysShowPlaceholderText', 'doNotDemarcateInvalidXml', 'saveXmlDataOnly', 'useXSLTWhenSaving', 'saveThroughXslt',
  'showXMLTags', 'alwaysMergeEmptyNamespace', 'updateFields', 'hdrShapeDefaults', 'footnotePr', 'endnotePr',
  'compat', 'docVars', 'rsids', 'mathPr',
];

const NOTE_PR_ORDER = ['pos', 'numFmt', 'numStart', 'numRestart', 'numId', 'suppressRef'];

function orderedProperty(parent: Element, name: string, order: string[]): Element {
  let result = children(parent, name)[0];
  if (!result) {
    result = wordElement(parent.ownerDocument!, name);
    const position = order.indexOf(name);
    const following = children(parent).find(child => order.indexOf(child.localName!) > position);
    parent.insertBefore(result, following ?? null);
  }
  return result;
}

export function noteRelationshipType(kind: NoteKind): string {
  return REL_TYPE[kind];
}

export function noteContentType(kind: NoteKind): string {
  return CONTENT_TYPE[kind];
}

export function noteReferenceName(kind: NoteKind): 'footnoteReference' | 'endnoteReference' {
  return kind === 'footnote' ? 'footnoteReference' : 'endnoteReference';
}

export function noteRefName(kind: NoteKind): 'footnoteRef' | 'endnoteRef' {
  return kind === 'footnote' ? 'footnoteRef' : 'endnoteRef';
}

export function noteReferenceStyle(kind: NoteKind): 'FootnoteReference' | 'EndnoteReference' {
  return kind === 'footnote' ? 'FootnoteReference' : 'EndnoteReference';
}

export function defaultNotePartXml(kind: NoteKind): string {
  const wrapper = `${kind}s`;
  const single = kind;
  return `<w:${wrapper} xmlns:w="${WORD_NS}">`
    + `<w:${single} w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:${single}>`
    + `<w:${single} w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:${single}>`
    + `</w:${wrapper}>`;
}

function readNumFmt(parent: Element): string | undefined {
  return wordValue(children(parent, 'numFmt')[0]);
}

function readNumStart(parent: Element): number | undefined {
  const value = Number(wordValue(children(parent, 'numStart')[0]));
  return Number.isSafeInteger(value) ? value : undefined;
}

function readNumRestart(parent: Element): NoteSettingsValue['numRestart'] {
  const value = wordValue(children(parent, 'numRestart')[0]);
  return ['continuous', 'eachSect'].includes(value ?? '') ? value as NoteSettingsValue['numRestart'] : undefined;
}

function readPos(parent: Element): NoteSettingsValue['pos'] {
  const value = wordValue(children(parent, 'pos')[0]);
  return ['pageBottom', 'beneathText', 'sectEnd', 'docEnd'].includes(value ?? '') ? value as NoteSettingsValue['pos'] : undefined;
}

function readNotePr(parent: Element | undefined): NoteSettingsValue {
  if (!parent) return {};
  return { pos: readPos(parent), numFmt: readNumFmt(parent), numStart: readNumStart(parent), numRestart: readNumRestart(parent) };
}

function mergeSetting(base: NoteSettingsValue, override: NoteSettingsValue): NoteSettingsValue {
  return { ...base, ...Object.fromEntries(Object.entries(override).filter(([, value]) => value !== undefined)) };
}

export function parseDocumentNoteSettings(settingsDocument: Document | null): NoteSettings {
  const root = settingsDocument?.documentElement;
  if (!root || root.namespaceURI !== WORD_NS || root.localName !== 'settings') {
    return { footnote: {}, endnote: {} };
  }
  return {
    footnote: readNotePr(children(root, 'footnotePr')[0]),
    endnote: readNotePr(children(root, 'endnotePr')[0]),
  };
}

export function parseSectionNoteSettings(sectPr: Element | undefined, base: NoteSettings): NoteSettings {
  if (!sectPr) return base;
  return {
    footnote: mergeSetting(base.footnote, readNotePr(children(sectPr, 'footnotePr')[0])),
    endnote: mergeSetting(base.endnote, readNotePr(children(sectPr, 'endnotePr')[0])),
  };
}

export function setNoteSettingsOn(settingsDocument: Document, update: Partial<NoteSettings>): void {
  const root = settingsDocument.documentElement;
  if (!root || root.namespaceURI !== WORD_NS || root.localName !== 'settings') throw new Error('Invalid settings.xml root.');
  for (const kind of ['footnote', 'endnote'] as const) {
    const patch = update[kind];
    if (!patch) continue;
    const tag = kind === 'footnote' ? 'footnotePr' : 'endnotePr';
    const pr = orderedProperty(root, tag, SETTINGS_ORDER);
    if (patch.pos !== undefined) setWordValue(orderedProperty(pr, 'pos', NOTE_PR_ORDER), patch.pos);
    if (patch.numFmt !== undefined) setWordValue(orderedProperty(pr, 'numFmt', NOTE_PR_ORDER), patch.numFmt);
    if (patch.numStart !== undefined) setWordValue(orderedProperty(pr, 'numStart', NOTE_PR_ORDER), String(patch.numStart));
    if (patch.numRestart !== undefined) setWordValue(orderedProperty(pr, 'numRestart', NOTE_PR_ORDER), patch.numRestart);
  }
}

export interface ParsedNoteEntry {
  id: number;
  type: string;
  element: Element;
}

export function parseNoteEntries(document: Document | null, kind: NoteKind): ParsedNoteEntry[] {
  if (!document?.documentElement || document.documentElement.namespaceURI !== WORD_NS) return [];
  const name = kind;
  return children(document.documentElement, name)
    .map((element): ParsedNoteEntry | null => {
      const id = Number(element.getAttributeNS(WORD_NS, 'id'));
      if (!Number.isSafeInteger(id)) return null;
      return { id, type: element.getAttributeNS(WORD_NS, 'type') ?? 'normal', element };
    })
    .filter((entry): entry is ParsedNoteEntry => Boolean(entry));
}

export function parseCustomMark(note: Element): string | undefined {
  const firstParagraph = descendants(note, 'p')[0];
  if (!firstParagraph) return undefined;
  let marker = '';
  let foundRef = false;
  for (const run of children(firstParagraph, 'r')) {
    if (children(run, 'footnoteRef').length || children(run, 'endnoteRef').length) {
      foundRef = true;
      break;
    }
    marker += descendants(run, 't').map(item => item.textContent ?? '').join('');
  }
  return foundRef ? marker || undefined : undefined;
}

function roman(value: number): string {
  const table: [number, string][] = [[1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'], [100, 'C'], [90, 'XC'],
    [50, 'L'], [40, 'XL'], [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I']];
  let rest = Math.max(1, Math.floor(value));
  let result = '';
  for (const [num, token] of table) while (rest >= num) { rest -= num; result += token; }
  return result;
}

function alpha(value: number): string {
  let rest = Math.max(1, Math.floor(value));
  let result = '';
  while (rest > 0) {
    rest--;
    result = String.fromCharCode((rest % 26) + 65) + result;
    rest = Math.floor(rest / 26);
  }
  return result;
}

export function formatNoteMarker(number: number, format = 'decimal'): string {
  switch (format) {
    case 'upperRoman': return roman(number);
    case 'lowerRoman': return roman(number).toLowerCase();
    case 'upperLetter': return alpha(number);
    case 'lowerLetter': return alpha(number).toLowerCase();
    case 'chicago': return ['*', '†', '‡', '§', '‖', '¶'][Math.max(0, (number - 1) % 6)] ?? String(number);
    default: return String(number);
  }
}
