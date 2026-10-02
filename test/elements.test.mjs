import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { DOMParser } from '@xmldom/xmldom';
import { PROPERTY_ORDER, property } from '../dist/internal/elements.js';
import { markRevision } from '../dist/revisions.js';

const WORD_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

test('the internal elements module stays a leaf with no browser DOM dependency', () => {
  const source = readFileSync(new URL('../src/internal/elements.ts', import.meta.url), 'utf8');
  assert.doesNotMatch(source, /\b(?:globalThis\.document|window)\./);
  assert.doesNotMatch(source, /HTMLElement|globalThis/);
  assert.deepEqual(
    [...source.matchAll(/from ['"]([^'"]+)['"]/g)].map((match) => match[1]),
    ['@xmldom/xmldom', '../types.js', '../xml.js'],
  );
});

test('PROPERTY_ORDER has one 12-parent definition shared by document and revision writers', () => {
  const sourceFiles = ['../src/document.ts', '../src/internal/elements.ts', '../src/revisions.ts']
    .map((path) => readFileSync(new URL(path, import.meta.url), 'utf8'));
  assert.equal(sourceFiles.reduce((count, source) => count + [...source.matchAll(/\bconst PROPERTY_ORDER\s*=/g)].length, 0), 1);
  assert.equal(Object.keys(PROPERTY_ORDER).length, 12);
  assert.deepEqual(Object.keys(PROPERTY_ORDER), [
    'pPr', 'rPr', 'paraRPr', 'style', 'tblPr', 'tblPrEx', 'trPr', 'tcPr', 'tblBorders', 'tcBorders', 'tblCellMar', 'tcMar',
  ]);

  const parseProperties = () => new DOMParser().parseFromString(
    `<w:pPr xmlns:w="${WORD_NS}"><w:jc w:val="center"/></w:pPr>`,
    'application/xml',
  ).documentElement;
  const documentProperties = parseProperties();
  property(documentProperties, 'spacing');
  property(documentProperties, 'pPrChange');
  const revisionProperties = parseProperties();
  property(revisionProperties, 'spacing');
  markRevision(revisionProperties, 'pPrChange');
  const childNames = (parent) => Array.from(parent.childNodes)
    .filter((child) => child.nodeType === 1)
    .map((child) => child.localName);
  assert.deepEqual(childNames(revisionProperties), childNames(documentProperties));
});

test('CommentContext borrows only caches and the six instance-bound helpers', () => {
  const source = readFileSync(new URL('../src/comments.ts', import.meta.url), 'utf8');
  const context = source.match(/export interface CommentContext[\s\S]*?\n}/)?.[0] ?? '';
  const borrowed = [...context.matchAll(/^\s{2}(?!\w+:\s*Bound<)(\w+)(?:\??)?(?:\(|:)/gm)].map((match) => match[1]);
  assert.deepEqual(borrowed, [
    'caches',
    'boundaryRun',
    'buildBlocksFromElement',
    'getStylesContext',
    'normalizeRangeOn',
    'normalizeDocumentRange',
    'splitRunAtOffset',
  ]);
});
