import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { customGeometryPath, presetGeometryPath } from '../dist/geometry.js';

const presets = [
  'rect', 'roundRect', 'ellipse', 'triangle', 'rtTriangle', 'diamond', 'parallelogram', 'trapezoid',
  'pentagon', 'hexagon', 'star5', 'rightArrow', 'leftArrow', 'upArrow', 'downArrow',
  'leftRightArrow', 'line', 'straightConnector1', 'wedgeRectCallout', 'cloudCallout',
];

test('preset shape geometry returns deterministic paths for supported names', () => {
  for (const name of presets) {
    const path = presetGeometryPath(name, 120, 80);
    assert.equal(typeof path, 'string', name);
    assert.equal(path, presetGeometryPath(name, 120, 80), name);
  }
  assert.equal(presetGeometryPath('flowChartMagneticDisk', 120, 80), undefined);
});

test('left arrow adjustments are clamped to keep the path within its bounds', () => {
  assert.equal(
    presetGeometryPath('leftArrow', 120, 80, new Map([['adj1', 5], ['adj2', 5]])),
    presetGeometryPath('leftArrow', 120, 80, new Map([['adj1', 0.8], ['adj2', 0.45]])),
  );
});

test('custom geometry scales supported path commands and rejects unsupported commands', () => {
  const geometry = {
    width: 100,
    height: 50,
    commands: [
      { type: 'moveTo', x: 0, y: 0 },
      { type: 'lnTo', x: 50, y: 25 },
      { type: 'cubicBezTo', x1: 60, y1: 20, x2: 80, y2: 10, x: 100, y: 50 },
      { type: 'close' },
    ],
  };
  assert.equal(customGeometryPath(geometry, 200, 100), 'M 0 0 L 100 50 C 120 40 160 20 200 100 Z');
  assert.equal(customGeometryPath({ ...geometry, commands: [{ type: 'arcTo' }] }, 200, 100), undefined);
});

test('geometry module imports only its type definitions and has no DOM dependencies', () => {
  const source = readFileSync(new URL('../src/geometry.ts', import.meta.url), 'utf8');
  assert.doesNotMatch(source, /\b(?:document|window)\.|HTMLElement/);
  assert.deepEqual([...source.matchAll(/from ['"]([^'"]+)['"]/g)].map((match) => match[1]), ['./types.js']);
});
