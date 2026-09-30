import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { axisTicks, barRects, pieSlicePath, valueToPx } from '../dist/chart.js';

test('axisTicks handles degenerate and non-finite ranges', () => {
  for (const [min, max] of [[0, 0], [-5, -5], [-5, 5], [2, 2], [NaN, Infinity]]) {
    const result = axisTicks(min, max, 5);
    assert.ok(Number.isFinite(result.min) && Number.isFinite(result.max) && Number.isFinite(result.step));
    assert.ok(result.ticks.length > 0);
    assert.ok(result.ticks.every(Number.isFinite));
  }
  assert.deepEqual(axisTicks(0, 10, 5).ticks, [0, 2, 4, 6, 8, 10]);
});

test('chart geometry calculates values, grouped and stacked bars, and slices', () => {
  assert.equal(valueToPx(5, { min: 0, max: 10 }, 100), 50);
  const clustered = barRects([[1, 2], [3, 4]], { min: 0, max: 4 }, 100, 80, { grouping: 'clustered' });
  assert.equal(clustered.length, 2);
  assert.ok(clustered[0][0].width > 0);
  const stacked = barRects([[1, 2], [3, 4]], { min: 0, max: 5 }, 100, 80, { grouping: 'stacked' });
  assert.ok(stacked[1][0].y < stacked[0][0].y);
  assert.match(pieSlicePath(0, Math.PI, 20, 20, 10), /^M 20 20 L/);
  assert.match(pieSlicePath(0, Math.PI, 20, 20, 10, 5), /A 5 5/);
});

test('chart geometry has no DOM dependencies', () => {
  const source = readFileSync(new URL('../src/chart.ts', import.meta.url), 'utf8');
  assert.doesNotMatch(source, /\b(?:document|window)\.|HTMLElement/);
  assert.deepEqual([...source.matchAll(/from ['"]([^'"]+)['"]/g)].map((match) => match[1]), []);
});
