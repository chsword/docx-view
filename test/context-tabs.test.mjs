import { test } from 'node:test';
import assert from 'node:assert/strict';
import { RibbonContextState } from '../examples/context-tabs.ts';

const fixedTabs = ['home', 'insert', 'layout', 'review', 'view'];
const contextTabs = ['table', 'image'];

function makeState() {
  return new RibbonContextState(fixedTabs, contextTabs);
}

test('context tab activation and restoration rules', async (t) => {
  const cases = [
    ['a newly available table tab activates and remembers the current tab', (state) => {
      state.manualActivate('insert');
      assert.equal(state.setContextVisible('table', true), 'table');
      assert.equal(state.activeTabId, 'table');
      assert.equal(state.preContextTabId, 'insert');
    }],
    ['a newly available image tab activates and remembers the current tab', (state) => {
      assert.equal(state.setContextVisible('image', true), 'image');
      assert.equal(state.preContextTabId, 'home');
    }],
    ['a disappearing active context tab restores its previous tab', (state) => {
      state.setContextVisible('table', true);
      assert.equal(state.setContextVisible('table', false), 'home');
      assert.equal(state.activeTabId, 'home');
    }],
    ['a disappearing inactive context tab does not change the active tab', (state) => {
      state.setContextVisible('table', true);
      state.manualActivate('review');
      assert.equal(state.setContextVisible('table', false), null);
      assert.equal(state.activeTabId, 'review');
    }],
    ['the last context event wins when both contexts are available', (state) => {
      state.setContextVisible('table', true);
      assert.equal(state.setContextVisible('image', true), 'image');
      assert.equal(state.activeTabId, 'image');
    }],
    ['the remaining context can be restored after the active context disappears', (state) => {
      state.setContextVisible('table', true);
      state.setContextVisible('image', true);
      assert.equal(state.setContextVisible('image', false), 'table');
      assert.equal(state.activeTabId, 'table');
    }],
    ['a removed previous context falls back to the last manual tab', (state) => {
      state.manualActivate('layout');
      state.setContextVisible('table', true);
      state.setContextVisible('image', true);
      state.setContextVisible('table', false);
      assert.equal(state.setContextVisible('image', false), 'layout');
    }],
    ['manual activation records fixed tabs without replacing the previous-context id', (state) => {
      state.setContextVisible('table', true);
      state.manualActivate('review');
      assert.equal(state.activeTabId, 'review');
      assert.equal(state.lastManualTabId, 'review');
      assert.equal(state.preContextTabId, 'home');
    }],
    ['manual activation cannot select a contextual tab', (state) => {
      state.setContextVisible('table', true);
      state.manualActivate('image');
      assert.equal(state.activeTabId, 'table');
    }],
    ['repeated visible notifications do not steal focus back from a manually selected tab', (state) => {
      state.setContextVisible('table', true);
      state.manualActivate('insert');
      assert.equal(state.setContextVisible('table', true), null);
      assert.equal(state.activeTabId, 'insert');
    }],
    ['unknown context tabs are ignored', (state) => {
      assert.equal(state.setContextVisible('unknown', true), null);
      assert.equal(state.activeTabId, 'home');
    }],
    ['a context that was never visible cannot cause a restoration jump', (state) => {
      assert.equal(state.setContextVisible('table', false), null);
      assert.equal(state.activeTabId, 'home');
    }],
  ];

  for (const [name, run] of cases) {
    await t.test(name, () => run(makeState()));
  }
});
