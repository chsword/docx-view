import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { initializeRibbon, ribbonNavigationTarget } from '../examples/ribbon.ts';

class FakeElement {
  constructor(id = '', attributes = {}) {
    this.id = id;
    this.attributes = new Map(Object.entries(attributes));
    this.listeners = new Map();
    this.hidden = false;
    this.tabIndex = -1;
    this.checked = true;
    this.dataset = {};
    this.classes = new Set();
    this.classList = {
      toggle: (name, force) => {
        if (force ?? !this.classes.has(name)) this.classes.add(name);
        else this.classes.delete(name);
      },
    };
  }

  getAttribute(name) {
    return this.attributes.get(name) ?? null;
  }

  setAttribute(name, value) {
    this.attributes.set(name, value);
  }

  addEventListener(name, listener) {
    const listeners = this.listeners.get(name) ?? [];
    listeners.push(listener);
    this.listeners.set(name, listeners);
  }

  dispatch(name, event = {}) {
    let prevented = false;
    for (const listener of this.listeners.get(name) ?? []) {
      listener({ preventDefault: () => { prevented = true; }, ...event });
    }
    return prevented;
  }

  focus() {
    this.focused = true;
  }
}

function makeRibbon(selectedIndex = 0) {
  const tabs = Array.from({ length: 5 }, (_, index) => new FakeElement(`ribbon-tab-${index}`, {
    'aria-controls': `ribbon-panel-${index}`,
    'aria-selected': String(index === selectedIndex),
  }));
  const panels = tabs.map((tab, index) => new FakeElement(`ribbon-panel-${index}`));
  const outline = new FakeElement('outline-panel');
  const toggle = new FakeElement('toggle-outline');
  toggle.dataset.ribbonToggle = 'outline-panel';
  const root = new FakeElement('ribbon');
  root.querySelectorAll = (selector) => {
    if (selector === '[role="tab"]') return tabs;
    if (selector === '[role="tabpanel"]') return panels;
    if (selector === '[data-ribbon-toggle]') return [toggle];
    return [];
  };
  root.ownerDocument = { getElementById: (id) => id === outline.id ? outline : null };
  return { root, tabs, panels, outline, toggle };
}

test('Ribbon activation keeps tab, panel and focus state synchronized', () => {
  const { root, tabs, panels } = makeRibbon(2);
  initializeRibbon(root);

  assert.equal(root.getAttribute('contenteditable'), 'false');
  assert.equal(tabs[2].getAttribute('aria-selected'), 'true');
  assert.equal(tabs[2].tabIndex, 0);
  assert.equal(panels[2].hidden, false);
  assert.equal(tabs.filter((tab) => tab.getAttribute('aria-selected') === 'true').length, 1);
  assert.equal(panels.filter((panel) => !panel.hidden).length, 1);

  assert.equal(tabs[2].dispatch('keydown', { key: 'ArrowRight' }), true);
  assert.equal(tabs[3].getAttribute('aria-selected'), 'true');
  assert.equal(tabs[3].tabIndex, 0);
  assert.equal(tabs[3].focused, true);
  assert.equal(panels[2].hidden, true);
  assert.equal(panels[3].hidden, false);

  tabs[0].dispatch('click');
  assert.equal(tabs[0].getAttribute('aria-selected'), 'true');
  assert.equal(tabs[0].tabIndex, 0);
  assert.equal(panels[0].hidden, false);
  assert.equal(tabs[3].tabIndex, -1);
});

test('Ribbon navigation keys wrap, jump to ends and ignore unrelated keys', async (t) => {
  const cases = [
    ['ArrowRight advances', 1, 5, 'ArrowRight', 2],
    ['ArrowRight wraps', 4, 5, 'ArrowRight', 0],
    ['ArrowLeft moves back', 3, 5, 'ArrowLeft', 2],
    ['ArrowLeft wraps', 0, 5, 'ArrowLeft', 4],
    ['Home selects first', 3, 5, 'Home', 0],
    ['End selects last', 1, 5, 'End', 4],
    ['unknown key is ignored', 2, 5, 'Tab', null],
    ['zero tabs have no target', 0, 0, 'ArrowRight', null],
    ['negative index is ignored', -1, 5, 'ArrowRight', null],
    ['past-end index is ignored', 5, 5, 'ArrowLeft', null],
  ];
  for (const [name, index, count, key, expected] of cases) {
    await t.test(name, () => assert.equal(ribbonNavigationTarget(index, count, key), expected));
  }
});

test('Ribbon view toggle only changes its target visibility', () => {
  const { root, outline, toggle } = makeRibbon();
  initializeRibbon(root);
  assert.equal(outline.hidden, false);
  toggle.checked = false;
  toggle.dispatch('change');
  assert.equal(outline.hidden, true);
  toggle.checked = true;
  toggle.dispatch('change');
  assert.equal(outline.hidden, false);
});

test('Ribbon panels contain every migrated control once and keep developer tabs separate', () => {
  const html = readFileSync(new URL('../examples/index.html', import.meta.url), 'utf8');
  const ribbon = html.slice(html.indexOf('id="ribbon"'), html.indexOf('<div class="paper-stage"'));
  const controls = [
    'format-bold', 'format-italic', 'format-underline', 'font-size', 'font-color', 'format-painter', 'clear-format',
    'paragraph-style', 'alignment', 'list-bullet', 'list-decimal', 'list-outdent', 'list-indent', 'add-paragraph',
    'table-rows', 'table-cols', 'add-table', 'insert-row', 'delete-row', 'insert-col', 'delete-col', 'merge-cells',
    'split-cell', 'cell-fill', 'apply-cell-style', 'insert-image', 'replace-image', 'delete-image', 'image-alt',
    'page-size', 'page-orientation', 'apply-page-setup', 'insert-section-break', 'insert-page-break', 'header-kind',
    'footer-kind', 'add-footnote', 'add-endnote', 'review-revision-view', 'review-show-revisions', 'review-show-comments',
    'review-select-all', 'review-clear-authors', 'reviewer-list', 'toggle-comments', 'comment-author-filter',
    'comment-resolved-filter', 'new-comment', 'reply-comment', 'resolve-comment', 'delete-comment',
  ];
  const tabs = [...ribbon.matchAll(/<button\b[^>]*role="tab"[^>]*>(.*?)<\/button>/g)].map((match) => match[1]);
  assert.deepEqual(tabs, ['开始', '插入', '布局', '审阅', '视图']);
  assert.doesNotMatch(ribbon, /引用/);
  for (const id of controls) {
    assert.equal((ribbon.match(new RegExp(`id="${id}"`, 'g')) ?? []).length, 1, `${id} occurs once in Ribbon`);
  }
  assert.match(ribbon, /id="add-footnote"[\s\S]*id="add-endnote"/);
  assert.match(html, /class="tabs" role="tablist" aria-label="开发者工具"/);
  assert.doesNotMatch(ribbon, /tab-agent|tab-xml|tab-snapshot/);
});
