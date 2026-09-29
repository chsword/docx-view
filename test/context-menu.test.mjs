import { test } from 'node:test';
import assert from 'node:assert/strict';
import {
  contextMenuEntries,
  hitTestContext,
  initializeContextMenu,
  menuNavigationTarget,
  pointIsInsideSelection,
} from '../examples/context-menu.ts';

function context(overrides = {}) {
  return {
    revisionView: 'markup',
    editable: true,
    selection: { paragraph: 0, range: null, format: null, collapsed: true },
    table: null,
    image: null,
    hyperlink: null,
    revisionsAtPoint: [],
    commentsAtPoint: [],
    clipboard: 'unknown',
    source: 'context-menu',
    ...overrides,
  };
}

function hit(position, flags = {}) {
  return {
    position,
    target: {},
    table: position === 'table',
    hyperlink: position === 'hyperlink',
    revision: position === 'revision',
    comment: position === 'comment',
    ...flags,
  };
}

function commands(entries) {
  return entries.flatMap((entry) =>
    entry.type === 'command' ? [entry.command] : entry.type === 'submenu' ? commands(entry.items) : []);
}

test('body menu has the specified command list', () => {
  assert.deepEqual(commands(contextMenuEntries(hit('body'), context())), [
    'clipboard.cut', 'clipboard.copy', 'clipboard.paste',
    'format.bold', 'format.italic', 'format.underline', 'format.fontSize', 'format.fontColor',
    'paragraph.alignment', 'list.indent', 'list.outdent', 'list.bullet', 'list.decimal',
    'comment.addAtSelection', 'hyperlink.insertAtSelection', 'paragraph.style',
  ]);
});

test('table menu prepends every specified table command to the body menu', () => {
  assert.deepEqual(commands(contextMenuEntries(hit('table'), context())), [
    'table.insertRowAbove', 'table.insertRowBelow', 'table.insertColumnLeft', 'table.insertColumnRight',
    'table.deleteRow', 'table.deleteColumn', 'table.deleteTable', 'table.mergeCells', 'table.splitCell',
    'table.applyCellStyle', 'table.applyStyle',
    'clipboard.cut', 'clipboard.copy', 'clipboard.paste',
    'format.bold', 'format.italic', 'format.underline', 'format.fontSize', 'format.fontColor',
    'paragraph.alignment', 'list.indent', 'list.outdent', 'list.bullet', 'list.decimal',
    'comment.addAtSelection', 'hyperlink.insertAtSelection', 'paragraph.style',
  ]);
});

test('image menu omits paste and unsupported size commands', () => {
  assert.deepEqual(commands(contextMenuEntries(hit('image'), context())), [
    'clipboard.cut', 'clipboard.copy', 'image.replace', 'image.setAlt', 'image.delete',
  ]);
});

test('image in a table appends the table command group', () => {
  const ids = commands(contextMenuEntries(hit('image', { table: true }), context()));
  assert.ok(ids.includes('table.applyCellStyle'));
  assert.equal(ids.includes('table.insertRowAbove'), false);
});

test('hyperlink menu includes link commands followed by body commands', () => {
  const ids = commands(contextMenuEntries(hit('hyperlink'), context({ hyperlink: { unsafe: false } })));
  assert.deepEqual(ids.slice(0, 4), ['hyperlink.open', 'hyperlink.copyAddress', 'hyperlink.edit', 'hyperlink.remove']);
  assert.ok(ids.includes('clipboard.paste'));
});

test('unsafe hyperlink menu includes a noninteractive warning', () => {
  const entries = contextMenuEntries(hit('hyperlink'), context({ hyperlink: { unsafe: true } }));
  assert.equal(entries[0].type, 'notice');
});

test('revision menu has no unrelated body commands', () => {
  assert.deepEqual(commands(contextMenuEntries(hit('revision'), context())), [
    'revision.acceptAtPoint', 'revision.rejectAtPoint', 'revision.acceptParagraph',
    'revision.rejectParagraph', 'review.options',
  ]);
});

test('multiple revisions are rendered as individually targeted submenus', () => {
  const revisionsAtPoint = [{ id: 1, kind: 'insertion', author: '甲' }, { id: 2, kind: 'deletion', author: '乙' }];
  const entries = contextMenuEntries(hit('revision'), context({ revisionsAtPoint }));
  assert.equal(entries[0].type, 'submenu');
  assert.deepEqual(entries[0].items.map((item) => item.context.revisionsAtPoint[0].id), [1, 2]);
  assert.equal(entries[1].type, 'submenu');
});

test('comment menu has exactly the four specified commands', () => {
  assert.deepEqual(commands(contextMenuEntries(hit('comment'), context())), [
    'comment.replyAtPoint', 'comment.toggleResolvedAtPoint', 'comment.deleteAtPoint', 'comment.focus',
  ]);
});

test('overlapping comments are rendered as independently targeted groups', () => {
  const entries = contextMenuEntries(hit('comment'), context({ commentsAtPoint: [4, 9] }));
  const groups = entries.filter((entry) => entry.type === 'submenu');
  assert.deepEqual(groups.map((entry) => entry.label), ['批注 4', '批注 9']);
  assert.deepEqual(groups.map((entry) => entry.items[0].context.commentsAtPoint), [[4], [9]]);
});

test('overlapping link revision and comment marks append all matching groups', () => {
  const ids = commands(contextMenuEntries(
    hit('hyperlink', { revision: true, comment: true }),
    context({ hyperlink: { unsafe: false } }),
  ));
  assert.ok(ids.includes('revision.acceptAtPoint'));
  assert.ok(ids.includes('comment.replyAtPoint'));
});

class FakeElement {
  constructor(ownerDocument, attributes = {}) {
    this.ownerDocument = ownerDocument;
    this.nodeType = 1;
    this.attributes = new Map(Object.entries(attributes));
    this.dataset = {};
    this.children = [];
    this.parentElement = null;
    this.listeners = new Map();
    this.style = {};
    this.hidden = false;
    this.tabIndex = -1;
  }
  setAttribute(name, value) { this.attributes.set(name, String(value)); }
  getAttribute(name) { return this.attributes.get(name) ?? null; }
  append(...nodes) { for (const node of nodes) { node.parentElement = this; this.children.push(node); } }
  replaceChildren(...nodes) { this.children = []; this.append(...nodes); }
  contains(node) { return node === this || this.children.some((child) => child.contains(node)); }
  querySelectorAll(selector) {
    if (selector !== '.docx-editor [role="menu"]') return [];
    const editors = [];
    const collectEditors = (node) => {
      if ((node.className ?? '').split(/\s+/).includes('docx-editor')) editors.push(node);
      for (const child of node.children) collectEditors(child);
    };
    collectEditors(this);
    return editors.flatMap((editor) => {
      const menus = [];
      const collectMenus = (node) => {
        if (node !== editor && node.getAttribute('role') === 'menu') menus.push(node);
        for (const child of node.children) collectMenus(child);
      };
      collectMenus(editor);
      return menus;
    });
  }
  addEventListener(name, fn) { this.listeners.set(name, [...(this.listeners.get(name) ?? []), fn]); }
  removeEventListener() {}
  dispatch(name, event = {}) {
    let prevented = false;
    for (const fn of this.listeners.get(name) ?? []) fn({
      target: this,
      preventDefault: () => { prevented = true; },
      ...event,
    });
    return prevented;
  }
  focus() { this.ownerDocument.activeElement = this; this.focused = true; }
  click() { this.dispatch('click'); }
  remove() {
    if (this.parentElement) this.parentElement.children = this.parentElement.children.filter((child) => child !== this);
  }
  closest(selector) {
    const match = /^\[data-([a-z-]+)(?:="([^"]+)")?\]$/.exec(selector);
    if (match) {
      const key = match[1].replace(/-([a-z])/g, (_, letter) => letter.toUpperCase());
      if (key in this.dataset && (match[2] === undefined || this.dataset[key] === match[2])) return this;
    }
    return this.parentElement?.closest(selector) ?? null;
  }
  getBoundingClientRect() { return { left: 10, right: 20, top: 10, bottom: 20 }; }
}

globalThis.HTMLElement = FakeElement;

class FakeDocument {
  constructor(range) {
    this.range = range;
    this.listeners = new Map();
    this.defaultView = { innerWidth: 1000, innerHeight: 700 };
    this.body = new FakeElement(this);
    this.activeElement = null;
    this.selection = {
      rangeCount: 1,
      isCollapsed: false,
      anchorNode: null,
      getRangeAt: () => this.range,
      removeAllRanges: () => { this.selection.rangeCount = 0; },
      addRange: (next) => { this.range = next; this.selection.rangeCount = 1; },
    };
  }
  createElement() { return new FakeElement(this); }
  getSelection() { return this.selection; }
  addEventListener(name, fn) { this.listeners.set(name, fn); }
  removeEventListener(name) { this.listeners.delete(name); }
}

function targetTree(...datasets) {
  const doc = new FakeDocument({
    getClientRects: () => [],
    getBoundingClientRect: () => ({ left: 10, right: 20, top: 10, bottom: 20 }),
    cloneRange() { return this; },
  });
  const root = new FakeElement(doc);
  root.className = 'docx-editor';
  doc.body.append(root);
  let current = root;
  for (const dataset of datasets) {
    const child = new FakeElement(doc);
    child.dataset = dataset;
    current.append(child);
    current = child;
  }
  doc.selection.anchorNode = current;
  return { doc, root, target: current };
}

test('image hit wins over all less-specific hit types', () => {
  const { root, target } = targetTree(
    { tableCell: 'true' },
    { docxLink: '1', docxRevisionIds: '1', docxCommentIds: '2' },
    { image: 'image-1' },
  );
  assert.equal(hitTestContext(target, root).position, 'image');
});

test('table hit wins over link revision and comment hits', () => {
  const { root, target } = targetTree({ tableCell: 'true' }, { docxLink: '1', docxRevisionIds: '1', docxCommentIds: '2' });
  assert.equal(hitTestContext(target, root).position, 'table');
});

test('hyperlink hit wins over revision and comment hits', () => {
  const { root, target } = targetTree({ docxLink: '1', docxRevisionIds: '1', docxCommentIds: '2' });
  assert.equal(hitTestContext(target, root).position, 'hyperlink');
});

test('revision hit wins over a comment hit', () => {
  const { root, target } = targetTree({ docxRevisionIds: '1', docxCommentIds: '2' });
  assert.equal(hitTestContext(target, root).position, 'revision');
});

test('comment hit is detected without a more specific target', () => {
  const { root, target } = targetTree({ docxCommentIds: '2' });
  assert.equal(hitTestContext(target, root).position, 'comment');
});

test('plain editor content resolves to body', () => {
  const { root, target } = targetTree({});
  assert.equal(hitTestContext(target, root).position, 'body');
});

test('point inside any client rect preserves the selection', () => {
  const selection = { rangeCount: 1, isCollapsed: false, getRangeAt: () => ({ getClientRects: () => [{ left: 5, right: 15, top: 5, bottom: 15 }] }) };
  assert.equal(pointIsInsideSelection(selection, 10, 10), true);
});

test('selection rectangle boundaries count as inside', () => {
  const selection = { rangeCount: 1, isCollapsed: false, getRangeAt: () => ({ getClientRects: () => [{ left: 5, right: 15, top: 5, bottom: 15 }] }) };
  assert.equal(pointIsInsideSelection(selection, 5, 15), true);
});

test('point outside all client rects moves the insertion point', () => {
  const selection = { rangeCount: 1, isCollapsed: false, getRangeAt: () => ({ getClientRects: () => [{ left: 5, right: 15, top: 5, bottom: 15 }] }) };
  assert.equal(pointIsInsideSelection(selection, 20, 20), false);
});

test('collapsed selections never retain an unrelated click position', () => {
  assert.equal(pointIsInsideSelection({ rangeCount: 1, isCollapsed: true }, 10, 10), false);
});

test('ArrowDown skips disabled menu items', () => {
  assert.equal(menuNavigationTarget(0, [true, false, true], 'ArrowDown'), 2);
});

test('ArrowUp wraps to the last enabled menu item', () => {
  assert.equal(menuNavigationTarget(0, [true, false, true], 'ArrowUp'), 2);
});

test('unrelated menu navigation keys are ignored', () => {
  assert.equal(menuNavigationTarget(0, [true], 'Tab'), null);
});

function controllerFixture() {
  const { doc, root, target } = targetTree({});
  const calls = [];
  const registry = {
    get(id) {
      return { id, title: id, group: 'test', enabled: () => id !== 'clipboard.cut', checked: id === 'format.bold' ? () => true : undefined, run: () => calls.push(id) };
    },
    list: () => [],
    run: async (id, ctx) => { calls.push([id, ctx.source]); },
  };
  const sources = [];
  const controller = initializeContextMenu({
    editorRoot: root,
    registry,
    buildContext(source) { sources.push(source); return context({ source }); },
  });
  return { doc, root, target, controller, sources, calls };
}

test('Shift+right-click leaves the native menu untouched', () => {
  const { root, target, controller } = controllerFixture();
  assert.equal(root.dispatch('contextmenu', { target, shiftKey: true, clientX: 1, clientY: 1 }), false);
  assert.equal(controller.element.hidden, true);
});

test('plain right-click prevents the native menu and mounts outside the editor', () => {
  const { doc, root, target, controller } = controllerFixture();
  assert.equal(root.dispatch('contextmenu', { target, shiftKey: false, clientX: 1, clientY: 1 }), true);
  assert.equal(controller.element.hidden, false);
  assert.equal(controller.element.getAttribute('role'), 'menu');
  assert.equal(controller.element.getAttribute('aria-label'), '文档编辑上下文菜单');
  assert.equal(root.contains(controller.element), false);
  assert.equal(doc.body.contains(controller.element), true);
  assert.equal(doc.body.querySelectorAll('.docx-editor [role="menu"]').length, 0);
  const bold = controller.element.children.find((item) => item.children.some((child) =>
    child.children.some((nested) => nested.dataset.command === 'format.bold')))
    ?.children[0]?.children.find((item) => item.dataset.command === 'format.bold');
  assert.equal(bold?.getAttribute('role'), 'menuitemcheckbox');
  assert.equal(bold?.getAttribute('aria-checked'), 'true');
});

test('right-click outside a selection asks to move the caret first', () => {
  const { root, target } = targetTree({});
  let moved = 0;
  initializeContextMenu({
    editorRoot: root,
    registry: {
      get: (id) => ({ id, title: id, group: 'test', enabled: () => true, run: () => {} }),
      list: () => [],
      run: async () => {},
    },
    buildContext: () => context(),
    moveCaretToPoint: () => { moved++; },
  });
  root.dispatch('contextmenu', { target, shiftKey: false, clientX: 30, clientY: 30 });
  assert.equal(moved, 1);
});

test('right-click inside a selection preserves it', () => {
  const { doc, root, target } = targetTree({});
  doc.range.getClientRects = () => [{ left: 0, right: 40, top: 0, bottom: 40 }];
  let moved = 0;
  initializeContextMenu({
    editorRoot: root,
    registry: {
      get: (id) => ({ id, title: id, group: 'test', enabled: () => true, run: () => {} }),
      list: () => [],
      run: async () => {},
    },
    buildContext: () => context(),
    moveCaretToPoint: () => { moved++; },
  });
  root.dispatch('contextmenu', { target, shiftKey: false, clientX: 20, clientY: 20 });
  assert.equal(moved, 0);
});

test('Shift+F10 opens through the keyboard CommandContext path', () => {
  const { root, target, controller, sources } = controllerFixture();
  assert.equal(root.dispatch('keydown', { target, key: 'F10', shiftKey: true }), true);
  assert.equal(controller.element.hidden, false);
  assert.equal(sources.at(-1), 'keyboard');
});

test('ContextMenu key opens through the same keyboard path', () => {
  const { root, target, controller, sources } = controllerFixture();
  assert.equal(root.dispatch('keydown', { target, key: 'ContextMenu', shiftKey: false }), true);
  assert.equal(controller.element.hidden, false);
  assert.equal(sources.at(-1), 'keyboard');
});

test('ArrowDown moves focus to the next enabled rendered item', () => {
  const { doc, root, target, controller } = controllerFixture();
  root.dispatch('keydown', { target, key: 'ContextMenu', shiftKey: false });
  const first = doc.activeElement;
  controller.element.dispatch('keydown', { target: first, key: 'ArrowDown' });
  assert.notEqual(doc.activeElement, first);
  assert.equal(doc.activeElement.dataset.command, 'clipboard.paste');
});

test('Escape closes the menu and returns focus', () => {
  const { doc, root, target, controller } = controllerFixture();
  doc.activeElement = target;
  root.dispatch('keydown', { target, key: 'ContextMenu', shiftKey: false });
  const focused = doc.activeElement;
  assert.notEqual(focused, target);
  controller.element.dispatch('keydown', { target: focused, key: 'Escape' });
  assert.equal(controller.element.hidden, true);
  assert.equal(doc.activeElement, target);
});

test('opening and closing without a command leaves document state unchanged', () => {
  const { root, target, controller } = controllerFixture();
  const model = { revision: 7, text: '正文' };
  const before = { ...model };
  root.dispatch('contextmenu', { target, shiftKey: false, clientX: 1, clientY: 1 });
  controller.close();
  const flush = () => {};
  flush();
  assert.deepEqual(model, before);
});
