import { test } from 'node:test';
import assert from 'node:assert/strict';
import { DocxDocument } from '../dist/document.js';
import { DocxEditor } from '../dist/editor.js';
import { sanitizeTextWithInfo, WORD_NS } from '../dist/xml.js';

function makeFlushEditor({ text, elementText = text, previous = '', options = {}, document = DocxDocument.create() }) {
  const editor = Object.create(DocxEditor.prototype);
  editor.destroyed = false;
  editor.document = document;
  editor.options = options;
  const element = { text: elementText };
  const content = { text };
  editor.paragraphs = new Map([[0, { element, content, text: previous, failed: false }]]);
  editor.readText = (entry) => entry.text;
  return { editor, element, content, document };
}

function makeSelectionEditor(cellByParagraph = {}) {
  const editor = Object.create(DocxEditor.prototype);
  const events = [];
  class FakeCustomEvent {
    constructor(type, init = {}) {
      this.type = type;
      this.detail = init.detail;
      this.bubbles = init.bubbles;
    }
  }
  editor.selected = null;
  editor.selectedImageInfo = null;
  editor.selectedTableCellInfo = null;
  editor.document = { getTableCellAt: (index) => cellByParagraph[index] ?? null };
  editor.root = {
    ownerDocument: { defaultView: { CustomEvent: FakeCustomEvent } },
    dispatchEvent: (event) => { events.push(event); return true; },
  };
  return { editor, events };
}

function makeRunRenderEditor({ showRevisions = true, revisionView = 'markup' } = {}) {
  const editor = Object.create(DocxEditor.prototype);
  editor.reviewFilter = { showRevisions, showComments: true, revisionView };
  editor.options = { showFormattingMarks: false };
  editor.commentParagraphIds = new Map();
  editor.commentRunIds = new Map();
  editor.revisionRunIds = new Map();
  const createElement = (tagName) => {
    const element = {
      nodeType: 1,
      tagName: tagName.toUpperCase(),
      dataset: {},
      style: {},
      childNodes: [],
      className: '',
      attributes: new Map(),
      append(child) { this.childNodes.push(child); },
      setAttribute(name, value) { this.attributes.set(name, value); },
    };
    element.classList = {
      add: (...names) => {
        for (const name of names) {
          if (!name) continue;
          element.className = element.className ? `${element.className} ${name}` : name;
        }
      },
    };
    return element;
  };
  editor.root = {
    ownerDocument: {
      createElement,
      createTextNode: (text) => ({ nodeType: 3, textContent: text }),
    },
  };
  return editor;
}

function appendRunToParagraph(editor, {
  paragraphIndex = 0,
  run,
  reviewContext = { deletedTextByRun: new Map(), revisionColors: new Map() },
} = {}) {
  const paragraphElement = editor.root.ownerDocument.createElement('span');
  const paragraph = { index: paragraphIndex };
  const offset = editor.appendRun(paragraphElement, paragraph, run, reviewContext, 720, 0);
  return { paragraphElement, runSpan: paragraphElement.childNodes[0], offset };
}

test('flush sanitizes disallowed control characters before commit', () => {
  const { editor, document } = makeFlushEditor({ text: 'a\u0001b' });
  assert.doesNotThrow(() => editor.flush());
  assert.equal(document.getParagraphs()[0].text, 'ab');
});

test('flush sanitizes isolated high surrogate before commit', () => {
  const { editor, document } = makeFlushEditor({ text: 'a\ud800b' });
  assert.doesNotThrow(() => editor.flush());
  assert.equal(document.getParagraphs()[0].text, 'ab');
});

test('flush sanitizes isolated low surrogate before commit', () => {
  const { editor, document } = makeFlushEditor({ text: 'a\udc00b' });
  assert.doesNotThrow(() => editor.flush());
  assert.equal(document.getParagraphs()[0].text, 'ab');
});

test('flush preserves tab and newline text', () => {
  const { editor, document } = makeFlushEditor({ text: 'a\tb\nc' });
  assert.doesNotThrow(() => editor.flush());
  assert.equal(document.getParagraphs()[0].text, 'a\tb\nc');
});

test('flush reads inner content text instead of outer paragraph wrapper text', () => {
  const { editor, document } = makeFlushEditor({ text: '用户文本', elementText: '1. 用户文本' });
  assert.doesNotThrow(() => editor.flush());
  assert.equal(document.getParagraphs()[0].text, '用户文本');
});

test('flush truncates overlong input, reports once, and commits legal text', () => {
  const calls = [];
  const overlong = `a${'x'.repeat(1_000_000)}`;
  const { editor, document } = makeFlushEditor({
    text: overlong,
    options: { onError: (error, context) => calls.push({ error, context }) },
  });
  assert.doesNotThrow(() => editor.flush());
  assert.equal(calls.length, 1);
  assert.equal(calls[0].context.paragraph, 0);
  assert.match(calls[0].error.message, /truncated at 1000000 characters/);
  assert.equal(document.getParagraphs()[0].text.length, 1_000_000);
  assert.equal(editor.paragraphs.get(0).text.length, 1_000_000);
});

test('flush does not repeatedly report unchanged overlong text after first truncation', () => {
  const calls = [];
  const overlong = `a${'x'.repeat(1_000_000)}`;
  const { editor } = makeFlushEditor({
    text: overlong,
    options: { onError: () => calls.push('error') },
  });
  editor.flush();
  editor.flush();
  assert.deepEqual(calls, ['error']);
});

test('flush falls back to console.error when onError is not provided', () => {
  const original = console.error;
  const calls = [];
  console.error = (...args) => calls.push(args);
  try {
    const overlong = `a${'x'.repeat(1_000_000)}`;
    const { editor } = makeFlushEditor({ text: overlong });
    assert.doesNotThrow(() => editor.flush());
    assert.equal(calls.length, 1);
  } finally {
    console.error = original;
  }
});

test('flush still logs original error when onError throws', () => {
  const original = console.error;
  const calls = [];
  console.error = (...args) => calls.push(args);
  try {
    const document = DocxDocument.create();
    document.setParagraphText = () => { throw new Error('commit failed'); };
    const { editor } = makeFlushEditor({
      text: 'changed',
      previous: '',
      document,
      options: { onError: () => { throw new Error('handler failed'); } },
    });
    assert.doesNotThrow(() => editor.flush());
    assert.equal(calls.length, 2);
    assert.match(String(calls[0][0]), /commit failed/);
    assert.match(String(calls[1][0]), /handler failed/);
  } finally {
    console.error = original;
  }
});

test('flush does not advance entry.text when commit fails', () => {
  const calls = [];
  const document = DocxDocument.create();
  document.setParagraphText = () => { throw new Error('commit failed'); };
  const { editor } = makeFlushEditor({
    text: 'new text',
    previous: 'old text',
    document,
    options: { onError: () => calls.push('error') },
  });
  assert.doesNotThrow(() => editor.flush());
  assert.equal(editor.paragraphs.get(0).text, 'old text');
  assert.equal(editor.paragraphs.get(0).failed, true);
  assert.deepEqual(calls, ['error']);
});

test('render completes even when flush commit throws', () => {
  const document = DocxDocument.create();
  document.setParagraphText = () => { throw new Error('commit failed'); };
  const { editor } = makeFlushEditor({ text: 'changed', previous: '', document, options: { onError: () => {} } });
  editor.composing = false;
  editor.renderAfterComposition = false;
  editor.selected = null;
  editor.root = {
    ownerDocument: {
      activeElement: null,
      getSelection: () => null,
      createDocumentFragment: () => ({ tag: 'fragment' }),
    },
    contains: () => false,
    replaceChildren: () => {},
  };
  editor.document.getBlocks = () => [];
  assert.doesNotThrow(() => editor.render());
});

test('setDocument and destroy remain usable after flush commit failures', () => {
  const document = DocxDocument.create();
  document.setParagraphText = () => { throw new Error('commit failed'); };
  const { editor } = makeFlushEditor({ text: 'changed', previous: '', document, options: { onError: () => {} } });
  let rendered = 0;
  editor.render = () => { rendered++; };
  const removed = [];
  let rootRemoved = false;
  editor.handleSelection = () => {};
  editor.handleRootKeydown = () => {};
  editor.root = {
    ownerDocument: { removeEventListener: (...args) => removed.push(args) },
    removeEventListener: (...args) => removed.push(args),
    remove: () => { rootRemoved = true; },
  };
  const next = DocxDocument.create();
  assert.doesNotThrow(() => editor.setDocument(next));
  assert.equal(editor.document, next);
  assert.equal(rendered, 1);
  assert.doesNotThrow(() => editor.destroy());
  assert.equal(editor.destroyed, true);
  assert.equal(editor.paragraphs.size, 0);
  assert.equal(rootRemoved, true);
  assert.equal(removed.length, 2);
  assert.equal(removed[0][0], 'selectionchange');
  assert.equal(removed[1][0], 'keydown');
});

test('sanitizeTextWithInfo preserves surrogate pairs when truncating at max length', () => {
  const input = `${'a'.repeat(999_999)}😀`;
  const sanitized = sanitizeTextWithInfo(input);
  assert.equal(sanitized.text.length, 999_999);
  assert.equal(sanitized.text, 'a'.repeat(999_999));
  assert.equal(sanitized.truncated, true);
  assert.equal(sanitized.truncatedAt, 999_999);
});

test('insertText sanitizes invalid paste-like input before insertion', () => {
  const editor = Object.create(DocxEditor.prototype);
  let inserted = '';
  const range = {
    startContainer: {},
    endContainer: {},
    deleteContents: () => {},
    insertNode: (node) => { inserted = node.textContent; },
    setStartAfter: () => {},
    collapse: () => {},
  };
  const selection = {
    rangeCount: 1,
    getRangeAt: () => range,
    removeAllRanges: () => {},
    addRange: () => {},
  };
  editor.root = {
    ownerDocument: {
      getSelection: () => selection,
      createTextNode: (text) => ({ textContent: text }),
    },
  };
  const element = { contains: () => true };
  assert.doesNotThrow(() => editor.insertText(element, 'a\u0001b\ud800c\r\nd'));
  assert.equal(inserted, 'abc\nd');
});

test('selectedTableCell getter and event update when selection enters a table cell', () => {
  const cell = { table: 0, row: 1, col: 2, rowSpan: 1, colSpan: 3, nested: false };
  const { editor, events } = makeSelectionEditor({ 4: cell });
  editor.selectParagraph(4);
  assert.equal(editor.selectedParagraph, 4);
  assert.deepEqual(editor.selectedTableCell, cell);
  assert.deepEqual(events.map((event) => [event.type, event.detail]), [
    ['docx-selectionchange', { index: 4 }],
    ['docx-tablecellchange', { cell }],
  ]);
});

test('table cell change dispatch is debounced within the same table cell while exposed state refreshes', () => {
  const first = { table: 0, row: 0, col: 0, rowSpan: 1, colSpan: 1, nested: false };
  const second = { table: 0, row: 0, col: 0, rowSpan: 2, colSpan: 1, nested: false };
  const { editor, events } = makeSelectionEditor({ 1: first, 2: second });
  editor.selectParagraph(1);
  editor.selectParagraph(2);
  assert.deepEqual(editor.selectedTableCell, second);
  assert.deepEqual(events.map((event) => event.type), [
    'docx-selectionchange',
    'docx-tablecellchange',
    'docx-selectionchange',
  ]);
});

test('selectedTableCell refreshes span changes for the same paragraph without redispatching table cell change', () => {
  const first = { table: 0, row: 0, col: 0, rowSpan: 1, colSpan: 1, nested: false };
  const second = { table: 0, row: 0, col: 0, rowSpan: 1, colSpan: 2, nested: false };
  let current = first;
  const { editor, events } = makeSelectionEditor();
  editor.document = { getTableCellAt: () => current };
  editor.selectParagraph(1);
  current = second;
  editor.selectParagraph(1);
  assert.equal(editor.selectedTableCell.colSpan, 2);
  assert.deepEqual(events.map((event) => event.type), [
    'docx-selectionchange',
    'docx-tablecellchange',
  ]);
});

test('table cell change dispatch fires when selection crosses table cells and when it leaves tables', () => {
  const a = { table: 0, row: 0, col: 0, rowSpan: 1, colSpan: 1, nested: false };
  const b = { table: 0, row: 0, col: 1, rowSpan: 1, colSpan: 1, nested: false };
  const { editor, events } = makeSelectionEditor({ 0: a, 1: b, 2: null });
  editor.selectParagraph(0);
  editor.selectParagraph(1);
  editor.selectParagraph(2);
  assert.equal(editor.selectedTableCell, null);
  assert.deepEqual(events.map((event) => [event.type, event.detail]), [
    ['docx-selectionchange', { index: 0 }],
    ['docx-tablecellchange', { cell: a }],
    ['docx-selectionchange', { index: 1 }],
    ['docx-tablecellchange', { cell: b }],
    ['docx-selectionchange', { index: 2 }],
    ['docx-tablecellchange', null],
  ]);
});

test('readText skips non-editable decorations unless explicitly marked as document content', () => {
  const editor = Object.create(DocxEditor.prototype);
  const text = (value) => ({ nodeType: 3, textContent: value });
  const element = (tagName, { dataset = {}, contentEditable = 'inherit' } = {}, childNodes = []) => ({
    nodeType: 1,
    tagName,
    dataset,
    contentEditable,
    childNodes,
  });
  const root = element('SPAN', {}, [
    element('SPAN', { contentEditable: 'false' }, [text('drop')]),
    element('SPAN', { contentEditable: 'false', dataset: { docxContent: '1' } }, [text('keep')]),
    element('SPAN', { contentEditable: 'false', dataset: { docxMark: '1' } }, [text('drop')]),
  ]);
  assert.equal(editor.readText(root), 'keep');
});

test('applyHistory flushes first and re-renders after undo/redo', () => {
  const calls = [];
  const editor = Object.create(DocxEditor.prototype);
  editor.destroyed = false;
  editor.selected = 0;
  editor.paragraphs = new Map([[0, { content: { focus: () => {} } }]]);
  editor.focusContent = () => {};
  editor.flush = () => calls.push('flush');
  editor.render = () => calls.push('render');
  editor.options = { onChange: () => calls.push('change') };
  editor.document = {
    revision: 1,
    undo: () => ({ revision: 2, paragraphs: [{ index: 0 }] }),
    redo: () => ({ revision: 3, paragraphs: [{ index: 0 }] }),
  };
  editor.applyHistory('undo');
  editor.applyHistory('redo');
  assert.deepEqual(calls, ['flush', 'render', 'change', 'flush', 'render', 'change']);
});

test('applyHistory is a no-op when revision does not change', () => {
  const calls = [];
  const editor = Object.create(DocxEditor.prototype);
  editor.destroyed = false;
  editor.selected = null;
  editor.paragraphs = new Map();
  editor.flush = () => calls.push('flush');
  editor.render = () => calls.push('render');
  editor.options = { onChange: () => calls.push('change') };
  editor.document = {
    revision: 2,
    undo: () => ({ revision: 2, paragraphs: [] }),
    redo: () => ({ revision: 2, paragraphs: [] }),
  };
  editor.applyHistory('undo');
  assert.deepEqual(calls, ['flush']);
});

test('handleHistoryShortcut maps Ctrl/Cmd+Z to undo', () => {
  const editor = Object.create(DocxEditor.prototype);
  const calls = [];
  editor.composing = false;
  editor.applyHistory = (direction) => calls.push(direction);
  const content = {};
  editor.root = { contains: (value) => value === content };
  const event = {
    isComposing: false,
    ctrlKey: true,
    metaKey: false,
    altKey: false,
    shiftKey: false,
    key: 'z',
    preventDefault: () => calls.push('prevent'),
    target: { nodeType: 1, closest: () => content },
  };
  assert.equal(editor.handleHistoryShortcut(event), true);
  assert.deepEqual(calls, ['prevent', 'undo']);
});

test('handleHistoryShortcut maps Cmd+Shift+Z and Ctrl+Y to redo', () => {
  const editor = Object.create(DocxEditor.prototype);
  editor.composing = false;
  const calls = [];
  editor.applyHistory = (direction) => calls.push(direction);
  const content = {};
  editor.root = { contains: (value) => value === content };
  const metaShiftZ = {
    isComposing: false,
    ctrlKey: false,
    metaKey: true,
    altKey: false,
    shiftKey: true,
    key: 'z',
    preventDefault: () => calls.push('prevent:z'),
    target: { nodeType: 1, closest: () => content },
  };
  const ctrlY = {
    isComposing: false,
    ctrlKey: true,
    metaKey: false,
    altKey: false,
    shiftKey: false,
    key: 'y',
    preventDefault: () => calls.push('prevent:y'),
    target: { nodeType: 1, closest: () => content },
  };
  assert.equal(editor.handleHistoryShortcut(metaShiftZ), true);
  assert.equal(editor.handleHistoryShortcut(ctrlY), true);
  assert.deepEqual(calls, ['prevent:z', 'redo', 'prevent:y', 'redo']);
});

test('parseClipboardFragment validates payload shape', () => {
  const editor = Object.create(DocxEditor.prototype);
  const valid = editor.parseClipboardFragment('{"version":1,"text":"x","paragraphs":[{"runs":[{"text":"x"}]}]}');
  const invalid = editor.parseClipboardFragment('{"version":2}');
  assert.equal(valid.version, 1);
  assert.equal(invalid, null);
});

test('writeClipboardFragment writes custom mime, plain html text, and plain text', () => {
  const editor = Object.create(DocxEditor.prototype);
  const written = new Map();
  const transfer = { setData: (type, value) => written.set(type, value) };
  const ok = editor.writeClipboardFragment(transfer, { version: 1, text: 'hello', paragraphs: [{ runs: [{ text: 'hello' }] }] });
  assert.equal(ok, true);
  assert.equal(written.get('application/x-docx-view+json').includes('"version":1'), true);
  assert.equal(written.get('text/plain'), 'hello');
  assert.equal(written.get('text/html'), 'hello');
});

test('mapExternalHtmlFragment strips unsafe href and external images', () => {
  const editor = Object.create(DocxEditor.prototype);
  const text = (value) => ({ nodeType: 3, textContent: value, childNodes: [] });
  const element = (tagName, attrs = {}, childNodes = []) => ({
    nodeType: 1,
    tagName,
    childNodes,
    children: childNodes.filter((child) => child.nodeType === 1),
    getAttribute: (name) => attrs[name] ?? null,
    querySelectorAll: () => [],
  });
  editor.root = {
    ownerDocument: {
      defaultView: {
        DOMParser: class {
          parseFromString() {
            return {
              body: {
                childNodes: [
                  element('DIV', {}, [
                    element('A', { href: 'javascript:alert(1)' }, [text('bad link')]),
                    element('IMG', { src: 'https://example.com/x.png', alt: 'ext' }, []),
                  ]),
                ],
                textContent: 'bad link',
              },
            };
          }
        },
      },
    },
  };
  const fragment = editor.mapExternalHtmlFragment('<div/>', 'fallback');
  assert.equal(fragment.paragraphs[0].runs.some((run) => run.hyperlink?.url), false);
  assert.equal(fragment.paragraphs[0].runs.some((run) => (run.images?.length ?? 0) > 0), false);
});

test('mapExternalHtmlFragment keeps data-image URIs as images', () => {
  const editor = Object.create(DocxEditor.prototype);
  const element = (tagName, attrs = {}, childNodes = []) => ({
    nodeType: 1,
    tagName,
    childNodes,
    children: childNodes.filter((child) => child.nodeType === 1),
    getAttribute: (name) => attrs[name] ?? null,
    querySelectorAll: () => [],
  });
  editor.root = {
    ownerDocument: {
      defaultView: {
        DOMParser: class {
          parseFromString() {
            return {
              body: {
                childNodes: [
                  element('DIV', {}, [
                    element('IMG', { src: 'data:image/png;base64,AAAA', alt: 'ok' }, []),
                  ]),
                ],
                textContent: '',
              },
            };
          }
        },
      },
    },
  };
  const fragment = editor.mapExternalHtmlFragment('<div/>', '');
  assert.equal(fragment.paragraphs[0].runs.some((run) => (run.images?.length ?? 0) === 1), true);
});

test('mapExternalHtmlFragment does not inject list marker text into document runs', () => {
  const editor = Object.create(DocxEditor.prototype);
  const text = (value) => ({ nodeType: 3, textContent: value, childNodes: [] });
  const element = (tagName, attrs = {}, childNodes = []) => ({
    nodeType: 1,
    tagName,
    childNodes,
    children: childNodes.filter((child) => child.nodeType === 1),
    getAttribute: (name) => attrs[name] ?? null,
    querySelectorAll: () => [],
  });
  editor.root = {
    ownerDocument: {
      defaultView: {
        DOMParser: class {
          parseFromString() {
            const li = element('LI', {}, [text('item')]);
            return {
              body: {
                childNodes: [element('UL', {}, [li])],
                textContent: 'item',
              },
            };
          }
        },
      },
    },
  };
  const fragment = editor.mapExternalHtmlFragment('<ul><li>item</li></ul>', 'item');
  assert.equal(fragment.paragraphs[0].runs.map((run) => run.text ?? '').join(''), 'item');
  assert.deepEqual(fragment.paragraphs[0].numbering, { kind: 'bullet', level: 0, listId: 1 });
});

test('mapExternalHtmlFragment maps html tables into table blocks', () => {
  const editor = Object.create(DocxEditor.prototype);
  const text = (value) => ({ nodeType: 3, textContent: value, childNodes: [] });
  const element = (tagName, attrs = {}, childNodes = []) => ({
    nodeType: 1,
    tagName,
    childNodes,
    children: childNodes.filter((child) => child.nodeType === 1),
    getAttribute: (name) => attrs[name] ?? null,
    querySelectorAll: () => [],
  });
  editor.root = {
    ownerDocument: {
      defaultView: {
        DOMParser: class {
          parseFromString() {
            return {
              body: {
                childNodes: [element('TABLE', {}, [
                  element('TR', {}, [
                    element('TD', {}, [text('A')]),
                    element('TD', {}, [text('B')]),
                  ]),
                ])],
                textContent: 'A B',
              },
            };
          }
        },
      },
    },
  };
  const fragment = editor.mapExternalHtmlFragment('<table><tr><td>A</td><td>B</td></tr></table>', '');
  assert.equal(fragment.blocks[0].type, 'table');
  assert.equal(fragment.blocks[0].table.rows[0][0].runs[0].text, 'A');
  assert.equal(fragment.blocks[0].table.rows[0][1].runs[0].text, 'B');
});

test('handleClipboardPaste prefers internal rich fragment and re-renders once', () => {
  const editor = Object.create(DocxEditor.prototype);
  const calls = [];
  editor.root = { ownerDocument: {}, contains: () => true };
  editor.captureDocumentRange = () => ({ start: { paragraph: 0, offset: 0 }, end: { paragraph: 0, offset: 0 } });
  editor.documentRange = (range) => range;
  editor.document = {
    revision: 1,
    pasteClipboardFragment: function () { this.revision = 2; return true; },
    getSnapshot: () => ({ revision: 2 }),
  };
  editor.parseClipboardFragment = () => ({ version: 1, text: 'x', paragraphs: [{ runs: [{ text: 'x' }] }] });
  editor.render = () => calls.push('render');
  editor.options = { onChange: () => calls.push('change') };
  const event = {
    clipboardData: { getData: () => '' },
    preventDefault: () => calls.push('prevent'),
  };
  editor.handleClipboardPaste(event, {});
  assert.deepEqual(calls, ['prevent', 'render', 'change']);
});

test('handleClipboardPaste does not trust HTML marker without custom mime payload', () => {
  const editor = Object.create(DocxEditor.prototype);
  const calls = [];
  editor.root = { ownerDocument: {}, contains: () => true };
  editor.captureDocumentRange = () => ({ start: { paragraph: 0, offset: 0 }, end: { paragraph: 0, offset: 0 } });
  editor.documentRange = (range) => range;
  editor.document = { revision: 1, pasteClipboardFragment: () => false, getSnapshot: () => ({ revision: 1 }) };
  editor.insertText = (_content, text) => calls.push(`plain:${text}`);
  editor.parseClipboardFragment = () => null;
  editor.mapExternalHtmlFragment = (html) => {
    calls.push(`html:${html.includes('data-docx-clip="1"')}`);
    return { version: 1, text: 'safe', paragraphs: [{ runs: [{ text: 'safe' }] }] };
  };
  const event = {
    clipboardData: {
      getData: (type) => type === 'text/plain' ? 'plain' : type === 'text/html' ? '<div data-docx-clip="1" data-docx-payload="..."></div>' : '',
    },
    preventDefault: () => calls.push('prevent'),
  };
  editor.handleClipboardPaste(event, {});
  assert.deepEqual(calls, ['html:true', 'prevent', 'plain:plain']);
});

test('handleClipboardCut wraps deletion in one history group', () => {
  const editor = Object.create(DocxEditor.prototype);
  const calls = [];
  editor.root = { ownerDocument: {}, contains: () => true };
  editor.captureDocumentRange = () => ({ start: { paragraph: 0, offset: 1 }, end: { paragraph: 0, offset: 2 } });
  editor.documentRange = (range) => range;
  editor.document = {
    revision: 3,
    copyClipboardFragment: () => ({ version: 1, text: 'x', paragraphs: [{ runs: [{ text: 'x' }] }] }),
    beginHistoryGroup: (label) => calls.push(`begin:${label}`),
    pasteClipboardFragment: function () { this.revision = 4; return true; },
    endHistoryGroup: () => calls.push('end'),
    getSnapshot: () => ({ revision: 4 }),
  };
  editor.writeClipboardFragment = () => true;
  editor.render = () => calls.push('render');
  editor.options = { onChange: () => calls.push('change') };
  editor.handleClipboardCut({ clipboardData: {}, preventDefault: () => calls.push('prevent') }, {});
  assert.deepEqual(calls, ['begin:cut', 'end', 'prevent', 'render', 'change']);
});

test('handleClipboardCut prevents default for unsupported multi-paragraph ranges', () => {
  const editor = Object.create(DocxEditor.prototype);
  let prevented = false;
  editor.captureDocumentRange = () => ({ start: { paragraph: 0, offset: 0 }, end: { paragraph: 1, offset: 0 } });
  editor.documentRange = (range) => range;
  editor.handleClipboardCut({ preventDefault: () => { prevented = true; } }, {});
  assert.equal(prevented, true);
});

test('setReviewFilter updates render state without triggering onChange', () => {
  const editor = Object.create(DocxEditor.prototype);
  let renderCalls = 0;
  let changes = 0;
  editor.destroyed = false;
  editor.reviewFilter = { showRevisions: true, showComments: true, revisionView: 'markup' };
  editor.options = { onChange: () => { changes++; } };
  editor.flush = () => {};
  editor.render = () => { renderCalls++; };
  editor.setReviewFilter({ authors: [{ kind: 'named', author: 'Alice' }], revisionView: 'final', showComments: false });
  assert.equal(renderCalls, 1);
  assert.equal(changes, 0);
  assert.deepEqual(editor.reviewFilter, {
    authors: [{ kind: 'named', author: 'Alice' }],
    showRevisions: true,
    showComments: false,
    revisionView: 'final',
  });
});

test('setReviewFilter is a no-op when filter is unchanged', () => {
  const editor = Object.create(DocxEditor.prototype);
  let renders = 0;
  editor.destroyed = false;
  editor.reviewFilter = { authors: [{ kind: 'named', author: 'Alice' }], showRevisions: true, showComments: true, revisionView: 'markup' };
  editor.render = () => { renders++; };
  editor.setReviewFilter({ authors: [{ kind: 'named', author: 'Alice' }], showRevisions: true, showComments: true, revisionView: 'markup' });
  assert.equal(renders, 0);
});

test('setReviewFilter merges partial updates without clearing existing fields', () => {
  const editor = Object.create(DocxEditor.prototype);
  editor.destroyed = false;
  editor.reviewFilter = { authors: [{ kind: 'named', author: 'Alice' }], showRevisions: true, showComments: true, revisionView: 'markup' };
  editor.flush = () => {};
  editor.render = () => {};
  editor.setReviewFilter({ showComments: false });
  assert.deepEqual(editor.reviewFilter, {
    authors: [{ kind: 'named', author: 'Alice' }],
    showRevisions: true,
    showComments: false,
    revisionView: 'markup',
  });
});

test('setReviewFilter validates authors as bounded text list', () => {
  const editor = Object.create(DocxEditor.prototype);
  editor.destroyed = false;
  editor.reviewFilter = { showRevisions: true, showComments: true, revisionView: 'markup' };
  editor.flush = () => {};
  editor.render = () => {};
  assert.throws(() => editor.setReviewFilter({ authors: 'Alice' }), /authors must be an array/);
  assert.throws(() => editor.setReviewFilter({ authors: [{ kind: 'named', author: 'A\u0000' }] }), /reviewFilter\.authors\[\]\.author/);
  assert.throws(() => editor.setReviewFilter({ authors: Array.from({ length: 1001 }, (_, index) => ({ kind: 'named', author: String(index) })) }), /at most 1000 items/);
});

test('setReviewFilter validates author kind and raw author consistency', () => {
  const editor = Object.create(DocxEditor.prototype);
  editor.destroyed = false;
  editor.reviewFilter = { showRevisions: true, showComments: true, revisionView: 'markup' };
  editor.flush = () => {};
  editor.render = () => {};
  assert.throws(() => editor.setReviewFilter({ authors: [{ kind: 'unattributed', author: 'Alice' }] }), /must be omitted for unattributed kind/);
  assert.throws(() => editor.setReviewFilter({ authors: [{ kind: 'named' }] }), /is required for named kind/);
  assert.throws(() => editor.setReviewFilter({ authors: [{ kind: 'empty', author: '  ' }] }), /must be an empty string for empty kind/);
  assert.throws(() => editor.setReviewFilter({ authors: [{ kind: 'blank', author: 'Alice' }] }), /must be whitespace-only for blank kind/);
});

test('setReviewFilter accepts getReviewers() author buckets without reshaping', () => {
  const doc = DocxDocument.create();
  doc.setPartXml(doc.mainDocumentPath, `<w:document xmlns:w="${WORD_NS}"><w:body><w:p><w:ins w:id="1"><w:r><w:t>u</w:t></w:r></w:ins><w:ins w:id="2" w:author=""><w:r><w:t>e</w:t></w:r></w:ins><w:ins w:id="3" w:author="   "><w:r><w:t>b</w:t></w:r></w:ins><w:ins w:id="4" w:author="Alice"><w:r><w:t>n</w:t></w:r></w:ins></w:p><w:sectPr/></w:body></w:document>`);
  const authors = doc.getReviewers().map((reviewer) =>
    (reviewer.author === undefined ? { kind: reviewer.kind } : { kind: reviewer.kind, author: reviewer.author }));
  const editor = Object.create(DocxEditor.prototype);
  editor.destroyed = false;
  editor.reviewFilter = { showRevisions: true, showComments: true, revisionView: 'markup' };
  editor.flush = () => {};
  editor.render = () => {};
  assert.doesNotThrow(() => editor.setReviewFilter({ authors }));
  assert.deepEqual(new Set(editor.reviewFilter.authors.map((item) => item.kind)), new Set(['unattributed', 'empty', 'blank', 'named']));
});

test('setReviewFilter validates show flags and revisionView enum', () => {
  const editor = Object.create(DocxEditor.prototype);
  editor.destroyed = false;
  editor.reviewFilter = { showRevisions: true, showComments: true, revisionView: 'markup' };
  editor.flush = () => {};
  editor.render = () => {};
  assert.throws(() => editor.setReviewFilter({ showComments: 'yes' }), /showComments must be boolean/);
  assert.throws(() => editor.setReviewFilter({ showRevisions: 1 }), /showRevisions must be boolean/);
  assert.throws(() => editor.setReviewFilter({ revisionView: 'other' }), /revisionView must be one of/);
});

test('setReviewFilter flushes pending edits across all filter dimensions', () => {
  const cases = [
    { name: 'showComments', filter: { showComments: false } },
    { name: 'showRevisions', filter: { showRevisions: false } },
    { name: 'authors', filter: { authors: [{ kind: 'named', author: 'Alice' }] } },
    { name: 'revisionView original', filter: { revisionView: 'original' } },
    { name: 'revisionView final', filter: { revisionView: 'final' } },
  ];
  for (const sample of cases) {
    const doc = DocxDocument.create();
    const { editor } = makeFlushEditor({
      text: `pending-${sample.name}`,
      previous: '',
      document: doc,
    });
    editor.reviewFilter = { showRevisions: true, showComments: true, revisionView: 'markup' };
    editor.options = {};
    editor.root = { ownerDocument: {} };
    editor.render = function () { this.flush(); };
    const before = doc.revision;
    editor.setReviewFilter(sample.filter);
    assert.equal(doc.getParagraphs()[0].text, `pending-${sample.name}`);
    assert.equal(doc.revision, before + 1);
  }
});

test('flush is a no-op in original review view to avoid projected-text writeback', () => {
  const doc = DocxDocument.create();
  doc.setTrackChanges(true);
  doc.setRevisionAuthor('Alice');
  doc.setParagraphText(0, 'ABXDEF');
  const beforeRevision = doc.revision;
  const beforeRevisions = doc.getRevisions().map((revision) => ({ ...revision }));
  const { editor } = makeFlushEditor({
    text: 'ABCDEFZ',
    previous: 'ABCDEF',
    document: doc,
  });
  editor.reviewFilter = { showRevisions: true, showComments: true, revisionView: 'original' };
  editor.options = {};
  editor.flush();
  assert.equal(doc.revision, beforeRevision);
  assert.deepEqual(doc.getRevisions(), beforeRevisions);
  assert.equal(doc.getParagraphs()[0].text, 'ABXDEF');
});

test('switching from non-markup back to markup does not flush preview DOM text', () => {
  const doc = DocxDocument.create();
  const { editor } = makeFlushEditor({
    text: '',
    previous: '',
    document: doc,
  });
  editor.reviewFilter = { showRevisions: true, showComments: true, revisionView: 'original' };
  editor.options = {};
  editor.root = { ownerDocument: {} };
  editor.render = function () { this.flush(); };
  const before = doc.revision;
  editor.setReviewFilter({ revisionView: 'markup' });
  assert.equal(doc.getParagraphs()[0].text, '');
  assert.equal(doc.revision, before);
});

test('toggling reviewFilter fields does not mutate document revision, text, or XML bytes', () => {
  const doc = DocxDocument.create();
  doc.setTrackChanges(true);
  doc.setRevisionAuthor('Alice');
  doc.setParagraphText(0, 'Alpha');
  const beforeRevision = doc.revision;
  const beforeText = doc.getParagraphs().map((paragraph) => paragraph.text);
  const beforeXml = doc.getPartXml(doc.mainDocumentPath);
  const editor = Object.create(DocxEditor.prototype);
  editor.destroyed = false;
  editor.reviewFilter = { showRevisions: true, showComments: true, revisionView: 'markup' };
  editor.flush = () => {};
  editor.render = () => {};
  editor.setReviewFilter({ showComments: false });
  editor.setReviewFilter({ showComments: true, revisionView: 'original' });
  editor.setReviewFilter({ showRevisions: false, revisionView: 'final' });
  assert.equal(doc.revision, beforeRevision);
  assert.deepEqual(doc.getParagraphs().map((paragraph) => paragraph.text), beforeText);
  assert.equal(doc.getPartXml(doc.mainDocumentPath), beforeXml);
});

test('reviewFilter author narrowing does not change unfiltered getRevisions output', () => {
  const doc = DocxDocument.create();
  doc.setTrackChanges(true);
  doc.setRevisionAuthor('Alice');
  doc.setParagraphText(0, 'Alice text');
  doc.setRevisionAuthor('Bob');
  doc.setParagraphText(0, 'Bob text');
  const before = doc.getRevisions().map((item) => item.id);
  const editor = Object.create(DocxEditor.prototype);
  editor.destroyed = false;
  editor.reviewFilter = { showRevisions: true, showComments: true, revisionView: 'markup' };
  editor.flush = () => {};
  editor.render = () => {};
  editor.setReviewFilter({ authors: [{ kind: 'named', author: 'Alice' }] });
  assert.deepEqual(doc.getRevisions().map((item) => item.id), before);
});

test('reviewScopedRun applies revisionView final/original semantics without mutating source run', () => {
  const editor = Object.create(DocxEditor.prototype);
  editor.reviewFilter = { showRevisions: true, showComments: true, revisionView: 'original' };
  const context = { deletedTextByRun: new Map([['0:0', 'deleted text']]) };
  const deletedRun = { index: 0, text: '', revisions: [{ id: 1, kind: 'deletion' }] };
  const insertionRun = { index: 1, text: 'inserted', revisions: [{ id: 2, kind: 'insertion' }] };
  const moveFromRun = { index: 2, text: '', revisions: [{ id: 3, kind: 'move', move: { name: 'm', side: 'from', pairedId: 4 } }] };
  const moveToRun = { index: 3, text: 'moved', revisions: [{ id: 4, kind: 'move', move: { name: 'm', side: 'to', pairedId: 3 } }] };
  const originalDeleted = editor.reviewScopedRun(0, deletedRun, context);
  const originalInserted = editor.reviewScopedRun(0, insertionRun, context);
  const originalMoveFrom = editor.reviewScopedRun(0, moveFromRun, { deletedTextByRun: new Map([['0:2', 'moved']]) });
  const originalMoveTo = editor.reviewScopedRun(0, moveToRun, context);
  assert.equal(originalDeleted.text, 'deleted text');
  assert.equal(originalInserted.text, '');
  assert.equal(originalMoveFrom.text, 'moved');
  assert.equal(originalMoveTo.text, '');
  assert.equal(deletedRun.text, '');
  assert.equal(insertionRun.text, 'inserted');
  editor.reviewFilter = { showRevisions: true, showComments: true, revisionView: 'final' };
  const finalDeleted = editor.reviewScopedRun(0, { index: 0, text: 'old', revisions: [{ id: 3, kind: 'deletion' }] }, context);
  const finalMoveFrom = editor.reviewScopedRun(0, moveFromRun, context);
  const finalMoveTo = editor.reviewScopedRun(0, moveToRun, context);
  assert.equal(finalDeleted.text, '');
  assert.equal(finalMoveFrom.text, '');
  assert.equal(finalMoveTo.text, 'moved');
});

test('reviewScopedRun author filtering does not collide named and unattributed buckets', () => {
  const editor = Object.create(DocxEditor.prototype);
  editor.reviewFilter = {
    authors: [{ kind: 'named', author: '(unattributed)' }],
    showRevisions: true,
    showComments: true,
    revisionView: 'markup',
  };
  const context = { authors: new Set(['named:(unattributed)']), deletedTextByRun: new Map() };
  const run = {
    index: 0,
    text: 'text',
    revisions: [
      { id: 1, kind: 'insertion', author: '(unattributed)' },
      { id: 2, kind: 'insertion' },
    ],
  };
  const filtered = editor.reviewScopedRun(0, run, context);
  assert.deepEqual(filtered.revisions?.map((revision) => revision.id), [1]);
});

test('revisionAriaDescription includes move label', () => {
  const editor = Object.create(DocxEditor.prototype);
  assert.equal(editor.revisionAriaDescription([{ id: 1, kind: 'move', author: 'Alice', move: { name: 'm', side: 'to' } }]).includes('移动'), true);
});

test('readText skips deleted-text visualization nodes', () => {
  const editor = Object.create(DocxEditor.prototype);
  const text = (value) => ({ nodeType: 3, textContent: value });
  const element = (tagName, { dataset = {}, contentEditable = 'inherit' } = {}, childNodes = []) => ({
    nodeType: 1,
    tagName,
    dataset,
    contentEditable,
    childNodes,
  });
  const root = element('SPAN', {}, [
    text('AB'),
    element('SPAN', { dataset: { docxDeleted: '1' }, contentEditable: 'false' }, [text('C')]),
    text('XDEF'),
  ]);
  assert.equal(editor.readText(root), 'ABXDEF');
});

test('focusRevision validates revision id input', () => {
  const editor = Object.create(DocxEditor.prototype);
  editor.filteredRevisionIds = () => [1];
  editor.setActiveRevision = () => {};
  editor.revisionRunIds = new Map();
  editor.revisionParagraphIds = new Map();
  assert.throws(() => editor.focusRevision(-1), /non-negative integer/);
});

test('focusRevision returns false for filtered-out revision', () => {
  const editor = Object.create(DocxEditor.prototype);
  editor.filteredRevisionIds = () => [2];
  editor.setActiveRevision = () => {};
  editor.revisionRunIds = new Map();
  editor.revisionParagraphIds = new Map();
  assert.equal(editor.focusRevision(1), false);
});

test('focusRevision activates and focuses mapped revision node', () => {
  const editor = Object.create(DocxEditor.prototype);
  const calls = [];
  editor.filteredRevisionIds = () => [3];
  editor.setActiveRevision = (id) => calls.push(`active:${id}`);
  editor.revisionRunIds = new Map([['3', [{ focus: () => calls.push('focus'), scrollIntoView: () => calls.push('scroll') }]]]);
  editor.revisionParagraphIds = new Map();
  assert.equal(editor.focusRevision(3), true);
  assert.deepEqual(calls, ['active:3', 'focus', 'scroll']);
});

test('focusNextRevision cycles through filtered revision ids', () => {
  const editor = Object.create(DocxEditor.prototype);
  const focused = [];
  editor.filteredRevisionIds = () => [10, 20];
  editor.setActiveRevision = (id) => { editor.activeRevisionId = id; };
  editor.revisionRunIds = new Map([
    ['10', [{ focus: () => focused.push(10), scrollIntoView: () => {} }]],
    ['20', [{ focus: () => focused.push(20), scrollIntoView: () => {} }]],
  ]);
  editor.revisionParagraphIds = new Map();
  editor.activeRevisionId = null;
  assert.equal(editor.focusNextRevision(), 10);
  assert.equal(editor.focusNextRevision(), 20);
  assert.deepEqual(focused, [10, 20]);
});

test('focusPreviousRevision cycles backwards through filtered revision ids', () => {
  const editor = Object.create(DocxEditor.prototype);
  const focused = [];
  editor.filteredRevisionIds = () => [10, 20];
  editor.setActiveRevision = (id) => { editor.activeRevisionId = id; };
  editor.revisionRunIds = new Map([
    ['10', [{ focus: () => focused.push(10), scrollIntoView: () => {} }]],
    ['20', [{ focus: () => focused.push(20), scrollIntoView: () => {} }]],
  ]);
  editor.revisionParagraphIds = new Map();
  editor.activeRevisionId = null;
  assert.equal(editor.focusPreviousRevision(), 20);
  assert.equal(editor.focusPreviousRevision(), 10);
  assert.deepEqual(focused, [20, 10]);
});

test('acceptRevision flushes, applies, rerenders, and emits onChange once', () => {
  const editor = Object.create(DocxEditor.prototype);
  const calls = [];
  editor.flush = () => calls.push('flush');
  editor.document = {
    revision: 1,
    acceptRevision: () => { editor.document.revision = 2; },
    getSnapshot: () => ({ revision: 2 }),
  };
  editor.render = () => calls.push('render');
  editor.options = { onChange: () => calls.push('change') };
  editor.setActiveRevision = () => calls.push('clear');
  assert.equal(editor.acceptRevision(1), true);
  assert.deepEqual(calls, ['flush', 'render', 'change', 'clear']);
});

test('acceptRevision returns false when document revision is unchanged', () => {
  const editor = Object.create(DocxEditor.prototype);
  const calls = [];
  editor.flush = () => calls.push('flush');
  editor.document = {
    revision: 3,
    acceptRevision: () => {},
    getSnapshot: () => ({ revision: 3 }),
  };
  editor.render = () => calls.push('render');
  editor.options = { onChange: () => calls.push('change') };
  editor.setActiveRevision = () => calls.push('clear');
  assert.equal(editor.acceptRevision(1), false);
  assert.deepEqual(calls, ['flush']);
});

test('rejectRevision clears active revision after successful mutation', () => {
  const editor = Object.create(DocxEditor.prototype);
  const calls = [];
  editor.flush = () => calls.push('flush');
  editor.document = {
    revision: 4,
    rejectRevision: () => { editor.document.revision = 5; },
    getSnapshot: () => ({ revision: 5 }),
  };
  editor.render = () => calls.push('render');
  editor.options = { onChange: () => calls.push('change') };
  editor.setActiveRevision = (id) => calls.push(`active:${id}`);
  assert.equal(editor.rejectRevision(2), true);
  assert.deepEqual(calls, ['flush', 'render', 'change', 'active:null']);
});

test('acceptAllRevisions and rejectAllRevisions are no-ops when revision is unchanged', () => {
  const editor = Object.create(DocxEditor.prototype);
  const calls = [];
  editor.flush = () => calls.push('flush');
  editor.document = {
    revision: 7,
    acceptAllRevisions: () => {},
    rejectAllRevisions: () => {},
    getSnapshot: () => ({ revision: 7 }),
  };
  editor.render = () => calls.push('render');
  editor.options = { onChange: () => calls.push('change') };
  editor.setActiveRevision = () => calls.push('clear');
  assert.equal(editor.acceptAllRevisions({}), false);
  assert.equal(editor.rejectAllRevisions({}), false);
  assert.deepEqual(calls, ['flush', 'flush']);
});

test('acceptAllRevisions honors non-named reviewer filters via atomic applyOperations', () => {
  const editor = Object.create(DocxEditor.prototype);
  editor.flush = () => {};
  const calls = [];
  editor.document = {
    revision: 1,
    getRevisions: () => [{ id: 2 }, { id: 3, author: '' }],
    applyOperations: ({ operations }) => {
      calls.push(operations.map((operation) => `${operation.type}:${operation.id}`).join(','));
      editor.document.revision = 2;
    },
    getSnapshot: () => ({ revision: 2 }),
  };
  editor.render = () => {};
  editor.options = {};
  editor.setActiveRevision = () => {};
  assert.equal(editor.acceptAllRevisions({ authors: [{ kind: 'unattributed' }, { kind: 'empty', author: '' }] }), true);
  assert.deepEqual(calls, ['acceptRevision:2,acceptRevision:3']);
});

test('rejectAllRevisions with reviewer filter skips mutation when nothing matches', () => {
  const editor = Object.create(DocxEditor.prototype);
  editor.flush = () => {};
  const calls = [];
  editor.document = {
    revision: 9,
    getRevisions: () => [{ id: 1, author: 'Alice' }],
    applyOperations: () => calls.push('apply'),
    getSnapshot: () => ({ revision: 9 }),
  };
  editor.render = () => calls.push('render');
  editor.options = { onChange: () => calls.push('change') };
  editor.setActiveRevision = () => calls.push('active');
  assert.equal(editor.rejectAllRevisions({ authors: [{ kind: 'named', author: 'Bob' }] }), false);
  assert.deepEqual(calls, []);
});

test('setActiveRevision toggles docx-revision-active class by id membership', () => {
  const editor = Object.create(DocxEditor.prototype);
  const states = [];
  const node = {
    dataset: { docxRevisionIds: '1,2' },
    classList: { toggle: (_name, active) => states.push(active) },
  };
  editor.root = { querySelectorAll: () => [node] };
  editor.setActiveRevision(2);
  editor.setActiveRevision(3);
  assert.deepEqual(states, [true, false]);
});

test('appendRun writes data-docx-run for runs without revisions', () => {
  const editor = makeRunRenderEditor();
  const { runSpan } = appendRunToParagraph(editor, { run: { index: 2, text: 'plain' } });
  assert.equal(runSpan.dataset.docxRun, '2');
  assert.equal('docxRevisionIds' in runSpan.dataset, false);
});

test('appendRun omits data-docx-revision-ids for empty revision arrays', () => {
  const editor = makeRunRenderEditor();
  const { runSpan } = appendRunToParagraph(editor, { run: { index: 1, text: 'plain', revisions: [] } });
  assert.equal(runSpan.dataset.docxRun, '1');
  assert.equal('docxRevisionIds' in runSpan.dataset, false);
});

test('appendRun preserves hyperlink and comment datasets alongside run dataset', () => {
  const editor = makeRunRenderEditor();
  editor.commentParagraphIds.set(0, [5]);
  editor.commentRunIds.set('0:3', [7, 5]);
  const { runSpan } = appendRunToParagraph(editor, {
    run: {
      index: 3,
      text: 'link',
      hyperlink: { url: 'https://example.com' },
    },
  });
  assert.equal(runSpan.dataset.docxRun, '3');
  assert.equal(runSpan.dataset.docxLink, '1');
  assert.equal(runSpan.dataset.docxUrl, 'https://example.com');
  assert.equal(runSpan.dataset.docxCommentIds, '5,7');
});

test('appendRun writes distinct revision ids for different runs in the same paragraph', () => {
  const editor = makeRunRenderEditor();
  const paragraphElement = editor.root.ownerDocument.createElement('span');
  const paragraph = { index: 0 };
  const context = { deletedTextByRun: new Map(), revisionColors: new Map() };
  editor.appendRun(paragraphElement, paragraph, { index: 0, text: 'A', revisions: [{ id: 11, kind: 'insertion' }] }, context, 720, 0);
  editor.appendRun(paragraphElement, paragraph, { index: 1, text: 'B', revisions: [{ id: 22, kind: 'deletion' }] }, context, 720, 0);
  assert.equal(paragraphElement.childNodes[0].dataset.docxRevisionIds, '11');
  assert.equal(paragraphElement.childNodes[1].dataset.docxRevisionIds, '22');
  assert.equal(paragraphElement.childNodes[0].dataset.docxRun, '0');
  assert.equal(paragraphElement.childNodes[1].dataset.docxRun, '1');
});

test('appendRun writes all revision ids for a run with multiple revisions', () => {
  const editor = makeRunRenderEditor();
  const { runSpan } = appendRunToParagraph(editor, {
    run: {
      index: 4,
      text: 'AB',
      revisions: [
        { id: 31, kind: 'insertion', author: 'Alice' },
        { id: 32, kind: 'move', author: 'Bob', move: { name: 'm', side: 'to', pairedId: 33 } },
      ],
    },
  });
  assert.equal(runSpan.dataset.docxRevisionIds, '31,32');
  assert.equal(runSpan.dataset.docxRun, '4');
});

test('appendRun omits revision ids when revisions are hidden', () => {
  const editor = makeRunRenderEditor({ showRevisions: false });
  const reviewContext = { deletedTextByRun: new Map(), revisionColors: new Map() };
  const visibleRun = editor.reviewScopedRun(0, { index: 0, text: 'A', revisions: [{ id: 41, kind: 'insertion' }] }, reviewContext);
  const { runSpan } = appendRunToParagraph(editor, { run: visibleRun, reviewContext });
  assert.equal('docxRevisionIds' in runSpan.dataset, false);
  assert.equal(editor.revisionRunIds.size, 0);
});

test('appendRun omits revision ids when revisions are filtered out by author', () => {
  const editor = makeRunRenderEditor();
  const reviewContext = {
    authors: new Set(['named:Bob']),
    deletedTextByRun: new Map(),
    revisionColors: new Map(),
  };
  const visibleRun = editor.reviewScopedRun(0, { index: 0, text: 'A', revisions: [{ id: 42, kind: 'insertion', author: 'Alice' }] }, reviewContext);
  const { runSpan } = appendRunToParagraph(editor, { run: visibleRun, reviewContext });
  assert.equal('docxRevisionIds' in runSpan.dataset, false);
  assert.equal(editor.revisionRunIds.size, 0);
});

test('appendRun exposes revision ids in final view', () => {
  const editor = makeRunRenderEditor({ revisionView: 'final' });
  const reviewContext = { deletedTextByRun: new Map(), revisionColors: new Map() };
  const visibleRun = editor.reviewScopedRun(0, { index: 6, text: 'A', revisions: [{ id: 51, kind: 'deletion' }] }, reviewContext);
  const { runSpan } = appendRunToParagraph(editor, { run: visibleRun, reviewContext });
  assert.equal(runSpan.dataset.docxRevisionIds, '51');
  assert.equal(runSpan.dataset.docxRun, '6');
});

test('appendRun refreshes data-docx-run after rerendered run insertion', () => {
  const editor = makeRunRenderEditor();
  const initial = appendRunToParagraph(editor, { run: { index: 0, text: 'A' } });
  const rerenderedParagraph = editor.root.ownerDocument.createElement('span');
  const paragraph = { index: 0 };
  const context = { deletedTextByRun: new Map(), revisionColors: new Map() };
  editor.appendRun(rerenderedParagraph, paragraph, { index: 0, text: 'X' }, context, 720, 0);
  editor.appendRun(rerenderedParagraph, paragraph, { index: 1, text: 'A' }, context, 720, 0);
  assert.equal(initial.runSpan.dataset.docxRun, '0');
  assert.equal(rerenderedParagraph.childNodes[0].dataset.docxRun, '0');
  assert.equal(rerenderedParagraph.childNodes[1].dataset.docxRun, '1');
});

test('appendRun refreshes data-docx-run after rerendered run deletion', () => {
  const editor = makeRunRenderEditor();
  const initialParagraph = editor.root.ownerDocument.createElement('span');
  const paragraph = { index: 0 };
  const context = { deletedTextByRun: new Map(), revisionColors: new Map() };
  editor.appendRun(initialParagraph, paragraph, { index: 0, text: 'A' }, context, 720, 0);
  editor.appendRun(initialParagraph, paragraph, { index: 1, text: 'B' }, context, 720, 0);
  const rerendered = appendRunToParagraph(editor, { run: { index: 0, text: 'B' } });
  assert.equal(initialParagraph.childNodes[1].dataset.docxRun, '1');
  assert.equal(rerendered.runSpan.dataset.docxRun, '0');
});

test('appendDeletedRunVisualization + readText keeps deleted visualization text out of writeback text', () => {
  const doc = DocxDocument.create();
  doc.setTrackChanges(true);
  doc.setRevisionAuthor('Alice');
  doc.setParagraphText(0, 'ABXDEF');
  const editor = Object.create(DocxEditor.prototype);
  editor.reviewFilter = { showRevisions: true, showComments: true, revisionView: 'markup' };
  editor.revisionRunIds = new Map();
  editor.root = {
    ownerDocument: {
      createElement: (tagName) => ({
        nodeType: 1,
        tagName: tagName.toUpperCase(),
        dataset: {},
        contentEditable: 'inherit',
        style: {},
        childNodes: [],
        append(child) { this.childNodes.push(child); },
        setAttribute: () => {},
      }),
      createTextNode: (value) => ({ nodeType: 3, textContent: value }),
    },
  };
  const content = {
    nodeType: 1,
    tagName: 'SPAN',
    dataset: {},
    contentEditable: 'inherit',
    childNodes: [
      { nodeType: 3, textContent: 'AB' },
      { nodeType: 3, textContent: 'DEF' },
    ],
    append(node) { this.childNodes.splice(1, 0, node); },
  };
  const run = { index: 0, revisions: [{ id: 1, kind: 'deletion', author: 'Alice' }] };
  const context = { deletedTextByRun: new Map([['0:0', 'X']]), revisionColors: new Map() };
  editor.appendDeletedRunVisualization(content, 0, run, context);
  assert.equal(editor.readText(content), 'ABDEF');
  editor.reviewFilter = { showRevisions: true, showComments: true, revisionView: 'final' };
  editor.appendDeletedRunVisualization(content, 0, run, context);
  assert.equal(content.childNodes.length, 3);
});

test('filteredRevisionIds respects showRevisions and author filters', () => {
  const editor = Object.create(DocxEditor.prototype);
  editor.document = {
    getRevisions: () => [
      { id: 1, author: 'Alice' },
      { id: 2, author: 'Bob' },
    ],
  };
  editor.reviewFilter = { showRevisions: false, showComments: true, revisionView: 'markup' };
  assert.deepEqual(editor.filteredRevisionIds(), []);
  editor.reviewFilter = { showRevisions: true, showComments: true, revisionView: 'markup', authors: [{ kind: 'named', author: 'Bob' }] };
  assert.deepEqual(editor.filteredRevisionIds(), [2]);
});

test('filteredRevisionIds treats named author filter as trimmed identity', () => {
  const editor = Object.create(DocxEditor.prototype);
  editor.document = {
    getRevisions: () => [
      { id: 1, author: 'Alice' },
      { id: 2, author: ' Alice' },
      { id: 3, author: 'Alice ' },
    ],
  };
  editor.reviewFilter = { showRevisions: true, showComments: true, revisionView: 'markup', authors: [{ kind: 'named', author: 'Alice' }] };
  assert.deepEqual(editor.filteredRevisionIds(), [1, 2, 3]);
});

test('reviewColor stays stable for each author after accept/reject changes reviewer counts', () => {
  const doc = DocxDocument.create();
  doc.setTrackChanges(true);
  doc.setRevisionAuthor('Bob');
  doc.setParagraphText(0, 'B1');
  doc.setParagraphText(0, 'B2');
  doc.setRevisionAuthor('Alice');
  doc.setParagraphText(0, 'A1');
  const editor = Object.create(DocxEditor.prototype);
  const beforeContext = { deletedTextByRun: new Map(), revisionColors: new Map() };
  const beforeBob = editor.reviewColor('Bob', beforeContext);
  const beforeAlice = editor.reviewColor('Alice', beforeContext);
  const bobRevision = doc.getRevisions().find((revision) => revision.author === 'Bob');
  assert.ok(bobRevision);
  doc.acceptRevision(bobRevision.id);
  const afterContext = { deletedTextByRun: new Map(), revisionColors: new Map() };
  assert.equal(editor.reviewColor('Bob', afterContext), beforeBob);
  assert.equal(editor.reviewColor('Alice', afterContext), beforeAlice);
});

test('appendDeletedRunVisualization emits a non-editable deleted marker only in markup view', () => {
  const editor = Object.create(DocxEditor.prototype);
  const appended = [];
  editor.reviewFilter = { showRevisions: true, showComments: true, revisionView: 'markup' };
  editor.root = {
    ownerDocument: {
      createElement: () => ({
        dataset: {},
        style: {},
        className: '',
        setAttribute: () => {},
      }),
    },
  };
  editor.registerRevisionNode = () => {};
  const parent = { append: (node) => appended.push(node) };
  editor.appendDeletedRunVisualization(parent, 0, { index: 0, revisions: [{ id: 1, kind: 'deletion', author: 'Alice' }] }, {
    deletedTextByRun: new Map([['0:0', 'old']]),
    revisionColors: new Map([['named:Alice', '#2E75B6']]),
  });
  assert.equal(appended.length, 1);
  editor.reviewFilter = { showRevisions: true, showComments: true, revisionView: 'final' };
  editor.appendDeletedRunVisualization(parent, 0, { index: 0, revisions: [{ id: 1, kind: 'deletion', author: 'Alice' }] }, {
    deletedTextByRun: new Map([['0:0', 'old']]),
    revisionColors: new Map([['named:Alice', '#2E75B6']]),
  });
  assert.equal(appended.length, 1);
});

test('appendDeletedRunVisualization marks move-from text distinctly in markup view', () => {
  const editor = Object.create(DocxEditor.prototype);
  const appended = [];
  editor.reviewFilter = { showRevisions: true, showComments: true, revisionView: 'markup' };
  editor.makeMark = (text, label) => ({ dataset: { docxMark: '1' }, textContent: text, title: label });
  editor.root = {
    ownerDocument: {
      createElement: () => ({
        dataset: {},
        style: {},
        className: '',
        children: [],
        setAttribute: () => {},
        append(node) { this.children.push(node); },
      }),
    },
  };
  editor.registerRevisionNode = () => {};
  const parent = { append: (node) => appended.push(node) };
  editor.appendDeletedRunVisualization(parent, 0, {
    index: 0,
    revisions: [{ id: 1, kind: 'move', author: 'Alice', move: { name: 'm', side: 'from', pairedId: 2 } }],
  }, {
    deletedTextByRun: new Map([['0:0', 'old']]),
    revisionColors: new Map([['named:Alice', '#2E75B6']]),
  });
  assert.equal(appended.length, 1);
  assert.equal(appended[0].style.textDecoration.includes('underline'), true);
});
