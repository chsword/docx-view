import { test } from 'node:test';
import assert from 'node:assert/strict';
import { DocxDocument } from '../dist/document.js';
import { DocxEditor } from '../dist/editor.js';
import { sanitizeTextWithInfo } from '../dist/xml.js';

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
  editor.render = () => { renderCalls++; };
  editor.setReviewFilter({ authors: ['Alice'], revisionView: 'final', showComments: false });
  assert.equal(renderCalls, 1);
  assert.equal(changes, 0);
  assert.deepEqual(editor.reviewFilter, {
    authors: ['Alice'],
    showRevisions: true,
    showComments: false,
    revisionView: 'final',
  });
});

test('setReviewFilter is a no-op when filter is unchanged', () => {
  const editor = Object.create(DocxEditor.prototype);
  let renders = 0;
  editor.destroyed = false;
  editor.reviewFilter = { authors: ['Alice'], showRevisions: true, showComments: true, revisionView: 'markup' };
  editor.render = () => { renders++; };
  editor.setReviewFilter({ authors: ['Alice'], showRevisions: true, showComments: true, revisionView: 'markup' });
  assert.equal(renders, 0);
});

test('setReviewFilter merges partial updates without clearing existing fields', () => {
  const editor = Object.create(DocxEditor.prototype);
  editor.destroyed = false;
  editor.reviewFilter = { authors: ['Alice'], showRevisions: true, showComments: true, revisionView: 'markup' };
  editor.render = () => {};
  editor.setReviewFilter({ showComments: false });
  assert.deepEqual(editor.reviewFilter, {
    authors: ['Alice'],
    showRevisions: true,
    showComments: false,
    revisionView: 'markup',
  });
});

test('setReviewFilter validates authors as bounded text list', () => {
  const editor = Object.create(DocxEditor.prototype);
  editor.destroyed = false;
  editor.reviewFilter = { showRevisions: true, showComments: true, revisionView: 'markup' };
  editor.render = () => {};
  assert.throws(() => editor.setReviewFilter({ authors: 'Alice' }), /authors must be an array/);
  assert.throws(() => editor.setReviewFilter({ authors: ['A\u0000'] }), /reviewFilter\.authors\[\]/);
  assert.throws(() => editor.setReviewFilter({ authors: Array.from({ length: 1001 }, (_, index) => String(index)) }), /at most 1000 items/);
});

test('setReviewFilter validates show flags and revisionView enum', () => {
  const editor = Object.create(DocxEditor.prototype);
  editor.destroyed = false;
  editor.reviewFilter = { showRevisions: true, showComments: true, revisionView: 'markup' };
  editor.render = () => {};
  assert.throws(() => editor.setReviewFilter({ showComments: 'yes' }), /showComments must be boolean/);
  assert.throws(() => editor.setReviewFilter({ showRevisions: 1 }), /showRevisions must be boolean/);
  assert.throws(() => editor.setReviewFilter({ revisionView: 'other' }), /revisionView must be one of/);
});

test('setReviewFilter flushes pending edits instead of discarding them', () => {
  const doc = DocxDocument.create();
  const { editor } = makeFlushEditor({
    text: 'pending',
    previous: '',
    document: doc,
  });
  editor.reviewFilter = { showRevisions: true, showComments: true, revisionView: 'markup' };
  editor.options = {};
  editor.root = { ownerDocument: {} };
  editor.render = function () { this.flush(); };
  const before = doc.revision;
  editor.setReviewFilter({ showComments: false });
  assert.equal(doc.getParagraphs()[0].text, 'pending');
  assert.equal(doc.revision, before + 1);
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
  editor.render = () => {};
  editor.setReviewFilter({ authors: ['Alice'] });
  assert.deepEqual(doc.getRevisions().map((item) => item.id), before);
});

test('reviewScopedRun applies revisionView final/original semantics without mutating source run', () => {
  const editor = Object.create(DocxEditor.prototype);
  editor.reviewFilter = { showRevisions: true, showComments: true, revisionView: 'original' };
  editor.activeReviewDeletedTextByRun = new Map([['0:0', 'deleted text']]);
  const deletedRun = { index: 0, text: '', revisions: [{ id: 1, kind: 'deletion' }] };
  const insertionRun = { index: 1, text: 'inserted', revisions: [{ id: 2, kind: 'insertion' }] };
  const originalDeleted = editor.reviewScopedRun(0, deletedRun);
  const originalInserted = editor.reviewScopedRun(0, insertionRun);
  assert.equal(originalDeleted.text, 'deleted text');
  assert.equal(originalInserted.text, '');
  assert.equal(deletedRun.text, '');
  assert.equal(insertionRun.text, 'inserted');
  editor.reviewFilter = { showRevisions: true, showComments: true, revisionView: 'final' };
  const finalDeleted = editor.reviewScopedRun(0, { index: 0, text: 'old', revisions: [{ id: 3, kind: 'deletion' }] });
  assert.equal(finalDeleted.text, '');
});
