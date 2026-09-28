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

test('writeClipboardFragment writes custom mime, html marker, and plain text', () => {
  const editor = Object.create(DocxEditor.prototype);
  const written = new Map();
  const transfer = { setData: (type, value) => written.set(type, value) };
  const ok = editor.writeClipboardFragment(transfer, { version: 1, text: 'hello', paragraphs: [{ runs: [{ text: 'hello' }] }] });
  assert.equal(ok, true);
  assert.equal(written.get('application/x-docx-view+json').includes('"version":1'), true);
  assert.equal(written.get('text/plain'), 'hello');
  assert.equal(/data-docx-clip="1"/.test(written.get('text/html')), true);
});

test('parseClipboardFragmentFromHtml reads embedded docx payload', () => {
  const editor = Object.create(DocxEditor.prototype);
  const payload = encodeURIComponent('{"version":1,"text":"ok","paragraphs":[]}');
  editor.root = {
    ownerDocument: {
      defaultView: {
        DOMParser: class {
          parseFromString() {
            return {
              querySelector: () => ({ dataset: { docxPayload: payload } }),
            };
          }
        },
      },
    },
  };
  const parsed = editor.parseClipboardFragmentFromHtml('<div data-docx-clip="1"></div>');
  assert.equal(parsed.text, 'ok');
});

test('isUnsafeHtmlHref blocks script-like schemes', () => {
  const editor = Object.create(DocxEditor.prototype);
  assert.equal(editor.isUnsafeHtmlHref('javascript:alert(1)'), true);
  assert.equal(editor.isUnsafeHtmlHref('vbscript:msgbox(1)'), true);
  assert.equal(editor.isUnsafeHtmlHref('data:text/html;base64,abcd'), true);
  assert.equal(editor.isUnsafeHtmlHref('https://example.com'), false);
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
  editor.parseClipboardFragmentFromHtml = () => null;
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
  editor.parseClipboardFragmentFromHtml = () => ({ version: 1, text: 'evil', paragraphs: [{ runs: [{ text: 'evil' }] }] });
  editor.mapExternalHtmlFragment = () => ({ version: 1, text: 'safe', paragraphs: [{ runs: [{ text: 'safe' }] }] });
  const event = {
    clipboardData: {
      getData: (type) => type === 'text/plain' ? 'plain' : type === 'text/html' ? '<div data-docx-clip="1" data-docx-payload="..."></div>' : '',
    },
    preventDefault: () => calls.push('prevent'),
  };
  editor.handleClipboardPaste(event, {});
  assert.deepEqual(calls, ['prevent', 'plain:plain']);
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
