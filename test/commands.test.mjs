import { test } from 'node:test';
import assert from 'node:assert/strict';
import { createCommandRegistry, createExampleCommandDescriptors } from '../examples/commands.ts';

function makeContext(overrides = {}) {
  return {
    revisionView: 'markup',
    editable: true,
    selection: {
      paragraph: null,
      range: null,
      format: null,
      collapsed: true,
    },
    table: null,
    image: null,
    hyperlink: null,
    revisionsAtPoint: [],
    commentsAtPoint: [],
    clipboard: 'unknown',
    source: 'ribbon',
    ...overrides,
    selection: {
      paragraph: null,
      range: null,
      format: null,
      collapsed: true,
      ...overrides.selection,
    },
  };
}

function makeRegistry(options = {}) {
  const calls = [];
  const state = {
    paragraphs: new Map(),
    revisionCount: 0,
    formatPainterActive: false,
    fontSizeValue: '',
    fontColorValue: '25334a',
    paragraphStyleValue: '',
    alignmentValue: 'left',
    cellFillValue: '#F3F6FB',
    imageAltValue: '',
    ...options,
  };
  const action = (name, transform = (value) => value) => (value) => calls.push([name, transform(value)]);
  const registry = createCommandRegistry(createExampleCommandDescriptors({
    actions: {
      toggleRunFormat: action('toggleRunFormat'),
      activateFormatPainter: action('activateFormatPainter'),
      clearFormat: () => calls.push(['clearFormat']),
      setFontSize: action('setFontSize'),
      setFontColor: action('setFontColor'),
      applyParagraphStyle: action('applyParagraphStyle'),
      setAlignment: action('setAlignment'),
      applyNumbering: action('applyNumbering'),
      changeNumberingLevel: action('changeNumberingLevel'),
      insertTableRow: () => calls.push(['insertTableRow']),
      deleteTableRow: () => calls.push(['deleteTableRow']),
      insertTableColumn: () => calls.push(['insertTableColumn']),
      deleteTableColumn: () => calls.push(['deleteTableColumn']),
      mergeCells: () => calls.push(['mergeCells']),
      splitCell: () => calls.push(['splitCell']),
      applyCellStyle: action('applyCellStyle'),
      startInsertImage: () => calls.push(['startInsertImage']),
      startReplaceImage: () => calls.push(['startReplaceImage']),
      deleteImage: () => calls.push(['deleteImage']),
      setImageAlt: action('setImageAlt'),
      focusPreviousRevision: () => calls.push(['focusPreviousRevision']),
      focusNextRevision: () => calls.push(['focusNextRevision']),
      acceptAllRevisions: () => calls.push(['acceptAllRevisions']),
      rejectAllRevisions: () => calls.push(['rejectAllRevisions']),
      addComment: () => calls.push(['addComment']),
      replyComment: action('replyComment'),
      toggleCommentResolved: action('toggleCommentResolved'),
      deleteComment: action('deleteComment'),
    },
    getParagraphInfo(paragraph) {
      return paragraph === null ? null : state.paragraphs.get(paragraph) ?? null;
    },
    getVisibleRevisionCount() {
      return state.revisionCount;
    },
    isFormatPainterActive() {
      return state.formatPainterActive;
    },
    getFontSizeValue() {
      return state.fontSizeValue;
    },
    getFontColorValue() {
      return state.fontColorValue;
    },
    getParagraphStyleValue() {
      return state.paragraphStyleValue;
    },
    getAlignmentValue() {
      return state.alignmentValue;
    },
    getCellFillValue() {
      return state.cellFillValue;
    },
    getImageAltValue() {
      return state.imageAltValue;
    },
  }));
  return { registry, state, calls };
}

test('command enabled predicates match migrated toolbar behavior', async (t) => {
  const cases = [
    ['format.bold requires a paragraph or range', 'format.bold', makeContext(), false],
    ['format.bold accepts a selected paragraph', 'format.bold', makeContext({ selection: { paragraph: 2 } }), true],
    ['format.italic accepts a selected range', 'format.italic', makeContext({ selection: { range: { start: { paragraph: 1, offset: 0 }, end: { paragraph: 1, offset: 2 } }, collapsed: false } }), true],
    ['format.clear blocks readonly edits', 'format.clear', makeContext({ editable: false, selection: { paragraph: 0 } }), false],
    ['format.clear keeps markup edits enabled', 'format.clear', makeContext({ selection: { paragraph: 0 } }), true],
    ['paragraph.style requires markup paragraph selection', 'paragraph.style', makeContext({ editable: false, selection: { paragraph: 0 } }), false],
    ['paragraph.alignment enables on markup paragraph selection', 'paragraph.alignment', makeContext({ selection: { paragraph: 0 } }), true],
    ['list.bullet disables outside a paragraph', 'list.bullet', makeContext(), false],
    ['list.bullet enables in markup view', 'list.bullet', makeContext({ selection: { paragraph: 0 } }), true],
    ['list.indent disables without numbering', 'list.indent', makeContext({ selection: { paragraph: 0 } }), false],
    ['list.indent disables at max level', 'list.indent', makeContext({ selection: { paragraph: 0 } }), false, { paragraphs: new Map([[0, { numbering: { level: 8, isBullet: true } }]]) }],
    ['list.indent enables below max level', 'list.indent', makeContext({ selection: { paragraph: 0 } }), true, { paragraphs: new Map([[0, { numbering: { level: 7, isBullet: true } }]]) }],
    ['list.outdent disables at base level', 'list.outdent', makeContext({ selection: { paragraph: 0 } }), false, { paragraphs: new Map([[0, { numbering: { level: 0, isBullet: true } }]]) }],
    ['list.outdent enables above base level', 'list.outdent', makeContext({ selection: { paragraph: 0 } }), true, { paragraphs: new Map([[0, { numbering: { level: 1, isBullet: false } }]]) }],
    ['table.insertRow remains enabled without table context', 'table.insertRow', makeContext(), true],
    ['image.replace disables without a selected image', 'image.replace', makeContext(), false],
    ['image.replace enables with a selected image', 'image.replace', makeContext({ image: { relationshipId: 'rId5' } }), true],
    ['comment.reply disables without a selected comment', 'comment.reply', makeContext(), false],
    ['comment.reply enables with a selected comment', 'comment.reply', makeContext({ commentsAtPoint: [42] }), true],
    ['review.acceptAll disables when no revisions are visible', 'review.acceptAll', makeContext(), false],
    ['review.acceptAll enables when revisions are visible', 'review.acceptAll', makeContext(), true, { revisionCount: 3 }],
  ];
  for (const [name, commandId, ctx, expected, options] of cases) {
    await t.test(name, () => {
      const { registry } = makeRegistry(options);
      assert.equal(registry.get(commandId).enabled(ctx), expected);
    });
  }
});

test('command checked predicates match migrated toolbar behavior', async (t) => {
  const cases = [
    ['format.bold mirrors selected range format', 'format.bold', makeContext({ selection: { format: { bold: true } } }), true],
    ['format.painter mirrors painter lock state', 'format.painter', makeContext({ selection: { paragraph: 0 } }), true, { formatPainterActive: true }],
    ['list.bullet reflects bullet numbering', 'list.bullet', makeContext({ selection: { paragraph: 0 } }), true, { paragraphs: new Map([[0, { numbering: { level: 0, isBullet: true } }]]) }],
    ['list.decimal reflects decimal numbering', 'list.decimal', makeContext({ selection: { paragraph: 0 } }), true, { paragraphs: new Map([[0, { numbering: { level: 0, isBullet: false } }]]) }],
  ];
  for (const [name, commandId, ctx, expected, options] of cases) {
    await t.test(name, () => {
      const { registry } = makeRegistry(options);
      assert.equal(registry.get(commandId).checked?.(ctx), expected);
    });
  }
});

test('command run delegates to the shared action registry', async (t) => {
  await t.test('format.fontSize parses the current control value', async () => {
    const { registry, calls } = makeRegistry({ fontSizeValue: '24' });
    await registry.run('format.fontSize', makeContext({ selection: { paragraph: 0 } }));
    assert.deepEqual(calls, [['setFontSize', 24]]);
  });

  await t.test('table.applyCellStyle passes the current cell fill value', async () => {
    const { registry, calls } = makeRegistry({ cellFillValue: '#ABCDEF' });
    await registry.run('table.applyCellStyle', makeContext());
    assert.deepEqual(calls, [['applyCellStyle', '#ABCDEF']]);
  });

  await t.test('comment.reply uses the selected comment id from context', async () => {
    const { registry, calls } = makeRegistry();
    await registry.run('comment.reply', makeContext({ commentsAtPoint: [7] }));
    assert.deepEqual(calls, [['replyComment', 7]]);
  });

  await t.test('review.nextRevision delegates to the review action', async () => {
    const { registry, calls } = makeRegistry();
    await registry.run('review.nextRevision', makeContext());
    assert.deepEqual(calls, [['focusNextRevision']]);
  });
});

test('command registry rejects duplicate and unknown command ids', async () => {
  assert.throws(() => createCommandRegistry([
    { id: 'duplicate', title: 'A', group: 'g', enabled: () => true, run: () => {} },
    { id: 'duplicate', title: 'B', group: 'g', enabled: () => true, run: () => {} },
  ]), /重复的命令 id/);

  const { registry } = makeRegistry();
  await assert.rejects(() => registry.run('missing.command', makeContext()), /未知命令/);
});
