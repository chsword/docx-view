import { test } from 'node:test';
import assert from 'node:assert/strict';
import {
  createCommandContextBuilder,
  createCommandRegistry,
  createExampleCommandDescriptors,
  getCommandControlState,
} from '../examples/commands.ts';

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
    shape: null,
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
      openStyleDialog: action('openStyleDialog'),
      openMathDialog: () => calls.push(['openMathDialog']),
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
      startReplaceImage: (image) => calls.push(['startReplaceImage', image]),
      deleteImage: (image) => calls.push(['deleteImage', image]),
      setImageAlt: (text, image) => calls.push(['setImageAlt', text, image]),
      focusPreviousRevision: () => calls.push(['focusPreviousRevision']),
      focusNextRevision: () => calls.push(['focusNextRevision']),
      acceptAllRevisions: () => calls.push(['acceptAllRevisions']),
      rejectAllRevisions: () => calls.push(['rejectAllRevisions']),
      acceptRevisions: (ctx, paragraph) => calls.push(['acceptRevisions', ctx.revisionsAtPoint[0]?.id, paragraph]),
      rejectRevisions: (ctx, paragraph) => calls.push(['rejectRevisions', ctx.revisionsAtPoint[0]?.id, paragraph]),
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
    isCommentResolved() {
      return false;
    },
  }));
  return { registry, state, calls };
}

test('command enabled predicates match migrated toolbar behavior', async (t) => {
  const cases = [
    ['format.bold requires a paragraph or range', 'format.bold', makeContext(), false],
    ['format.bold accepts a selected paragraph', 'format.bold', makeContext({ selection: { paragraph: 2 } }), true],
    ['format.italic accepts a selected range', 'format.italic', makeContext({ selection: { range: { start: { paragraph: 1, offset: 0 }, end: { paragraph: 1, offset: 2 } }, collapsed: false } }), true],
    ['format.underline accepts a selected paragraph', 'format.underline', makeContext({ selection: { paragraph: 2 } }), true],
    ['format.clear blocks readonly edits', 'format.clear', makeContext({ editable: false, selection: { paragraph: 0 } }), false],
    ['format.clear keeps markup edits enabled', 'format.clear', makeContext({ selection: { paragraph: 0 } }), true],
    ['format.fontColor disables without selection target', 'format.fontColor', makeContext(), false],
    ['format.fontColor enables with a selected paragraph', 'format.fontColor', makeContext({ selection: { paragraph: 0 } }), true],
    ['paragraph.style requires markup paragraph selection', 'paragraph.style', makeContext({ editable: false, selection: { paragraph: 0 } }), false],
    ['style.modify is enabled in markup view without a selection', 'style.modify', makeContext(), true],
    ['style.modify is disabled in read-only views', 'style.modify', makeContext({ editable: false }), false],
    ['style.new needs a paragraph to apply the new style to', 'style.new', makeContext(), false],
    ['math.insert needs a paragraph to insert into', 'math.insert', makeContext(), false],
    ['math.insert enables on markup paragraph selection', 'math.insert', makeContext({ selection: { paragraph: 0 } }), true],
    ['style.new enables on markup paragraph selection', 'style.new', makeContext({ selection: { paragraph: 0 } }), true],
    ['paragraph.alignment enables on markup paragraph selection', 'paragraph.alignment', makeContext({ selection: { paragraph: 0 } }), true],
    ['list.bullet disables outside a paragraph', 'list.bullet', makeContext(), false],
    ['list.bullet enables in markup view', 'list.bullet', makeContext({ selection: { paragraph: 0 } }), true],
    ['list.indent disables without numbering', 'list.indent', makeContext({ selection: { paragraph: 0 } }), false],
    ['list.indent disables at max level', 'list.indent', makeContext({ selection: { paragraph: 0 } }), false, { paragraphs: new Map([[0, { numbering: { level: 8, isBullet: true } }]]) }],
    ['list.indent enables below max level', 'list.indent', makeContext({ selection: { paragraph: 0 } }), true, { paragraphs: new Map([[0, { numbering: { level: 7, isBullet: true } }]]) }],
    ['list.outdent disables at base level', 'list.outdent', makeContext({ selection: { paragraph: 0 } }), false, { paragraphs: new Map([[0, { numbering: { level: 0, isBullet: true } }]]) }],
    ['list.outdent enables above base level', 'list.outdent', makeContext({ selection: { paragraph: 0 } }), true, { paragraphs: new Map([[0, { numbering: { level: 1, isBullet: false } }]]) }],
    ['table.deleteRow disables without table context', 'table.deleteRow', makeContext(), false],
    ['table.insertRow disables without table context', 'table.insertRow', makeContext(), false],
    ['table.insertRow enables with table context', 'table.insertRow', makeContext({ table: { table: 0, row: 1, col: 2, rowSpan: 1, colSpan: 1, nested: false } }), true],
    ['table.insertColumn enables with table context', 'table.insertColumn', makeContext({ table: { table: 0, row: 1, col: 2, rowSpan: 1, colSpan: 1, nested: false } }), true],
    ['table.deleteColumn disables without table context', 'table.deleteColumn', makeContext(), false],
    ['table.mergeCells enables with table context', 'table.mergeCells', makeContext({ table: { table: 0, row: 1, col: 2, rowSpan: 1, colSpan: 1, nested: false } }), true],
    ['table.mergeCells disables in nested tables', 'table.mergeCells', makeContext({ table: { table: 0, row: 1, col: 2, rowSpan: 1, colSpan: 1, nested: true } }), false],
    ['table.splitCell disables without table context', 'table.splitCell', makeContext(), false],
    ...['table.insertRow', 'table.deleteRow', 'table.insertColumn', 'table.deleteColumn', 'table.mergeCells', 'table.splitCell', 'table.applyCellStyle']
      .map((commandId) => [`${commandId} disables in nested tables`, commandId, makeContext({ table: { table: 0, row: 1, col: 2, rowSpan: 1, colSpan: 1, nested: true } }), false]),
    ['image.replace disables without a selected image', 'image.replace', makeContext(), false],
    ['image.insert is always enabled', 'image.insert', makeContext(), true],
    ['image.delete disables without a selected image', 'image.delete', makeContext(), false],
    ['image.delete enables with a selected image', 'image.delete', makeContext({ image: { relationshipId: 'rId7' } }), true],
    ['image.replace enables with a selected image', 'image.replace', makeContext({ image: { relationshipId: 'rId5' } }), true],
    ['image.setAlt disables without a selected image', 'image.setAlt', makeContext(), false],
    ['image.setAlt enables with a selected image', 'image.setAlt', makeContext({ image: { relationshipId: 'rId6' } }), true],
    ['hyperlink insertion disables at a collapsed caret', 'hyperlink.insertAtSelection', makeContext({ selection: { paragraph: 0, collapsed: true } }), false],
    ['hyperlink insertion enables for a same-paragraph range', 'hyperlink.insertAtSelection', makeContext({ selection: { paragraph: 0, range: { start: { paragraph: 0, offset: 0 }, end: { paragraph: 0, offset: 2 } }, collapsed: false } }), true],
    ['comment.new is always enabled', 'comment.new', makeContext(), true],
    ['comment.reply disables without a selected comment', 'comment.reply', makeContext(), false],
    ['comment.reply enables with a selected comment target', 'comment.reply', makeContext({ commentsAtPoint: [42] }), true],
    ['comment.reply enables with an active sidebar comment', 'comment.reply', makeContext({ activeCommentId: 43 }), true],
    ['review.acceptAll disables when no revisions are visible', 'review.acceptAll', makeContext(), false],
    ['review.acceptAll enables when revisions are visible', 'review.acceptAll', makeContext(), true, { revisionCount: 3 }],
    ['review.rejectAll disables when no revisions are visible', 'review.rejectAll', makeContext(), false],
    ['review.rejectAll enables when revisions are visible', 'review.rejectAll', makeContext(), true, { revisionCount: 2 }],
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

  await t.test('image commands use the image captured in command context', async () => {
    const { registry, calls } = makeRegistry({ imageAltValue: 'right-click image' });
    const image = { id: 'image-2', relationshipId: 'rId2' };
    const ctx = makeContext({ image });
    await registry.run('image.replace', ctx);
    await registry.run('image.delete', ctx);
    await registry.run('image.setAlt', ctx);
    assert.deepEqual(calls, [
      ['startReplaceImage', image],
      ['deleteImage', image],
      ['setImageAlt', 'right-click image', image],
    ]);
  });

  await t.test('comment.reply uses the selected comment id from context', async () => {
    const { registry, calls } = makeRegistry();
    await registry.run('comment.reply', makeContext({ commentsAtPoint: [7] }));
    assert.deepEqual(calls, [['replyComment', 7]]);
  });

  await t.test('comment.toggleResolved uses the active sidebar comment id', async () => {
    const { registry, calls } = makeRegistry();
    await registry.run('comment.toggleResolved', makeContext({ activeCommentId: 8 }));
    assert.deepEqual(calls, [['toggleCommentResolved', 8]]);
  });

  await t.test('comment.delete uses the active sidebar comment id', async () => {
    const { registry, calls } = makeRegistry();
    await registry.run('comment.delete', makeContext({ activeCommentId: 9 }));
    assert.deepEqual(calls, [['deleteComment', 9]]);
  });

  await t.test('review.previousRevision delegates to the review action', async () => {
    const { registry, calls } = makeRegistry();
    await registry.run('review.previousRevision', makeContext());
    assert.deepEqual(calls, [['focusPreviousRevision']]);
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

test('comment action controls derive final disabled state from registry only', () => {
  const { registry } = makeRegistry();
  const selectionCtx = makeContext({ commentsAtPoint: [42], activeCommentId: null });
  for (const commandId of ['comment.reply', 'comment.toggleResolved', 'comment.delete']) {
    const state = getCommandControlState(registry.get(commandId), selectionCtx);
    assert.equal(state.disabled, false);
  }

  const emptyCtx = makeContext({ commentsAtPoint: [], activeCommentId: null });
  for (const commandId of ['comment.reply', 'comment.toggleResolved', 'comment.delete']) {
    const state = getCommandControlState(registry.get(commandId), emptyCtx);
    assert.equal(state.disabled, true);
  }
});

const readOnlyWriteCommands = [
  'format.bold', 'format.italic', 'format.underline', 'format.painter',
  'table.insertRow', 'table.insertRowAbove', 'table.insertRowBelow', 'table.deleteRow',
  'table.insertColumn', 'table.insertColumnLeft', 'table.insertColumnRight', 'table.deleteColumn',
  'table.mergeCells', 'table.splitCell', 'table.applyCellStyle', 'table.deleteTable', 'table.applyStyle',
  'image.insert', 'image.replace', 'image.delete', 'image.setAlt',
  'comment.new', 'comment.addAtSelection', 'comment.reply', 'comment.replyAtPoint',
  'comment.toggleResolved', 'comment.toggleResolvedAtPoint', 'comment.delete', 'comment.deleteAtPoint',
];

function populatedContext(revisionView = 'markup') {
  return makeContext({
    revisionView,
    editable: revisionView === 'markup',
    selection: {
      paragraph: 0,
      range: { start: { paragraph: 0, offset: 0 }, end: { paragraph: 0, offset: 1 } },
      collapsed: false,
    },
    table: { table: 0, row: 0, col: 0, rowSpan: 1, colSpan: 1, nested: false },
    image: { relationshipId: 'rId1' },
    hyperlink: { url: 'https://example.com', unsafe: false },
    revisionsAtPoint: [{ id: 1, kind: 'insertion' }],
    commentsAtPoint: [1],
    activeCommentId: 1,
  });
}

test('all document-writing commands are disabled in both read-only preview views', async (t) => {
  const { registry } = makeRegistry({ revisionCount: 1 });
  for (const view of ['final', 'original']) {
    for (const commandId of readOnlyWriteCommands) {
      await t.test(`${commandId} is disabled in ${view}`, () => {
        assert.equal(registry.get(commandId).enabled(populatedContext(view)), false);
      });
    }
  }
});

test('editable command states recover after switching between markup and preview views', async (t) => {
  const { registry } = makeRegistry({ revisionCount: 1 });
  const commandIds = [
    'format.bold', 'format.painter', 'table.insertRow', 'table.mergeCells',
    'image.insert', 'image.delete', 'comment.new', 'comment.reply',
  ];
  for (const view of ['markup', 'final', 'original', 'markup']) {
    for (const commandId of commandIds) {
      await t.test(`${commandId} in ${view}`, () => {
        assert.equal(registry.get(commandId).enabled(populatedContext(view)), view === 'markup');
      });
    }
  }
});

test('read-only, navigation, and revision-decision commands remain enabled in preview views', async (t) => {
  const { registry } = makeRegistry({ revisionCount: 1 });
  const commandIds = [
    'clipboard.copy', 'hyperlink.open', 'hyperlink.copyAddress',
    'review.previousRevision', 'review.nextRevision', 'review.options',
    'review.acceptAll', 'review.rejectAll', 'revision.acceptAtPoint', 'revision.rejectAtPoint',
    'revision.acceptParagraph', 'revision.rejectParagraph', 'comment.focus',
  ];
  for (const view of ['final', 'original']) {
    for (const commandId of commandIds) {
      await t.test(`${commandId} remains enabled in ${view}`, () => {
        assert.equal(registry.get(commandId).enabled(populatedContext(view)), true);
      });
    }
  }
});

test('Ribbon and context-menu enabled states match for every registered command and view', async (t) => {
  const { registry } = makeRegistry({ revisionCount: 1 });
  assert.equal(registry.list().length, 62);
  for (const view of ['markup', 'final', 'original']) {
    const userState = populatedContext(view);
    const targetStates = new Map(
      ['selected-target', 'invoked-target'].map((id) => [id, {
        selection: structuredClone(userState.selection),
        table: userState.table && { ...userState.table },
        image: userState.image && { ...userState.image },
        shape: userState.shape && { ...userState.shape },
        hyperlink: userState.hyperlink && { ...userState.hyperlink },
        revisionsAtPoint: userState.revisionsAtPoint.map((revision) => ({ ...revision })),
        commentsAtPoint: [...userState.commentsAtPoint],
      }]),
    );
    const targetState = (target) => targetStates.get(target.id);
    const buildContext = createCommandContextBuilder({
      getRevisionView: () => userState.revisionView,
      getSelection: (target) => targetState(target).selection,
      getTable: (target) => targetState(target).table,
      getImage: (target) => targetState(target).image,
      getShape: (target) => targetState(target).shape,
      getHyperlink: (target) => targetState(target).hyperlink,
      getRevisionsAtPoint: (target) => targetState(target).revisionsAtPoint,
      getCommentsAtPoint: (target) => targetState(target).commentsAtPoint,
      getActiveCommentId: () => userState.activeCommentId,
      getClipboard: () => userState.clipboard,
    });
    const ribbonCtx = buildContext('ribbon', { id: 'selected-target' });
    const contextMenuCtx = buildContext('context-menu', { id: 'invoked-target' });
    const relevantFields = (ctx) => ({
      revisionView: ctx.revisionView,
      editable: ctx.editable,
      selection: ctx.selection,
      table: ctx.table,
      image: ctx.image,
      hyperlink: ctx.hyperlink,
      revisionsAtPoint: ctx.revisionsAtPoint,
      commentsAtPoint: ctx.commentsAtPoint,
      activeCommentId: ctx.activeCommentId,
      clipboard: ctx.clipboard,
    });
    assert.notEqual(ribbonCtx, contextMenuCtx);
    assert.equal(ribbonCtx.source, 'ribbon');
    assert.equal(contextMenuCtx.source, 'context-menu');
    assert.deepEqual(relevantFields(ribbonCtx), relevantFields(contextMenuCtx));
    for (const command of registry.list()) {
      await t.test(`${view}: ${command.id}`, () => {
        const ribbonDisabled = getCommandControlState(command, ribbonCtx).disabled;
        const contextMenuDisabled = !command.enabled(contextMenuCtx);
        assert.equal(ribbonDisabled, contextMenuDisabled);
      });
    }
  }
});

test('revision acceptance commands still execute in read-only preview views', async (t) => {
  for (const view of ['final', 'original']) {
    await t.test(view, async () => {
      const { registry, calls } = makeRegistry({ revisionCount: 1 });
      await registry.run('revision.acceptAtPoint', populatedContext(view));
      assert.deepEqual(calls, [['acceptRevisions', 1, false]]);
    });
  }
});
