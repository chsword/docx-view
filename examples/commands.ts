import type {
  DocumentRange,
  HyperlinkInfo,
  ImageInfo,
  ParagraphFormat,
  ParagraphInfo,
  RevisionMark,
  RunFormat,
  ShapeInfo,
  TableCellLocation,
} from '../src/index.js';

export interface CommandContext {
  revisionView: 'markup' | 'final' | 'original';
  editable: boolean;
  selection: {
    paragraph: number | null;
    range: DocumentRange | null;
    format: RunFormat | null;
    collapsed: boolean;
  };
  table: TableCellLocation | null;
  image: ImageInfo | null;
  shape: ShapeInfo | null;
  hyperlink: HyperlinkInfo | null;
  revisionsAtPoint: RevisionMark[];
  commentsAtPoint: number[];
  activeCommentId?: number | null;
  clipboard: 'unknown' | 'empty' | 'has-content';
  source: 'ribbon' | 'context-menu' | 'keyboard';
}

export interface CommandContextBuilderDependencies {
  getRevisionView(): CommandContext['revisionView'];
  getSelection(target: HTMLElement | null): CommandContext['selection'];
  getTable(target: HTMLElement | null, paragraph: number | null): TableCellLocation | null;
  getImage(target: HTMLElement | null): ImageInfo | null;
  getShape(target: HTMLElement | null): ShapeInfo | null;
  getHyperlink(target: HTMLElement | null): HyperlinkInfo | null;
  getRevisionsAtPoint(target: HTMLElement | null): RevisionMark[];
  getCommentsAtPoint(target: HTMLElement | null): number[];
  getActiveCommentId(): number | null | undefined;
  getClipboard(): CommandContext['clipboard'];
}

export function createCommandContextBuilder(deps: CommandContextBuilderDependencies) {
  return (
    source: CommandContext['source'] = 'ribbon',
    target: HTMLElement | null = null,
  ): CommandContext => {
    const revisionView = deps.getRevisionView();
    const selection = deps.getSelection(target);
    return {
      revisionView,
      editable: revisionView === 'markup',
      selection,
      table: deps.getTable(target, selection.paragraph),
      image: deps.getImage(target),
      shape: deps.getShape(target),
      hyperlink: deps.getHyperlink(target),
      revisionsAtPoint: deps.getRevisionsAtPoint(target),
      commentsAtPoint: deps.getCommentsAtPoint(target),
      activeCommentId: deps.getActiveCommentId(),
      clipboard: deps.getClipboard(),
      source,
    };
  };
}

export interface CommandDescriptor {
  id: string;
  title: string;
  group: string;
  shortcut?: string;
  enabled(ctx: CommandContext): boolean;
  checked?(ctx: CommandContext): boolean;
  visibleInMenu?(ctx: CommandContext): boolean;
  run(ctx: CommandContext): void | Promise<void>;
}

export interface ExampleCommandActions {
  toggleViewMode(): void;
  clipboard(action: 'cut' | 'copy' | 'paste'): void;
  toggleRunFormat(kind: 'bold' | 'italic' | 'underline'): void;
  activateFormatPainter(locked: boolean): void;
  clearFormat(): void;
  setFontSize(value: number): void;
  setFontColor(color: string): void;
  applyParagraphStyle(style: string): void;
  setAlignment(value: ParagraphFormat['alignment']): void;
  applyNumbering(kind: 'bullet' | 'decimal'): void;
  changeNumberingLevel(delta: number): void;
  insertTableRow(): void;
  insertTableRowAt(where: 'above' | 'below'): void;
  deleteTableRow(): void;
  insertTableColumn(): void;
  insertTableColumnAt(where: 'left' | 'right'): void;
  deleteTableColumn(): void;
  mergeCells(): void;
  splitCell(): void;
  applyCellStyle(fill: string): void;
  unavailableTableAction(): void;
  startInsertImage(): void;
  startReplaceImage(image?: ImageInfo | null): void;
  deleteImage(image?: ImageInfo | null): void;
  setImageAlt(text: string, image?: ImageInfo | null): void;
  focusPreviousRevision(): void;
  focusNextRevision(): void;
  acceptAllRevisions(): void;
  rejectAllRevisions(): void;
  insertHyperlink(ctx: CommandContext): void;
  openHyperlink(link: HyperlinkInfo): void;
  copyHyperlinkAddress(link: HyperlinkInfo): void;
  editHyperlink(link: HyperlinkInfo): void;
  removeHyperlink(link: HyperlinkInfo): void;
  acceptRevisions(ctx: CommandContext, paragraph: boolean): void;
  rejectRevisions(ctx: CommandContext, paragraph: boolean): void;
  openReviewOptions(): void;
  addComment(): void;
  replyComment(commentId: number): void;
  toggleCommentResolved(commentId: number): void;
  deleteComment(commentId: number): void;
  focusComment(commentId: number): void;
}

export interface ExampleCommandDeps {
  actions: ExampleCommandActions;
  getParagraphInfo(paragraph: number | null): ParagraphInfo | null;
  getVisibleRevisionCount(): number;
  isFormatPainterActive(): boolean;
  getFontSizeValue(): string;
  getFontColorValue(): string;
  getParagraphStyleValue(): string;
  getAlignmentValue(): ParagraphFormat['alignment'];
  getCellFillValue(): string;
  getImageAltValue(): string;
  isCommentResolved(commentId: number): boolean;
}

export interface CommandRegistry {
  get(id: string): CommandDescriptor;
  list(): CommandDescriptor[];
  run(id: string, ctx: CommandContext): Promise<void>;
}

export interface CommandControlState {
  disabled: boolean;
  pressed?: boolean;
}

function hasSelectionTarget(ctx: CommandContext): boolean {
  return ctx.selection.paragraph !== null || ctx.selection.range !== null;
}

function firstCommentId(ctx: CommandContext): number {
  if (ctx.activeCommentId !== undefined && ctx.activeCommentId !== null) return ctx.activeCommentId;
  const commentId = ctx.commentsAtPoint[0];
  if (commentId === undefined) throw new Error('请先选择一条批注。');
  return commentId;
}

export function createExampleCommandDescriptors(deps: ExampleCommandDeps): CommandDescriptor[] {
  const paragraph = (ctx: CommandContext) => deps.getParagraphInfo(ctx.selection.paragraph);
  const hasLink = (ctx: CommandContext) => ctx.hyperlink !== null;
  const hasPointRevision = (ctx: CommandContext) => ctx.revisionsAtPoint.length > 0;
  const hasPointComment = (ctx: CommandContext) => ctx.commentsAtPoint.length > 0;
  const tableIsAddressable = (ctx: CommandContext) => ctx.editable && ctx.table !== null && !ctx.table.nested;
  return [
    {
      id: 'view.mode',
      title: '切换视图',
      group: 'view',
      enabled: () => true,
      run: () => deps.actions.toggleViewMode(),
    },
    {
      id: 'clipboard.cut',
      title: '剪切',
      group: 'clipboard',
      enabled: (ctx) => ctx.editable && hasSelectionTarget(ctx),
      run: () => deps.actions.clipboard('cut'),
    },
    {
      id: 'clipboard.copy',
      title: '复制',
      group: 'clipboard',
      enabled: hasSelectionTarget,
      run: () => deps.actions.clipboard('copy'),
    },
    {
      id: 'clipboard.paste',
      title: '粘贴',
      group: 'clipboard',
      enabled: (ctx) => ctx.editable && ctx.clipboard !== 'empty',
      run: () => deps.actions.clipboard('paste'),
    },
    {
      id: 'shape.copyText',
      title: '复制替换文本',
      group: 'shape',
      enabled: (ctx) => ctx.shape !== null,
      run: () => undefined,
    },
    {
      id: 'shape.viewProperties',
      title: '查看属性',
      group: 'shape',
      enabled: (ctx) => ctx.shape !== null,
      run: () => undefined,
    },
    {
      id: 'format.bold',
      title: '加粗',
      group: 'font',
      shortcut: 'Ctrl+B',
      enabled: (ctx) => ctx.editable && hasSelectionTarget(ctx),
      checked: (ctx) => Boolean(ctx.selection.format?.bold),
      run: () => deps.actions.toggleRunFormat('bold'),
    },
    {
      id: 'format.italic',
      title: '斜体',
      group: 'font',
      shortcut: 'Ctrl+I',
      enabled: (ctx) => ctx.editable && hasSelectionTarget(ctx),
      checked: (ctx) => Boolean(ctx.selection.format?.italic),
      run: () => deps.actions.toggleRunFormat('italic'),
    },
    {
      id: 'format.underline',
      title: '下划线',
      group: 'font',
      shortcut: 'Ctrl+U',
      enabled: (ctx) => ctx.editable && hasSelectionTarget(ctx),
      checked: (ctx) => Boolean(ctx.selection.format?.underline),
      run: () => deps.actions.toggleRunFormat('underline'),
    },
    {
      id: 'format.painter',
      title: '格式刷',
      group: 'font',
      enabled: (ctx) => ctx.editable && hasSelectionTarget(ctx),
      checked: () => deps.isFormatPainterActive(),
      run: () => deps.actions.activateFormatPainter(false),
    },
    {
      id: 'format.clear',
      title: '清除格式',
      group: 'font',
      enabled: (ctx) => ctx.editable && hasSelectionTarget(ctx),
      run: () => deps.actions.clearFormat(),
    },
    {
      id: 'format.fontSize',
      title: '字号',
      group: 'font',
      enabled: (ctx) => ctx.editable && hasSelectionTarget(ctx),
      run: () => {
        const value = deps.getFontSizeValue();
        if (value) deps.actions.setFontSize(Number(value));
      },
    },
    {
      id: 'format.fontColor',
      title: '文字颜色',
      group: 'font',
      enabled: (ctx) => ctx.editable && hasSelectionTarget(ctx),
      run: () => deps.actions.setFontColor(deps.getFontColorValue()),
    },
    {
      id: 'paragraph.style',
      title: '段落样式',
      group: 'paragraph',
      enabled: (ctx) => ctx.editable && ctx.selection.paragraph !== null,
      run: () => {
        const value = deps.getParagraphStyleValue();
        if (value) deps.actions.applyParagraphStyle(value);
      },
    },
    {
      id: 'paragraph.alignment',
      title: '段落对齐',
      group: 'paragraph',
      enabled: (ctx) => ctx.editable && ctx.selection.paragraph !== null,
      run: () => deps.actions.setAlignment(deps.getAlignmentValue()),
    },
    {
      id: 'list.bullet',
      title: '项目符号列表',
      group: 'paragraph',
      enabled: (ctx) => ctx.editable && ctx.selection.paragraph !== null,
      checked: (ctx) => Boolean(paragraph(ctx)?.numbering?.isBullet),
      run: () => deps.actions.applyNumbering('bullet'),
    },
    {
      id: 'list.decimal',
      title: '编号列表',
      group: 'paragraph',
      enabled: (ctx) => ctx.editable && ctx.selection.paragraph !== null,
      checked: (ctx) => Boolean(paragraph(ctx)?.numbering && !paragraph(ctx)?.numbering?.isBullet),
      run: () => deps.actions.applyNumbering('decimal'),
    },
    {
      id: 'list.indent',
      title: '提高列表级别',
      group: 'paragraph',
      enabled: (ctx) => {
        const current = paragraph(ctx);
        return ctx.editable && Boolean(current?.numbering) && (current?.numbering?.level ?? 8) < 8;
      },
      run: () => deps.actions.changeNumberingLevel(1),
    },
    {
      id: 'list.outdent',
      title: '降低列表级别',
      group: 'paragraph',
      enabled: (ctx) => {
        const current = paragraph(ctx);
        return ctx.editable && Boolean(current?.numbering) && (current?.numbering?.level ?? 0) > 0;
      },
      run: () => deps.actions.changeNumberingLevel(-1),
    },
    {
      id: 'table.insertRow',
      title: '插入行',
      group: 'table',
      enabled: tableIsAddressable,
      run: () => deps.actions.insertTableRow(),
    },
    {
      id: 'table.insertRowAbove',
      title: '在上方插入行',
      group: 'table',
      enabled: tableIsAddressable,
      visibleInMenu: (ctx) => ctx.table !== null,
      run: () => deps.actions.insertTableRowAt('above'),
    },
    {
      id: 'table.insertRowBelow',
      title: '在下方插入行',
      group: 'table',
      enabled: tableIsAddressable,
      visibleInMenu: (ctx) => ctx.table !== null,
      run: () => deps.actions.insertTableRowAt('below'),
    },
    {
      id: 'table.deleteRow',
      title: '删除行',
      group: 'table',
      enabled: tableIsAddressable,
      run: () => deps.actions.deleteTableRow(),
    },
    {
      id: 'table.insertColumn',
      title: '插入列',
      group: 'table',
      enabled: tableIsAddressable,
      run: () => deps.actions.insertTableColumn(),
    },
    {
      id: 'table.insertColumnLeft',
      title: '在左侧插入列',
      group: 'table',
      enabled: tableIsAddressable,
      visibleInMenu: (ctx) => ctx.table !== null,
      run: () => deps.actions.insertTableColumnAt('left'),
    },
    {
      id: 'table.insertColumnRight',
      title: '在右侧插入列',
      group: 'table',
      enabled: tableIsAddressable,
      visibleInMenu: (ctx) => ctx.table !== null,
      run: () => deps.actions.insertTableColumnAt('right'),
    },
    {
      id: 'table.deleteColumn',
      title: '删除列',
      group: 'table',
      enabled: tableIsAddressable,
      run: () => deps.actions.deleteTableColumn(),
    },
    {
      id: 'table.mergeCells',
      title: '合并单元格',
      group: 'table',
      enabled: tableIsAddressable,
      visibleInMenu: (ctx) => ctx.table !== null,
      run: () => deps.actions.mergeCells(),
    },
    {
      id: 'table.splitCell',
      title: '拆分单元格',
      group: 'table',
      enabled: tableIsAddressable,
      visibleInMenu: (ctx) => ctx.table !== null,
      run: () => deps.actions.splitCell(),
    },
    {
      id: 'table.applyCellStyle',
      title: '单元格边框与底纹',
      group: 'table',
      enabled: tableIsAddressable,
      run: () => deps.actions.applyCellStyle(deps.getCellFillValue()),
    },
    {
      id: 'table.deleteTable',
      title: '删除表格',
      group: 'table',
      enabled: () => false,
      visibleInMenu: (ctx) => ctx.table !== null,
      run: () => deps.actions.unavailableTableAction(),
    },
    {
      id: 'table.applyStyle',
      title: '应用表格样式…',
      group: 'table',
      enabled: () => false,
      visibleInMenu: (ctx) => ctx.table !== null,
      run: () => deps.actions.unavailableTableAction(),
    },
    {
      id: 'image.insert',
      title: '插入图片',
      group: 'image',
      enabled: (ctx) => ctx.editable,
      run: () => deps.actions.startInsertImage(),
    },
    {
      id: 'image.replace',
      title: '替换图片',
      group: 'image',
      enabled: (ctx) => ctx.editable && ctx.image !== null,
      visibleInMenu: (ctx) => ctx.image !== null,
      run: (ctx) => deps.actions.startReplaceImage(ctx.image),
    },
    {
      id: 'image.delete',
      title: '删除图片',
      group: 'image',
      enabled: (ctx) => ctx.editable && ctx.image !== null,
      visibleInMenu: (ctx) => ctx.image !== null,
      run: (ctx) => deps.actions.deleteImage(ctx.image),
    },
    {
      id: 'image.setAlt',
      title: '设置图片替代文本',
      group: 'image',
      enabled: (ctx) => ctx.editable && ctx.image !== null,
      visibleInMenu: (ctx) => ctx.image !== null,
      run: (ctx) => deps.actions.setImageAlt(deps.getImageAltValue(), ctx.image),
    },
    {
      id: 'review.previousRevision',
      title: '上一条修订',
      group: 'review',
      enabled: () => deps.getVisibleRevisionCount() > 0,
      run: () => deps.actions.focusPreviousRevision(),
    },
    {
      id: 'review.nextRevision',
      title: '下一条修订',
      group: 'review',
      enabled: () => deps.getVisibleRevisionCount() > 0,
      run: () => deps.actions.focusNextRevision(),
    },
    {
      id: 'review.acceptAll',
      title: '全部接受修订',
      group: 'review',
      enabled: () => deps.getVisibleRevisionCount() > 0,
      run: () => deps.actions.acceptAllRevisions(),
    },
    {
      id: 'review.rejectAll',
      title: '全部拒绝修订',
      group: 'review',
      enabled: () => deps.getVisibleRevisionCount() > 0,
      run: () => deps.actions.rejectAllRevisions(),
    },
    {
      id: 'hyperlink.insertAtSelection',
      title: '插入超链接…',
      group: 'hyperlink',
      enabled: (ctx) => ctx.editable && !ctx.selection.collapsed &&
        ctx.selection.range !== null &&
        ctx.selection.range.start.paragraph === ctx.selection.range.end.paragraph,
      run: (ctx) => deps.actions.insertHyperlink(ctx),
    },
    {
      id: 'hyperlink.open',
      title: '打开超链接',
      group: 'hyperlink',
      enabled: (ctx) => Boolean(ctx.hyperlink && !ctx.hyperlink.unsafe),
      visibleInMenu: hasLink,
      run: (ctx) => { if (ctx.hyperlink) deps.actions.openHyperlink(ctx.hyperlink); },
    },
    {
      id: 'hyperlink.copyAddress',
      title: '复制超链接地址',
      group: 'hyperlink',
      enabled: (ctx) => Boolean(ctx.hyperlink && !ctx.hyperlink.unsafe),
      visibleInMenu: hasLink,
      run: (ctx) => { if (ctx.hyperlink) deps.actions.copyHyperlinkAddress(ctx.hyperlink); },
    },
    {
      id: 'hyperlink.edit',
      title: '编辑超链接…',
      group: 'hyperlink',
      enabled: (ctx) => ctx.editable && hasLink(ctx),
      visibleInMenu: hasLink,
      run: (ctx) => { if (ctx.hyperlink) deps.actions.editHyperlink(ctx.hyperlink); },
    },
    {
      id: 'hyperlink.remove',
      title: '取消超链接',
      group: 'hyperlink',
      enabled: (ctx) => ctx.editable && hasLink(ctx),
      visibleInMenu: hasLink,
      run: (ctx) => { if (ctx.hyperlink) deps.actions.removeHyperlink(ctx.hyperlink); },
    },
    {
      id: 'revision.acceptAtPoint',
      title: '接受修订',
      group: 'review',
      enabled: hasPointRevision,
      visibleInMenu: hasPointRevision,
      run: (ctx) => deps.actions.acceptRevisions(ctx, false),
    },
    {
      id: 'revision.rejectAtPoint',
      title: '拒绝修订',
      group: 'review',
      enabled: hasPointRevision,
      visibleInMenu: hasPointRevision,
      run: (ctx) => deps.actions.rejectRevisions(ctx, false),
    },
    {
      id: 'revision.acceptParagraph',
      title: '接受此段的所有修订',
      group: 'review',
      enabled: (ctx) => ctx.selection.paragraph !== null,
      visibleInMenu: hasPointRevision,
      run: (ctx) => deps.actions.acceptRevisions(ctx, true),
    },
    {
      id: 'revision.rejectParagraph',
      title: '拒绝此段的所有修订',
      group: 'review',
      enabled: (ctx) => ctx.selection.paragraph !== null,
      visibleInMenu: hasPointRevision,
      run: (ctx) => deps.actions.rejectRevisions(ctx, true),
    },
    {
      id: 'review.options',
      title: '修订选项…',
      group: 'review',
      enabled: () => true,
      visibleInMenu: hasPointRevision,
      run: () => deps.actions.openReviewOptions(),
    },
    {
      id: 'comment.new',
      title: '新建批注',
      group: 'comment',
      enabled: (ctx) => ctx.editable,
      run: () => deps.actions.addComment(),
    },
    {
      id: 'comment.addAtSelection',
      title: '插入批注',
      group: 'comment',
      enabled: (ctx) => ctx.editable && !ctx.selection.collapsed,
      run: () => deps.actions.addComment(),
    },
    {
      id: 'comment.reply',
      title: '回复批注',
      group: 'comment',
      enabled: (ctx) => ctx.editable && (ctx.activeCommentId != null || ctx.commentsAtPoint.length > 0),
      visibleInMenu: hasPointComment,
      run: (ctx) => deps.actions.replyComment(firstCommentId(ctx)),
    },
    {
      id: 'comment.replyAtPoint',
      title: '回复批注…',
      group: 'comment',
      enabled: (ctx) => ctx.editable && hasPointComment(ctx),
      visibleInMenu: hasPointComment,
      run: (ctx) => deps.actions.replyComment(firstCommentId(ctx)),
    },
    {
      id: 'comment.toggleResolved',
      title: '解决或取消批注',
      group: 'comment',
      enabled: (ctx) => ctx.editable && (ctx.activeCommentId != null || ctx.commentsAtPoint.length > 0),
      visibleInMenu: hasPointComment,
      run: (ctx) => deps.actions.toggleCommentResolved(firstCommentId(ctx)),
    },
    {
      id: 'comment.toggleResolvedAtPoint',
      title: '标记为已解决 / 取消已解决',
      group: 'comment',
      enabled: (ctx) => ctx.editable && hasPointComment(ctx),
      checked: (ctx) => hasPointComment(ctx) && deps.isCommentResolved(firstCommentId(ctx)),
      visibleInMenu: hasPointComment,
      run: (ctx) => deps.actions.toggleCommentResolved(firstCommentId(ctx)),
    },
    {
      id: 'comment.delete',
      title: '删除批注',
      group: 'comment',
      enabled: (ctx) => ctx.editable && (ctx.activeCommentId != null || ctx.commentsAtPoint.length > 0),
      visibleInMenu: hasPointComment,
      run: (ctx) => deps.actions.deleteComment(firstCommentId(ctx)),
    },
    {
      id: 'comment.deleteAtPoint',
      title: '删除批注',
      group: 'comment',
      enabled: (ctx) => ctx.editable && hasPointComment(ctx),
      visibleInMenu: hasPointComment,
      run: (ctx) => deps.actions.deleteComment(firstCommentId(ctx)),
    },
    {
      id: 'comment.focus',
      title: '定位到批注面板',
      group: 'comment',
      enabled: hasPointComment,
      visibleInMenu: hasPointComment,
      run: (ctx) => deps.actions.focusComment(firstCommentId(ctx)),
    },
  ];
}

export function createCommandRegistry(commands: CommandDescriptor[]): CommandRegistry {
  const map = new Map<string, CommandDescriptor>();
  for (const command of commands) {
    if (map.has(command.id)) throw new Error(`重复的命令 id：${command.id}`);
    map.set(command.id, command);
  }
  return {
    get(id) {
      const command = map.get(id);
      if (!command) throw new Error(`未知命令：${id}`);
      return command;
    },
    list() {
      return [...map.values()];
    },
    async run(id, ctx) {
      const command = map.get(id);
      if (!command) throw new Error(`未知命令：${id}`);
      await command.run(ctx);
    },
  };
}

export function getCommandControlState(
  command: CommandDescriptor,
  ctx: CommandContext,
  options: { pressed?: boolean } = {},
): CommandControlState {
  return {
    disabled: !command.enabled(ctx),
    ...(options.pressed ? { pressed: Boolean(command.checked?.(ctx)) } : {}),
  };
}
