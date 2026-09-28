import type {
  DocumentRange,
  ImageInfo,
  ParagraphFormat,
  ParagraphInfo,
  RevisionMark,
  RunFormat,
} from '../src/index.js';

export interface TableCellLocation {
  table: number;
  row: number;
  col: number;
  rowSpan: number;
  colSpan: number;
}

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
  hyperlink: { paragraph: number; runs: number[]; url?: string; anchor?: string; unsafe: boolean } | null;
  revisionsAtPoint: RevisionMark[];
  commentsAtPoint: number[];
  clipboard: 'unknown' | 'empty' | 'has-content';
  source: 'ribbon' | 'context-menu' | 'keyboard';
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
  deleteTableRow(): void;
  insertTableColumn(): void;
  deleteTableColumn(): void;
  mergeCells(): void;
  splitCell(): void;
  applyCellStyle(fill: string): void;
  startInsertImage(): void;
  startReplaceImage(): void;
  deleteImage(): void;
  setImageAlt(text: string): void;
  focusPreviousRevision(): void;
  focusNextRevision(): void;
  acceptAllRevisions(): void;
  rejectAllRevisions(): void;
  addComment(): void;
  replyComment(commentId: number): void;
  toggleCommentResolved(commentId: number): void;
  deleteComment(commentId: number): void;
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
}

export interface CommandRegistry {
  get(id: string): CommandDescriptor;
  list(): CommandDescriptor[];
  run(id: string, ctx: CommandContext): Promise<void>;
}

function hasSelectionTarget(ctx: CommandContext): boolean {
  return ctx.selection.paragraph !== null || ctx.selection.range !== null;
}

function firstCommentId(ctx: CommandContext): number {
  const commentId = ctx.commentsAtPoint[0];
  if (commentId === undefined) throw new Error('请先选择一条批注。');
  return commentId;
}

export function createExampleCommandDescriptors(deps: ExampleCommandDeps): CommandDescriptor[] {
  const paragraph = (ctx: CommandContext) => deps.getParagraphInfo(ctx.selection.paragraph);
  return [
    {
      id: 'format.bold',
      title: '加粗',
      group: 'font',
      shortcut: 'Ctrl+B',
      enabled: hasSelectionTarget,
      checked: (ctx) => Boolean(ctx.selection.format?.bold),
      run: () => deps.actions.toggleRunFormat('bold'),
    },
    {
      id: 'format.italic',
      title: '斜体',
      group: 'font',
      shortcut: 'Ctrl+I',
      enabled: hasSelectionTarget,
      checked: (ctx) => Boolean(ctx.selection.format?.italic),
      run: () => deps.actions.toggleRunFormat('italic'),
    },
    {
      id: 'format.underline',
      title: '下划线',
      group: 'font',
      shortcut: 'Ctrl+U',
      enabled: hasSelectionTarget,
      checked: (ctx) => Boolean(ctx.selection.format?.underline),
      run: () => deps.actions.toggleRunFormat('underline'),
    },
    {
      id: 'format.painter',
      title: '格式刷',
      group: 'font',
      enabled: hasSelectionTarget,
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
      enabled: () => true,
      run: () => deps.actions.insertTableRow(),
    },
    {
      id: 'table.deleteRow',
      title: '删除行',
      group: 'table',
      enabled: () => true,
      run: () => deps.actions.deleteTableRow(),
    },
    {
      id: 'table.insertColumn',
      title: '插入列',
      group: 'table',
      enabled: () => true,
      run: () => deps.actions.insertTableColumn(),
    },
    {
      id: 'table.deleteColumn',
      title: '删除列',
      group: 'table',
      enabled: () => true,
      run: () => deps.actions.deleteTableColumn(),
    },
    {
      id: 'table.mergeCells',
      title: '合并单元格',
      group: 'table',
      enabled: () => true,
      run: () => deps.actions.mergeCells(),
    },
    {
      id: 'table.splitCell',
      title: '拆分单元格',
      group: 'table',
      enabled: () => true,
      run: () => deps.actions.splitCell(),
    },
    {
      id: 'table.applyCellStyle',
      title: '单元格边框与底纹',
      group: 'table',
      enabled: () => true,
      run: () => deps.actions.applyCellStyle(deps.getCellFillValue()),
    },
    {
      id: 'image.insert',
      title: '插入图片',
      group: 'image',
      enabled: () => true,
      run: () => deps.actions.startInsertImage(),
    },
    {
      id: 'image.replace',
      title: '替换图片',
      group: 'image',
      enabled: (ctx) => ctx.image !== null,
      run: () => deps.actions.startReplaceImage(),
    },
    {
      id: 'image.delete',
      title: '删除图片',
      group: 'image',
      enabled: (ctx) => ctx.image !== null,
      run: () => deps.actions.deleteImage(),
    },
    {
      id: 'image.setAlt',
      title: '设置图片替代文本',
      group: 'image',
      enabled: (ctx) => ctx.image !== null,
      run: () => deps.actions.setImageAlt(deps.getImageAltValue()),
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
      id: 'comment.new',
      title: '新建批注',
      group: 'comment',
      enabled: () => true,
      run: () => deps.actions.addComment(),
    },
    {
      id: 'comment.reply',
      title: '回复批注',
      group: 'comment',
      enabled: (ctx) => ctx.commentsAtPoint.length > 0,
      run: (ctx) => deps.actions.replyComment(firstCommentId(ctx)),
    },
    {
      id: 'comment.toggleResolved',
      title: '解决或取消批注',
      group: 'comment',
      enabled: (ctx) => ctx.commentsAtPoint.length > 0,
      run: (ctx) => deps.actions.toggleCommentResolved(firstCommentId(ctx)),
    },
    {
      id: 'comment.delete',
      title: '删除批注',
      group: 'comment',
      enabled: (ctx) => ctx.commentsAtPoint.length > 0,
      run: (ctx) => deps.actions.deleteComment(firstCommentId(ctx)),
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
