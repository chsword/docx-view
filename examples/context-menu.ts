import type { CommandContext, CommandRegistry } from './commands.js';

export type ContextMenuPosition = 'body' | 'table' | 'image' | 'hyperlink' | 'revision' | 'comment';

export interface ContextMenuHit {
  position: ContextMenuPosition;
  target: HTMLElement;
  table: boolean;
  hyperlink: boolean;
  revision: boolean;
  comment: boolean;
}

export type ContextMenuEntry =
  | { type: 'command'; command: string; label?: string; context?: CommandContext }
  | { type: 'submenu'; label: string; items: ContextMenuEntry[] }
  | { type: 'separator' }
  | { type: 'notice'; label: string };

export interface ContextMenuOptions {
  editorRoot: HTMLElement;
  registry: CommandRegistry;
  buildContext(source: 'context-menu' | 'keyboard', target: HTMLElement): CommandContext;
  moveCaretToPoint?(x: number, y: number): void;
  beforeRun?(): void;
  onError?(error: unknown): void;
}

const separator = (): ContextMenuEntry => ({ type: 'separator' });
const command = (id: string, label?: string, context?: CommandContext): ContextMenuEntry => ({
  type: 'command',
  command: id,
  ...(label ? { label } : {}),
  ...(context ? { context } : {}),
});
const submenu = (label: string, ids: string[]): ContextMenuEntry => ({
  type: 'submenu',
  label,
  items: ids.map((id) => command(id)),
});

const BODY_MENU: ContextMenuEntry[] = [
  command('clipboard.cut'),
  command('clipboard.copy'),
  command('clipboard.paste'),
  separator(),
  submenu('字体…', ['format.bold', 'format.italic', 'format.underline', 'format.fontSize', 'format.fontColor']),
  submenu('段落…', ['paragraph.alignment', 'list.indent', 'list.outdent']),
  submenu('项目符号 / 编号', ['list.bullet', 'list.decimal']),
  separator(),
  command('comment.addAtSelection'),
  command('hyperlink.insertAtSelection'),
  separator(),
  submenu('段落样式', ['paragraph.style']),
];

const TABLE_MENU: ContextMenuEntry[] = [
  submenu('插入', ['table.insertRowAbove', 'table.insertRowBelow', 'table.insertColumnLeft', 'table.insertColumnRight']),
  submenu('删除', ['table.deleteRow', 'table.deleteColumn', 'table.deleteTable']),
  command('table.mergeCells'),
  command('table.splitCell'),
  command('table.applyCellStyle'),
  command('table.applyStyle'),
];

const IMAGE_MENU: ContextMenuEntry[] = [
  command('clipboard.cut'),
  command('clipboard.copy'),
  command('image.replace'),
  command('image.setAlt'),
  command('image.delete'),
];

const HYPERLINK_MENU: ContextMenuEntry[] = [
  command('hyperlink.open'),
  command('hyperlink.copyAddress'),
  command('hyperlink.edit'),
  command('hyperlink.remove'),
];

const REVISION_MENU: ContextMenuEntry[] = [
  command('revision.acceptAtPoint'),
  command('revision.rejectAtPoint'),
  command('revision.acceptParagraph'),
  command('revision.rejectParagraph'),
  separator(),
  command('review.options'),
];

const COMMENT_MENU: ContextMenuEntry[] = [
  command('comment.replyAtPoint'),
  command('comment.toggleResolvedAtPoint'),
  command('comment.deleteAtPoint'),
  command('comment.focus'),
];

function revisionMenu(ctx: CommandContext): ContextMenuEntry[] {
  if (ctx.revisionsAtPoint.length <= 1) return copyEntries(REVISION_MENU);
  const item = (id: string, revision: CommandContext['revisionsAtPoint'][number]) =>
    command(id, `${revision.kind}：${revision.author ?? '未知作者'}${revision.date ? ` ${revision.date.slice(0, 10)}` : ''}`, {
      ...ctx,
      revisionsAtPoint: [revision],
    });
  return [
    {
      type: 'submenu',
      label: '接受此处修订',
      items: ctx.revisionsAtPoint.map((revision) => item('revision.acceptAtPoint', revision)),
    },
    {
      type: 'submenu',
      label: '拒绝此处修订',
      items: ctx.revisionsAtPoint.map((revision) => item('revision.rejectAtPoint', revision)),
    },
    ...copyEntries(REVISION_MENU.slice(2)),
  ];
}

function commentMenu(ctx: CommandContext): ContextMenuEntry[] {
  if (ctx.commentsAtPoint.length <= 1) return copyEntries(COMMENT_MENU);
  return ctx.commentsAtPoint.flatMap((id, index) => [
    ...(index ? [separator()] : []),
    {
      type: 'submenu' as const,
      label: `批注 ${id}`,
      items: COMMENT_MENU.map((entry) => entry.type === 'command'
        ? command(entry.command, undefined, { ...ctx, commentsAtPoint: [id], activeCommentId: id })
        : entry),
    },
  ]);
}

function copyEntries(entries: ContextMenuEntry[]): ContextMenuEntry[] {
  return entries.map((entry) => entry.type === 'submenu'
    ? { ...entry, items: copyEntries(entry.items) }
    : { ...entry });
}

function appendGroup(entries: ContextMenuEntry[], group: ContextMenuEntry[]): void {
  if (entries.length && entries.at(-1)?.type !== 'separator') entries.push(separator());
  entries.push(...copyEntries(group));
}

export function contextMenuEntries(hit: ContextMenuHit, ctx: CommandContext): ContextMenuEntry[] {
  const entries: ContextMenuEntry[] = [];
  if (hit.position === 'image') {
    entries.push(...copyEntries(IMAGE_MENU));
    if (hit.table) appendGroup(entries, [command('table.applyCellStyle')]);
  } else if (hit.position === 'table') {
    entries.push(...copyEntries(TABLE_MENU), separator(), ...copyEntries(BODY_MENU));
  } else if (hit.position === 'hyperlink') {
    if (ctx.hyperlink?.unsafe) entries.push({ type: 'notice', label: '此链接目标不安全，已禁用打开/复制' });
    entries.push(...copyEntries(HYPERLINK_MENU), separator(), ...copyEntries(BODY_MENU));
  } else if (hit.position === 'revision') {
    entries.push(...revisionMenu(ctx));
  } else if (hit.position === 'comment') {
    entries.push(...commentMenu(ctx));
  } else {
    entries.push(...copyEntries(BODY_MENU));
  }

  if (hit.hyperlink && hit.position !== 'hyperlink') appendGroup(entries, HYPERLINK_MENU);
  if (hit.revision && hit.position !== 'revision') appendGroup(entries, revisionMenu(ctx));
  if (hit.comment && hit.position !== 'comment') appendGroup(entries, commentMenu(ctx));
  return entries;
}

function closestWithin(target: HTMLElement, selector: string, root: HTMLElement): HTMLElement | null {
  const found = target.closest<HTMLElement>(selector);
  return found && root.contains(found) ? found : null;
}

export function hitTestContext(target: HTMLElement, root: HTMLElement): ContextMenuHit {
  const image = closestWithin(target, '[data-image]', root);
  const table = closestWithin(target, '[data-table-cell="true"]', root);
  const hyperlink = closestWithin(target, '[data-docx-link="1"]', root);
  const revision = closestWithin(target, '[data-docx-revision-ids]', root);
  const comment = closestWithin(target, '[data-docx-comment-ids]', root);
  return {
    position: image ? 'image' : table ? 'table' : hyperlink ? 'hyperlink' : revision ? 'revision' : comment ? 'comment' : 'body',
    target,
    table: Boolean(table),
    hyperlink: Boolean(hyperlink),
    revision: Boolean(revision),
    comment: Boolean(comment),
  };
}

export function pointIsInsideSelection(selection: Selection | null, x: number, y: number): boolean {
  if (!selection?.rangeCount || selection.isCollapsed) return false;
  return Array.from(selection.getRangeAt(0).getClientRects()).some((rect) =>
    x >= rect.left && x <= rect.right && y >= rect.top && y <= rect.bottom);
}

export function menuNavigationTarget(index: number, enabled: boolean[], key: string): number | null {
  if (!enabled.length || !['ArrowDown', 'ArrowUp', 'Home', 'End'].includes(key)) return null;
  const direction = key === 'ArrowUp' ? -1 : 1;
  let candidate = key === 'Home' ? -1 : key === 'End' ? 0 : index;
  for (let attempts = 0; attempts < enabled.length; attempts++) {
    candidate = key === 'End'
      ? (candidate - 1 + enabled.length) % enabled.length
      : (candidate + direction + enabled.length) % enabled.length;
    if (enabled[candidate]) return candidate;
  }
  return null;
}

function elementTarget(value: EventTarget | null): HTMLElement | null {
  if (!value) return null;
  const node = value as Node;
  return node.nodeType === 1 ? node as HTMLElement : node.parentElement;
}

function keyboardTarget(root: HTMLElement): HTMLElement {
  const active = root.ownerDocument.activeElement;
  if (active instanceof HTMLElement && root.contains(active) &&
      ['[data-image]', '[data-table-cell="true"]', '[data-docx-link="1"]', '[data-docx-revision-ids]', '[data-docx-comment-ids]']
        .some((selector) => active.closest(selector))) {
    return active;
  }
  const selection = root.ownerDocument.getSelection();
  const target = elementTarget(selection?.anchorNode ?? null);
  if (target && root.contains(target)) return target;
  return active instanceof HTMLElement && root.contains(active) ? active : root;
}

export function initializeContextMenu(options: ContextMenuOptions): {
  element: HTMLElement;
  close(): void;
  destroy(): void;
} {
  const { editorRoot, registry } = options;
  const doc = editorRoot.ownerDocument;
  const menu = doc.createElement('div');
  menu.className = 'context-menu';
  menu.contentEditable = 'false';
  menu.setAttribute('role', 'menu');
  menu.setAttribute('aria-label', '文档编辑上下文菜单');
  menu.hidden = true;
  doc.body.append(menu);

  let restoreRange: Range | null = null;
  let restoreFocus: HTMLElement | null = null;

  const close = (): void => {
    if (menu.hidden) return;
    menu.hidden = true;
    menu.replaceChildren();
    restoreFocus?.focus({ preventScroll: true });
    if (restoreRange) {
      const selection = doc.getSelection();
      selection?.removeAllRanges();
      selection?.addRange(restoreRange);
    }
    restoreRange = null;
    restoreFocus = null;
  };

  const directItems = (parent: HTMLElement): HTMLElement[] =>
    Array.from(parent.children).filter((child): child is HTMLElement =>
      child instanceof HTMLElement && /^menuitem/.test(child.getAttribute('role') ?? ''));

  const focusFirst = (parent: HTMLElement): void => {
    directItems(parent).find((item) => item.getAttribute('aria-disabled') !== 'true')?.focus();
  };

  const renderEntries = (parent: HTMLElement, entries: ContextMenuEntry[], ctx: CommandContext): void => {
    for (const entry of entries) {
      if (entry.type === 'separator') {
        const line = doc.createElement('div');
        line.setAttribute('role', 'separator');
        parent.append(line);
        continue;
      }
      if (entry.type === 'notice') {
        const notice = doc.createElement('div');
        notice.className = 'context-menu-notice';
        notice.setAttribute('role', 'menuitem');
        notice.setAttribute('aria-disabled', 'true');
        notice.tabIndex = -1;
        notice.textContent = entry.label;
        parent.append(notice);
        continue;
      }
      if (entry.type === 'submenu') {
        const item = doc.createElement('div');
        item.className = 'context-menu-item';
        item.setAttribute('role', 'menuitem');
        item.setAttribute('aria-haspopup', 'menu');
        item.setAttribute('aria-expanded', 'false');
        item.tabIndex = -1;
        item.textContent = `${entry.label} ›`;
        const child = doc.createElement('div');
        child.className = 'context-submenu';
        child.setAttribute('role', 'menu');
        child.hidden = true;
        renderEntries(child, entry.items, ctx);
        item.append(child);
        const expand = (): void => {
          item.setAttribute('aria-expanded', 'true');
          child.hidden = false;
        };
        item.addEventListener('mouseenter', expand);
        item.addEventListener('mouseleave', () => {
          item.setAttribute('aria-expanded', 'false');
          child.hidden = true;
        });
        item.addEventListener('click', expand);
        parent.append(item);
        continue;
      }
      const descriptor = registry.get(entry.command);
      const commandContext = entry.context ?? ctx;
      if (descriptor.visibleInMenu && !descriptor.visibleInMenu(commandContext)) continue;
      const item = doc.createElement('div');
      const checked = descriptor.checked?.(commandContext);
      item.className = 'context-menu-item';
      item.dataset.command = descriptor.id;
      item.setAttribute('role', checked === undefined ? 'menuitem' : 'menuitemcheckbox');
      if (checked !== undefined) item.setAttribute('aria-checked', String(checked));
      item.setAttribute('aria-disabled', String(!descriptor.enabled(commandContext)));
      item.tabIndex = -1;
      item.textContent = entry.label ?? (descriptor.shortcut ? `${descriptor.title}\t${descriptor.shortcut}` : descriptor.title);
      item.addEventListener('click', () => {
        if (item.getAttribute('aria-disabled') === 'true') return;
        close();
        options.beforeRun?.();
        void registry.run(descriptor.id, commandContext).catch(options.onError);
      });
      parent.append(item);
    }
  };

  const open = (target: HTMLElement, x: number, y: number, source: 'context-menu' | 'keyboard'): void => {
    close();
    const selection = doc.getSelection();
    restoreRange = selection?.rangeCount ? selection.getRangeAt(0).cloneRange() : null;
    restoreFocus = doc.activeElement instanceof HTMLElement ? doc.activeElement : null;
    const hit = hitTestContext(target, editorRoot);
    const ctx = options.buildContext(source, target);
    menu.replaceChildren();
    renderEntries(menu, contextMenuEntries(hit, ctx), ctx);
    menu.hidden = false;
    const viewport = doc.defaultView;
    menu.style.left = `${Math.max(0, Math.min(x, (viewport?.innerWidth ?? 1024) - 280))}px`;
    menu.style.top = `${Math.max(0, Math.min(y, (viewport?.innerHeight ?? 768) - 360))}px`;
    focusFirst(menu);
  };

  const handleContextMenu = (event: MouseEvent): void => {
    if (event.shiftKey) return;
    const target = elementTarget(event.target);
    if (!target || !editorRoot.contains(target)) return;
    event.preventDefault();
    if (!pointIsInsideSelection(doc.getSelection(), event.clientX, event.clientY)) {
      options.moveCaretToPoint?.(event.clientX, event.clientY);
    }
    open(target, event.clientX, event.clientY, 'context-menu');
  };

  const handleEditorKeydown = (event: KeyboardEvent): void => {
    if (!(event.key === 'ContextMenu' || (event.key === 'F10' && event.shiftKey))) return;
    event.preventDefault();
    const target = keyboardTarget(editorRoot);
    const selection = doc.getSelection();
    const rect = selection?.rangeCount ? selection.getRangeAt(0).getBoundingClientRect() : editorRoot.getBoundingClientRect();
    open(target, rect.left, rect.bottom, 'keyboard');
  };

  const handleMenuKeydown = (event: KeyboardEvent): void => {
    if (event.key === 'Escape' || event.key === 'Esc') {
      event.preventDefault();
      close();
      return;
    }
    const current = elementTarget(event.target);
    if (!current) return;
    if (event.key === 'ArrowRight' && current.getAttribute('aria-haspopup') === 'menu') {
      event.preventDefault();
      current.setAttribute('aria-expanded', 'true');
      const child = Array.from(current.children).find((node) => node.getAttribute('role') === 'menu') as HTMLElement | undefined;
      if (child) {
        child.hidden = false;
        focusFirst(child);
      }
      return;
    }
    if (event.key === 'ArrowLeft') {
      const parentMenu = current.parentElement;
      const parentItem = parentMenu?.parentElement;
      if (parentMenu !== menu && parentItem?.getAttribute('aria-haspopup') === 'menu') {
        event.preventDefault();
        parentMenu!.hidden = true;
        parentItem.setAttribute('aria-expanded', 'false');
        parentItem.focus();
      }
      return;
    }
    if (event.key === 'Enter' || event.key === ' ') {
      event.preventDefault();
      current.click();
      return;
    }
    const parent = current.parentElement;
    if (!parent) return;
    const items = directItems(parent);
    const next = menuNavigationTarget(
      items.indexOf(current),
      items.map((item) => item.getAttribute('aria-disabled') !== 'true'),
      event.key,
    );
    if (next !== null) {
      event.preventDefault();
      items[next]?.focus();
    }
  };

  const handleOutsidePointer = (event: Event): void => {
    const target = elementTarget(event.target);
    if (!menu.hidden && target && !menu.contains(target)) close();
  };

  editorRoot.addEventListener('contextmenu', handleContextMenu);
  editorRoot.addEventListener('keydown', handleEditorKeydown);
  menu.addEventListener('keydown', handleMenuKeydown);
  doc.addEventListener('pointerdown', handleOutsidePointer);

  return {
    element: menu,
    close,
    destroy() {
      editorRoot.removeEventListener('contextmenu', handleContextMenu);
      editorRoot.removeEventListener('keydown', handleEditorKeydown);
      doc.removeEventListener('pointerdown', handleOutsidePointer);
      menu.remove();
    },
  };
}
