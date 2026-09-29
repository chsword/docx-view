export function ribbonNavigationTarget(index: number, count: number, key: string): number | null {
  if (count <= 0 || index < 0 || index >= count) return null;
  if (key === 'ArrowRight') return (index + 1) % count;
  if (key === 'ArrowLeft') return (index + count - 1) % count;
  if (key === 'Home') return 0;
  if (key === 'End') return count - 1;
  return null;
}

export interface RibbonController {
  activate(tabId: string, focus?: boolean): void;
  setTabVisible(tabId: string, visible: boolean): void;
}

export function initializeRibbon(
  root: HTMLElement,
  options: { onManualActivate?: (tabId: string) => void } = {},
): RibbonController {
  root.setAttribute('contenteditable', 'false');
  const tabs = Array.from(root.querySelectorAll<HTMLElement>('[role="tab"]'));
  const panels = Array.from(root.querySelectorAll<HTMLElement>('[role="tabpanel"]'));
  const panelForTab = new Map(tabs.map((tab) => [
    tab.id,
    panels.find((panel) => panel.id === tab.getAttribute('aria-controls')),
  ]));
  if (!tabs.length || Array.from(panelForTab.values()).some((panel) => !panel)) throw new Error('Ribbon 选项卡与面板不匹配。');

  const activate = (tabId: string, focus = false): void => {
    const selectedTab = tabs.find((tab) => tab.id === tabId && !tab.hidden);
    if (!selectedTab) return;
    tabs.forEach((tab) => {
      const selected = tab === selectedTab;
      tab.setAttribute('aria-selected', String(selected));
      tab.tabIndex = selected ? 0 : -1;
      tab.classList.toggle('active', selected);
      panelForTab.get(tab.id)!.hidden = !selected;
    });
    if (focus) selectedTab.focus();
  };

  const visibleTabs = (): HTMLElement[] => tabs.filter((tab) => !tab.hidden);
  const setTabVisible = (tabId: string, visible: boolean): void => {
    const tab = tabs.find((candidate) => candidate.id === tabId);
    if (!tab) return;
    tab.hidden = !visible;
    const panel = panelForTab.get(tabId)!;
    if (!visible) panel.hidden = true;
  };

  const initialTab = tabs.find((tab) => tab.getAttribute('aria-selected') === 'true' && !tab.hidden)
    ?? visibleTabs()[0];
  if (!initialTab) throw new Error('Ribbon 至少需要一个可见选项卡。');
  activate(initialTab.id);
  tabs.forEach((tab) => {
    tab.addEventListener('click', () => {
      activate(tab.id);
      if (tab.dataset.contextualTab === undefined) options.onManualActivate?.(tab.id);
    });
    tab.addEventListener('keydown', (event) => {
      const availableTabs = visibleTabs();
      const index = availableTabs.indexOf(tab);
      const target = ribbonNavigationTarget(index, availableTabs.length, event.key);
      if (target === null) return;
      event.preventDefault();
      const next = availableTabs[target]!;
      activate(next.id, true);
      if (next.dataset.contextualTab === undefined) options.onManualActivate?.(next.id);
    });
  });

  root.querySelectorAll<HTMLInputElement>('[data-ribbon-toggle]').forEach((toggle) => {
    const target = root.ownerDocument.getElementById(toggle.dataset.ribbonToggle ?? '');
    if (!target) return;
    const update = (): void => { target.hidden = !toggle.checked; };
    toggle.addEventListener('change', update);
    update();
  });
  return { activate, setTabVisible };
}
