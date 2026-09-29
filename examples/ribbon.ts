export function ribbonNavigationTarget(index: number, count: number, key: string): number | null {
  if (count <= 0 || index < 0 || index >= count) return null;
  if (key === 'ArrowRight') return (index + 1) % count;
  if (key === 'ArrowLeft') return (index + count - 1) % count;
  if (key === 'Home') return 0;
  if (key === 'End') return count - 1;
  return null;
}

export function initializeRibbon(root: HTMLElement): void {
  root.setAttribute('contenteditable', 'false');
  const tabs = Array.from(root.querySelectorAll<HTMLElement>('[role="tab"]'));
  const panels = Array.from(root.querySelectorAll<HTMLElement>('[role="tabpanel"]'));
  const panelForTab = tabs.map((tab) => panels.find((panel) => panel.id === tab.getAttribute('aria-controls')));
  if (!tabs.length || panelForTab.some((panel) => !panel)) throw new Error('Ribbon 选项卡与面板不匹配。');

  const activate = (index: number, focus: boolean): void => {
    tabs.forEach((tab, tabIndex) => {
      const selected = tabIndex === index;
      tab.setAttribute('aria-selected', String(selected));
      tab.tabIndex = selected ? 0 : -1;
      tab.classList.toggle('active', selected);
      panelForTab[tabIndex]!.hidden = !selected;
    });
    if (focus) tabs[index]!.focus();
  };

  const initialIndex = tabs.findIndex((tab) => tab.getAttribute('aria-selected') === 'true');
  activate(initialIndex < 0 ? 0 : initialIndex, false);
  tabs.forEach((tab, index) => {
    tab.addEventListener('click', () => activate(index, false));
    tab.addEventListener('keydown', (event) => {
      const target = ribbonNavigationTarget(index, tabs.length, event.key);
      if (target === null) return;
      event.preventDefault();
      activate(target, true);
    });
  });

  root.querySelectorAll<HTMLInputElement>('[data-ribbon-toggle]').forEach((toggle) => {
    const target = root.ownerDocument.getElementById(toggle.dataset.ribbonToggle ?? '');
    if (!target) return;
    const update = (): void => { target.hidden = !toggle.checked; };
    toggle.addEventListener('change', update);
    update();
  });
}
