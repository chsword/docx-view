export class RibbonContextState {
  activeTabId: string;
  lastManualTabId: string;
  preContextTabId: string | null = null;
  private readonly fixedTabIds: string[];
  private readonly contextTabs = new Map<string, boolean>();

  constructor(fixedTabIds: string[], contextTabIds: string[], initialTabId = fixedTabIds[0]!) {
    if (!fixedTabIds.length || !fixedTabIds.includes(initialTabId)) {
      throw new Error('Ribbon 需要一个有效的固定选项卡。');
    }
    this.fixedTabIds = [...fixedTabIds];
    this.activeTabId = initialTabId;
    this.lastManualTabId = initialTabId;
    for (const tabId of contextTabIds) this.contextTabs.set(tabId, false);
  }

  manualActivate(tabId: string): void {
    if (!this.fixedTabIds.includes(tabId)) return;
    this.activeTabId = tabId;
    this.lastManualTabId = tabId;
  }

  activateContext(tabId: string): void {
    if (this.contextTabs.get(tabId) === true) this.activeTabId = tabId;
  }

  setContextVisible(tabId: string, visible: boolean): string | null {
    const previous = this.contextTabs.get(tabId);
    if (previous === undefined || previous === visible) return null;
    this.contextTabs.set(tabId, visible);

    if (visible) {
      this.preContextTabId = this.activeTabId;
      this.activeTabId = tabId;
      return this.activeTabId;
    }

    if (this.activeTabId !== tabId) return null;
    const fallback = this.preContextTabId;
    this.activeTabId = fallback && this.isAvailable(fallback)
      ? fallback
      : this.isAvailable(this.lastManualTabId) ? this.lastManualTabId : this.fixedTabIds[0]!;
    return this.activeTabId;
  }

  private isAvailable(tabId: string): boolean {
    return this.fixedTabIds.includes(tabId) || this.contextTabs.get(tabId) === true;
  }
}
