/**
 * Keyboard Shortcuts - Handle keyboard shortcuts for Numeriq
 */

export interface ShortcutConfig {
  exploreFormula: string;
  navigateBack: string;
  setReference: string;
  setComparator: string;
  toggleFormulaMap: string;
  traceDependents: string;
  tracePrecedents: string;
  calculationFlow: string;
}

export type ShortcutHandler = () => void | Promise<void>;

export class KeyboardShortcutManager {
  private static shortcuts: Map<string, ShortcutHandler> = new Map();
  private static config: ShortcutConfig = {
    exploreFormula: 'Ctrl+Q',
    navigateBack: 'Ctrl+Backspace',
    setReference: 'Ctrl+Shift+S',
    setComparator: 'Ctrl+Shift+C',
    toggleFormulaMap: 'Ctrl+Shift+M',
    traceDependents: 'Ctrl+Shift+Q',
    tracePrecedents: 'Ctrl+Q',
    calculationFlow: 'Ctrl+Shift+F'
  };

  private static navigationHistory: Array<{ sheet: string; address: string }> = [];
  private static currentHistoryIndex: number = -1;
  private static maxHistorySize: number = 100;

  /**
   * Initialize keyboard shortcuts
   */
  static initialize(): void {
    document.addEventListener('keydown', this.handleKeyDown.bind(this));
  }

  /**
   * Register a shortcut handler
   */
  static registerShortcut(shortcut: string, handler: ShortcutHandler): void {
    this.shortcuts.set(shortcut, handler);
  }

  /**
   * Unregister a shortcut handler
   */
  static unregisterShortcut(shortcut: string): void {
    this.shortcuts.delete(shortcut);
  }

  /**
   * Handle keydown events
   */
  private static handleKeyDown(event: KeyboardEvent): void {
    const shortcut = this.getShortcutString(event);
    const handler = this.shortcuts.get(shortcut);

    if (handler) {
      event.preventDefault();
      event.stopPropagation();
      handler();
    }
  }

  /**
   * Get shortcut string from keyboard event
   */
  private static getShortcutString(event: KeyboardEvent): string {
    const parts: string[] = [];

    if (event.ctrlKey) parts.push('Ctrl');
    if (event.shiftKey) parts.push('Shift');
    if (event.altKey) parts.push('Alt');
    if (event.metaKey) parts.push('Meta');

    // Add the key (but not modifier keys themselves)
    if (!['Control', 'Shift', 'Alt', 'Meta'].includes(event.key)) {
      parts.push(event.key === ' ' ? 'Space' : event.key);
    }

    return parts.join('+');
  }

  /**
   * Update shortcut configuration
   */
  static updateConfig(newConfig: Partial<ShortcutConfig>): void {
    this.config = { ...this.config, ...newConfig };
  }

  /**
   * Get current shortcut configuration
   */
  static getConfig(): ShortcutConfig {
    return { ...this.config };
  }

  /**
   * Add location to navigation history
   */
  static addToHistory(sheet: string, address: string): void {
    // Remove any history after current index
    this.navigationHistory = this.navigationHistory.slice(0, this.currentHistoryIndex + 1);

    // Add new location
    this.navigationHistory.push({ sheet, address });

    // Limit history size
    if (this.navigationHistory.length > this.maxHistorySize) {
      this.navigationHistory.shift();
    } else {
      this.currentHistoryIndex++;
    }
  }

  /**
   * Navigate back in history
   */
  static async navigateBack(context: Excel.RequestContext): Promise<boolean> {
    if (this.currentHistoryIndex <= 0) {
      return false;
    }

    this.currentHistoryIndex--;
    const location = this.navigationHistory[this.currentHistoryIndex];

    try {
      const sheet = context.workbook.worksheets.getItem(location.sheet);
      const range = sheet.getRange(location.address);
      range.select();
      await context.sync();
      return true;
    } catch (error) {
      console.error('Error navigating back:', error);
      return false;
    }
  }

  /**
   * Navigate forward in history
   */
  static async navigateForward(context: Excel.RequestContext): Promise<boolean> {
    if (this.currentHistoryIndex >= this.navigationHistory.length - 1) {
      return false;
    }

    this.currentHistoryIndex++;
    const location = this.navigationHistory[this.currentHistoryIndex];

    try {
      const sheet = context.workbook.worksheets.getItem(location.sheet);
      const range = sheet.getRange(location.address);
      range.select();
      await context.sync();
      return true;
    } catch (error) {
      console.error('Error navigating forward:', error);
      return false;
    }
  }

  /**
   * Get navigation history
   */
  static getHistory(): Array<{ sheet: string; address: string }> {
    return [...this.navigationHistory];
  }

  /**
   * Clear navigation history
   */
  static clearHistory(): void {
    this.navigationHistory = [];
    this.currentHistoryIndex = -1;
  }

  /**
   * Get current history index
   */
  static getCurrentHistoryIndex(): number {
    return this.currentHistoryIndex;
  }
}
