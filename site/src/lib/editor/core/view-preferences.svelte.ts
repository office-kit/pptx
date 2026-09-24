/** Application display preferences are deliberately outside document history. */
const KEY = 'office-guide-settings';
export interface GuidePreferences {
  grid: boolean;
  smart: boolean;
  drawing: boolean | null;
}
export class ViewPreferences {
  grid = $state(false);
  smart = $state(true);
  drawing = $state<boolean | null>(null);
  constructor() {
    this.reload();
  }
  reload() {
    try {
      const value = JSON.parse(localStorage.getItem(KEY) ?? '{}');
      this.grid = value.grid === true;
      this.smart = value.smart !== false;
      this.drawing = typeof value.drawing === 'boolean' ? value.drawing : null;
    } catch {
      /* Storage can be unavailable in embedded/private contexts. */
    }
  }
  save(value: GuidePreferences) {
    this.grid = value.grid;
    this.smart = value.smart;
    this.drawing = value.drawing;
    try {
      localStorage.setItem(KEY, JSON.stringify(value));
    } catch {
      /* Keep session settings. */
    }
  }
}
