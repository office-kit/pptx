/** Application display preferences are deliberately outside document history. */
const KEY = 'office-guide-settings';
export interface GuidePreferences {
  ruler: boolean;
  grid: boolean;
  smart: boolean;
  drawing: boolean | null;
}
export class ViewPreferences {
  ruler = $state(false);
  grid = $state(false);
  smart = $state(true);
  drawing = $state<boolean | null>(null);
  constructor() {
    this.reload();
  }
  reload() {
    try {
      const value = JSON.parse(localStorage.getItem(KEY) ?? '{}');
      this.ruler = value.ruler === true;
      this.grid = value.grid === true;
      this.smart = value.smart !== false;
      this.drawing = typeof value.drawing === 'boolean' ? value.drawing : null;
    } catch {
      /* Storage can be unavailable in embedded/private contexts. */
    }
  }
  save(update: Partial<GuidePreferences>) {
    const value = {
      ruler: this.ruler,
      grid: this.grid,
      smart: this.smart,
      drawing: this.drawing,
      ...update,
    };
    this.ruler = value.ruler;
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
