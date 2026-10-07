// Collapse steps for the contextual tabs. Mac PowerPoint switches a tab to its
// compact layout by window width (the 1200 pt captures already show it while
// the 1512 pt ones do not), the same width at which the Home tab shrinks its
// Slides/Insert groups. Longer labels (Japanese) can still overflow, so a tab
// also takes the next step while its content does not fit.
import { getLocale } from '../i18n/i18n.svelte.ts';

const COMPACT_BELOW = 1300;

export class RibbonCollapse {
  width = $state(typeof window === 'undefined' ? Number.POSITIVE_INFINITY : window.innerWidth);
  node = $state<HTMLElement>();
  #extra = $state(0);
  readonly #max: number;

  /** Call during component initialisation; `max` is the tab's last step. */
  constructor(max: number) {
    this.#max = max;
    $effect.pre(() => {
      this.width;
      getLocale();
      this.#extra = 0;
    });
    $effect(() => {
      this.step;
      this.width;
      getLocale();
      const node = this.node;
      if (node && this.step < this.#max && node.scrollWidth > node.clientWidth + 1)
        this.#extra += 1;
    });
  }

  get step(): number {
    return Math.min(this.#max, (this.width < COMPACT_BELOW ? 1 : 0) + this.#extra);
  }
}
