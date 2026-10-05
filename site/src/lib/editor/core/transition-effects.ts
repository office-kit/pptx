import type { TransitionEffect } from '@office-kit/pptx';

// Every token setSlideTransition writes; a deck can carry others (p14
// extensions), which timing edits must not rewrite.
export const WRITABLE_TRANSITIONS: ReadonlySet<string> = new Set<TransitionEffect>([
  'none',
  'blinds',
  'checker',
  'circle',
  'dissolve',
  'comb',
  'cover',
  'cut',
  'diamond',
  'fade',
  'newsflash',
  'plus',
  'pull',
  'push',
  'random',
  'randomBar',
  'split',
  'strips',
  'wedge',
  'wheel',
  'wipe',
  'zoom',
]);
