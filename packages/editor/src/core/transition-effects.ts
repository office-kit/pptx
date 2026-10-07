import type {
  SlideTransition,
  TransitionEffect,
  TransitionOptions,
  TransitionPreset,
} from '@office-kit/pptx';

// Every token setSlideTransition writes; a deck can carry others (another
// vendor's extension elements), which timing edits must not rewrite.
const WRITABLE_EFFECTS: ReadonlySet<string> = new Set<TransitionEffect>([
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
  'conveyor',
  'doors',
  'ferris',
  'flash',
  'flip',
  'flythrough',
  'gallery',
  'glitter',
  'honeycomb',
  'pan',
  'prism',
  'reveal',
  'ripple',
  'shred',
  'switch',
  'vortex',
  'warp',
  'wheelReverse',
  'window',
  'prstTrans',
  'morph',
]);

const PRESETS: ReadonlySet<string> = new Set<TransitionPreset>([
  'fallOver',
  'drape',
  'curtains',
  'wind',
  'prestige',
  'fracture',
  'crush',
  'peelOff',
  'pageCurlDouble',
  'pageCurlSingle',
  'airplane',
  'origami',
]);

const isEffect = (token: string): token is TransitionEffect => WRITABLE_EFFECTS.has(token);
const isPreset = (token: string | undefined): token is TransitionPreset =>
  token !== undefined && PRESETS.has(token);

/**
 * The options that write a transition read from the deck back unchanged, so a
 * timing edit can keep its effect — or `null` when its effect (or, for a
 * preset transition, its preset) is not one setSlideTransition writes. A slide
 * without a transition reads as "None".
 */
export const writableTransition = (
  transition: SlideTransition | null,
): TransitionOptions | null => {
  if (transition === null) return { effect: 'none' };
  const { effect, preset, ...rest } = transition;
  if (!isEffect(effect)) return null;
  if (effect !== 'prstTrans') return { ...rest, effect };
  return isPreset(preset) ? { ...rest, effect, preset } : null;
};
