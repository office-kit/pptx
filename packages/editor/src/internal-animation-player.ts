// UNSTABLE, NOT PUBLIC API — see internal.ts. A separate entry so the dev
// preview page and presenter window load the player without the editor. The
// dev server serves this module as `/animation-player.js`; its browser tests
// also import `buildAnimationStops` from there.

/** @internal */
export { buildAnimationStops, createAnimationPlayer } from './core/animation-player.ts';
