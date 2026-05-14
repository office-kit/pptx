// Public API surface. The only directory users are intended to import from.
// Subsequent phases (P1+) will populate this with Presentation, Slide, etc.

export { type Emu, inches, cm, mm, pt, emu } from './units.ts';

// Library version. Replaced at build time by the package version.
export const VERSION = '0.0.0';
