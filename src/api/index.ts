// Public API surface. The only directory users are intended to import from.

export { type Emu, cm, emu, inches, mm, pt } from './units.ts';
export { Presentation, type PresentationInput, _internalPackageOf } from './presentation.ts';
export { Slide, type SlideShape } from './slide.ts';

// Library version. Replaced at build time by the package version.
export const VERSION = '0.0.0';
