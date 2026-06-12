// Presentation I/O: load, save, and create.

import { type BlankDeckAspect, OpcPackage, buildBlankDeck } from '../../internal/parts/index.ts';
import { INTERNAL_PACKAGE, type PresentationData } from '../_internal-symbols.ts';

/**
 * Anything that can be turned into a `Uint8Array` of PPTX bytes:
 *
 *   - `Uint8Array` — used as-is.
 *   - `ArrayBuffer` — wrapped without copying.
 *   - `Blob` or `File` — read via `arrayBuffer()`.
 */
export type PresentationInput = Uint8Array | ArrayBuffer | Blob;

const normalize = async (input: PresentationInput): Promise<Uint8Array> => {
  if (input instanceof Uint8Array) return input;
  if (input instanceof ArrayBuffer) return new Uint8Array(input);
  if (typeof Blob !== 'undefined' && input instanceof Blob) {
    return new Uint8Array(await input.arrayBuffer());
  }
  throw new TypeError('loadPresentation: expected Uint8Array | ArrayBuffer | Blob');
};

/**
 * Loads an existing `.pptx` and returns a `PresentationData` value.
 */
export const loadPresentation = async (input: PresentationInput): Promise<PresentationData> => {
  const bytes = await normalize(input);
  const pkg = OpcPackage.load(bytes);
  return { [INTERNAL_PACKAGE]: pkg, _slidesCache: null };
};

/** Slide-canvas aspect ratio for {@link createPresentation}. */
export type PresentationSize = '16:9' | '4:3';

/**
 * Creates a fresh, immediately-authorable `PresentationData`.
 *
 * The returned deck carries a slide master, the Office theme, and three
 * slide layouts — `'Blank'`, `'Title Slide'`, and `'Title and Content'` —
 * but no slides yet. Add one with {@link addSlide} (or the
 * `addBlankSlide` / `addTitleSlide` / `addContentSlide` helpers), then
 * `savePresentation` to emit a `.pptx` that opens cleanly in PowerPoint,
 * Keynote, Google Slides, and LibreOffice.
 *
 * @param options.size Slide aspect ratio. `'16:9'` (12192000×6858000 EMU,
 *   PowerPoint's modern default) or `'4:3'` (9144000×6858000). Defaults to
 *   `'16:9'`.
 */
export const createPresentation = (options: { size?: PresentationSize } = {}): PresentationData => {
  const aspect: BlankDeckAspect = options.size ?? '16:9';
  const pkg = buildBlankDeck(aspect);
  return { [INTERNAL_PACKAGE]: pkg, _slidesCache: null };
};

/**
 * Serializes a presentation back to PPTX bytes.
 */
export const savePresentation = (pres: PresentationData): Promise<Uint8Array> => {
  return Promise.resolve(pres[INTERNAL_PACKAGE].save());
};
