// Tree-shakeable free-function entry points.
//
// Every operation in this module is a standalone export that operates on
// the opaque `PresentationData` interface (shared with the class-based
// API via `_internal-symbols.ts`). Crucially, none of these functions
// references any class — when a consumer imports only what they need
// from this module, modern bundlers drop the class definitions in
// `presentation.ts` / `slide.ts` / etc. entirely.
//
// The class-based API in those files is preserved as a legacy facade
// (re-exported from `./index.ts` as `Presentation`, `Slide`, etc.) so
// existing tests and downstream consumers continue to work. As class
// methods migrate to live exclusively here, the class definitions will
// shrink and eventually disappear.

import { OpcPackage } from '../internal/parts/index.ts';
import { INTERNAL_PACKAGE, type PresentationData } from './_internal-symbols.ts';

const TEXT_DECODER = new TextDecoder();

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
 *
 * Throws `TypeError` for unsupported input shapes and a descriptive
 * `Error` if the bytes do not form a valid OPC package.
 */
export const loadPresentation = async (input: PresentationInput): Promise<PresentationData> => {
  const bytes = await normalize(input);
  const pkg = OpcPackage.load(bytes);
  return { [INTERNAL_PACKAGE]: pkg, _slidesCache: null };
};

/**
 * Creates a fresh, empty `PresentationData`. The result is NOT yet a
 * valid PPTX — it carries only the OPC defaults. A future helper will
 * produce the canonical PowerPoint skeleton.
 */
export const createPresentation = (): PresentationData => {
  const pkg = OpcPackage.empty();
  return { [INTERNAL_PACKAGE]: pkg, _slidesCache: null };
};

/**
 * Serializes a presentation back to PPTX bytes. Internally re-emits
 * `[Content_Types].xml` and every part, preserving entry order.
 */
export const savePresentation = (pres: PresentationData): Promise<Uint8Array> => {
  return Promise.resolve(pres[INTERNAL_PACKAGE].save());
};

// Suppress unused-import warning for the decoder — kept available for
// follow-up free functions that need text decoding inline.
void TEXT_DECODER;
