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

import { replaceTokensInTree } from '../internal/drawingml/index.ts';
import { OpcPackage } from '../internal/parts/index.ts';
import { readSlideLayoutPart } from '../internal/presentationml/index.ts';
import { parseXml, serializeXml } from '../internal/xml/index.ts';
import {
  INTERNAL_PACKAGE,
  LAYOUT_PART,
  LAYOUT_PART_NAME,
  type PresentationData,
  type SlideLayoutData,
} from './_internal-symbols.ts';

const TEXT_DECODER = new TextDecoder();
const TEXT_ENCODER = new TextEncoder();
const decode = (b: Uint8Array): string => TEXT_DECODER.decode(b);
const encode = (s: string): Uint8Array => TEXT_ENCODER.encode(s);

const SLIDE_LAYOUT_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml';
const SLIDE_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.presentationml.slide+xml';

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

/**
 * Enumerates every slide layout in the package. Returns plain
 * `SlideLayoutData` values that work as inputs to other authoring
 * functions (`addSlide`, etc.) once those land.
 */
export const getSlideLayouts = (pres: PresentationData): ReadonlyArray<SlideLayoutData> => {
  const pkg = pres[INTERNAL_PACKAGE];
  const out: SlideLayoutData[] = [];
  for (const part of pkg.parts) {
    if (part.contentType !== SLIDE_LAYOUT_CONTENT_TYPE) continue;
    const root = parseXml(decode(part.data)).root;
    out.push({
      [LAYOUT_PART_NAME]: part.name,
      [LAYOUT_PART]: readSlideLayoutPart(root),
    });
  }
  return out;
};

/**
 * Replaces `{{key}}` tokens on every slide. Returns the total number of
 * substitutions performed.
 *
 * Walks XML parts directly so no slide model is required — the function
 * stays minimal-bundle-friendly when the caller imports only this and
 * `loadPresentation` / `savePresentation`.
 */
export const replaceTokensInPresentation = (
  pres: PresentationData,
  tokens: Record<string, string>,
): number => {
  const pkg = pres[INTERNAL_PACKAGE];
  let count = 0;
  for (const part of pkg.parts) {
    if (part.contentType !== SLIDE_CONTENT_TYPE) continue;
    const doc = parseXml(decode(part.data));
    const n = replaceTokensInTree(doc.root, tokens);
    if (n > 0) {
      part.data = encode(serializeXml(doc));
      count += n;
    }
  }
  pres._slidesCache = null;
  return count;
};

void TEXT_DECODER;
