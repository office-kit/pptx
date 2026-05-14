// Public `Presentation` entry point.
//
// The class wraps an internal `OpcPackage` and currently exposes load,
// save, and a read-only `slides` view. Authoring methods land as the
// internal PresentationML / DrawingML layers grow real width.

import { partName } from '../internal/opc/index.ts';
import { OpcPackage } from '../internal/parts/index.ts';
import { readPresentationPart } from '../internal/presentationml/index.ts';
import { parseXml } from '../internal/xml/index.ts';
import { _internalCreateSlide, type Slide } from './slide.ts';

/**
 * Anything that can be turned into a `Uint8Array` of PPTX bytes:
 *
 *   - `Uint8Array` — used as-is.
 *   - `ArrayBuffer` — wrapped without copying.
 *   - `Blob` or `File` — read via `arrayBuffer()`. Supported in both Node
 *     (since Node 18) and all current browsers.
 *
 * Path strings are NOT accepted here to keep the browser bundle free of
 * `node:fs`. Use the `loadFile` helper from `pptx-kit/node` instead.
 */
export type PresentationInput = Uint8Array | ArrayBuffer | Blob;

const normalize = async (input: PresentationInput): Promise<Uint8Array> => {
  if (input instanceof Uint8Array) return input;
  if (input instanceof ArrayBuffer) return new Uint8Array(input);
  if (typeof Blob !== 'undefined' && input instanceof Blob) {
    return new Uint8Array(await input.arrayBuffer());
  }
  throw new TypeError('Presentation.load: expected Uint8Array | ArrayBuffer | Blob');
};

/**
 * Internal symbol used to hand the underlying `OpcPackage` to subclasses
 * (such as the Node-specific extension in `pptx-kit/node`) without
 * exposing it on the public type. Treat as `@internal`.
 */
const INTERNAL_PACKAGE = Symbol('pptx-kit.package');

export class Presentation {
  /** @internal */
  [INTERNAL_PACKAGE]: OpcPackage;

  /** @internal */
  protected constructor(pkg: OpcPackage) {
    this[INTERNAL_PACKAGE] = pkg;
  }

  /**
   * Loads an existing `.pptx` and returns a `Presentation`.
   *
   * Throws `TypeError` if the input is not a supported type, and
   * `Error` (with a descriptive message) if the bytes are not a valid
   * OPC package.
   */
  static async load(input: PresentationInput): Promise<Presentation> {
    const bytes = await normalize(input);
    return new Presentation(OpcPackage.load(bytes));
  }

  /**
   * Creates a fresh, empty `Presentation`. The result is NOT yet a valid
   * PPTX — it carries only the OPC defaults. Subsequent phases will add
   * a `Presentation.create()` that produces the canonical PowerPoint
   * skeleton (presentation.xml + at least one slide / layout / master /
   * theme).
   */
  static create(): Presentation {
    return new Presentation(OpcPackage.empty());
  }

  /**
   * Serializes the presentation back to PPTX bytes. Internally re-emits
   * `[Content_Types].xml` and every part, preserving entry order.
   */
  save(): Promise<Uint8Array> {
    return Promise.resolve(this[INTERNAL_PACKAGE].save());
  }

  /** @internal — populated lazily on first `slides` read; reused for mutation. */
  private _slidesCache: Slide[] | null = null;

  /**
   * Enumerates slides in presentation order. Returns cached `Slide`
   * instances so that mutations made through one handle (e.g.
   * `slide.shapes[0].setText('...')`) persist for the lifetime of this
   * `Presentation`.
   *
   * Throws if any referenced slide part is missing — a structurally
   * invalid PPTX cannot honor the L1 contract.
   */
  get slides(): ReadonlyArray<Slide> {
    if (this._slidesCache !== null) return this._slidesCache;

    const pkg = this[INTERNAL_PACKAGE];
    const presPart = pkg.getPart(partName('/ppt/presentation.xml'));
    if (presPart === null) {
      this._slidesCache = [];
      return this._slidesCache;
    }
    const presRels = pkg.getRels(partName('/ppt/presentation.xml'));
    if (presRels === null) {
      this._slidesCache = [];
      return this._slidesCache;
    }

    const presRoot = parseXml(new TextDecoder().decode(presPart.data)).root;
    const presModel = readPresentationPart(presRoot);

    const out: Slide[] = [];
    for (const sld of presModel.slides) {
      const rel = presRels.items.find((r) => r.id === sld.rId);
      if (!rel) {
        throw new Error(`presentation.xml.rels missing entry for ${sld.rId}`);
      }
      const target = rel.target;
      // Quick relative-resolve: presentation.xml lives at /ppt/; slide
      // targets are like "slides/slideN.xml".
      const slideName = partName(target.startsWith('/') ? target : `/ppt/${target}`);
      const slidePart = pkg.getPart(slideName);
      if (slidePart === null) {
        throw new Error(`slide part ${slideName} not found`);
      }
      out.push(_internalCreateSlide(pkg, slideName, slidePart.data));
    }
    this._slidesCache = out;
    return out;
  }

  /**
   * Replaces `{{key}}` tokens on every slide with the value at
   * `tokens[key]`. Returns the total number of substitutions performed.
   *
   * Convenience over `for (const s of pres.slides) s.replaceTokens(tokens)`.
   * Useful for the common "fill the template once" workflow.
   */
  replaceTokens(tokens: Record<string, string>): number {
    let n = 0;
    for (const slide of this.slides) {
      n += slide.replaceTokens(tokens);
    }
    return n;
  }

  /** @internal */
  static _fromPackage(pkg: OpcPackage): Presentation {
    return new Presentation(pkg);
  }
}

/**
 * @internal — used by `pptx-kit/node` to mount fs-backed helpers without
 * importing internal modules directly.
 */
export const _internalPackageOf = (p: Presentation): OpcPackage => p[INTERNAL_PACKAGE];

// `Presentation._fromPackage` (declared inside the class body) is the
// supported way for sibling modules in this package — notably the Node
// subclass — to construct a Presentation from a pre-parsed OpcPackage
// without bypassing the protected constructor.
