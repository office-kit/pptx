// Public, read-only view of a slide layout.
//
// Layouts are reusable templates that slides inherit from. The PPTX picker
// shows them as `Title Slide`, `Title and Content`, `Two Content`, etc.
// `pres.slideLayouts` lists every layout in the package; each `Slide`
// resolves its layout via `slide.layout`.

import type { PartName } from '../internal/opc/index.ts';
import type { SlideLayoutPart, SlideLayoutType } from '../internal/presentationml/index.ts';

export class SlideLayout {
  /** @internal */
  readonly _part: SlideLayoutPart;
  /** @internal */
  readonly _partName: PartName;

  /** @internal */
  constructor(partName: PartName, part: SlideLayoutPart) {
    this._partName = partName;
    this._part = part;
  }

  /** PowerPoint's user-visible layout name (`Title Slide`, etc.). */
  get name(): string {
    return this._part.name;
  }

  /**
   * Layout type token, when present (`title`, `obj`, `twoObj`, ...). `null`
   * when omitted — the spec default for that case is `cust`.
   *
   * Typed as `SlideLayoutType | string | null` because real-world templates
   * occasionally embed vendor-specific tokens not enumerated in the spec.
   */
  get layoutType(): SlideLayoutType | string | null {
    return this._part.layoutType;
  }
}
