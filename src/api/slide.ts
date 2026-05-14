// `Slide` and `SlideShape` — read + mutate a single slide.
//
// Mutations bubble back into the underlying `OpcPackage` immediately: the
// slide's XML is re-serialized into its part on every change. This keeps
// the model simple (`pres.save()` just serializes the package as-is) at
// the cost of one serialization per mutation. For typical scripts — "open
// template, fill a few placeholders, save" — that's a handful of round
// trips, which is fine.

import { setTextBody } from '../internal/drawingml/index.ts';
import type { PartName } from '../internal/opc/index.ts';
import type { OpcPackage } from '../internal/parts/index.ts';
import {
  type ShapeKind,
  type SlidePart,
  readSlidePart,
  slideText,
} from '../internal/presentationml/index.ts';
import {
  NS,
  type XmlDocument,
  type XmlElement,
  firstChildElement,
  parseXml,
  qname,
  serializeXml,
} from '../internal/xml/index.ts';

const NAME_TX_BODY = qname('p', 'txBody', NS.pml);

const encoder = new TextEncoder();
const decoder = new TextDecoder();

/**
 * Common placeholder type tokens defined by ECMA-376 Part 1 §19.7.10 (`ST_PlaceholderType`).
 * Listed here as a union for autocompletion; callers may pass any string
 * since custom layouts can introduce additional types.
 */
export type PlaceholderType =
  | 'title'
  | 'body'
  | 'ctrTitle'
  | 'subTitle'
  | 'dt'
  | 'sldNum'
  | 'ftr'
  | 'hdr'
  | 'obj'
  | 'chart'
  | 'tbl'
  | 'clipArt'
  | 'dgm'
  | 'media'
  | 'sldImg'
  | 'pic';

export class Slide {
  /** @internal */
  readonly _pkg: OpcPackage;
  /** @internal */
  readonly _partName: PartName;
  /** @internal */
  _document: XmlDocument;
  /** @internal */
  _part: SlidePart;
  /** @internal */
  _shapes: SlideShape[];

  /** @internal */
  constructor(pkg: OpcPackage, partName: PartName, document: XmlDocument) {
    this._pkg = pkg;
    this._partName = partName;
    this._document = document;
    this._part = readSlidePart(document.root);
    this._shapes = this._part.shapes.map((s) => new SlideShape(this, s.element, s));
  }

  /** Shapes on this slide in document order, group children flattened. */
  get shapes(): ReadonlyArray<SlideShape> {
    return this._shapes;
  }

  /** Concatenated visible text from every shape. */
  get text(): string {
    return slideText(this._part);
  }

  /**
   * Returns the first placeholder shape with the given `type`. Returns
   * `null` if no match. Shapes whose `<p:ph>` lacks an explicit `type`
   * attribute are treated as the catch-all body placeholder and reachable
   * via `findPlaceholder('body')`.
   */
  findPlaceholder(type: PlaceholderType | string): SlideShape | null {
    for (const shape of this._shapes) {
      if (shape.placeholderType === type) return shape;
      // ECMA-376 §19.7.10: omitted type defaults to "body".
      if (type === 'body' && shape.placeholderType === null && shape.placeholderIdx !== null) {
        return shape;
      }
    }
    return null;
  }

  /** @internal — re-parse the typed view after the underlying XML mutates. */
  _refresh(): void {
    this._part = readSlidePart(this._document.root);
    // Rebuild shape handles in place so existing references keep tracking the
    // same shape across mutations.
    for (let i = 0; i < this._shapes.length; i++) {
      const next = this._part.shapes[i];
      const existing = this._shapes[i];
      if (!next || !existing) continue;
      existing._element = next.element;
      existing._snapshot = next;
    }
  }

  /** @internal — serialize the in-memory XML back into the package. */
  _commit(): void {
    const xml = serializeXml(this._document);
    const part = this._pkg.getPart(this._partName);
    if (!part) throw new Error(`slide part missing: ${this._partName}`);
    part.data = encoder.encode(xml);
  }
}

export class SlideShape {
  /** @internal */
  readonly _slide: Slide;
  /** @internal */
  _element: XmlElement;
  /** @internal */
  _snapshot: SlidePart['shapes'][number];

  /** @internal */
  constructor(slide: Slide, element: XmlElement, snapshot: SlidePart['shapes'][number]) {
    this._slide = slide;
    this._element = element;
    this._snapshot = snapshot;
  }

  get kind(): ShapeKind {
    return this._snapshot.kind;
  }

  get id(): number {
    return this._snapshot.id;
  }

  get name(): string {
    return this._snapshot.name;
  }

  get placeholderType(): string | null {
    return this._snapshot.placeholderType;
  }

  get placeholderIdx(): number | null {
    return this._snapshot.placeholderIdx;
  }

  get text(): string {
    return this._snapshot.text;
  }

  /**
   * Replaces this shape's visible text with `value`. Newlines start a new
   * paragraph. The first existing run/paragraph properties are cloned so
   * font, color, size, alignment, and bullet style are preserved.
   *
   * Throws if the shape is not a `<p:sp>` with a text body. Pictures,
   * graphic frames, groups, and connectors are not text-bearing.
   */
  setText(value: string): void {
    if (this._snapshot.kind !== 'shape') {
      throw new Error(
        `setText only works on text-bearing shapes; ${this._snapshot.kind} is not one`,
      );
    }
    const txBody = firstChildElement(this._element, NAME_TX_BODY);
    if (txBody === null) {
      throw new Error(`shape "${this._snapshot.name}" has no <p:txBody>`);
    }
    setTextBody(txBody, value);
    this._slide._commit();
    this._slide._refresh();
  }
}

/** @internal — used by `Presentation` to construct slides without leaking internals. */
export const _internalCreateSlide = (
  pkg: OpcPackage,
  partName: PartName,
  bytes: Uint8Array,
): Slide => {
  const doc = parseXml(decoder.decode(bytes));
  return new Slide(pkg, partName, doc);
};
