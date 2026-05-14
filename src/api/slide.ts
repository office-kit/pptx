// `Slide` and `SlideShape` — read + mutate a single slide.
//
// Mutations bubble back into the underlying `OpcPackage` immediately: the
// slide's XML is re-serialized into its part on every change. This keeps
// the model simple (`pres.save()` just serializes the package as-is) at
// the cost of one serialization per mutation. For typical scripts — "open
// template, fill a few placeholders, save" — that's a handful of round
// trips, which is fine.

import {
  getPictureEmbedRId,
  type Position,
  readPosition,
  readSize,
  replaceTokensInTree,
  setTextBody,
  type Size,
} from '../internal/drawingml/index.ts';
import type { Emu } from './units.ts';
import { REL_TYPES, buildTextBox, readSlideLayoutPart } from '../internal/presentationml/index.ts';
import { SlideLayout } from './slide-layout.ts';
import {
  type ImageFormat,
  type PartName,
  contentTypeForFormat,
  detectImageFormat,
  extensionForFormat,
  partName,
  resolveTarget,
} from '../internal/opc/index.ts';
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
   * The slide layout this slide is bound to, or `null` if the slide has no
   * layout relationship (which is malformed for a PPTX but tolerated here).
   *
   * Looked up via the slide's `.rels` graph: the unique rel whose type is
   * the `slideLayout` relationship type points at the layout part.
   */
  get layout(): SlideLayout | null {
    const rels = this._pkg.getRels(this._partName);
    if (rels === null) return null;
    const layoutRel = rels.items.find((r) => r.type === REL_TYPES.slideLayout);
    if (!layoutRel) return null;
    const layoutName = layoutRel.target.startsWith('/')
      ? partName(layoutRel.target)
      : resolveTarget(this._partName, layoutRel.target);
    const layoutPart = this._pkg.getPart(layoutName);
    if (layoutPart === null) return null;
    const root = parseXml(decoder.decode(layoutPart.data)).root;
    return new SlideLayout(layoutName, readSlideLayoutPart(root));
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

  /**
   * Replaces `{{key}}` tokens in every text-bearing shape on this slide
   * with the value at `tokens[key]`. Tokens whose key is missing from
   * `tokens` are left untouched. Returns the number of substitutions
   * performed.
   *
   * Limitation: tokens must fit entirely within a single text run (see
   * `replaceTokensInTree` in `drawingml/`). For tokens that have been
   * fragmented by PowerPoint's interactive editing, use `setText()` on
   * the affected shape instead.
   */
  replaceTokens(tokens: Record<string, string>): number {
    const n = replaceTokensInTree(this._document.root, tokens);
    if (n > 0) {
      this._commit();
      this._refresh();
    }
    return n;
  }

  /**
   * Adds a free-form text box to this slide. Returns the new `SlideShape`.
   *
   * The box is a plain rectangle with no fill or outline, containing one
   * paragraph with one run of `text`. Position and size are in EMU; use
   * the unit helpers (`inches`, `cm`, `pt`) from the public API to spell
   * those out.
   *
   * The shape id is allocated as one more than the current max id on the
   * slide, so the new shape never collides with existing shapes.
   */
  addTextBox(opts: { x: Emu; y: Emu; w: Emu; h: Emu; text: string; name?: string }): SlideShape {
    // Allocate the next shape id (max existing + 1).
    let maxId = 0;
    for (const s of this._part.shapes) {
      if (s.id > maxId) maxId = s.id;
    }
    // Also count the root group (which classify() doesn't surface). The
    // canonical root has id=1; respect it.
    const newId = Math.max(maxId, 1) + 1;

    const sp = buildTextBox({
      id: newId,
      ...(opts.name !== undefined ? { name: opts.name } : {}),
      x: opts.x,
      y: opts.y,
      w: opts.w,
      h: opts.h,
      text: opts.text,
    });

    // Find the spTree and append the new shape.
    const cSld = firstChildElement(this._document.root, qname('p', 'cSld', NS.pml));
    if (!cSld) throw new Error('slide has no <p:cSld>');
    const spTree = firstChildElement(cSld, qname('p', 'spTree', NS.pml));
    if (!spTree) throw new Error('slide has no <p:spTree>');
    spTree.children.push(sp);

    this._commit();
    // Rebuild shape handles from the updated part — the existing handles
    // remain valid; we just append a new one. Easiest: refresh whole list.
    const previousLength = this._shapes.length;
    this._part = readSlidePart(this._document.root);
    this._shapes = this._part.shapes.map((s) => new SlideShape(this, s.element, s));
    const created = this._shapes[previousLength];
    if (!created) throw new Error('addTextBox: post-condition failed');
    return created;
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
   * The shape's position on the slide in EMU, or `null` when the shape
   * inherits its position from a layout / master placeholder. To resolve
   * the inherited value, walk the slide → layout → master chain by hand
   * for now; a built-in helper lands when the layout reader is in place.
   */
  get position(): Position | null {
    return readPosition(this._element, this._snapshot.kind);
  }

  /**
   * The shape's size in EMU, or `null` when the shape inherits its size
   * from a layout / master placeholder. Same caveat as `position`.
   */
  get size(): Size | null {
    return readSize(this._element, this._snapshot.kind);
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

  /**
   * Replaces the picture's media with `bytes`. The format is detected from
   * the magic bytes; pass `options.format` to override (useful for
   * SVG-from-XML or odd file extensions). The original geometry — crop,
   * sizing, transform — is preserved.
   *
   * Replacement strategy:
   *
   *   - Same format as the existing media: the bytes are written in place,
   *     no rels or content-types change.
   *   - Different format: a new media part is allocated under
   *     `/ppt/media/imageN.<ext>`, the content type is registered as
   *     a Default if `<ext>` is not yet present, and the slide's `r:embed`
   *     rel is repointed at the new media. The old media part is left in
   *     place (it may be referenced by other slides).
   *
   * Throws if the shape is not a `<p:pic>` or if the format cannot be
   * detected and was not provided explicitly.
   */
  setImage(bytes: Uint8Array, options: { format?: ImageFormat } = {}): void {
    if (this._snapshot.kind !== 'picture') {
      throw new Error(`setImage only works on picture shapes; ${this._snapshot.kind} is not one`);
    }
    const format = options.format ?? detectImageFormat(bytes);
    if (format === null) {
      throw new Error('setImage: could not detect image format. Pass options.format explicitly.');
    }

    const rEmbed = getPictureEmbedRId(this._element);
    if (rEmbed === null) {
      throw new Error(`picture "${this._snapshot.name}" has no r:embed (external reference?)`);
    }
    const pkg = this._slide._pkg;
    const rels = pkg.getRels(this._slide._partName);
    if (rels === null) {
      throw new Error(`slide ${this._slide._partName} has no rels`);
    }
    const rel = rels.items.find((r) => r.id === rEmbed);
    if (!rel) {
      throw new Error(`slide rels missing entry for r:embed="${rEmbed}"`);
    }

    const mediaName = rel.target.startsWith('/')
      ? partName(rel.target)
      : resolveTarget(this._slide._partName, rel.target);
    const newExtension = extensionForFormat(format);
    const newContentType = contentTypeForFormat(format);

    const dotIdx = mediaName.lastIndexOf('.');
    const currentExtension = dotIdx >= 0 ? mediaName.slice(dotIdx + 1).toLowerCase() : '';

    if (currentExtension === newExtension) {
      const part = pkg.getPart(mediaName);
      if (!part) throw new Error(`media part missing: ${mediaName}`);
      part.data = bytes;
      part.contentType = newContentType;
      return;
    }

    // Cross-format: allocate a new media part.
    let nextN = 1;
    const mediaPathRegex = /^\/ppt\/media\/image(\d+)\./;
    for (const p of pkg.parts) {
      const m = p.name.match(mediaPathRegex);
      if (m?.[1] !== undefined) {
        const n = Number.parseInt(m[1], 10);
        if (Number.isFinite(n) && n >= nextN) nextN = n + 1;
      }
    }
    const newPartName = partName(`/ppt/media/image${nextN}.${newExtension}`);

    // Make sure Content_Types covers the new extension. If a Default for the
    // extension exists, addPart won't need an Override; otherwise it adds one.
    const hasDefault = pkg.contentTypes.defaults.some(
      (d) => d.extension.toLowerCase() === newExtension,
    );
    if (!hasDefault) {
      pkg.contentTypes.defaults.push({
        extension: newExtension,
        contentType: newContentType,
      });
    }

    pkg.addPart(newPartName, newContentType, bytes);

    // Repoint the rel at the new media part. The slide lives at
    // `/ppt/slides/slideN.xml`, media at `/ppt/media/imageM.<ext>` — relative
    // target is `../media/imageM.<ext>`.
    rel.target = `../media/image${nextN}.${newExtension}`;
    pkg.setRels(this._slide._partName, rels);
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
