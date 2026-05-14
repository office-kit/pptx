// `Slide` and `SlideShape` — read + mutate a single slide.
//
// Mutations bubble back into the underlying `OpcPackage` immediately: the
// slide's XML is re-serialized into its part on every change. This keeps
// the model simple (`pres.save()` just serializes the package as-is) at
// the cost of one serialization per mutation. For typical scripts — "open
// template, fill a few placeholders, save" — that's a handful of round
// trips, which is fine.

import {
  applyAlignmentToAllParagraphs,
  applyBulletToAllParagraphs,
  applyFormatToAllRuns,
  applyHyperlinkToAllRuns,
  type BulletStyle,
  clearFill,
  clearStroke,
  getPictureEmbedRId,
  type ParagraphAlignment,
  type Position,
  readPosition,
  readSize,
  replaceTokensInTree,
  setNoFill,
  setNoStroke,
  setPosition as writePosition,
  setSize as writeSize,
  setSolidFill,
  setSolidStroke,
  setTextBody,
  type Size,
  type StrokeOptions,
  type TextFormat,
} from '../internal/drawingml/index.ts';
import type { Emu } from './units.ts';
import {
  REL_TYPES,
  buildPicture,
  buildTable,
  buildTextBox,
  readSlideLayoutPart,
} from '../internal/presentationml/index.ts';
import { SlideLayout } from './slide-layout.ts';
import {
  type ImageFormat,
  type PartName,
  contentTypeForFormat,
  detectImageFormat,
  extensionForFormat,
  nextRelId,
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
   * Adds a new picture to this slide from raw bytes. Returns the new shape.
   *
   * Side effects:
   *   - A new media part is allocated under `/ppt/media/imageN.<ext>`.
   *   - If the image's extension isn't yet covered by a `Content_Types`
   *     `Default`, one is registered.
   *   - The slide's `.rels` gets a new `image` relationship pointing at
   *     the media part. The slide's `r:embed` references that rel.
   *   - A `<p:pic>` element is appended to the slide's `<p:spTree>` with
   *     the given geometry.
   *
   * Format detection follows the same magic-byte sniff as `setImage`. Pass
   * `options.format` to override (useful for SVG or when the bytes are
   * not yet realized).
   */
  addImage(
    bytes: Uint8Array,
    opts: { x: Emu; y: Emu; w: Emu; h: Emu; format?: ImageFormat; name?: string },
  ): SlideShape {
    const format = opts.format ?? detectImageFormat(bytes);
    if (format === null) {
      throw new Error('addImage: could not detect image format. Pass options.format explicitly.');
    }
    const contentType = contentTypeForFormat(format);
    const extension = extensionForFormat(format);

    const pkg = this._pkg;

    // 1. Allocate the next /ppt/media/imageN.<ext> part name.
    let nextN = 1;
    const mediaPattern = /^\/ppt\/media\/image(\d+)\./;
    for (const p of pkg.parts) {
      const m = p.name.match(mediaPattern);
      if (m?.[1] !== undefined) {
        const n = Number.parseInt(m[1], 10);
        if (Number.isFinite(n) && n >= nextN) nextN = n + 1;
      }
    }
    const newMediaName = partName(`/ppt/media/image${nextN}.${extension}`);

    // 2. Register a Content_Types Default if this extension isn't covered.
    const hasDefault = pkg.contentTypes.defaults.some(
      (d) => d.extension.toLowerCase() === extension,
    );
    if (!hasDefault) {
      pkg.contentTypes.defaults.push({ extension, contentType });
    }

    // 3. Add the media part.
    pkg.addPart(newMediaName, contentType, bytes);

    // 4. Add the slide→image rel.
    const rels = pkg.getRels(this._partName) ?? { items: [] };
    const newRId = nextRelId(rels.items.map((r) => r.id));
    rels.items.push({
      id: newRId,
      type: REL_TYPES.image,
      target: `../media/image${nextN}.${extension}`,
      targetMode: 'Internal',
    });
    pkg.setRels(this._partName, rels);

    // 5. Allocate the next shape id within this slide.
    let maxId = 0;
    for (const s of this._part.shapes) {
      if (s.id > maxId) maxId = s.id;
    }
    const newId = Math.max(maxId, 1) + 1;

    // 6. Build and append the <p:pic> element.
    const pic = buildPicture({
      id: newId,
      ...(opts.name !== undefined ? { name: opts.name } : {}),
      rEmbed: newRId,
      x: opts.x,
      y: opts.y,
      w: opts.w,
      h: opts.h,
    });
    const cSld = firstChildElement(this._document.root, qname('p', 'cSld', NS.pml));
    if (!cSld) throw new Error('slide has no <p:cSld>');
    const spTree = firstChildElement(cSld, qname('p', 'spTree', NS.pml));
    if (!spTree) throw new Error('slide has no <p:spTree>');
    spTree.children.push(pic);

    this._commit();
    const previousLength = this._shapes.length;
    this._part = readSlidePart(this._document.root);
    this._shapes = this._part.shapes.map((s) => new SlideShape(this, s.element, s));
    const created = this._shapes[previousLength];
    if (!created) throw new Error('addImage: post-condition failed');
    return created;
  }

  /**
   * Sets a solid fill on the slide's background. Creates `<p:bg>` inside
   * `<p:cSld>` if absent. Any existing background reference (`p:bgRef`)
   * or `p:bgPr` is replaced.
   */
  setBackground(color: string): void {
    this._setBackground((bgPr) => setSolidFill(bgPr, color));
  }

  /** Clears any explicit slide background, restoring layout inheritance. */
  clearBackground(): void {
    const cSld = firstChildElement(this._document.root, qname('p', 'cSld', NS.pml));
    if (!cSld) return;
    cSld.children = cSld.children.filter(
      (c) => !(c.kind === 'element' && c.name.namespaceURI === NS.pml && c.name.localName === 'bg'),
    );
    this._commit();
    this._refresh();
  }

  private _setBackground(configure: (bgPr: XmlElement) => void): void {
    const cSld = firstChildElement(this._document.root, qname('p', 'cSld', NS.pml));
    if (!cSld) throw new Error('slide has no <p:cSld>');

    const bgName = qname('p', 'bg', NS.pml);
    const bgPrName = qname('p', 'bgPr', NS.pml);
    let bg = firstChildElement(cSld, bgName);
    if (bg === null) {
      bg = { kind: 'element', name: bgName, attrs: [], prefixDecls: new Map(), children: [] };
      // Per the schema, <p:bg> comes BEFORE <p:spTree> inside <p:cSld>.
      cSld.children.unshift(bg);
    }
    // Replace any existing bg child with a fresh bgPr.
    bg.children = [];
    const bgPr: XmlElement = {
      kind: 'element',
      name: bgPrName,
      attrs: [],
      prefixDecls: new Map(),
      children: [],
    };
    bg.children.push(bgPr);
    configure(bgPr);
    this._commit();
    this._refresh();
  }

  /**
   * Adds a table to this slide. Returns the new `SlideShape` (kind
   * `graphicFrame`). Table cells render as plain text with default
   * theme-aware styling; the table's `firstRow` and `bandRow` flags drive
   * PowerPoint's banded-header look unless `options` says otherwise.
   *
   * Cell content is row-major (`rows[i][j]` is row `i`, column `j`). All
   * rows must have the same length. Column widths and row heights default
   * to equal distribution of the frame's geometry.
   */
  addTable(opts: {
    x: Emu;
    y: Emu;
    w: Emu;
    h: Emu;
    rows: ReadonlyArray<ReadonlyArray<string>>;
    colWidths?: ReadonlyArray<Emu>;
    rowHeights?: ReadonlyArray<Emu>;
    firstRow?: boolean;
    bandRow?: boolean;
    name?: string;
  }): SlideShape {
    let maxId = 0;
    for (const s of this._part.shapes) {
      if (s.id > maxId) maxId = s.id;
    }
    const newId = Math.max(maxId, 1) + 1;

    const frame = buildTable({
      id: newId,
      ...(opts.name !== undefined ? { name: opts.name } : {}),
      x: opts.x,
      y: opts.y,
      w: opts.w,
      h: opts.h,
      rows: opts.rows,
      ...(opts.colWidths !== undefined ? { colWidths: opts.colWidths } : {}),
      ...(opts.rowHeights !== undefined ? { rowHeights: opts.rowHeights } : {}),
      ...(opts.firstRow !== undefined ? { firstRow: opts.firstRow } : {}),
      ...(opts.bandRow !== undefined ? { bandRow: opts.bandRow } : {}),
    });

    const cSld = firstChildElement(this._document.root, qname('p', 'cSld', NS.pml));
    if (!cSld) throw new Error('slide has no <p:cSld>');
    const spTree = firstChildElement(cSld, qname('p', 'spTree', NS.pml));
    if (!spTree) throw new Error('slide has no <p:spTree>');
    spTree.children.push(frame);

    this._commit();
    const previousLength = this._shapes.length;
    this._part = readSlidePart(this._document.root);
    this._shapes = this._part.shapes.map((s) => new SlideShape(this, s.element, s));
    const created = this._shapes[previousLength];
    if (!created) throw new Error('addTable: post-condition failed');
    return created;
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

  /**
   * @internal — return an `rId` on this slide's `.rels` that points at
   * `url` as an external hyperlink. Adds the relationship if no matching
   * one exists; otherwise reuses the existing rId.
   */
  _ensureHyperlinkRel(url: string): string {
    const rels = this._pkg.getRels(this._partName) ?? { items: [] };
    const existing = rels.items.find(
      (r) => r.type === REL_TYPES.hyperlink && r.target === url && r.targetMode === 'External',
    );
    if (existing) return existing.id;
    const nextId = nextRelId(rels.items.map((r) => r.id));
    rels.items.push({
      id: nextId,
      type: REL_TYPES.hyperlink,
      target: url,
      targetMode: 'External',
    });
    this._pkg.setRels(this._partName, rels);
    return nextId;
  }

  /**
   * @internal — drop the given shape XML element from the shape tree.
   * Used by `SlideShape.remove`.
   */
  _removeShape(element: XmlElement): void {
    const cSld = firstChildElement(this._document.root, qname('p', 'cSld', NS.pml));
    if (!cSld) return;
    const spTree = firstChildElement(cSld, qname('p', 'spTree', NS.pml));
    if (!spTree) return;
    const idx = spTree.children.indexOf(element);
    if (idx < 0) return;
    spTree.children.splice(idx, 1);
    this._commit();
    // Rebuild shape handles; the removed slot disappears and indices shift.
    this._part = readSlidePart(this._document.root);
    this._shapes = this._part.shapes.map((s) => new SlideShape(this, s.element, s));
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
   * Sets a solid-color outline (`<a:ln>`) on this shape. `widthEmu` is the
   * stroke width in EMU; use the unit helpers (`pt(1)`, `mm(0.5)`) to
   * spell it out at the call site.
   *
   * Width and color are independently optional — pass only what you want
   * to set. Existing outline attributes not addressed are preserved.
   */
  setStroke(options: { color?: string; widthEmu?: number }): void {
    const spPr = this._ensureSpPr();
    setSolidStroke(spPr, options as StrokeOptions);
    this._slide._commit();
    this._slide._refresh();
  }

  /** Sets an explicit "no outline" on this shape. */
  setNoStroke(): void {
    const spPr = this._ensureSpPr();
    setNoStroke(spPr);
    this._slide._commit();
    this._slide._refresh();
  }

  /** Removes any outline override, restoring inheritance. */
  clearStroke(): void {
    const spPr = this._ensureSpPr();
    clearStroke(spPr);
    this._slide._commit();
    this._slide._refresh();
  }

  /**
   * Sets a solid fill on this shape. Accepts the same color formats as
   * `setTextFormat({ color })`: `#RRGGBB`, `RRGGBB`, scheme tokens
   * (`accent1`, `bg1`, ...).
   *
   * The shape's `<p:spPr>` is created if it didn't yet exist. Any prior
   * fill (`noFill`, `gradFill`, `blipFill`, ...) is replaced.
   */
  setFill(color: string): void {
    const spPr = this._ensureSpPr();
    setSolidFill(spPr, color);
    this._slide._commit();
    this._slide._refresh();
  }

  /** Sets `<a:noFill>` on this shape, leaving it transparent. */
  setNoFill(): void {
    const spPr = this._ensureSpPr();
    setNoFill(spPr);
    this._slide._commit();
    this._slide._refresh();
  }

  /**
   * Clears any fill choice from this shape. The shape then inherits its
   * fill from the layout / master placeholder it descends from.
   */
  clearFill(): void {
    const spPr = this._ensureSpPr();
    clearFill(spPr);
    this._slide._commit();
    this._slide._refresh();
  }

  private _ensureSpPr(): XmlElement {
    // Only sp / pic / cxnSp use <p:spPr>; group shapes use grpSpPr and
    // graphic frames have no fill of their own.
    if (
      this._snapshot.kind !== 'shape' &&
      this._snapshot.kind !== 'picture' &&
      this._snapshot.kind !== 'connector'
    ) {
      throw new Error(`fill is not supported on ${this._snapshot.kind} shapes`);
    }
    const spPrName = qname('p', 'spPr', NS.pml);
    let spPr = firstChildElement(this._element, spPrName);
    if (spPr === null) {
      spPr = { kind: 'element', name: spPrName, attrs: [], prefixDecls: new Map(), children: [] };
      this._element.children.push(spPr);
    }
    return spPr;
  }

  /**
   * Applies `format` to every run in this shape's text body. Existing
   * run-property attributes not addressed by `format` are preserved, so
   * partial updates compose:
   *
   *     shape.setTextFormat({ bold: true });
   *     shape.setTextFormat({ color: '#FF0000' });
   *     // Result: bold + red.
   *
   * Throws if the shape has no text body.
   */
  setTextFormat(format: TextFormat): void {
    const txBody = firstChildElement(this._element, NAME_TX_BODY);
    if (txBody === null) {
      throw new Error(`shape "${this._snapshot.name}" has no <p:txBody>`);
    }
    applyFormatToAllRuns(txBody, format);
    this._slide._commit();
    this._slide._refresh();
  }

  /**
   * Replaces this shape's visible text with `value`. Newlines start a new
   * paragraph. The first existing run/paragraph properties are cloned so
   * font, color, size, alignment, and bullet style are preserved.
   *
   * Pass `options.bullets` to override the layout's bullet style for every
   * paragraph; common shorthands are `'bullet'` (• prefix) and `'number'`
   * (`1.`, `2.`, ...). For custom characters or numbering schemes, pass an
   * object: `{ char: '◆' }` or `{ autoNum: 'romanLcPeriod' }`.
   *
   * Throws if the shape is not a `<p:sp>` with a text body. Pictures,
   * graphic frames, groups, and connectors are not text-bearing.
   */
  setText(value: string, options: { bullets?: BulletStyle } = {}): void {
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
    if (options.bullets !== undefined) {
      applyBulletToAllParagraphs(txBody, options.bullets);
    }
    this._slide._commit();
    this._slide._refresh();
  }

  /**
   * Sets the horizontal alignment of every paragraph in this shape's
   * text body. Accepts plain names (`left` / `center` / `right` /
   * `justify` / `distribute`) and ECMA-376 tokens (`l` / `ctr` / `r` /
   * `just` / `dist` / `justLow` / `thaiDist`).
   */
  setAlignment(align: ParagraphAlignment): void {
    const txBody = this._requireTxBody();
    applyAlignmentToAllParagraphs(txBody, align);
    this._slide._commit();
    this._slide._refresh();
  }

  /**
   * Sets an external hyperlink on every run in this shape's text body.
   * Allocates a new `hyperlink` relationship on the slide's `.rels` (or
   * reuses an existing one with the same target). Pass `null` to clear.
   *
   * The hyperlink covers the entire shape text; per-run hyperlinks
   * require post-mutation editing of the AST directly for now.
   */
  setHyperlink(url: string | null): void {
    const txBody = this._requireTxBody();
    if (url === null) {
      applyHyperlinkToAllRuns(txBody, null);
    } else {
      const rId = this._slide._ensureHyperlinkRel(url);
      applyHyperlinkToAllRuns(txBody, rId);
    }
    this._slide._commit();
    this._slide._refresh();
  }

  private _requireTxBody(): XmlElement {
    if (this._snapshot.kind !== 'shape') {
      throw new Error(
        `text operations require a shape kind; ${this._snapshot.kind} is not text-bearing`,
      );
    }
    const txBody = firstChildElement(this._element, NAME_TX_BODY);
    if (txBody === null) {
      throw new Error(`shape "${this._snapshot.name}" has no <p:txBody>`);
    }
    return txBody;
  }

  /**
   * Sets the bullet style on every paragraph in this shape's text body
   * without touching the text content. Use to add or change bullets on
   * existing text.
   */
  setBullets(style: BulletStyle): void {
    if (this._snapshot.kind !== 'shape') {
      throw new Error(
        `setBullets only works on text-bearing shapes; ${this._snapshot.kind} is not one`,
      );
    }
    const txBody = firstChildElement(this._element, NAME_TX_BODY);
    if (txBody === null) {
      throw new Error(`shape "${this._snapshot.name}" has no <p:txBody>`);
    }
    applyBulletToAllParagraphs(txBody, style);
    this._slide._commit();
    this._slide._refresh();
  }

  /**
   * Sets this shape's position in EMU. Creates the `<a:xfrm>` host (and
   * its parent `<p:spPr>` / `<p:grpSpPr>` / `<p:xfrm>` as appropriate) if
   * the shape was previously inheriting position from its layout.
   */
  setPosition(x: Emu, y: Emu): void {
    writePosition(this._element, this._snapshot.kind, x, y);
    this._slide._commit();
    this._slide._refresh();
  }

  /** Sets this shape's size in EMU. Companion to `setPosition`. */
  setSize(w: Emu, h: Emu): void {
    writeSize(this._element, this._snapshot.kind, w, h);
    this._slide._commit();
    this._slide._refresh();
  }

  /**
   * Removes this shape from its slide's shape tree. Subsequent property
   * reads on this handle reflect the stale snapshot — discard the handle
   * after calling.
   *
   * Removing a picture does NOT delete the underlying media part — it may
   * be referenced from other slides. Use `_internalPackageOf(pres)` if
   * you need to garbage-collect orphan media.
   */
  remove(): void {
    this._slide._removeShape(this._element);
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
