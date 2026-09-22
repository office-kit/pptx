// The format painter's clipboard: what "copy formatting" reads off a shape or
// a text range, and what "paste formatting" writes onto another.
//
// Everything here goes through the library's own readers and writers, so a
// copied format is the same data an author could have set by hand — never a
// lift of the source's XML. That is also the limit: a property the library can
// write but not read back cannot be copied, and `copiedFormatLimits` names the
// ones that apply to a given source so the UI can say so instead of silently
// dropping them.

import {
  clearShapeEffects,
  clearShapeFill,
  clearShapeStroke,
  getParagraphPropertiesEffective,
  getShapeEffects,
  getShapeFill,
  getShapeGradientFill,
  getShapeParagraphCount,
  getShapePatternFill,
  getShapeStroke,
  getShapeStrokeArrow,
  getShapeStrokeCap,
  getShapeStrokeCompound,
  getShapeStrokeDash,
  getShapeStrokeJoin,
  getTableCellParagraphs,
  setParagraphAlignment,
  setParagraphBullet,
  setParagraphLevel,
  setParagraphLineSpacing,
  setParagraphSpacing,
  setShapeFill,
  setShapeGlow,
  setShapeGradientFill,
  setShapeNoFill,
  setShapeNoStroke,
  setShapePatternFill,
  setShapeShadow,
  setShapeStroke,
  setShapeStrokeArrow,
  setShapeStrokeCap,
  setShapeStrokeCompound,
  setShapeStrokeDash,
  setShapeStrokeJoin,
  setShapeTextFormat,
  setTableCellTextFormat,
  type GradientFillOptions,
  type PresentationData,
  type SlideShapeData,
  type TableCellData,
  type TextFormat,
} from '@office-kit/pptx';

/** The paint half: everything that lives in a shape's `<p:spPr>`. */
interface ShapePaint {
  readonly fill:
    | { readonly kind: 'solid'; readonly color: string }
    | { readonly kind: 'gradient'; readonly gradient: GradientFillOptions }
    | {
        readonly kind: 'pattern';
        readonly pattern: NonNullable<ReturnType<typeof getShapePatternFill>>;
      }
    | { readonly kind: 'none' }
    | { readonly kind: 'inherit' }
    | null;
  readonly stroke:
    | { readonly kind: 'solid'; readonly color: string; readonly widthEmu?: number }
    | { readonly kind: 'none' }
    | { readonly kind: 'inherit' };
  readonly dash: ReturnType<typeof getShapeStrokeDash>;
  readonly cap: ReturnType<typeof getShapeStrokeCap>;
  readonly join: ReturnType<typeof getShapeStrokeJoin>;
  readonly compound: ReturnType<typeof getShapeStrokeCompound>;
  readonly headArrow: ReturnType<typeof getShapeStrokeArrow>;
  readonly tailArrow: ReturnType<typeof getShapeStrokeArrow>;
  readonly shadow: {
    color: string;
    blurEmu: number;
    offsetEmu: number;
    angleDeg: number;
    opacity?: number;
  } | null;
  readonly glow: { color: string; radiusEmu: number } | null;
}

/** Paragraph properties the editor can both read and write back. */
interface ParagraphFormat {
  readonly align: ReturnType<typeof getParagraphPropertiesEffective>['align'];
  readonly bullet: ReturnType<typeof getParagraphPropertiesEffective>['bullet'];
  readonly level: number;
  readonly lineSpacing: ReturnType<typeof getParagraphPropertiesEffective>['lineSpacing'];
  readonly spcBefPts: number | null;
  readonly spcAftPts: number | null;
}

export interface CopiedFormat {
  /** `null` when the format was copied from a text range, which has no paint. */
  readonly paint: ShapePaint | null;
  /** Both `null` when the source has no text: pasting it then leaves the
   *  target's own text formatting alone instead of clearing it. */
  readonly character: TextFormat | null;
  readonly paragraph: ParagraphFormat | null;
}

/** A shape, or one table cell inside one — both accept the same writers. */
export type FormatTarget = SlideShapeData | TableCellData;

const paragraphFormatOf = (
  pres: PresentationData,
  target: FormatTarget,
  index: number,
): ParagraphFormat => {
  const props = getParagraphPropertiesEffective(pres, target, index);
  return {
    align: props.align,
    bullet: props.bullet,
    level: props.level,
    lineSpacing: props.lineSpacing,
    spcBefPts: props.spcBefPts,
    spcAftPts: props.spcAftPts,
  };
};

const paintOf = (pres: PresentationData, shape: SlideShapeData): ShapePaint => {
  const fill = getShapeFill(shape);
  const effects = getShapeEffects(pres, shape);
  // `setShapeShadow` writes an outer shadow; an inner one has no writer, so it
  // is left behind rather than turned into something it is not.
  const shadow = effects.find((effect) => effect.kind === 'outerShdw');
  const glow = effects.find((effect) => effect.kind === 'glow');
  return {
    fill:
      fill.kind === 'solid'
        ? { kind: 'solid', color: fill.color }
        : fill.kind === 'gradient'
          ? gradientFillOf(shape)
          : fill.kind === 'pattern'
            ? patternFillOf(pres, shape)
            : fill.kind === 'image'
              ? // A picture fill is the shape's own image, not a format.
                null
              : { kind: fill.kind },
    stroke: getShapeStroke(shape),
    dash: getShapeStrokeDash(shape),
    cap: getShapeStrokeCap(shape),
    join: getShapeStrokeJoin(shape),
    compound: getShapeStrokeCompound(shape),
    headArrow: getShapeStrokeArrow(shape, 'head'),
    tailArrow: getShapeStrokeArrow(shape, 'tail'),
    shadow: shadow
      ? {
          color: shadow.color,
          blurEmu: shadow.blurEmu,
          offsetEmu: shadow.distEmu,
          angleDeg: shadow.angleDeg,
          ...(shadow.opacity === undefined ? {} : { opacity: shadow.opacity }),
        }
      : null,
    glow: glow ? { color: glow.color, radiusEmu: glow.radiusEmu } : null,
  };
};

const gradientFillOf = (shape: SlideShapeData): ShapePaint['fill'] => {
  const gradient = getShapeGradientFill(shape);
  return gradient ? { kind: 'gradient', gradient } : null;
};

const patternFillOf = (pres: PresentationData, shape: SlideShapeData): ShapePaint['fill'] => {
  const pattern = getShapePatternFill(pres, shape);
  return pattern ? { kind: 'pattern', pattern } : null;
};

/**
 * The whole format of a shape: its paint, and the character and paragraph
 * formatting its text starts with. A shape with no text still copies its
 * paint, and pasting it leaves the target's text alone.
 */
export function readShapeFormat(
  pres: PresentationData,
  shape: SlideShapeData,
  character: TextFormat | null,
): CopiedFormat {
  const hasText = getShapeParagraphCount(shape) > 0;
  return {
    paint: paintOf(pres, shape),
    character: hasText ? (character ?? {}) : null,
    paragraph: hasText ? paragraphFormatOf(pres, shape, 0) : null,
  };
}

/**
 * The format at a text selection: the character format the caller sampled from
 * the range, and the properties of the paragraph it starts in.
 */
export function readTextFormat(
  pres: PresentationData,
  target: FormatTarget,
  paragraphIndex: number,
  character: TextFormat | null,
): CopiedFormat {
  return { paint: null, character, paragraph: paragraphFormatOf(pres, target, paragraphIndex) };
}

/** What could not be copied off this shape, for the UI to report. */
export function copiedFormatLimits(pres: PresentationData, shape: SlideShapeData): string[] {
  const limits: string[] = [];
  if (getShapeFill(shape).kind === 'image') limits.push('picture fill');
  if (getShapeEffects(pres, shape).some((effect) => effect.kind === 'innerShdw'))
    limits.push('inner shadow');
  if (
    getShapeEffects(pres, shape).some(
      (effect) =>
        effect.kind === 'reflection' || effect.kind === 'softEdge' || effect.kind === 'blur',
    )
  )
    limits.push('reflection, soft edge or blur');
  return limits;
}

const applyPaint = (shape: SlideShapeData, paint: ShapePaint): void => {
  if (paint.fill !== null) {
    if (paint.fill.kind === 'solid') setShapeFill(shape, paint.fill.color);
    else if (paint.fill.kind === 'gradient') setShapeGradientFill(shape, paint.fill.gradient);
    else if (paint.fill.kind === 'pattern') setShapePatternFill(shape, paint.fill.pattern);
    else if (paint.fill.kind === 'none') setShapeNoFill(shape);
    else clearShapeFill(shape);
  }
  if (paint.stroke.kind === 'solid') {
    setShapeStroke(shape, {
      color: paint.stroke.color,
      ...(paint.stroke.widthEmu === undefined ? {} : { widthEmu: paint.stroke.widthEmu }),
    });
    // Stroke detail only means anything once the outline exists, and each
    // writer is skipped when the source said nothing — writing a default
    // would invent an outline style the source never had.
    if (paint.dash) setShapeStrokeDash(shape, paint.dash);
    if (paint.cap) setShapeStrokeCap(shape, paint.cap);
    if (paint.join) setShapeStrokeJoin(shape, paint.join);
    if (paint.compound) setShapeStrokeCompound(shape, paint.compound);
    if (paint.headArrow) setShapeStrokeArrow(shape, 'head', paint.headArrow);
    if (paint.tailArrow) setShapeStrokeArrow(shape, 'tail', paint.tailArrow);
  } else if (paint.stroke.kind === 'none') setShapeNoStroke(shape);
  else clearShapeStroke(shape);
  // One clear, then whichever effects the source had: pasting a format
  // replaces the target's effects rather than merging into them.
  clearShapeEffects(shape);
  if (paint.shadow) setShapeShadow(shape, paint.shadow);
  if (paint.glow) setShapeGlow(shape, paint.glow);
};

const applyParagraph = (target: FormatTarget, index: number, paragraph: ParagraphFormat): void => {
  // `null` means the source authored nothing anywhere up its cascade, and
  // there is no writer for "inherit" — so the target keeps what it inherits,
  // which is the closest thing to what the source is showing.
  if (paragraph.align !== null) setParagraphAlignment(target, index, paragraph.align);
  if (paragraph.bullet !== null) setParagraphBullet(target, index, paragraph.bullet);
  setParagraphLevel(target, index, paragraph.level);
  setParagraphLineSpacing(target, index, paragraph.lineSpacing);
  setParagraphSpacing(target, index, {
    beforePts: paragraph.spcBefPts,
    afterPts: paragraph.spcAftPts,
  });
};

const paragraphCount = (target: FormatTarget, isCell: boolean): number =>
  isCell
    ? getTableCellParagraphs(target as TableCellData).length
    : getShapeParagraphCount(target as SlideShapeData);

/**
 * Pastes a copied format onto a whole shape: its paint, and its text's
 * character and paragraph formatting. A format copied from a text range has no
 * paint and only restyles the text.
 */
export function applyShapeFormat(
  shape: SlideShapeData,
  format: CopiedFormat,
  options: { readonly text: boolean } = { text: true },
): void {
  if (format.paint) applyPaint(shape, format.paint);
  if (!options.text || format.character === null || format.paragraph === null) return;
  const count = getShapeParagraphCount(shape);
  if (count === 0) return;
  setShapeTextFormat(shape, format.character, { reset: true });
  for (let index = 0; index < count; index++) applyParagraph(shape, index, format.paragraph);
}

/**
 * Pastes a copied format onto one text range: the character format over the
 * range, and the paragraph properties onto each paragraph it touches. Paint is
 * ignored — a range has nowhere to put it.
 */
export function applyTextFormat(
  target: FormatTarget,
  range: { readonly start: number; readonly end: number },
  paragraphIndices: readonly number[],
  format: CopiedFormat,
  isCell: boolean,
): void {
  if (format.character !== null && range.end > range.start) {
    if (isCell)
      setTableCellTextFormat(target as TableCellData, format.character, { range, reset: true });
    else setShapeTextFormat(target as SlideShapeData, format.character, { range, reset: true });
  }
  if (format.paragraph === null) return;
  const paragraph = format.paragraph;
  const count = paragraphCount(target, isCell);
  for (const index of paragraphIndices) {
    if (index < count) applyParagraph(target, index, paragraph);
  }
}
