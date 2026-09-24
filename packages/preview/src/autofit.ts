// The factor the preview shrinks a shape's text by, exposed on its own.
//
// `<a:normAutofit/>` without a baked `fontScale` means "shrink to fit", and
// PowerPoint computes that reduction at display time — so does the renderer,
// once per shape, with the measurer it lays text out with. An editing surface
// that paints its own text over the preview has to shrink by the SAME factor
// or the glyphs jump the moment the caret appears.

import {
  getPresentationTheme,
  getShapeBoundsResolved,
  getShapePlaceholderType,
  type PresentationData,
  type SlideShapeData,
} from '@office-kit/pptx';
import { resolveTextBodyModel } from './render-slide.ts';
import { defaultMeasurer, type TextMeasurer } from './text-layout.ts';

export interface ShapeAutoFitScaleOptions {
  /** The box to fit into, in EMU. Defaults to the shape's resolved bounds.
   *  Pass it for a shape being edited, whose box on screen is not (yet) the
   *  one the model carries. */
  readonly bounds?: {
    readonly x: number;
    readonly y: number;
    readonly w: number;
    readonly h: number;
  };
  /** Measurer for run widths, defaulting to the heuristic one the renderer
   *  uses in a browser. Pass the same measurer the render call uses, or the
   *  two disagree about how much text fits. */
  readonly measureText?: TextMeasurer;
}

/**
 * How much the preview shrinks `shape`'s text to fit its box: a multiplier on
 * every authored size — font size, line spacing in points, paragraph spacing
 * and indents.
 *
 * `1` for a shape with no `<a:normAutofit>`: `<a:noAutofit>` overflows and
 * `<a:spAutoFit>` grows the box instead, so neither shrinks the text. A shape
 * with a baked `fontScale` reports that, unchanged — it is what PowerPoint
 * itself applied.
 */
export function shapeAutoFitScale(
  pres: PresentationData,
  shape: SlideShapeData,
  options: ShapeAutoFitScaleOptions = {},
): number {
  const bounds = options.bounds ?? getShapeBoundsResolved(pres, shape);
  if (!bounds) return 1;
  const model = resolveTextBodyModel(
    pres,
    shape,
    { x: bounds.x as number, y: bounds.y as number, w: bounds.w as number, h: bounds.h as number },
    getPresentationTheme(pres),
    getShapePlaceholderType(shape),
    options.measureText ?? defaultMeasurer,
    '#000000', // colors don't affect metrics
  );
  return model?.autoFitScale ?? 1;
}

/** Translation in preview pixels applied to the entire text block by centered
 * text anchoring. Paragraph widths and alignment remain unchanged. Editing
 * surfaces must apply the same offset to keep the caret over the preview. */
export function shapeTextAnchorOffset(
  pres: PresentationData,
  shape: SlideShapeData,
  options: ShapeAutoFitScaleOptions = {},
): { readonly x: number; readonly y: number } {
  const bounds = options.bounds ?? getShapeBoundsResolved(pres, shape);
  if (!bounds) return { x: 0, y: 0 };
  return (
    resolveTextBodyModel(
      pres,
      shape,
      {
        x: bounds.x as number,
        y: bounds.y as number,
        w: bounds.w as number,
        h: bounds.h as number,
      },
      getPresentationTheme(pres),
      getShapePlaceholderType(shape),
      options.measureText ?? defaultMeasurer,
      '#000000',
    )?.anchorOffset ?? { x: 0, y: 0 }
  );
}
