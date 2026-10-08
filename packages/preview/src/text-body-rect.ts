import {
  getPresetGeometry,
  getShapeAdjustValues,
  getShapeBounds,
  getShapeCustomGeometry,
  getShapePreset,
  type SlideShapeData,
} from '@office-kit/pptx';

/** A text rectangle as fractions of the shape's extent. */
export interface TextRectFractions {
  readonly l: number;
  readonly t: number;
  readonly r: number;
  readonly b: number;
}

/**
 * Where a shape lays out its text, as fractions of its own extent: the
 * `<a:rect>` its geometry states, written in its `<a:custGeom>` or in
 * ECMA-376's definition of its preset (an ellipse's inscribed rectangle, a
 * triangle's lower middle, an arrow's shaft). `null` lays text into the whole
 * box: a shape that inherits its geometry, or whose rectangle has no room.
 *
 * The rectangle is in the shape's own `<a:ext>` space, so dividing by that —
 * not by on-screen bounds, which carry a group's or autofit's scale — is what
 * lets the result travel with the shape.
 */
export function shapeTextRect(shape: SlideShapeData): TextRectFractions | null {
  const extent = getShapeBounds(shape);
  if (extent === null || extent.w <= 0 || extent.h <= 0) return null;
  const preset = getShapePreset(shape);
  const geometry =
    preset === null
      ? getShapeCustomGeometry(shape)
      : getPresetGeometry(preset, extent, getShapeAdjustValues(shape));
  const rect = geometry?.textRect;
  if (!rect || rect.r <= rect.l || rect.b <= rect.t) return null;
  return {
    l: rect.l / extent.w,
    t: rect.t / extent.h,
    r: rect.r / extent.w,
    b: rect.b / extent.h,
  };
}

/**
 * Text layout rectangle used by the preview: `region` (from
 * {@link shapeTextRect}, or `null` for the whole box) placed in `bounds`, less
 * the body insets. All coordinates and margins must use the same unit
 * (normally EMU). When margins collapse a region, the region is kept without
 * them, as the reference desktop app does.
 */
export function resolveTextBodyRect(
  bounds: { x: number; y: number; w: number; h: number },
  margins: { left: number; top: number; right: number; bottom: number },
  region: TextRectFractions | null,
): { x: number; y: number; w: number; h: number } {
  const x = bounds.x + (region?.l ?? 0) * bounds.w;
  const y = bounds.y + (region?.t ?? 0) * bounds.h;
  const w = ((region?.r ?? 1) - (region?.l ?? 0)) * bounds.w;
  const h = ((region?.b ?? 1) - (region?.t ?? 0)) * bounds.h;
  const innerW = w - margins.left - margins.right;
  const innerH = h - margins.top - margins.bottom;
  if (region && (innerW <= 0 || innerH <= 0)) return { x, y, w, h };
  return { x: x + margins.left, y: y + margins.top, w: innerW, h: innerH };
}
