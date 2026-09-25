// Non-rectangular autoshapes inscribe their text in a rect narrower than the
// bounding box, so a label that fits the box still wraps inside the shape (a
// triangle's text sits in its lower-middle, a diamond's in its center square,
// etc.). Only shapes whose text rect is materially narrower than the box are
// listed; everything else (rect / roundRect / ellipse / hexagon / octagon /
// star8 / rightArrow …) keeps the full box. Values are fractions that track the
// polygon PRESET_POINTS actually draws (so text stays inside the rendered ink).
const presetTextRect = (
  preset: string | null,
): { l: number; t: number; r: number; b: number } | null => {
  switch (preset) {
    case 'triangle':
      return { l: 0.25, t: 0.5, r: 0.75, b: 1.0 };
    case 'diamond':
      return { l: 0.25, t: 0.25, r: 0.75, b: 0.75 };
    case 'pentagon':
      return { l: 0.191, t: 0.236, r: 0.809, b: 1.0 };
    case 'star5':
      return { l: 0.309, t: 0.382, r: 0.691, b: 0.764 };
    case 'leftRightArrow':
      // Matches the fixed leftRightArrow polygon's central shaft (x 0.18..0.82,
      // y 0.35..0.65), which is not size-aware like the cardinal arrows.
      return { l: 0.18, t: 0.35, r: 0.82, b: 0.65 };
    default:
      return null;
  }
};

/**
 * The text rectangle a custom-geometry shape states for itself, as fractions
 * of its own extent, or `null` for a preset shape and for a custom one that
 * states none. `<a:rect>` is written in EMU against the shape's extents, so
 * dividing by them here is what lets the result travel with a shape that a
 * group has scaled.
 */
export function shapeCustomTextRect(
  custom: { textRect: { l: number; t: number; r: number; b: number } | null } | null,
  extent: { w: number; h: number } | null,
): { l: number; t: number; r: number; b: number } | null {
  const rect = custom?.textRect;
  if (!rect || extent === null || extent.w <= 0 || extent.h <= 0) return null;
  return {
    l: rect.l / extent.w,
    t: rect.t / extent.h,
    r: rect.r / extent.w,
    b: rect.b / extent.h,
  };
}

/**
 * Text layout rectangle used by the preview, including preset geometry and body
 * insets. All coordinates and margins must use the same unit (normally EMU).
 * Preset regions follow the preview's current geometry approximations. When
 * margins collapse a preset region, retain that region without margins.
 *
 * A custom-geometry shape states its own rectangle in `<a:custGeom><a:rect>`.
 * Pass it as `custom`, as fractions of the shape's extent so that it scales
 * with `bounds` the way a preset region does, and it wins: unlike the table
 * above it is not an approximation of anything — it is what the file says.
 */
export function resolveTextBodyRect(
  preset: string | null,
  bounds: { x: number; y: number; w: number; h: number },
  margins: { left: number; top: number; right: number; bottom: number },
  custom?: { l: number; t: number; r: number; b: number } | null,
): { x: number; y: number; w: number; h: number } {
  const region = custom ?? presetTextRect(preset);
  const x = bounds.x + (region?.l ?? 0) * bounds.w;
  const y = bounds.y + (region?.t ?? 0) * bounds.h;
  const w = ((region?.r ?? 1) - (region?.l ?? 0)) * bounds.w;
  const h = ((region?.b ?? 1) - (region?.t ?? 0)) * bounds.h;
  const innerW = w - margins.left - margins.right;
  const innerH = h - margins.top - margins.bottom;
  if (region && (innerW <= 0 || innerH <= 0)) return { x, y, w, h };
  return { x: x + margins.left, y: y + margins.top, w: innerW, h: innerH };
}
