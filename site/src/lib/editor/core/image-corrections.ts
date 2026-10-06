/** SVG transfer coefficients for DrawingML luminance, in sRGB channel space. */
export function imageCorrections(
  brightness: number,
  contrast: number,
): { slope: number; intercept: number } {
  // PowerPoint applies half the brightness before contrast and half afterwards.
  // LibreOffice Bitmap::Adjust(msoBrightness=true) documents this compatibility rule.
  const midpoint = 128 / 255;
  const contrastScale = 127 / 128;
  const slope = contrast >= 0 ? 1 / (1 - contrastScale * contrast) : 1 + contrastScale * contrast;
  return { slope, intercept: midpoint * (1 - slope) + (brightness * (1 + slope)) / 2 };
}
