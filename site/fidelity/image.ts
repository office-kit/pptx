// Minimal RGBA image utilities for the fidelity harness. Pure (no Node or
// browser APIs) so the SSIM metric can be unit-tested in the root vitest
// suite without LibreOffice / resvg present.

export interface RgbaImage {
  readonly width: number;
  readonly height: number;
  /** Row-major RGBA, 4 bytes per pixel, length === width * height * 4. */
  readonly data: Uint8Array;
}

/** ITU-R BT.601 luma. Returns a width*height array in [0, 255]. */
export const toGray = (img: RgbaImage): Float64Array => {
  const { width, height, data } = img;
  const out = new Float64Array(width * height);
  for (let i = 0, p = 0; i < out.length; i++, p += 4) {
    out[i] = 0.299 * data[p]! + 0.587 * data[p + 1]! + 0.114 * data[p + 2]!;
  }
  return out;
};

/** Pack a tightly-packed RGB buffer (3 bytes/pixel, e.g. from a PPM) into RGBA. */
export const rgbToRgba = (rgb: Uint8Array, width: number, height: number): RgbaImage => {
  const data = new Uint8Array(width * height * 4);
  for (let i = 0, s = 0, d = 0; i < width * height; i++, s += 3, d += 4) {
    data[d] = rgb[s]!;
    data[d + 1] = rgb[s + 1]!;
    data[d + 2] = rgb[s + 2]!;
    data[d + 3] = 255;
  }
  return { width, height, data };
};

// Bilinear resize. The ground-truth raster (from PDF) and our SVG raster can
// land a pixel or two apart even when asked for the same dimensions; SSIM
// needs identical dimensions, so we coerce ours onto the ground truth's grid.
export const resizeRgba = (img: RgbaImage, dstW: number, dstH: number): RgbaImage => {
  if (img.width === dstW && img.height === dstH) return img;
  const { width: srcW, height: srcH, data: src } = img;
  const dst = new Uint8Array(dstW * dstH * 4);
  // Map dst pixel centers back into src space.
  const sx = srcW / dstW;
  const sy = srcH / dstH;
  for (let y = 0; y < dstH; y++) {
    const fy = (y + 0.5) * sy - 0.5;
    const y0 = Math.max(0, Math.floor(fy));
    const y1 = Math.min(srcH - 1, y0 + 1);
    const wy = fy - y0;
    for (let x = 0; x < dstW; x++) {
      const fx = (x + 0.5) * sx - 0.5;
      const x0 = Math.max(0, Math.floor(fx));
      const x1 = Math.min(srcW - 1, x0 + 1);
      const wx = fx - x0;
      const i00 = (y0 * srcW + x0) * 4;
      const i01 = (y0 * srcW + x1) * 4;
      const i10 = (y1 * srcW + x0) * 4;
      const i11 = (y1 * srcW + x1) * 4;
      const d = (y * dstW + x) * 4;
      for (let c = 0; c < 4; c++) {
        const top = src[i00 + c]! * (1 - wx) + src[i01 + c]! * wx;
        const bot = src[i10 + c]! * (1 - wx) + src[i11 + c]! * wx;
        dst[d + c] = Math.round(top * (1 - wy) + bot * wy);
      }
    }
  }
  return { width: dstW, height: dstH, data: dst };
};
