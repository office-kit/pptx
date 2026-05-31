// Structural-similarity metric for the fidelity harness.
//
// SSIM is the standard perceptual image-similarity measure (Wang et al.,
// 2004). It correlates with "do these look the same to a human" far better
// than per-pixel error, which is why the preview-fidelity roadmap uses it as
// the gate. We compute mean SSIM over a sliding window on the luma channel,
// plus a coarse pixel-difference percentage that is easier to reason about
// when triaging a single slide.
//
// This module is pure (no Node/browser deps) so it is unit-tested directly.

import type { RgbaImage } from './image.ts';
import { toGray } from './image.ts';

export interface CompareResult {
  /** Mean SSIM in [-1, 1]; 1.0 === identical. */
  readonly ssim: number;
  /**
   * Foreground-weighted SSIM. Each window is weighted by how much "ink" the
   * ground truth has there (darkness away from white), so blank slide areas —
   * which dominate plain SSIM and reward a renderer for drawing *nothing* —
   * barely count. This is the metric that actually tracks content fidelity:
   * missing text scores low here, correctly rendered text scores high.
   */
  readonly fgSsim: number;
  /** Mean absolute luma error normalized to [0, 1]. */
  readonly meanAbsError: number;
  /** Fraction of pixels whose luma differs by more than `diffThreshold`. */
  readonly diffPercent: number;
}

// 8.0 default constants from the SSIM paper, scaled to an 8-bit range.
const C1 = (0.01 * 255) ** 2;
const C2 = (0.03 * 255) ** 2;

export interface CompareOptions {
  /** Window edge length in pixels. */
  readonly window?: number;
  /** Window stride in pixels. */
  readonly stride?: number;
  /** Luma delta (0-255) above which a pixel counts toward diffPercent. */
  readonly diffThreshold?: number;
}

/** Compare two equally-sized images. Throws if the dimensions differ. */
export const compareImages = (
  a: RgbaImage,
  b: RgbaImage,
  opts: CompareOptions = {},
): CompareResult => {
  if (a.width !== b.width || a.height !== b.height) {
    throw new Error(
      `compareImages: dimension mismatch ${a.width}x${a.height} vs ${b.width}x${b.height}`,
    );
  }
  const window = opts.window ?? 8;
  const stride = opts.stride ?? 4;
  const diffThreshold = opts.diffThreshold ?? 16;
  const { width, height } = a;
  const ga = toGray(a);
  const gb = toGray(b);

  let ssimSum = 0;
  let windowCount = 0;
  let fgWeightedSum = 0;
  let fgWeight = 0;
  // Slide the window; clamp the final window flush against the right/bottom
  // edge so every pixel is covered even when (size - window) % stride !== 0.
  for (let wy = 0; wy <= height - window; wy += stepTo(wy, height, window, stride)) {
    for (let wx = 0; wx <= width - window; wx += stepTo(wx, width, window, stride)) {
      const w = windowSsim(ga, gb, width, wx, wy, window);
      ssimSum += w.ssim;
      windowCount++;
      // Ground-truth ink weight: 0 for a white window, →1 as it darkens.
      const ink = Math.max(0, (255 - w.muA) / 255);
      fgWeightedSum += w.ssim * ink;
      fgWeight += ink;
      if (wx === width - window) break;
    }
    if (wy === height - window) break;
  }

  let absErr = 0;
  let diffCount = 0;
  for (let i = 0; i < ga.length; i++) {
    const d = Math.abs(ga[i]! - gb[i]!);
    absErr += d;
    if (d > diffThreshold) diffCount++;
  }

  return {
    ssim: windowCount > 0 ? ssimSum / windowCount : 1,
    fgSsim: fgWeight > 0 ? fgWeightedSum / fgWeight : 1,
    meanAbsError: ga.length > 0 ? absErr / (ga.length * 255) : 0,
    diffPercent: ga.length > 0 ? diffCount / ga.length : 0,
  };
};

// Advance by `stride`, but if the next step would overshoot the edge, jump
// exactly to the edge so the last window is flush.
const stepTo = (pos: number, size: number, window: number, stride: number): number => {
  const edge = size - window;
  if (pos === edge) return stride; // loop guard handles the break
  return pos + stride > edge ? edge - pos : stride;
};

const windowSsim = (
  ga: Float64Array,
  gb: Float64Array,
  width: number,
  wx: number,
  wy: number,
  window: number,
): { ssim: number; muA: number } => {
  const n = window * window;
  let sumA = 0;
  let sumB = 0;
  let sumAA = 0;
  let sumBB = 0;
  let sumAB = 0;
  for (let y = 0; y < window; y++) {
    const row = (wy + y) * width + wx;
    for (let x = 0; x < window; x++) {
      const va = ga[row + x]!;
      const vb = gb[row + x]!;
      sumA += va;
      sumB += vb;
      sumAA += va * va;
      sumBB += vb * vb;
      sumAB += va * vb;
    }
  }
  const muA = sumA / n;
  const muB = sumB / n;
  const varA = sumAA / n - muA * muA;
  const varB = sumBB / n - muB * muB;
  const covAB = sumAB / n - muA * muB;
  const num = (2 * muA * muB + C1) * (2 * covAB + C2);
  const den = (muA * muA + muB * muB + C1) * (varA + varB + C2);
  return { ssim: num / den, muA };
};

// A human-readable diff: faded grayscale of the ground truth, with pixels
// that differ painted toward red in proportion to the luma delta.
export const diffImage = (truth: RgbaImage, ours: RgbaImage): RgbaImage => {
  const gt = toGray(truth);
  const go = toGray(ours);
  const data = new Uint8Array(truth.width * truth.height * 4);
  for (let i = 0, p = 0; i < gt.length; i++, p += 4) {
    const base = Math.round(gt[i]! * 0.25 + 191); // wash out to a light gray
    const delta = Math.min(255, Math.abs(gt[i]! - go[i]!));
    data[p] = Math.min(255, base + delta);
    data[p + 1] = Math.max(0, base - delta);
    data[p + 2] = Math.max(0, base - delta);
    data[p + 3] = 255;
  }
  return { width: truth.width, height: truth.height, data };
};
