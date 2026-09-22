import { describe, expect, it } from 'vitest';
import { resolveTextBodyRect } from '../packages/preview/src/text-body-rect.ts';

describe('preview text body rectangle', () => {
  const bounds = { x: 100, y: 200, w: 400, h: 200 };
  const margins = { left: 10, right: 20, top: 5, bottom: 15 };
  it('places triangle text in its lower middle, with asymmetric margins', () => {
    expect(resolveTextBodyRect('triangle', bounds, margins)).toEqual({
      x: 210,
      y: 305,
      w: 170,
      h: 80,
    });
  });
  it('keeps the full body for rectangular and unknown geometry', () => {
    for (const preset of ['rect', 'custom', null]) {
      expect(resolveTextBodyRect(preset, bounds, margins)).toEqual({
        x: 110,
        y: 205,
        w: 370,
        h: 180,
      });
    }
  });
  it('drops all insets when either dimension collapses a preset region', () => {
    expect(resolveTextBodyRect('diamond', bounds, { ...margins, top: 200 })).toEqual({
      x: 200,
      y: 250,
      w: 200,
      h: 100,
    });
    expect(resolveTextBodyRect('diamond', bounds, { ...margins, left: 200 })).toEqual({
      x: 200,
      y: 250,
      w: 200,
      h: 100,
    });
  });
  it('leaves degenerate rectangular regions for the renderer to reject', () => {
    expect(resolveTextBodyRect('rect', bounds, { ...margins, left: 400 }).w).toBe(-20);
  });
});
