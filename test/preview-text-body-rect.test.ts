import { describe, expect, it } from 'vitest';
import {
  resolveTextBodyRect,
  shapeCustomTextRect,
} from '../packages/preview/src/text-body-rect.ts';

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

  // A custom shape says where its own text goes, so nothing is approximated
  // and no preset is consulted — even one that has a region of its own.
  describe('custom geometry', () => {
    const custom = { l: 0.5, t: 0, r: 1, b: 0.5 };
    it('is used as stated, in place of any preset region', () => {
      for (const preset of [null, 'triangle']) {
        expect(resolveTextBodyRect(preset, bounds, margins, custom)).toEqual({
          x: 310,
          y: 205,
          w: 170,
          h: 80,
        });
      }
    });
    it('travels with bounds a group has scaled', () => {
      const scaled = { x: 0, y: 0, w: 800, h: 400 };
      expect(resolveTextBodyRect(null, scaled, margins, custom)).toEqual({
        x: 410,
        y: 5,
        w: 370,
        h: 180,
      });
    });
    it('drops its insets rather than collapsing, like a preset region', () => {
      expect(resolveTextBodyRect(null, bounds, { ...margins, top: 200 }, custom)).toEqual({
        x: 300,
        y: 200,
        w: 200,
        h: 100,
      });
    });
  });
});

describe('a custom shape’s own text rectangle', () => {
  it('is taken as fractions of the extent it was written against', () => {
    expect(
      shapeCustomTextRect({ textRect: { l: 100, t: 200, r: 900, b: 700 } }, { w: 1000, h: 800 }),
    ).toEqual({ l: 0.1, t: 0.25, r: 0.9, b: 0.875 });
  });
  it('is null without a rect, without an extent, or with an empty one', () => {
    expect(shapeCustomTextRect(null, { w: 1000, h: 800 })).toBeNull();
    expect(shapeCustomTextRect({ textRect: null }, { w: 1000, h: 800 })).toBeNull();
    expect(shapeCustomTextRect({ textRect: { l: 0, t: 0, r: 1, b: 1 } }, null)).toBeNull();
    expect(
      shapeCustomTextRect({ textRect: { l: 0, t: 0, r: 1, b: 1 } }, { w: 0, h: 800 }),
    ).toBeNull();
  });
});
