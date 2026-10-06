import { describe, expect, it } from 'vitest';
import {
  addBlankSlide,
  addSlideShape,
  createPresentation,
  emu,
  inches,
  setShapeAdjustValues,
  setShapeCustomGeometry,
  type Emu,
} from '@office-kit/pptx';
import { resolveTextBodyRect, shapeTextRect } from '../packages/preview/src/text-body-rect.ts';

describe('preview text body rectangle', () => {
  const bounds = { x: 100, y: 200, w: 400, h: 200 };
  const margins = { left: 10, right: 20, top: 5, bottom: 15 };
  const region = { l: 0.5, t: 0, r: 1, b: 0.5 };
  it('places the region in the bounds, less asymmetric margins', () => {
    expect(resolveTextBodyRect(bounds, margins, region)).toEqual({ x: 310, y: 205, w: 170, h: 80 });
  });
  it('keeps the full body without a region', () => {
    expect(resolveTextBodyRect(bounds, margins, null)).toEqual({ x: 110, y: 205, w: 370, h: 180 });
  });
  it('travels with bounds a group has scaled', () => {
    const scaled = { x: 0, y: 0, w: 800, h: 400 };
    expect(resolveTextBodyRect(scaled, margins, region)).toEqual({ x: 410, y: 5, w: 370, h: 180 });
  });
  it('drops the insets rather than collapsing a region', () => {
    expect(resolveTextBodyRect(bounds, { ...margins, top: 200 }, region)).toEqual({
      x: 300,
      y: 200,
      w: 200,
      h: 100,
    });
  });
  it('leaves a degenerate full body for the renderer to reject', () => {
    expect(resolveTextBodyRect(bounds, { ...margins, left: 400 }, null).w).toBe(-20);
  });
});

describe('a shape’s own text rectangle', () => {
  const shape = (preset: string, w: Emu = inches(4), h: Emu = inches(2)) =>
    addSlideShape(addBlankSlide(createPresentation()), { preset, x: emu(0), y: emu(0), w, h });

  it('is the inscribed rectangle of an ellipse, from its preset definition', () => {
    const rect = shapeTextRect(shape('ellipse'))!;
    // ECMA-376 inscribes the text at cos 45° of each radius from the center.
    const inset = (1 - Math.SQRT1_2) / 2;
    expect(rect.l).toBeCloseTo(inset, 4);
    expect(rect.t).toBeCloseTo(inset, 4);
    expect(rect.r).toBeCloseTo(1 - inset, 4);
    expect(rect.b).toBeCloseTo(1 - inset, 4);
  });
  it('is the lower middle of a triangle', () => {
    expect(shapeTextRect(shape('triangle'))).toEqual({ l: 0.25, t: 0.5, r: 0.75, b: 1 });
  });
  it('follows the adjust handles', () => {
    const arrow = shape('rightArrow');
    const shaft = shapeTextRect(arrow)!;
    setShapeAdjustValues(arrow, { adj1: 80000 });
    const wider = shapeTextRect(arrow)!;
    expect(wider.b - wider.t).toBeGreaterThan(shaft.b - shaft.t);
  });
  it('is the whole box for a rectangle', () => {
    expect(shapeTextRect(shape('rect'))).toEqual({ l: 0, t: 0, r: 1, b: 1 });
  });
  it('is what a custom shape states', () => {
    const custom = shape('rect', emu(1000), emu(800));
    setShapeCustomGeometry(custom, {
      paths: [{ w: 1000, h: 800, commands: [{ kind: 'moveTo', pt: { x: 0, y: 0 } }] }],
    });
    // `setShapeCustomGeometry` states the whole box.
    expect(shapeTextRect(custom)).toEqual({ l: 0, t: 0, r: 1, b: 1 });
  });
});
