import { describe, expect, it } from 'vitest';
import {
  addMissingSlidePlaceholders,
  addSlide,
  addSlideTextBox,
  createPresentation,
  findSlideLayout,
  getSlideShapes,
  getSlides,
  getShapeText,
  getShapeId,
  getShapeBoundsResolved,
  getShapeXmlString,
  isShapePlaceholder,
  removeShape,
  setShapeText,
  setShapePosition,
  setShapeSize,
  savePresentation,
  loadPresentation,
  inches,
  groupShapes,
} from '../src/api/index.ts';

describe('addMissingSlidePlaceholders', () => {
  it('restores only deleted slots with inherited geometry, empty content and unique IDs', async () => {
    const pres = createPresentation();
    const slide = addSlide(pres, { layout: findSlideLayout(pres, 'Title and Content')! });
    const [title, body] = getSlideShapes(slide);
    const expected = getShapeBoundsResolved(pres, body!);
    setShapeText(title!, 'Keep 日本語 / English');
    setShapePosition(title!, inches(3), inches(2));
    const retained = getShapeXmlString(title!);
    removeShape(body!);
    const decoration = addSlideTextBox(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(2),
      h: inches(1),
      text: 'Decoration',
    });
    const decorationXml = getShapeXmlString(decoration);
    expect(addMissingSlidePlaceholders(slide)).toBe(1);
    expect(getShapeXmlString(title!)).toBe(retained);
    expect(getShapeXmlString(decoration)).toBe(decorationXml);
    const shapes = getSlideShapes(slide);
    expect(new Set(shapes.map(getShapeId)).size).toBe(shapes.length);
    const restored = shapes.filter(isShapePlaceholder)[1]!;
    expect(getShapeText(restored)).toBe('');
    expect(getShapeBoundsResolved(pres, restored)).toEqual(expected);
    expect(getShapeXmlString(restored)).not.toContain('xfrm');
    expect(addMissingSlidePlaceholders(slide)).toBe(0);
    setShapeText(restored, 'Restored 本文');
    const reloaded = await loadPresentation(await savePresentation(pres));
    expect(getSlideShapes(getSlides(reloaded)[0]!).map(getShapeText)).toEqual([
      'Keep 日本語 / English',
      'Decoration',
      'Restored 本文',
    ]);
  });

  it('restores an empty slide and leaves blank layouts alone', () => {
    const pres = createPresentation();
    const slide = addSlide(pres, { layout: findSlideLayout(pres, 'Title and Content')! });
    const count = getSlideShapes(slide).length;
    for (const shape of getSlideShapes(slide)) removeShape(shape);
    expect(addMissingSlidePlaceholders(slide)).toBe(count);
    expect(addMissingSlidePlaceholders(slide)).toBe(0);
    const blank = addSlide(pres, { layout: findSlideLayout(pres, 'Blank')! });
    expect(addMissingSlidePlaceholders(blank)).toBe(0);
  });

  it('does not duplicate placeholders already inside groups', () => {
    const pres = createPresentation();
    const slide = addSlide(pres, { layout: findSlideLayout(pres, 'Title and Content')! });
    for (const shape of getSlideShapes(slide)) {
      const bounds = getShapeBoundsResolved(pres, shape)!;
      setShapePosition(shape, bounds.x, bounds.y);
      setShapeSize(shape, bounds.w, bounds.h);
    }
    groupShapes([...getSlideShapes(slide)]);
    const before = getSlideShapes(slide).map(getShapeXmlString);
    expect(addMissingSlidePlaceholders(slide)).toBe(0);
    expect(getSlideShapes(slide).map(getShapeXmlString)).toEqual(before);
  });
});
