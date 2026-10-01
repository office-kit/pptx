import assert from 'node:assert/strict';
import test from 'node:test';
import {
  addSlide,
  addSlideTextBox,
  createPresentation,
  getSlideLayouts,
  getSlideLayoutName,
  getSlideShapes,
  getShapeId,
  getShapePlaceholderType,
  setShapeText,
  savePresentation,
  loadPresentation,
  getSlides,
} from '@office-kit/pptx';
import { outlineShapes } from '../src/lib/editor/core/outline.ts';

test('outline keeps empty title/body placeholders but excludes ordinary text boxes', async () => {
  const pres = createPresentation();
  const layout = getSlideLayouts(pres).find(
    (item) => getSlideLayoutName(item) === 'Title and Content',
  );
  assert.ok(layout);
  const slide = addSlide(pres, { layout });
  const placeholders = getSlideShapes(slide);
  for (const shape of placeholders) setShapeText(shape, '');
  addSlideTextBox(slide, { x: 0, y: 0, w: 914400, h: 914400, text: 'Ordinary text' });
  const expected = placeholders.map((shape) => ({
    id: getShapeId(shape),
    title: getShapePlaceholderType(shape) === 'title',
  }));
  assert.equal(expected.length, 2);
  assert.deepEqual(outlineShapes(slide), expected);
  const loaded = await loadPresentation(await savePresentation(pres));
  assert.deepEqual(outlineShapes(getSlides(loaded)[0]), expected);
});
