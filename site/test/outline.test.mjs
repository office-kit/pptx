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
  getSlideTitle,
  getShapeText,
  setParagraphLevel,
  getParagraphLevel,
  setShapeRunHyperlink,
  getShapeRunHyperlink,
} from '@office-kit/pptx';
import { outlineShapes, promoteOutlineBody } from '../src/lib/editor/core/outline.ts';

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

test('promotes each selected root paragraph into a title and retains the following body and links', async () => {
  const pres = createPresentation();
  const layout = getSlideLayouts(pres).find(
    (item) => getSlideLayoutName(item) === 'Title and Content',
  );
  const slide = addSlide(pres, { layout });
  const body = getSlideShapes(slide).find((shape) =>
    ['obj', 'body', null].includes(getShapePlaceholderType(shape)),
  );
  assert.ok(body);
  setShapeText(body, 'First\nSecond\nThird\nFollowing');
  setShapeRunHyperlink(body, 2, 0, 'https://example.com/third');
  const added = promoteOutlineBody(pres, slide, body, { start: 6, end: 18 });
  assert.equal(added.length, 2);
  assert.equal(getShapeText(body), 'First');
  const loaded = await loadPresentation(await savePresentation(pres));
  const slides = getSlides(loaded);
  assert.deepEqual(slides.slice(1).map(getSlideTitle), ['Second', 'Third']);
  const last = getSlideShapes(slides[2]);
  assert.equal(
    getShapeText(
      last.find((shape) => ['obj', 'body', null].includes(getShapePlaceholderType(shape))),
    ),
    'Following',
  );
  assert.equal(
    getShapeRunHyperlink(
      last.find((shape) => getShapePlaceholderType(shape) === 'title'),
      0,
      0,
    ),
    'https://example.com/third',
  );
});

test('promotes a nested paragraph one level without creating slides', () => {
  const pres = createPresentation();
  const layout = getSlideLayouts(pres).find(
    (item) => getSlideLayoutName(item) === 'Title and Content',
  );
  const slide = addSlide(pres, { layout });
  const body = getSlideShapes(slide).find((shape) =>
    ['obj', 'body', null].includes(getShapePlaceholderType(shape)),
  );
  setShapeText(body, 'First\nSecond\nThird');
  setParagraphLevel(body, 1, 2);
  assert.deepEqual(promoteOutlineBody(pres, slide, body, { start: 6, end: 6 }), []);
  assert.deepEqual(getParagraphLevel(body, { start: 0, end: 18 }), [0, 1, 0]);
  assert.equal(getSlides(pres).length, 1);
});

test('promotes mixed root and nested paragraphs like Mac PowerPoint and preserves an empty final paragraph', async () => {
  const pres = createPresentation();
  const layout = getSlideLayouts(pres).find(
    (item) => getSlideLayoutName(item) === 'Title and Content',
  );
  const slide = addSlide(pres, { layout });
  const body = getSlideShapes(slide).find((shape) =>
    ['obj', 'body', null].includes(getShapePlaceholderType(shape)),
  );
  setShapeText(body, 'First\nSecond\nThird\n');
  setParagraphLevel(body, 2, 1);
  const added = promoteOutlineBody(pres, slide, body, { start: 6, end: 18 });
  assert.equal(added.length, 1);
  const loaded = await loadPresentation(await savePresentation(pres));
  const slides = getSlides(loaded);
  assert.equal(getSlideTitle(slides[1]), 'Second');
  const moved = getSlideShapes(slides[1]).find((shape) =>
    ['obj', 'body', null].includes(getShapePlaceholderType(shape)),
  );
  assert.equal(getShapeText(moved), 'Third\n');
  assert.equal(getParagraphLevel(moved, 0), 0);
  assert.equal(getParagraphLevel(moved, 1), 0);
  assert.equal(getShapeText(body), 'First');
});
