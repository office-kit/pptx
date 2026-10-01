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
import {
  outlineShapes,
  promoteOutlineBody,
  outlineParagraphMove,
} from '../src/lib/editor/core/outline.ts';

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

test('demotes a title and body into the previous slide while preserving paragraph levels and links', async () => {
  const { demoteOutlineTitle } = await import('../src/lib/editor/core/outline.ts');
  const pres = createPresentation();
  const layout = getSlideLayouts(pres).find(
    (item) => getSlideLayoutName(item) === 'Title and Content',
  );
  const first = addSlide(pres, { layout });
  const second = addSlide(pres, { layout });
  const [firstTitle, firstBody] = getSlideShapes(first);
  const [secondTitle, secondBody] = getSlideShapes(second);
  setShapeText(firstTitle, 'First title');
  setShapeText(firstBody, 'Existing body');
  setShapeText(secondTitle, 'Second title');
  setShapeText(secondBody, 'Child\nNested');
  setParagraphLevel(secondBody, 1, 2);
  setShapeRunHyperlink(secondTitle, 0, 0, 'https://example.com/heading');
  assert.equal(demoteOutlineTitle(pres, first), null);
  assert.equal(demoteOutlineTitle(pres, second), firstBody);
  const loaded = await loadPresentation(await savePresentation(pres));
  const slides = getSlides(loaded);
  assert.equal(slides.length, 1);
  const [title, body] = getSlideShapes(slides[0]);
  assert.equal(getShapeText(title), 'First title');
  assert.equal(getShapeText(body), 'Existing body\nSecond title\nChild\nNested');
  assert.equal(getParagraphLevel(body, 3), 2);
  assert.equal(getShapeRunHyperlink(body, 1, 0), 'https://example.com/heading');
});

test('rejects demotion with additional slide objects before changing either slide', async () => {
  const { demoteOutlineTitle } = await import('../src/lib/editor/core/outline.ts');
  const pres = createPresentation();
  const layout = getSlideLayouts(pres).find(
    (item) => getSlideLayoutName(item) === 'Title and Content',
  );
  const first = addSlide(pres, { layout });
  const second = addSlide(pres, { layout });
  const body = getSlideShapes(first)[1];
  setShapeText(body, 'Kept');
  addSlideTextBox(second, { x: 0, y: 0, w: 914400, h: 914400, text: 'Additional object' });
  assert.throws(() => demoteOutlineTitle(pres, second), /additional objects/);
  assert.equal(getShapeText(body), 'Kept');
  assert.equal(getSlides(pres).length, 2);
});

test('outline paragraph movement uses whole selected paragraphs and leaves nested followers in place', () => {
  const pres = createPresentation();
  const slide = addSlide(pres, { layout: getSlideLayouts(pres)[0] });
  const body = addSlideTextBox(slide, {
    x: 0,
    y: 0,
    w: 914400,
    h: 914400,
    text: 'First\nSecond\nThird\n',
  });
  setParagraphLevel(body, 2, 1);
  const up = outlineParagraphMove(body, { start: 7, end: 9 }, -1);
  assert.deepEqual(up, {
    ranges: [
      { start: 6, end: 12 },
      { start: 0, end: 5 },
      { start: 13, end: 18 },
      { start: 19, end: 19 },
    ],
    selection: { start: 0, end: 6 },
  });
  assert.equal(outlineParagraphMove(body, { start: 0, end: 6 }, -1), null);
  assert.equal(outlineParagraphMove(body, { start: 19, end: 19 }, 1), null);
  const down = outlineParagraphMove(body, { start: 0, end: 13 }, 1);
  assert.deepEqual(down.selection, { start: 6, end: 18 });
  assert.deepEqual(down.ranges, [
    { start: 13, end: 18 },
    { start: 0, end: 5 },
    { start: 6, end: 12 },
    { start: 19, end: 19 },
  ]);
});
