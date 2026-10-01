import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import test from 'node:test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import {
  addSlide,
  removeShape,
  getSlideLayout,
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
  getSlideMasterPartName,
} from '@office-kit/pptx';
import { newSlideLayout } from '../src/lib/editor/core/new-slide.ts';
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

test('generic New Slide keeps content layout and advances title slide to content layout', () => {
  const pres = createPresentation();
  const layouts = getSlideLayouts(pres);
  const content = layouts.find((item) => getSlideLayoutName(item) === 'Title and Content');
  const title = layouts.find((item) => getSlideLayoutName(item) === 'Title Slide');
  assert.ok(content);
  assert.ok(title);
  const contentSlide = addSlide(pres, { layout: content });
  const titleSlide = addSlide(pres, { layout: title });
  assert.equal(getSlideLayoutName(newSlideLayout(pres, contentSlide)), 'Title and Content');
  assert.equal(getSlideLayoutName(newSlideLayout(pres, titleSlide)), 'Title and Content');
});

test('generic New Slide does not cross masters when the content layout is unused', async () => {
  const zip = unzipSync(
    await readFile(new URL('../../test/fixtures/minimal/blank.pptx', import.meta.url)),
  );
  zip['ppt/slideMasters/slideMaster2.xml'] = zip['ppt/slideMasters/slideMaster1.xml'];
  zip['ppt/slideMasters/_rels/slideMaster2.xml.rels'] =
    zip['ppt/slideMasters/_rels/slideMaster1.xml.rels'];
  zip['[Content_Types].xml'] = strToU8(
    strFromU8(zip['[Content_Types].xml']).replace(
      '</Types>',
      '<Override PartName="/ppt/slideMasters/slideMaster2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/></Types>',
    ),
  );
  zip['ppt/presentation.xml'] = strToU8(
    strFromU8(zip['ppt/presentation.xml']).replace(
      '</p:sldMasterIdLst>',
      '<p:sldMasterId id="2147483649" r:id="rIdOutlineMaster"/></p:sldMasterIdLst>',
    ),
  );
  zip['ppt/_rels/presentation.xml.rels'] = strToU8(
    strFromU8(zip['ppt/_rels/presentation.xml.rels']).replace(
      '</Relationships>',
      '<Relationship Id="rIdOutlineMaster" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster2.xml"/></Relationships>',
    ),
  );
  const titleLayoutRel = 'ppt/slideLayouts/_rels/slideLayout1.xml.rels';
  zip[titleLayoutRel] = strToU8(
    strFromU8(zip[titleLayoutRel]).replace('slideMaster1.xml', 'slideMaster2.xml'),
  );
  const pres = await loadPresentation(zipSync(zip));
  const title = getSlideLayouts(pres).find((item) => getSlideLayoutName(item) === 'Title Slide');
  assert.ok(title);
  const slide = addSlide(pres, { layout: title });
  const content = getSlideLayouts(pres).find(
    (item) => getSlideLayoutName(item) === 'Title and Content',
  );
  assert.ok(content);
  assert.notEqual(getSlideMasterPartName(slide), getSlideMasterPartName(content));
  assert.equal(getSlideLayoutName(newSlideLayout(pres, slide)), 'Title Slide');
});

test('generic New Slide can use an unused object layout from the current master', async () => {
  const zip = unzipSync(
    await readFile(new URL('../../test/fixtures/minimal/blank.pptx', import.meta.url)),
  );
  zip['ppt/slideMasters/slideMaster2.xml'] = zip['ppt/slideMasters/slideMaster1.xml'];
  zip['ppt/slideMasters/_rels/slideMaster2.xml.rels'] =
    zip['ppt/slideMasters/_rels/slideMaster1.xml.rels'];
  zip['[Content_Types].xml'] = strToU8(
    strFromU8(zip['[Content_Types].xml']).replace(
      '</Types>',
      '<Override PartName="/ppt/slideMasters/slideMaster2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/></Types>',
    ),
  );
  for (const layout of ['slideLayout1.xml', 'slideLayout2.xml']) {
    const rels = `ppt/slideLayouts/_rels/${layout}.rels`;
    zip[rels] = strToU8(strFromU8(zip[rels]).replace('slideMaster1.xml', 'slideMaster2.xml'));
  }
  zip['ppt/presentation.xml'] = strToU8(
    strFromU8(zip['ppt/presentation.xml']).replace(
      '</p:sldMasterIdLst>',
      '<p:sldMasterId id="2147483649" r:id="rIdOutlineMaster"/></p:sldMasterIdLst>',
    ),
  );
  zip['ppt/_rels/presentation.xml.rels'] = strToU8(
    strFromU8(zip['ppt/_rels/presentation.xml.rels']).replace(
      '</Relationships>',
      '<Relationship Id="rIdOutlineMaster" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster2.xml"/></Relationships>',
    ),
  );
  const pres = await loadPresentation(zipSync(zip));
  const title = getSlideLayouts(pres).find((item) => getSlideLayoutName(item) === 'Title Slide');
  const content = getSlideLayouts(pres).find(
    (item) => getSlideLayoutName(item) === 'Title and Content',
  );
  assert.ok(title);
  assert.ok(content);
  const slide = addSlide(pres, { layout: title });
  assert.equal(getSlideMasterPartName(slide), getSlideMasterPartName(content));
  assert.equal(getSlideLayoutName(newSlideLayout(pres, slide)), 'Title and Content');
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

test('confirmed demotion removes additional objects while retaining outline text', async () => {
  const { demoteOutlineTitle, outlineDemotionNeedsConfirmation } =
    await import('../src/lib/editor/core/outline.ts');
  const pres = createPresentation();
  const layout = getSlideLayouts(pres).find(
    (item) => getSlideLayoutName(item) === 'Title and Content',
  );
  const first = addSlide(pres, { layout });
  const second = addSlide(pres, { layout });
  const body = getSlideShapes(first)[1];
  setShapeText(body, 'Kept');
  addSlideTextBox(second, { x: 0, y: 0, w: 914400, h: 914400, text: 'Additional object' });
  setShapeText(getSlideShapes(second)[0], 'Merged title');
  assert.equal(outlineDemotionNeedsConfirmation(first), false);
  assert.equal(outlineDemotionNeedsConfirmation(second), true);
  demoteOutlineTitle(pres, second);
  const loaded = await loadPresentation(await savePresentation(pres));
  assert.equal(getSlides(loaded).length, 1);
  const shapes = getSlideShapes(getSlides(loaded)[0]);
  assert.equal(getShapeText(shapes[1]), 'Kept\nMerged title');
  assert.equal(
    shapes.some((shape) => getShapeText(shape).includes('Additional object')),
    false,
  );
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

for (const direction of [-1, 1]) {
  test(`title movement shifts the body boundary (${direction}) and preserves paragraph links`, async () => {
    const { moveOutlineTitle } = await import('../src/lib/editor/core/outline.ts');
    const pres = createPresentation();
    const layout = getSlideLayouts(pres).find(
      (item) => getSlideLayoutName(item) === 'Title and Content',
    );
    const first = addSlide(pres, { layout });
    const second = addSlide(pres, { layout });
    const bodies = [first, second].map((slide) =>
      getSlideShapes(slide).find((shape) =>
        ['obj', 'body', null].includes(getShapePlaceholderType(shape)),
      ),
    );
    setShapeText(bodies[0], 'A\nB');
    setShapeText(bodies[1], 'C\nD');
    const source = bodies[direction === -1 ? 0 : 1];
    const paragraph = direction === -1 ? 1 : 0;
    setParagraphLevel(source, paragraph, 2);
    setShapeRunHyperlink(source, paragraph, 0, 'https://example.com/moved');
    assert.equal(moveOutlineTitle(pres, first, direction), false);
    assert.equal(moveOutlineTitle(pres, second, direction), true);
    assert.deepEqual(
      bodies.map(getShapeText),
      direction === -1 ? ['A', 'B\nC\nD'] : ['A\nB\nC', 'D'],
    );
    const loaded = await loadPresentation(await savePresentation(pres));
    const savedBodies = getSlides(loaded).map((slide) =>
      getSlideShapes(slide).find((shape) =>
        ['obj', 'body', null].includes(getShapePlaceholderType(shape)),
      ),
    );
    assert.deepEqual(savedBodies.map(getShapeText), bodies.map(getShapeText));
    const target = savedBodies[direction === -1 ? 1 : 0];
    const movedIndex = direction === -1 ? 0 : 2;
    assert.equal(getParagraphLevel(target, movedIndex), 2);
    assert.equal(getShapeRunHyperlink(target, movedIndex, 0), 'https://example.com/moved');
  });
}

for (const layoutName of ['Title and Content', 'Title Only']) {
  test(`demotion restores a missing body on ${layoutName} without changing layout`, async () => {
    const { demoteOutlineTitle } = await import('../src/lib/editor/core/outline.ts');
    const pres = await loadPresentation(
      await readFile(new URL('../../test/fixtures/minimal/blank.pptx', import.meta.url)),
    );
    const layout = getSlideLayouts(pres).find((item) => getSlideLayoutName(item) === layoutName);
    assert.ok(layout);
    const first = addSlide(pres, { layout });
    for (const slot of outlineShapes(first).filter((item) => !item.title)) {
      removeShape(getSlideShapes(first).find((shape) => getShapeId(shape) === slot.id));
    }
    const content = getSlideLayouts(pres).find(
      (item) => getSlideLayoutName(item) === 'Title and Content',
    );
    const second = addSlide(pres, { layout: content });
    setShapeText(getSlideShapes(second)[0], 'Second title');
    setShapeText(getSlideShapes(second)[1], 'Child');
    demoteOutlineTitle(pres, second);
    const loaded = await loadPresentation(await savePresentation(pres));
    const slides = getSlides(loaded);
    assert.equal(slides.length, 1);
    assert.equal(getSlideLayoutName(getSlideLayout(slides[0])), layoutName);
    const bodySlot = outlineShapes(slides[0]).find((item) => !item.title);
    assert.ok(bodySlot);
    const body = getSlideShapes(slides[0]).find((shape) => getShapeId(shape) === bodySlot.id);
    assert.equal(getShapeText(body), 'Second title\nChild');
  });
}
