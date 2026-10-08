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
  isShapePlaceholder,
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
import { newSlideLayout } from '../src/core/new-slide.ts';
import {
  splitOutlineTitleRange,
  outlineShapes,
  promoteOutlineBody,
  outlineParagraphMove,
} from '../src/core/outline.ts';
import { OutlineSelectionModel } from '../src/core/outline-selection.ts';

globalThis.Node = { TEXT_NODE: 3 };
globalThis.HTMLElement = class {};
globalThis.HTMLBRElement = class extends HTMLElement {};

test('outline selection retains its anchor across fields and replaces as one transaction', () => {
  const roots = [0, 1, 2].map((order) => ({
    ownerDocument: { getSelection: () => null },
    childNodes: [],
    compareDocumentPosition(other) {
      return order < other.order ? 4 : 2;
    },
    order,
  }));
  const values = ['Title', 'Body😀', 'Next'];
  const applied = [];
  let transactions = 0;
  const model = new OutlineSelectionModel();
  const fields = roots.map((root, index) =>
    model.register({
      key: String(index),
      root,
      text: () => values[index],
      copy: (start, end) => ({ text: values[index].slice(start, end), formats: [] }),
      flush: () => [],
      apply: (edits) => applied.push({ index, edits }),
      transact: (_label, fn) => {
        transactions++;
        fn();
      },
      focus: () => {},
    }),
  );
  void fields;
  const registered = model.fields();
  model.update(registered[0], 4, 5);
  assert.equal(model.extend(registered[0], 1), true);
  model.update(registered[1], 0, values[1].length, true);
  assert.equal(model.extend(registered[1], 1), true);
  assert.deepEqual(model.copy(), { text: 'e\nBody😀\n', formats: [] });
  assert.equal(model.replace('X', [], 'Cut'), true);
  assert.equal(transactions, 1);
  const edits = applied.filter(({ edits }) => edits.length);
  assert.equal(edits.length, 3);
  assert.equal(edits[0].index, 2);
  assert.equal(edits[1].index, 1);
  assert.equal(edits[2].index, 0);
});

test('outline selection extends backward across fields', () => {
  const roots = [0, 1, 2].map((order) => ({
    ownerDocument: { getSelection: () => null },
    childNodes: [],
    order,
    compareDocumentPosition(other) {
      return order < other.order ? 4 : 2;
    },
  }));
  const values = ['A', 'B', 'C'];
  const model = new OutlineSelectionModel();
  roots.forEach((root, index) =>
    model.register({
      key: String(index),
      root,
      text: () => values[index],
      copy: (start, end) => ({ text: values[index].slice(start, end), formats: [] }),
      flush: () => [],
      apply: () => {},
      transact: (_label, fn) => fn(),
      focus: () => {},
    }),
  );
  const fields = model.fields();
  model.update(fields[2], 0, 0);
  assert.equal(model.extend(fields[2], -1), true);
  assert.equal(model.extend(fields[1], -1), true);
  assert.equal(model.extend(fields[1], -1), true);
  assert.deepEqual(model.copy(), { text: '\nB\n', formats: [] });
});

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
    await readFile(new URL('../../../test/fixtures/minimal/blank.pptx', import.meta.url)),
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
    await readFile(new URL('../../../test/fixtures/minimal/blank.pptx', import.meta.url)),
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

test('promotes mixed root and nested paragraphs like the reference desktop app (Mac) and preserves an empty final paragraph', async () => {
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
  const { demoteOutlineTitle } = await import('../src/core/outline.ts');
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
    await import('../src/core/outline.ts');
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
    const { moveOutlineTitle } = await import('../src/core/outline.ts');
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
    const { demoteOutlineTitle } = await import('../src/core/outline.ts');
    const pres = await loadPresentation(
      await readFile(new URL('../../../test/fixtures/minimal/blank.pptx', import.meta.url)),
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

test('title-to-body outline split promotes the suffix and preserves links through save/load', async () => {
  const pres = createPresentation();
  const layout = getSlideLayouts(pres).find(
    (item) => getSlideLayoutName(item) === 'Title and Content',
  );
  const slide = addSlide(pres, { layout });
  const title = getSlideShapes(slide).find((shape) => getShapePlaceholderType(shape) === 'title');
  const body = getSlideShapes(slide).find((shape) => getShapePlaceholderType(shape) !== 'title');
  setShapeText(title, 'Heading');
  setShapeText(body, 'Body');
  setShapeRunHyperlink(body, 0, 0, 'https://example.com/body');
  assert.equal(
    splitOutlineTitleRange(
      pres,
      slide,
      { id: getShapeId(title), offset: 2 },
      { id: getShapeId(body), offset: 2 },
    ),
    1,
  );
  const loaded = await loadPresentation(await savePresentation(pres));
  const slides = getSlides(loaded);
  assert.deepEqual(slides.map(getSlideTitle), ['He', 'dy']);
  const heading = getSlideShapes(slides[1]).find(
    (shape) => getShapePlaceholderType(shape) === 'title',
  );
  assert.equal(getShapeRunHyperlink(heading, 0, 0), 'https://example.com/body');
  assert.equal(
    getShapeText(
      getSlideShapes(slides[0]).find((shape) => getShapePlaceholderType(shape) !== 'title'),
    ),
    '',
  );
});

test('outline split keeps later paragraphs in the new body as in the reference desktop app (Mac)', async () => {
  const pres = createPresentation();
  const layout = getSlideLayouts(pres).find(
    (item) => getSlideLayoutName(item) === 'Title and Content',
  );
  const slide = addSlide(pres, { layout });
  const title = getSlideShapes(slide).find((shape) => getShapePlaceholderType(shape) === 'title');
  const body = getSlideShapes(slide).find((shape) => getShapePlaceholderType(shape) !== 'title');
  setShapeText(title, 'Outline title');
  setShapeText(body, 'Body\nFollowing');
  setParagraphLevel(body, 1, 2);
  setShapeRunHyperlink(body, 1, 0, 'https://example.com/following');
  splitOutlineTitleRange(
    pres,
    slide,
    { id: getShapeId(title), offset: 2 },
    { id: getShapeId(body), offset: 2 },
  );
  const slides = getSlides(await loadPresentation(await savePresentation(pres)));
  assert.deepEqual(slides.map(getSlideTitle), ['Ou', 'dy']);
  const nextBody = getSlideShapes(slides[1]).find(
    (shape) => getShapePlaceholderType(shape) !== 'title',
  );
  assert.equal(getShapeText(nextBody), 'Following');
  assert.equal(getParagraphLevel(nextBody, 0), 2);
  assert.equal(getShapeRunHyperlink(nextBody, 0, 0), 'https://example.com/following');
});

test('Enter across adjacent outline titles retains both slides and the remaining title runs', async () => {
  const pres = createPresentation();
  const layout = getSlideLayouts(pres).find(
    (item) => getSlideLayoutName(item) === 'Title and Content',
  );
  const first = addSlide(pres, { layout });
  const second = addSlide(pres, { layout });
  const [title, body] = getSlideShapes(first);
  const [nextTitle, nextBody] = getSlideShapes(second);
  setShapeText(title, 'Outline title');
  setShapeText(body, 'Body');
  setShapeText(nextTitle, 'Next');
  setShapeText(nextBody, 'Following');
  setShapeRunHyperlink(nextTitle, 0, 0, 'https://example.com/title');
  addSlideTextBox(first, { x: 0, y: 0, w: 914400, h: 914400, text: 'Ordinary text' });
  assert.equal(
    splitOutlineTitleRange(
      pres,
      first,
      { id: getShapeId(title), offset: 2 },
      { id: getShapeId(nextTitle), offset: 2, slide: second },
    ),
    1,
  );
  const slides = getSlides(await loadPresentation(await savePresentation(pres)));
  assert.deepEqual(slides.map(getSlideTitle), ['Ou', 'xt']);
  assert.deepEqual(getSlideShapes(slides[0]).map(getShapeText), ['Ou', '', 'Ordinary text']);
  assert.deepEqual(getSlideShapes(slides[1]).map(getShapeText), ['xt', 'Following']);
  assert.equal(
    getShapeRunHyperlink(getSlideShapes(slides[1])[0], 0, 0),
    'https://example.com/title',
  );
});

for (const bodyText of ['Body', 'Body\nFollowing']) {
  test(`outline deletion joins a body suffix to its title (${JSON.stringify(bodyText)})`, async () => {
    const { deleteOutlineTitleBodyRange } = await import('../src/core/outline.ts');
    const pres = createPresentation();
    const layout = getSlideLayouts(pres).find(
      (item) => getSlideLayoutName(item) === 'Title and Content',
    );
    const slide = addSlide(pres, { layout });
    const [title, body] = getSlideShapes(slide);
    setShapeText(title, 'Outline title');
    setShapeText(body, bodyText);
    setShapeRunHyperlink(body, 0, 0, 'https://example.com/body');
    if (bodyText.includes('\n')) {
      setParagraphLevel(body, 1, 2);
      setShapeRunHyperlink(body, 1, 0, 'https://example.com/following');
    }
    assert.equal(
      deleteOutlineTitleBodyRange(
        slide,
        { id: getShapeId(title), offset: 2 },
        { id: getShapeId(body), offset: 2 },
      ),
      true,
    );
    const slides = getSlides(await loadPresentation(await savePresentation(pres)));
    assert.equal(slides.length, 1);
    const [savedTitle, savedBody] = getSlideShapes(slides[0]);
    assert.equal(getShapeText(savedTitle), 'Oudy');
    assert.equal(getShapeRunHyperlink(savedTitle, 0, 1), 'https://example.com/body');
    assert.equal(getShapeText(savedBody), bodyText.includes('\n') ? 'Following' : '');
    if (bodyText.includes('\n')) {
      assert.equal(getParagraphLevel(savedBody, 0), 2);
      assert.equal(getShapeRunHyperlink(savedBody, 0, 0), 'https://example.com/following');
    }
  });
}

for (const bodyText of ['Body', 'Body\nFollowing']) {
  test(`outline deletion retains earlier title paragraphs (${JSON.stringify(bodyText)})`, async () => {
    const { deleteOutlineTitleBodyRange } = await import('../src/core/outline.ts');
    const pres = createPresentation();
    const layout = getSlideLayouts(pres).find(
      (item) => getSlideLayoutName(item) === 'Title and Content',
    );
    const slide = addSlide(pres, { layout });
    const [title, body] = getSlideShapes(slide);
    setShapeText(title, 'First\nSecond');
    setShapeText(body, bodyText);
    setShapeRunHyperlink(body, 0, 0, 'https://example.com/body');
    if (bodyText.includes('\n')) {
      setParagraphLevel(body, 1, 2);
      setShapeRunHyperlink(body, 1, 0, 'https://example.com/following');
    }
    assert.equal(
      deleteOutlineTitleBodyRange(
        slide,
        { id: getShapeId(title), offset: 8 },
        { id: getShapeId(body), offset: 2 },
      ),
      true,
    );
    const slides = getSlides(await loadPresentation(await savePresentation(pres)));
    assert.equal(slides.length, 1);
    const [savedTitle, savedBody] = getSlideShapes(slides[0]);
    assert.equal(getShapeText(savedTitle), 'First\nSedy');
    assert.equal(getShapeRunHyperlink(savedTitle, 1, 1), 'https://example.com/body');
    assert.equal(getShapeText(savedBody), bodyText.includes('\n') ? 'Following' : '');
    if (bodyText.includes('\n')) {
      assert.equal(getParagraphLevel(savedBody, 0), 2);
      assert.equal(getShapeRunHyperlink(savedBody, 0, 0), 'https://example.com/following');
    }
  });
}

for (const withBody of [false, true]) {
  test(`outline cross-slide title deletion joins titles and transfers body (${withBody})`, async () => {
    const { deleteOutlineSlideRange } = await import('../src/core/outline.ts');
    const pres = createPresentation();
    const layout = getSlideLayouts(pres).find(
      (item) => getSlideLayoutName(item) === 'Title and Content',
    );
    const first = addSlide(pres, { layout });
    const last = addSlide(pres, { layout });
    const [title, body] = getSlideShapes(first);
    const [next, following] = getSlideShapes(last);
    setShapeText(title, 'Outline title');
    if (withBody) setShapeText(body, 'Body');
    else removeShape(body);
    setShapeText(next, 'Next');
    setShapeRunHyperlink(next, 0, 0, 'https://example.com/title');
    setShapeText(following, 'Following');
    setShapeRunHyperlink(following, 0, 0, 'https://example.com/body');
    setParagraphLevel(following, 0, 2);
    addSlideTextBox(first, { text: 'Keep graphic', x: 0, y: 0, w: 914400, h: 914400 });
    assert.equal(
      deleteOutlineSlideRange(
        pres,
        { slide: first, id: getShapeId(title), offset: 2 },
        { slide: last, id: getShapeId(next), offset: 2 },
      ),
      true,
    );
    const saved = getSlides(await loadPresentation(await savePresentation(pres)));
    assert.equal(saved.length, 1);
    const shapes = getSlideShapes(saved[0]);
    const savedTitle = shapes.find((shape) => getShapePlaceholderType(shape) === 'title');
    const savedBody = shapes.find(
      (shape) =>
        isShapePlaceholder(shape) &&
        ['body', 'obj'].includes(getShapePlaceholderType(shape) ?? 'obj'),
    );
    assert.equal(getShapeText(savedTitle), 'Ouxt');
    assert.equal(getShapeRunHyperlink(savedTitle, 0, 1), 'https://example.com/title');
    assert.equal(getShapeText(savedBody), 'Following');
    assert.equal(getShapeRunHyperlink(savedBody, 0, 0), 'https://example.com/body');
    assert.equal(getParagraphLevel(savedBody, 0), 2);
    assert.ok(shapes.some((shape) => getShapeText(shape) === 'Keep graphic'));
  });
}

function contentDeck(slides) {
  const pres = createPresentation();
  const layout = getSlideLayouts(pres).find(
    (item) => getSlideLayoutName(item) === 'Title and Content',
  );
  const created = slides.map(([title, body, levels = []]) => {
    const slide = addSlide(pres, { layout });
    const [titleShape, bodyShape] = getSlideShapes(slide);
    setShapeText(titleShape, title);
    setShapeText(bodyShape, body);
    levels.forEach((level, index) => setParagraphLevel(bodyShape, index, level));
    return slide;
  });
  return { pres, slides: created };
}

function outlineTexts(pres) {
  return getSlides(pres).map((slide) =>
    outlineShapes(slide).map(({ id }) =>
      getShapeText(getSlideShapes(slide).find((shape) => getShapeId(shape) === id)),
    ),
  );
}

function shapesOf(slide) {
  const [title, body] = getSlideShapes(slide);
  return { title, body, titleId: getShapeId(title), bodyId: getShapeId(body) };
}

function levelsOf(shape) {
  return getParagraphLevel(shape, { start: 0, end: getShapeText(shape).length });
}

test('deleting from a body into a later title joins the title suffix and moves its body', async () => {
  const { deleteOutlineSlideRange } = await import('../src/core/outline.ts');
  const { pres, slides } = contentDeck([
    ['One', 'Alpha\nBeta'],
    ['Two', 'Middle'],
    ['Three', 'Gamma\nDelta', [1, 2]],
  ]);
  const first = shapesOf(slides[0]);
  const last = shapesOf(slides[2]);
  setShapeRunHyperlink(last.body, 1, 0, 'https://example.com/delta');
  assert.equal(
    deleteOutlineSlideRange(
      pres,
      { slide: slides[0], id: first.bodyId, offset: 8 },
      { slide: slides[2], id: last.titleId, offset: 2 },
    ),
    true,
  );
  const saved = await loadPresentation(await savePresentation(pres));
  assert.deepEqual(outlineTexts(saved), [['One', 'Alpha\nBeree\nGamma\nDelta']]);
  const body = getSlideShapes(getSlides(saved)[0])[1];
  assert.deepEqual(levelsOf(body), [0, 0, 1, 2]);
  assert.equal(getShapeRunHyperlink(body, 3, 0), 'https://example.com/delta');
});

test('deleting from a title into a later body keeps the join in the title', async () => {
  const { deleteOutlineSlideRange } = await import('../src/core/outline.ts');
  const { pres, slides } = contentDeck([
    ['One', 'Alpha'],
    ['Two', 'Gamma\nDelta', [0, 1]],
  ]);
  deleteOutlineSlideRange(
    pres,
    { slide: slides[0], id: shapesOf(slides[0]).titleId, offset: 1 },
    { slide: slides[1], id: shapesOf(slides[1]).bodyId, offset: 2 },
  );
  assert.deepEqual(outlineTexts(pres), [['Omma', 'Delta']]);
  assert.equal(getParagraphLevel(getSlideShapes(getSlides(pres)[0])[1], 0), 1);
});

test('deleting between bodies of different slides keeps the start paragraph level', async () => {
  const { deleteOutlineSlideRange } = await import('../src/core/outline.ts');
  const { pres, slides } = contentDeck([
    ['One', 'Alpha', [2]],
    ['Two', 'Gamma\nDelta'],
  ]);
  deleteOutlineSlideRange(
    pres,
    { slide: slides[0], id: shapesOf(slides[0]).bodyId, offset: 5 },
    { slide: slides[1], id: shapesOf(slides[1]).bodyId, offset: 0 },
  );
  assert.deepEqual(outlineTexts(pres), [['One', 'AlphaGamma\nDelta']]);
  assert.deepEqual(levelsOf(getSlideShapes(getSlides(pres)[0])[1]), [2, 0]);
});

test('deleteOutlineRange deletes within a field and joins a title to its body', async () => {
  const { deleteOutlineRange } = await import('../src/core/outline.ts');
  const { pres, slides } = contentDeck([['Heading', 'Alpha\nBeta']]);
  const { titleId, bodyId } = shapesOf(slides[0]);
  deleteOutlineRange(
    pres,
    { slide: slides[0], id: bodyId, offset: 1 },
    { slide: slides[0], id: bodyId, offset: 7 },
  );
  assert.deepEqual(outlineTexts(pres), [['Heading', 'Aeta']]);
  deleteOutlineRange(
    pres,
    { slide: slides[0], id: titleId, offset: 4 },
    { slide: slides[0], id: bodyId, offset: 1 },
  );
  assert.deepEqual(outlineTexts(pres), [['Headeta', '']]);
});

test('splitOutlineTitle moves the title suffix and the body to a new slide', async () => {
  const { splitOutlineTitle } = await import('../src/core/outline.ts');
  const { pres, slides } = contentDeck([['Heading', 'Alpha']]);
  assert.equal(
    splitOutlineTitle(pres, slides[0], shapesOf(slides[0]).title, { start: 4, end: 4 }),
    1,
  );
  assert.deepEqual(outlineTexts(pres), [
    ['Head', ''],
    ['ing', 'Alpha'],
  ]);
});

test('a bullet block spans the following deeper paragraphs', async () => {
  const { outlineParagraphBlock, outlineParagraphRange } = await import('../src/core/outline.ts');
  const { slides } = contentDeck([['One', 'A\nB\nC\nD', [0, 1, 2, 0]]]);
  const { body } = shapesOf(slides[0]);
  assert.deepEqual(outlineParagraphBlock(body, 0), { first: 0, last: 2 });
  assert.deepEqual(outlineParagraphBlock(body, 1), { first: 1, last: 2 });
  assert.deepEqual(outlineParagraphBlock(body, 3), { first: 3, last: 3 });
  assert.deepEqual(outlineParagraphRange(body, 1, 2), { start: 2, end: 5 });
});

test('moving paragraphs across slides keeps their XML and shifts levels', async () => {
  const { moveOutlineParagraphs } = await import('../src/core/outline.ts');
  const { pres, slides } = contentDeck([
    ['One', 'A\nB\nC', [0, 1, 0]],
    ['Two', 'X\nY'],
  ]);
  const first = shapesOf(slides[0]);
  const second = shapesOf(slides[1]);
  setShapeRunHyperlink(first.body, 1, 0, 'https://example.com/b');
  const moved = moveOutlineParagraphs(
    { slide: slides[0], id: first.bodyId, first: 0, last: 1 },
    { slide: slides[1], id: second.bodyId, index: 1 },
    1,
  );
  assert.equal(getShapeId(moved.shape), second.bodyId);
  assert.deepEqual([moved.first, moved.last], [1, 2]);
  const saved = await loadPresentation(await savePresentation(pres));
  assert.deepEqual(outlineTexts(saved), [
    ['One', 'C'],
    ['Two', 'X\nA\nB\nY'],
  ]);
  const body = getSlideShapes(getSlides(saved)[1])[1];
  assert.deepEqual(levelsOf(body), [0, 1, 2, 0]);
  assert.equal(getShapeRunHyperlink(body, 2, 0), 'https://example.com/b');
});

test('moving paragraphs within a body, into an empty body and to a slide without one', async () => {
  const { moveOutlineParagraphs } = await import('../src/core/outline.ts');
  const { pres, slides } = contentDeck([
    ['One', 'A\nB\nC'],
    ['Two', ''],
    ['Three', ''],
  ]);
  const first = shapesOf(slides[0]);
  moveOutlineParagraphs(
    { slide: slides[0], id: first.bodyId, first: 0, last: 0 },
    { slide: slides[0], id: first.bodyId, index: 3 },
  );
  assert.deepEqual(outlineTexts(pres)[0], ['One', 'B\nC\nA']);
  assert.equal(
    moveOutlineParagraphs(
      { slide: slides[0], id: first.bodyId, first: 1, last: 1 },
      { slide: slides[0], id: first.bodyId, index: 2 },
    ),
    null,
  );
  moveOutlineParagraphs(
    { slide: slides[0], id: first.bodyId, first: 2, last: 2 },
    { slide: slides[1], id: shapesOf(slides[1]).bodyId, index: 0 },
  );
  removeShape(shapesOf(slides[2]).body);
  moveOutlineParagraphs(
    { slide: slides[0], id: first.bodyId, first: 0, last: 1 },
    { slide: slides[2], id: null, index: 0 },
  );
  assert.deepEqual(outlineTexts(pres), [
    ['One', ''],
    ['Two', 'A'],
    ['Three', 'B\nC'],
  ]);
});
