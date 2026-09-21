import { unzipSync } from 'fflate';
import { readFile } from 'node:fs/promises';
import { expect, it } from 'vitest';
import {
  addSlideTable,
  addSlideTextBox,
  emu,
  inches,
  getSlides,
  getSlideShapes,
  getShapeBounds,
  getShapeBoundsResolved,
  getShapeBodyPrEffective,
  getShapeImageBytes,
  getShapeImageCrop,
  getShapeTextMargins,
  getSlideXmlString,
  getTableColumnWidths,
  getTableRowHeights,
  groupShapes,
  getGroupChildren,
  loadPresentation,
  savePresentation,
  setSlideSize,
  setShapeTextFormat,
  setShapeTextMargins,
  setShapeStroke,
  setShapeImageCrop,
  setShapeText,
  getShapeText,
  SLIDE_SIZE_4_3,
} from '../src/api/index.ts';

const load = async () =>
  loadPresentation(
    await readFile(new URL('./fixtures/minimal/one-image-slide.pptx', import.meta.url)),
  );

it('fits pictures, text and tables proportionally and keeps cached handles editable after saving', async () => {
  const pres = await load();
  setSlideSize(pres, SLIDE_SIZE_4_3);
  const slide = getSlides(pres)[0]!;
  let picture = getSlideShapes(slide)[1]!;
  const original = getShapeBounds(picture)!;
  const image = getShapeImageBytes(picture);
  setShapeImageCrop(picture, { left: 0.2, right: 0.1 });
  let text = addSlideTextBox(slide, {
    x: inches(1),
    y: inches(1),
    w: inches(4),
    h: inches(1),
    text: 'Fit content',
  });
  setShapeTextFormat(text, { size: 24 });
  setShapeTextMargins(text, { left: 100000, top: 50000, right: 200000, bottom: 25000 });
  setShapeStroke(text, { color: '123456', widthEmu: 50800 });
  const table = addSlideTable(slide, {
    x: inches(1),
    y: inches(3),
    w: inches(4),
    h: inches(2),
    rows: [
      ['a', 'b'],
      ['c', 'd'],
    ],
  });
  picture = getSlideShapes(slide)[1]!;
  text = getSlideShapes(slide)[2]!;
  const widths = getTableColumnWidths(table);
  const heights = getTableRowHeights(table);
  setSlideSize(pres, { width: inches(20), height: inches(20) }, { content: 'fit' });
  const dy = inches(2.5);
  expect(getShapeBounds(picture)).toEqual({
    x: original.x * 2,
    y: original.y * 2 + dy,
    w: original.w * 2,
    h: original.h * 2,
  });
  expect(getShapeBounds(text)).toEqual({
    x: inches(2),
    y: inches(4.5),
    w: inches(8),
    h: inches(2),
  });
  expect(getShapeTextMargins(text)).toEqual({
    left: 200000,
    top: 100000,
    right: 400000,
    bottom: 50000,
  });
  expect(getShapeBounds(table)).toEqual({
    x: inches(2),
    y: inches(8.5),
    w: inches(8),
    h: inches(4),
  });
  expect(getTableColumnWidths(table)).toEqual(widths.map((v) => v * 2));
  expect(getTableRowHeights(table)).toEqual(heights.map((v) => v * 2));
  expect(getSlideXmlString(slide)).toContain('sz="4800"');
  expect(getSlideXmlString(slide)).toContain('w="101600"');
  setShapeText(text, 'Still editable');
  const restored = await loadPresentation(await savePresentation(pres));
  const shapes = getSlideShapes(getSlides(restored)[0]!);
  expect(getShapeBounds(shapes[1]!)).toEqual(getShapeBounds(picture));
  expect(getShapeImageBytes(shapes[1]!)).toEqual(image);
  expect(getShapeImageCrop(shapes[1]!)).toEqual({ left: 0.2, right: 0.1, top: 0, bottom: 0 });
  expect(shapes.some((shape) => getShapeText(shape) === 'Still editable')).toBe(true);
});

it('centers a group once while scaling its child coordinate space', async () => {
  const pres = await load();
  setSlideSize(pres, SLIDE_SIZE_4_3);
  const slide = getSlides(pres)[0]!;
  const make = (x: number) =>
    addSlideTextBox(slide, {
      x: inches(x),
      y: inches(1),
      w: inches(1),
      h: inches(1),
      text: 'Group',
    });
  const group = groupShapes([make(1), make(3)]);
  const children = getGroupChildren(group).map((shape) => getShapeBounds(shape)!);
  setSlideSize(pres, { width: inches(20), height: inches(20) }, { content: 'fit' });
  expect(getShapeBounds(group)).toEqual({
    x: inches(2),
    y: inches(4.5),
    w: inches(6),
    h: inches(2),
  });
  expect(getGroupChildren(group).map((shape) => getShapeBounds(shape))).toEqual(
    children.map((b) => ({ x: b.x * 2, y: b.y * 2, w: b.w * 2, h: b.h * 2 })),
  );
});

it('rejects a fit that would produce invalid font sizes without writing any parts', async () => {
  const pres = await load();
  setSlideSize(pres, SLIDE_SIZE_4_3);
  const text = addSlideTextBox(getSlides(pres)[0]!, {
    x: emu(0),
    y: emu(0),
    w: inches(1),
    h: inches(1),
    text: 'Large',
  });
  setShapeTextFormat(text, { size: 4000 });
  const before = unzipSync(await savePresentation(pres));
  expect(() =>
    setSlideSize(pres, { width: inches(20), height: inches(15) }, { content: 'fit' }),
  ).toThrow(RangeError);
  expect(unzipSync(await savePresentation(pres))).toEqual(before);
});

it('fits inherited placeholder bounds and margins through save and reload', async () => {
  const pres = await load();
  setSlideSize(pres, SLIDE_SIZE_4_3);
  const title = getSlideShapes(getSlides(pres)[0]!)[0]!;
  expect(getShapeBounds(title)).toBeNull();
  const bounds = getShapeBoundsResolved(pres, title)!;
  expect(bounds).not.toBeNull();
  const margins = getShapeBodyPrEffective(pres, title).margins;
  setSlideSize(pres, { width: inches(20), height: inches(20) }, { content: 'fit' });
  const expectedBounds = {
    x: bounds.x * 2,
    y: bounds.y * 2 + inches(2.5),
    w: bounds.w * 2,
    h: bounds.h * 2,
  };
  const expectedMargins = {
    left: (margins.left ?? 91440) * 2,
    right: (margins.right ?? 91440) * 2,
    top: (margins.top ?? 45720) * 2,
    bottom: (margins.bottom ?? 45720) * 2,
  };
  expect(getShapeBoundsResolved(pres, title)).toEqual(expectedBounds);
  expect(getShapeBodyPrEffective(pres, title).margins).toEqual(expectedMargins);
  const restored = await loadPresentation(await savePresentation(pres));
  const restoredTitle = getSlideShapes(getSlides(restored)[0]!)[0]!;
  expect(getShapeBoundsResolved(restored, restoredTitle)).toEqual(expectedBounds);
  expect(getShapeBodyPrEffective(restored, restoredTitle).margins).toEqual(expectedMargins);
});
