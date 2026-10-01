import { expect, it } from 'vitest';
import {
  createPresentation,
  addSlideTable,
  addSlideImage,
  addSlideLine,
  addSlideMedia,
  _internalPackageOf,
  addBlankSlide,
  addSlideTextBox,
  inches,
  groupShapes,
  getGroupChildren,
  getSlideShapes,
  getSlides,
  isShapeLocked,
  isShapeAspectRatioLocked,
  setShapeLocked,
  setShapeAspectRatioLocked,
  savePresentation,
  loadPresentation,
  getShapeText,
  setShapeText,
} from '../src/api/index.ts';
import { buildPng } from './lib/build-png.ts';
import { expectSchemaValid, isSchemaValidationAvailable } from './lib/expect-schema-valid.ts';
import { SHAPE_ELEMENT } from '../src/api/_internal-symbols.ts';
import { NS, parseXml, serializeFragment } from '../src/internal/xml/index.ts';

function fixture() {
  const pres = createPresentation();
  const slide = addBlankSlide(pres);
  const shapes = ['A', 'B', 'C'].map((text) =>
    addSlideTextBox(slide, { x: inches(1), y: inches(1), w: inches(1), h: inches(1), text }),
  );
  return { pres, slide, shapes };
}

it('round-trips native geometry locks without preventing text editing', async () => {
  const { pres, shapes } = fixture();
  expect(isShapeLocked(shapes[0]!)).toBe(false);
  setShapeLocked(shapes, true);
  setShapeText(shapes[0]!, 'Edited');
  const loaded = await loadPresentation(await savePresentation(pres));
  const saved = getSlideShapes(getSlides(loaded)[0]!);
  expect(saved.every(isShapeLocked)).toBe(true);
  expect(getShapeText(saved[0]!)).toBe('Edited');
  const xml = serializeFragment(saved[0]![SHAPE_ELEMENT]);
  for (const key of [
    'noGrp',
    'noRot',
    'noMove',
    'noResize',
    'noEditPoints',
    'noAdjustHandles',
    'noChangeArrowheads',
    'noChangeShapeType',
  ])
    expect(xml).toContain(`${key}="1"`);
  expect(xml).not.toContain('noTextEdit=');
  setShapeLocked(saved, false);
  const reopened = await loadPresentation(await savePresentation(loaded));
  expect(getSlideShapes(getSlides(reopened)[0]!).some(isShapeLocked)).toBe(false);
});

it('locks groups with schema-appropriate flags and leaves child locks independent', () => {
  const { shapes } = fixture();
  const group = groupShapes(shapes.slice(0, 2));
  setShapeLocked(group, true);
  const xml = serializeFragment(group[SHAPE_ELEMENT]);
  expect(xml).toContain('<a:grpSpLocks');
  expect(xml).toContain('noUngrp="1"');
  expect(xml).not.toContain('noEditPoints=');
  expect(getGroupChildren(group).some(isShapeLocked)).toBe(false);
  setShapeLocked(getGroupChildren(group), true);
  setShapeLocked(group, false);
  expect(getGroupChildren(group).every(isShapeLocked)).toBe(true);
});

it('preserves aspect ratio, text locks, extensions and unrelated nonvisual properties', () => {
  const { shapes } = fixture();
  const shape = shapes[0]!;
  const nv = shape[SHAPE_ELEMENT].children.find(
    (e) => e.kind === 'element' && e.name.localName === 'nvSpPr',
  );
  if (!nv || nv.kind !== 'element') throw new Error('missing nvSpPr');
  const props = nv.children.find((e) => e.kind === 'element' && e.name.localName === 'cNvSpPr');
  if (!props || props.kind !== 'element') throw new Error('missing cNvSpPr');
  props.children.unshift(
    parseXml(
      `<a:spLocks xmlns:a="${NS.dml}" noChangeAspect="1" noTextEdit="true"><a:extLst><a:ext uri="keep"/></a:extLst></a:spLocks>`,
    ).root,
  );
  setShapeLocked(shape, true);
  setShapeLocked(shape, false);
  const xml = serializeFragment(shape[SHAPE_ELEMENT]);
  expect(xml).toContain('noChangeAspect="1"');
  expect(xml).toContain('noTextEdit="true"');
  expect(xml).toContain('uri="keep"');
  expect(xml).toContain('txBox="1"');
  expect(xml).not.toContain('noMove=');
});

it.skipIf(!isSchemaValidationAvailable())(
  'locks every drawable kind with schema-valid nonvisual properties',
  async () => {
    const { pres, slide, shapes } = fixture();
    const bounds = { x: inches(1), y: inches(2), w: inches(2), h: inches(1) };
    const picture = addSlideImage(slide, buildPng(1, 1, [255, 0, 0]), bounds);
    const table = addSlideTable(slide, { ...bounds, rows: [['A']] });
    const line = addSlideLine(slide, {
      from: { x: inches(1), y: inches(1) },
      to: { x: inches(2), y: inches(2) },
    });
    const group = groupShapes(shapes.slice(0, 2));
    setShapeLocked([picture, table, line, group, ...shapes], true);
    const loaded = await loadPresentation(await savePresentation(pres));
    expect(getSlideShapes(getSlides(loaded)[0]!).every(isShapeLocked)).toBe(true);
    for (const part of _internalPackageOf(loaded).parts) {
      if (/^\/ppt\/slides\/slide\d+\.xml$/.test(part.name))
        expectSchemaValid(new TextDecoder().decode(part.data), 'pml');
    }
  },
);

it('validates a whole batch before modifying any member', () => {
  const { shapes } = fixture();
  shapes[1]![SHAPE_ELEMENT].name = { ...shapes[1]![SHAPE_ELEMENT].name, localName: 'unsupported' };
  expect(() => setShapeLocked(shapes, true)).toThrow(/nonvisual/);
  expect(isShapeLocked(shapes[0]!)).toBe(false);
});

it('reads missing and explicit aspect-ratio lock values', () => {
  const { slide } = fixture();
  const values = ['missing', 'false', '0', '1', 'true'];
  const shapes = values.map((value, index) => {
    const shape = addSlideTextBox(slide, {
      x: inches(index + 1),
      y: inches(1),
      w: inches(1),
      h: inches(1),
      text: value,
    });
    if (value !== 'missing') {
      const nv = shape[SHAPE_ELEMENT].children.find(
        (e) => e.kind === 'element' && e.name.localName === 'nvSpPr',
      );
      if (!nv || nv.kind !== 'element') throw new Error('missing nvSpPr');
      const props = nv.children.find((e) => e.kind === 'element' && e.name.localName === 'cNvSpPr');
      if (!props || props.kind !== 'element') throw new Error('missing cNvSpPr');
      props.children.unshift(
        parseXml(`<a:spLocks xmlns:a="${NS.dml}" noChangeAspect="${value}"/>`).root,
      );
    }
    return shape;
  });
  expect(shapes.map(isShapeAspectRatioLocked)).toEqual([false, false, false, true, true]);
});

it('toggles aspect-ratio locks for picture, video, group and graphic-frame shapes', async () => {
  const pres = createPresentation();
  const slide = addBlankSlide(pres);
  const bounds = { x: inches(1), y: inches(2), w: inches(2), h: inches(1) };
  const text = addSlideTextBox(slide, { ...bounds, text: 'text' });
  const groupText = addSlideTextBox(slide, { ...bounds, text: 'group text' });
  const groupText2 = addSlideTextBox(slide, { ...bounds, text: 'group text 2' });
  const picture = addSlideImage(slide, buildPng(1, 1, [255, 0, 0]), bounds);
  const video = addSlideMedia(slide, {
    kind: 'video',
    data: new Uint8Array([0, 0, 0, 0x18, 0x66, 0x74, 0x79, 0x70, 0x6d, 0x70, 0x34, 0x32]),
    ...bounds,
  });
  const table = addSlideTable(slide, { ...bounds, rows: [['A']] });
  const group = groupShapes([groupText, groupText2]);
  const shapes = [text, picture, video, table, group];

  setShapeAspectRatioLocked(shapes, false);
  expect(shapes.every((shape) => !isShapeAspectRatioLocked(shape))).toBe(true);
  setShapeAspectRatioLocked(shapes, true);
  expect(shapes.every(isShapeAspectRatioLocked)).toBe(true);

  const loaded = await loadPresentation(await savePresentation(pres));
  const loadedShapes = getSlideShapes(getSlides(loaded)[0]!);
  expect(loadedShapes.filter(isShapeAspectRatioLocked)).toHaveLength(5);
  const loadedGroup = loadedShapes.find((shape) => shape[SHAPE_ELEMENT].name.localName === 'grpSp');
  expect(loadedGroup && getGroupChildren(loadedGroup).every(isShapeAspectRatioLocked)).toBe(false);
  setShapeAspectRatioLocked(loadedShapes, false);
  expect(loadedShapes.every((shape) => !isShapeAspectRatioLocked(shape))).toBe(true);
});

it('prevalidates an aspect-ratio lock batch before changing any member', () => {
  const { shapes } = fixture();
  shapes[1]![SHAPE_ELEMENT].name = { ...shapes[1]![SHAPE_ELEMENT].name, localName: 'unsupported' };
  expect(() => setShapeAspectRatioLocked(shapes, true)).toThrow(/nonvisual/);
  expect(isShapeAspectRatioLocked(shapes[0]!)).toBe(false);
});

it('round-trips aspect-ratio locks and preserves explicit false and unknown XML', async () => {
  const pres = createPresentation();
  const slide = addBlankSlide(pres);
  const shape = addSlideTextBox(slide, {
    x: inches(1),
    y: inches(1),
    w: inches(2),
    h: inches(1),
    text: 'A',
  });
  const nv = shape[SHAPE_ELEMENT].children.find(
    (e) => e.kind === 'element' && e.name.localName === 'nvSpPr',
  );
  if (!nv || nv.kind !== 'element') throw new Error('missing nvSpPr');
  const props = nv.children.find((e) => e.kind === 'element' && e.name.localName === 'cNvSpPr');
  if (!props || props.kind !== 'element') throw new Error('missing cNvSpPr');
  props.children.unshift(
    parseXml(
      `<a:spLocks xmlns:a="${NS.dml}" noChangeAspect="false"><a:extLst><a:ext uri="keep"/></a:extLst></a:spLocks>`,
    ).root,
  );
  expect(isShapeAspectRatioLocked(shape)).toBe(false);
  expect(serializeFragment(shape[SHAPE_ELEMENT])).toContain('noChangeAspect="false"');
  expect(serializeFragment(shape[SHAPE_ELEMENT])).toContain('uri="keep"');
  setShapeAspectRatioLocked(shape, true);
  expect(isShapeAspectRatioLocked(shape)).toBe(true);
  const loaded = await loadPresentation(await savePresentation(pres));
  const roundTripped = getSlideShapes(getSlides(loaded)[0]!)[0]!;
  expect(isShapeAspectRatioLocked(roundTripped)).toBe(true);
  setShapeAspectRatioLocked(roundTripped, false);
  const xml = serializeFragment(roundTripped[SHAPE_ELEMENT]);
  expect(isShapeAspectRatioLocked(roundTripped)).toBe(false);
  expect(xml).not.toContain('noChangeAspect=');
  expect(xml).toContain('uri="keep"');
});
