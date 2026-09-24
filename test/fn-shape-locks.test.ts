import { expect, it } from 'vitest';
import {
  createPresentation,
  addSlideTable,
  addSlideImage,
  addSlideLine,
  _internalPackageOf,
  addBlankSlide,
  addSlideTextBox,
  inches,
  groupShapes,
  getGroupChildren,
  getSlideShapes,
  getSlides,
  isShapeLocked,
  setShapeLocked,
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
