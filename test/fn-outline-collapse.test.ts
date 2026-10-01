import { INTERNAL_PACKAGE } from '../src/api/_internal-symbols.ts';
import { partName } from '../src/internal/opc/index.ts';
import { decode } from '../src/api/fn/_helpers.ts';
import { expectSchemaValid, isSchemaValidationAvailable } from './lib/expect-schema-valid.ts';
import { expect, it } from 'vitest';
import {
  createPresentation,
  addBlankSlide,
  getSlides,
  savePresentation,
  loadPresentation,
  getCollapsedOutlineSlides,
  setSlideOutlineCollapsed,
  removeSlide,
  setGridSpacing,
  getGridSpacing,
} from '../src/api/index.ts';

it('round-trips collapse per slide without changing grid preferences or slide text', async () => {
  const pres = createPresentation();
  const first = addBlankSlide(pres),
    second = addBlankSlide(pres);
  setGridSpacing(pres, { x: 72008, y: 180000 });
  expect(getCollapsedOutlineSlides(pres).includes(first)).toBe(false);
  setSlideOutlineCollapsed(first, true);
  expect(getCollapsedOutlineSlides(pres).includes(first)).toBe(true);
  if (isSchemaValidationAvailable())
    expectSchemaValid(
      decode(pres[INTERNAL_PACKAGE].getPart(partName('/ppt/viewProps.xml'))!.data),
      'pml',
    );
  expect(getCollapsedOutlineSlides(pres).includes(second)).toBe(false);
  const restored = await loadPresentation(await savePresentation(pres));
  expect(getCollapsedOutlineSlides(restored).includes(getSlides(restored)[0]!)).toBe(true);
  expect(getCollapsedOutlineSlides(restored).includes(getSlides(restored)[1]!)).toBe(false);
  expect(getGridSpacing(restored)).toEqual({ x: 72008, y: 180000 });
  setSlideOutlineCollapsed(getSlides(restored)[0]!, false);
  expect(
    getCollapsedOutlineSlides(await loadPresentation(await savePresentation(restored))),
  ).toEqual([]);
});

it('removes the outline entry when its slide is deleted', () => {
  const pres = createPresentation();
  const slide = addBlankSlide(pres);
  setSlideOutlineCollapsed(slide, true);
  removeSlide(pres, slide);
  const xml = decode(pres[INTERNAL_PACKAGE].getPart(partName('/ppt/viewProps.xml'))!.data);
  expect(xml).not.toContain('<p:sld ');
  addBlankSlide(pres);
  expect(getCollapsedOutlineSlides(pres)).toEqual([]);
});
