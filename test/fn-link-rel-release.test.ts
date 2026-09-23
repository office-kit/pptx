// Removing or replacing a link must release the relationship it owned.

import { expect, it } from 'vitest';
import {
  _internalPackageOf,
  addBlankSlide,
  addSlideTextBox,
  createPresentation,
  duplicateSlide,
  getSlidePartName,
  getSlides,
  inches,
  loadPresentation,
  removeSlide,
  savePresentation,
  setShapeClickAction,
  setShapeHyperlink,
  setShapeRunHyperlink,
  validatePresentation,
} from '../src/api/index.ts';
import { partName } from '../src/internal/opc/index.ts';
import { REL_TYPES } from '../src/internal/presentationml/index.ts';

const box = { x: inches(1), y: inches(1), w: inches(3), h: inches(0.5) } as const;

const linkRelTargets = (pres: ReturnType<typeof createPresentation>, slideIndex: number) => {
  const slide = getSlides(pres)[slideIndex]!;
  const rels = _internalPackageOf(pres).getRels(partName(getSlidePartName(slide)));
  return (rels?.items ?? [])
    .filter((rel) => rel.type === REL_TYPES.hyperlink || rel.type === REL_TYPES.slide)
    .map((rel) => rel.target);
};

it('releases the slide rel when a slide-jump click action is removed', () => {
  const pres = createPresentation();
  const target = addBlankSlide(pres);
  const source = addBlankSlide(pres);
  const shape = addSlideTextBox(source, { ...box, text: 'back' });
  setShapeClickAction(shape, { kind: 'slide', slide: target });
  expect(linkRelTargets(pres, 1)).toEqual(['../slides/slide1.xml']);

  setShapeClickAction(shape, null);

  expect(linkRelTargets(pres, 1)).toEqual([]);
  // A leftover slide rel makes the removed slide a live dependency of this one.
  removeSlide(pres, target);
  expect(() => duplicateSlide(pres, source)).not.toThrow();
});

it('releases the previous rel when a click action is replaced', () => {
  const pres = createPresentation();
  const target = addBlankSlide(pres);
  const source = addBlankSlide(pres);
  const shape = addSlideTextBox(source, { ...box, text: 'back' });
  setShapeClickAction(shape, { kind: 'slide', slide: target });
  setShapeClickAction(shape, { kind: 'url', url: 'https://example.com/' });

  expect(linkRelTargets(pres, 1)).toEqual(['https://example.com/']);
});

it('releases the hyperlink rel when shape and run links are cleared', async () => {
  const pres = createPresentation();
  const slide = addBlankSlide(pres);
  const shape = addSlideTextBox(slide, { ...box, text: 'docs' });
  setShapeHyperlink(shape, 'https://example.com/a');
  setShapeHyperlink(shape, null);
  expect(linkRelTargets(pres, 0)).toEqual([]);

  setShapeRunHyperlink(shape, 0, 0, 'https://example.com/b');
  setShapeRunHyperlink(shape, 0, 0, null);
  expect(linkRelTargets(pres, 0)).toEqual([]);

  const reloaded = await loadPresentation(await savePresentation(pres));
  expect(validatePresentation(reloaded).filter((issue) => issue.severity === 'error')).toEqual([]);
});

it('keeps a rel that another run still points at', () => {
  const pres = createPresentation();
  const slide = addBlankSlide(pres);
  // Two lines give two paragraphs, so the shape-level setter reuses one rel for two runs.
  const shape = addSlideTextBox(slide, { ...box, text: 'one\ntwo' });
  setShapeHyperlink(shape, 'https://example.com/shared');
  setShapeRunHyperlink(shape, 0, 0, null);

  expect(linkRelTargets(pres, 0)).toEqual(['https://example.com/shared']);
});
