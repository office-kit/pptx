import { expect, it } from 'vitest';
import * as pptx from '../src/api/index.ts';
import { expectSchemaValid, isSchemaValidationAvailable } from './lib/expect-schema-valid.ts';

it('reset and layout changes preserve content, IDs and valid placeholder XML', async () => {
  const p = pptx.createPresentation();
  const layouts = pptx.getSlideLayouts(p);
  const content = layouts.find((l) => pptx.getSlideLayoutType(l) === 'obj')!;
  const blank = layouts.find((l) => pptx.getSlideLayoutType(l) === 'blank')!;
  const slide = pptx.addSlide(p, { layout: content });
  const title = pptx.getSlideShapes(slide)[0]!;
  const id = pptx.getShapeId(title);
  pptx.setShapeText(title, 'Keep me');
  const box = pptx.getShapeBoundsResolved(p, title);
  pptx.applySlideLayout(p, slide, blank);
  expect(pptx.getShapeBoundsResolved(p, pptx.getSlideShapes(slide)[0]!)).toEqual(box);
  expect(pptx.getShapeText(pptx.getSlideShapes(slide)[0]!)).toBe('Keep me');
  pptx.applySlideLayout(p, slide, content);
  pptx.resetSlideLayout(p, slide);
  expect(pptx.getSlideShapes(slide)).toHaveLength(2);
  expect(pptx.getShapeId(pptx.getSlideShapes(slide)[0]!)).toBe(id);
  expect(pptx.getShapeBoundsResolved(p, pptx.getSlideShapes(slide)[0]!)).toEqual(box);
  if (isSchemaValidationAvailable()) expectSchemaValid(pptx.getSlideXmlString(slide), 'pml');
  const saved = await pptx.loadPresentation(await pptx.savePresentation(p));
  expect(pptx.getSlideXmlString(pptx.getSlides(saved)[0]!)).toBe(pptx.getSlideXmlString(slide));
});
