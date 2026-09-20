import { describe, expect, it } from 'vitest';
import * as api from '../src/api/index.ts';
import { Fill, Presentation, Slide, compile } from '../packages/dsl/src/index.ts';
import { auditTextLayout } from '../packages/preview/src/index.ts';
import { buildFontkitMeasurer } from '../packages/preview/src/node.ts';

const LONG_TITLE =
  'Monthly churn rose to 2.3 percent in the second half, up 0.4 points from 1.9 percent in the first half, driven by price-sensitive monthly subscribers on the standard plan';

const titleSlide = (autoFit?: api.TextAutoFit) =>
  compile(
    Presentation({
      children: Slide({
        layout: { type: 'title' },
        children: Fill({
          target: { placeholder: { type: 'ctrTitle' } },
          children: LONG_TITLE,
          ...(autoFit ? { autoFit } : {}),
        }),
      }),
    }),
  );

const titleOf = (pres: api.PresentationData) =>
  api
    .getSlideShapes(api.getSlides(pres)[0]!)
    .find((shape) => api.getShapePlaceholderType(shape) === 'ctrTitle')!;

describe('Fill autoFit', () => {
  const measureText = buildFontkitMeasurer();

  it('shrinks an overflowing placeholder with autoFit="normal"', async () => {
    const pres = await titleSlide('normal');
    expect(api.getShapeTextAutoFit(titleOf(pres))).toBe('normal');
    expect(auditTextLayout(pres, { measureText })).toEqual([]);
  });

  it('keeps the template bodyPr when autoFit is omitted', async () => {
    const pres = await titleSlide();
    expect(api.getShapeTextAutoFit(titleOf(pres))).toBeNull();
    expect(auditTextLayout(pres, { measureText }).map((issue) => issue.kind)).toEqual([
      'overflow-y',
    ]);
  });

  it('leaves a non-default bodyPr of an existing deck untouched when autoFit is omitted', async () => {
    const original = api.createPresentation();
    const box = api.addSlideTextBox(api.addBlankSlide(original), {
      name: 'Body',
      x: api.inches(1),
      y: api.inches(1),
      w: api.inches(4),
      h: api.inches(1),
      text: 'before',
    });
    api.setShapeTextAutoFit(box, 'shape');
    const source = await api.savePresentation(original);
    const bodyPr = (pres: api.PresentationData) =>
      /<a:bodyPr[\s\S]*?<\/a:bodyPr>/.exec(
        api.getShapeXmlString(api.getSlideShapes(api.getSlides(pres)[0]!)[0]!),
      )?.[0];

    const edited = await compile(
      Presentation({
        source,
        mode: 'edit',
        children: Slide({
          target: { index: 0 },
          children: Fill({ target: { name: 'Body' }, children: 'after' }),
        }),
      }),
    );
    expect(bodyPr(edited)).toContain('<a:spAutoFit/>');
    expect(bodyPr(edited)).toBe(bodyPr(await api.loadPresentation(source)));
    expect(api.getShapeText(api.getSlideShapes(api.getSlides(edited)[0]!)[0]!)).toBe('after');
  });

  it('survives a save and reload', async () => {
    const reloaded = await api.loadPresentation(
      await api.savePresentation(await titleSlide('normal')),
    );
    expect(api.getShapeTextAutoFit(titleOf(reloaded))).toBe('normal');
    expect(api.validatePresentation(reloaded).filter((i) => i.severity === 'error')).toEqual([]);
  });
});
