// Drawing a layout or master on its own, as the editor's Slide Master view
// does: background and decoration, without any slide content.

import { readFile } from 'node:fs/promises';
import { expect, it } from 'vitest';
import {
  getSlideLayout,
  getSlideLayoutShapes,
  getSlideMasterShapes,
  getSlides,
  getShapeName,
  loadPresentation,
  setSlideLayoutBackgroundGraphicsHidden,
} from '../src/api/index.ts';
import { renderSlideLayoutToSvg, renderSlideToSvg } from '../packages/preview/src/index.ts';

const fixture = () =>
  readFile(new URL('./fixtures/minimal/layout-decoration.pptx', import.meta.url));

it('draws a layout’s and its master’s decoration without the slide', async () => {
  const pres = await loadPresentation(await fixture());
  const slide = getSlides(pres)[0]!;
  const layout = getSlideLayout(slide)!;
  const slideSvg = renderSlideToSvg(pres, slide);
  const layoutSvg = renderSlideLayoutToSvg(pres, layout);

  expect(slideSvg).toContain('pptx-kit sample 01');
  expect(layoutSvg).not.toContain('pptx-kit sample 01');
  // Every decoration the slide shows comes from the layout or master.
  expect(layoutSvg).toContain('TEMPLATE');
  expect(layoutSvg).toContain('<image');
  expect(layoutSvg).toContain('#2E75B6');
  expect(
    getSlideLayoutShapes(pres, layout).length + getSlideMasterShapes(pres, layout).length,
  ).toBeGreaterThan(0);
});

it('draws the master alone, and a layout without the master’s graphics when it hides them', async () => {
  const pres = await loadPresentation(await fixture());
  const layout = getSlideLayout(getSlides(pres)[0]!)!;
  // The fixture's layout carries the logo picture and a bar; the master a footer bar.
  expect(getSlideLayoutShapes(pres, layout).map(getShapeName)).toEqual(['Logo', 'Template Bar']);
  expect(getSlideMasterShapes(pres, layout).map(getShapeName)).toEqual(['Master Footer Bar']);
  const masterSvg = renderSlideLayoutToSvg(pres, layout, { master: true });
  const withMaster = renderSlideLayoutToSvg(pres, layout);

  setSlideLayoutBackgroundGraphicsHidden(layout, true);
  const hidden = renderSlideLayoutToSvg(pres, layout);
  // Hiding drops exactly what the master contributes; the master view still has it.
  expect(hidden.length).toBeLessThan(withMaster.length);
  expect(hidden).toContain('<image');
  expect(masterSvg).not.toContain('<image');
  expect(renderSlideLayoutToSvg(pres, layout, { master: true })).toBe(masterSvg);
});
