import { readFile } from 'node:fs/promises';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { expect, it } from 'vitest';
import {
  getSlides,
  getSlideLayout,
  getSlideLayouts,
  duplicateSlide,
  applySlideBackgroundToAll,
  isSlideLayoutBackgroundGraphicsHidden,
  getSlideXmlString,
  isSlideBackgroundGraphicsHidden,
  loadPresentation,
  savePresentation,
  setSlideBackground,
  setSlideBackgroundGraphicsHidden,
} from '../src/api/index.ts';
import { renderSlideToSvg } from '../packages/preview/src/index.ts';

const fixture = () =>
  readFile(new URL('./fixtures/minimal/layout-decoration.pptx', import.meta.url));

it('hides inherited decoration while preserving slide content, fill, and template parts', async () => {
  const pres = await loadPresentation(await fixture());
  const slide = getSlides(pres)[0]!;
  setSlideBackground(slide, '#ABCDEF');
  const originalXml = getSlideXmlString(slide);
  const originalParts = unzipSync(await savePresentation(pres));
  const before = renderSlideToSvg(pres, slide);
  expect(before).toContain('TEMPLATE');
  expect(before).toContain('<image');
  expect(before).toContain('#2E75B6');
  expect(isSlideBackgroundGraphicsHidden(slide)).toBe(false);

  setSlideBackgroundGraphicsHidden(slide, true);
  const hidden = renderSlideToSvg(pres, slide);
  expect(hidden).not.toContain('TEMPLATE');
  expect(hidden).not.toContain('<image');
  expect(hidden).not.toContain('#2E75B6');
  expect(hidden).toContain('pptx-kit sample 01');
  expect(hidden).toContain('#ABCDEF');
  expect(getSlideXmlString(slide)).toContain('showMasterSp="0"');

  const saved = await savePresentation(pres);
  const parts = unzipSync(saved);
  for (const name of Object.keys(originalParts).filter((name) =>
    /^ppt\/(slideMasters|slideLayouts|media)\//.test(name),
  )) {
    expect(parts[name]).toEqual(originalParts[name]);
  }
  const reloaded = await loadPresentation(saved);
  const target = getSlides(reloaded)[0]!;
  expect(isSlideBackgroundGraphicsHidden(target)).toBe(true);
  expect(renderSlideToSvg(reloaded, target)).toBe(hidden);
  setSlideBackgroundGraphicsHidden(target, false);
  expect(getSlideXmlString(target)).toBe(originalXml);
  expect(renderSlideToSvg(reloaded, target)).toBe(before);
});

it.each(['0', 'false', '1', 'true'])(
  'reads imported showMasterSp="%s" and preserves unrelated root attributes',
  async (value) => {
    const parts = unzipSync(await fixture());
    const name = 'ppt/slides/slide1.xml';
    parts[name] = strToU8(
      strFromU8(parts[name]!).replace('<p:sld ', `<p:sld show="0" showMasterSp="${value}" `),
    );
    const pres = await loadPresentation(zipSync(parts));
    const slide = getSlides(pres)[0]!;
    expect(isSlideBackgroundGraphicsHidden(slide)).toBe(value === '0' || value === 'false');
    setSlideBackgroundGraphicsHidden(slide, true);
    setSlideBackgroundGraphicsHidden(slide, true);
    expect(getSlideXmlString(slide).match(/showMasterSp=/g)).toHaveLength(1);
    setSlideBackgroundGraphicsHidden(slide, false);
    expect(getSlideXmlString(slide)).toContain('show="0"');
    expect(getSlideXmlString(slide)).not.toContain('showMasterSp');
  },
);

it('applies the graphics setting to slides and layouts, with independent slide re-display', async () => {
  const pres = await loadPresentation(await fixture());
  const slide = getSlides(pres)[0]!;
  const other = duplicateSlide(pres, slide);
  setSlideBackgroundGraphicsHidden(slide, true);
  applySlideBackgroundToAll(pres, slide);
  expect(getSlides(pres).every(isSlideBackgroundGraphicsHidden)).toBe(true);
  expect(getSlideLayouts(pres).every(isSlideLayoutBackgroundGraphicsHidden)).toBe(true);
  setSlideBackgroundGraphicsHidden(other, false);
  const svg = renderSlideToSvg(pres, other);
  expect(svg).toContain('TEMPLATE');
  expect(svg).toContain('<image');
  expect(svg).not.toContain('#2E75B6');
  const reloaded = await loadPresentation(await savePresentation(pres));
  const target = getSlides(reloaded)[1]!;
  expect(isSlideBackgroundGraphicsHidden(target)).toBe(false);
  expect(isSlideLayoutBackgroundGraphicsHidden(getSlideLayout(target)!)).toBe(true);
  expect(renderSlideToSvg(reloaded, target)).toBe(svg);
  applySlideBackgroundToAll(reloaded, target);
  expect(getSlides(reloaded).some(isSlideBackgroundGraphicsHidden)).toBe(false);
  expect(getSlideLayouts(reloaded).some(isSlideLayoutBackgroundGraphicsHidden)).toBe(false);
  expect(renderSlideToSvg(reloaded, target)).toContain('#2E75B6');
});
