import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
// Placeholder-type equivalence for inheritance. A `ctrTitle` must inherit from
// a `title` placeholder (and `subTitle` from `body`) when walking the layout /
// master cascade — otherwise a centered title on a title-slide layout drops the
// master title placeholder's bodyPr (anchor), lstStyle and geometry. Regression
// guard for the matcher shared by the bounds / bodyPr / rPr / pPr resolvers.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { matchPlaceholderShape, placeholderTypeCandidates } from '../src/api/fn/shape-read-base.ts';
import {
  addTitleSlide,
  getShapeBodyPrEffective,
  getShapePlaceholderType,
  getSlideShapes,
  loadPresentation,
  savePresentation,
  getSlides,
  setShapeTextDirection,
  setShapeTextAnchor,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('placeholderTypeCandidates', () => {
  it('treats ctrTitle/title and subTitle/body as equivalent', () => {
    expect(placeholderTypeCandidates('ctrTitle')).toEqual(['ctrTitle', 'title']);
    expect(placeholderTypeCandidates('title')).toEqual(['title', 'ctrTitle']);
    expect(placeholderTypeCandidates('subTitle')).toEqual(['subTitle', 'body']);
    expect(placeholderTypeCandidates('body')).toEqual(['body']);
    expect(placeholderTypeCandidates('ftr')).toEqual(['ftr']);
    expect(placeholderTypeCandidates(null)).toEqual([]);
  });
});

describe('matchPlaceholderShape', () => {
  const shape = (placeholderType: string | null, placeholderIdx: number | null = null) => ({
    placeholderType,
    placeholderIdx,
  });

  it('matches by idx before type', () => {
    const shapes = [shape('title', 1), shape('body', 2)];
    expect(matchPlaceholderShape(shapes, 2, 'title')?.placeholderType).toBe('body');
  });

  it('falls back from ctrTitle to a title placeholder', () => {
    const shapes = [shape('title'), shape('body')];
    expect(matchPlaceholderShape(shapes, null, 'ctrTitle')?.placeholderType).toBe('title');
  });

  it('prefers the exact type over the equivalent one', () => {
    const shapes = [shape('title'), shape('ctrTitle')];
    expect(matchPlaceholderShape(shapes, null, 'ctrTitle')?.placeholderType).toBe('ctrTitle');
  });

  it('falls back from subTitle to body', () => {
    const shapes = [shape('body')];
    expect(matchPlaceholderShape(shapes, null, 'subTitle')?.placeholderType).toBe('body');
  });

  it('returns undefined when nothing matches', () => {
    expect(matchPlaceholderShape([shape('ftr')], null, 'ctrTitle')).toBeUndefined();
  });
});

describe('ctrTitle inherits the master title bodyPr', () => {
  it('resolves the centered-title vertical anchor from the master title placeholder', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const slide = addTitleSlide(pres, 'Hello');
    const ctr = getSlideShapes(slide).find((s) => getShapePlaceholderType(s) === 'ctrTitle');
    expect(ctr).toBeDefined();
    // The master title placeholder carries anchor="ctr"; a ctrTitle must inherit
    // it (it returned null before the placeholder-type equivalence fix).
    expect(getShapeBodyPrEffective(pres, ctr!).anchor).toBe('center');
  });
});

describe('horizontal direction overrides vertical inheritance', () => {
  const verticalMaster = async (horizontalLayout = false) => {
    const parts = unzipSync(await readFile(fixture('blank.pptx')));
    const master = 'ppt/slideMasters/slideMaster1.xml';
    parts[master] = strToU8(strFromU8(parts[master]!).replaceAll('vert="horz"', 'vert="eaVert"'));
    if (horizontalLayout) {
      const layout = 'ppt/slideLayouts/slideLayout1.xml';
      parts[layout] = strToU8(
        strFromU8(parts[layout]!).replaceAll('<a:bodyPr/>', '<a:bodyPr vert="horz"/>'),
      );
    }
    return loadPresentation(zipSync(parts));
  };
  it('stops at an explicit horizontal layout instead of inheriting vertical master text', async () => {
    const pres = await verticalMaster(true);
    const title = getSlideShapes(addTitleSlide(pres, 'Horizontal'))[0]!;
    expect(getShapeBodyPrEffective(pres, title).vert).toBeNull();
  });
  it('persists a horizontal override and restores inheritance only when cleared', async () => {
    const pres = await verticalMaster();
    const title = getSlideShapes(addTitleSlide(pres, 'Horizontal'))[0]!;
    expect(getShapeBodyPrEffective(pres, title).vert).toBe('eaVert');
    setShapeTextDirection(title, 'horz');
    expect(getShapeBodyPrEffective(pres, title).vert).toBeNull();
    const restored = await loadPresentation(await savePresentation(pres));
    const restoredTitle = getSlideShapes(getSlides(restored).at(-1)!)[0]!;
    expect(getShapeBodyPrEffective(restored, restoredTitle).vert).toBeNull();
    setShapeTextDirection(restoredTitle, null);
    expect(getShapeBodyPrEffective(restored, restoredTitle).vert).toBe('eaVert');
  });
});

describe('centered anchor inheritance', () => {
  it('overrides inherited centering explicitly and restores it when cleared', async () => {
    const parts = unzipSync(await readFile(fixture('blank.pptx')));
    const master = 'ppt/slideMasters/slideMaster1.xml';
    const xml = strFromU8(parts[master]!);
    parts[master] = strToU8(xml.replaceAll('<a:bodyPr ', '<a:bodyPr anchorCtr="true" '));
    const pres = await loadPresentation(zipSync(parts));
    const title = getSlideShapes(addTitleSlide(pres, 'Centered'))[0]!;
    expect(getShapeBodyPrEffective(pres, title).anchorCentered).toBe(true);
    setShapeTextAnchor(title, 'top', { centered: false });
    const restored = await loadPresentation(await savePresentation(pres));
    const copy = getSlideShapes(getSlides(restored).at(-1)!)[0]!;
    expect(getShapeBodyPrEffective(restored, copy).anchorCentered).toBe(false);
    setShapeTextAnchor(copy, 'top', { centered: null });
    expect(getShapeBodyPrEffective(restored, copy).anchorCentered).toBe(true);
  });
});
