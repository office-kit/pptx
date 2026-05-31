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
