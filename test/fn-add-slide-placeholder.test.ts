// `addSlidePlaceholder` restores one named slot from the layout — the single
// slot an "insert slide number" (or date, or footer) command needs, without
// also bringing back every other placeholder the author deleted.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  addSlidePlaceholder,
  createPresentation,
  findSlideLayout,
  findSlidePlaceholder,
  getShapeBoundsResolved,
  getShapeId,
  getShapePlaceholderType,
  getShapeText,
  getShapeXmlString,
  getSlideLayouts,
  getSlideLayoutName,
  getSlides,
  getSlideShapes,
  isShapePlaceholder,
  loadPresentation,
  removeShape,
  savePresentation,
  setShapeText,
} from '../src/api/index.ts';

const fixturePath = fileURLToPath(new URL('./fixtures/minimal/blank.pptx', import.meta.url));

// A real template: its layouts reserve `dt`, `ftr` and `sldNum` slots that no
// slide starts with, which is the case this function exists for.
const templateSlide = async (layoutName = 'Title and Content') => {
  const pres = await loadPresentation(await readFile(fixturePath));
  const layout = getSlideLayouts(pres).find((l) => getSlideLayoutName(l) === layoutName)!;
  return { pres, slide: addSlide(pres, { layout }) };
};

describe('addSlidePlaceholder', () => {
  it('adds only the requested slot, empty and inheriting its geometry', async () => {
    const { pres, slide } = await templateSlide();
    // `addSlide` copies every slot the layout defines, so the three footer
    // placeholders go first — the state a deck authored in PowerPoint without
    // footers arrives in.
    for (const type of ['sldNum', 'ftr', 'dt'] as const)
      removeShape(findSlidePlaceholder(slide, type)!);
    const before = getSlideShapes(slide).length;

    const added = addSlidePlaceholder(slide, 'sldNum');

    expect(added).not.toBeNull();
    expect(getShapePlaceholderType(added!)).toBe('sldNum');
    expect(getShapeText(added!)).toBe('');
    expect(isShapePlaceholder(added!)).toBe(true);
    expect(getSlideShapes(slide).length).toBe(before + 1);
    // The layout reserves `dt` and `ftr` too; neither may come along.
    expect(findSlidePlaceholder(slide, 'ftr')).toBeNull();
    expect(findSlidePlaceholder(slide, 'dt')).toBeNull();
    // Geometry comes from the cascade, not from a copied `<a:xfrm>`.
    expect(getShapeXmlString(added!)).not.toContain('xfrm');
    expect(getShapeBoundsResolved(pres, added!)).not.toBeNull();
  });

  it('returns the existing slot instead of adding a second one', async () => {
    const { slide } = await templateSlide();
    const first = addSlidePlaceholder(slide, 'sldNum')!;
    const count = getSlideShapes(slide).length;

    const again = addSlidePlaceholder(slide, 'sldNum');

    expect(again && getShapeId(again)).toBe(getShapeId(first));
    expect(getSlideShapes(slide).length).toBe(count);
  });

  it('restores a slot the author deleted, with its content gone', async () => {
    const { slide } = await templateSlide();
    const title = findSlidePlaceholder(slide, 'title')!;
    setShapeText(title, 'Quarterly review');
    removeShape(title);

    const restored = addSlidePlaceholder(slide, 'title');

    expect(restored).not.toBeNull();
    expect(getShapeText(restored!)).toBe('');
  });

  it('reports null when the layout reserves no such slot', () => {
    // `createPresentation`'s own layouts carry title and body slots only.
    const pres = createPresentation();
    const slide = addSlide(pres, { layout: findSlideLayout(pres, 'Title and Content')! });

    expect(addSlidePlaceholder(slide, 'sldNum')).toBeNull();
    expect(addSlidePlaceholder(slide, 'ftr')).toBeNull();
  });

  it('leaves the other shapes byte-identical and survives the round trip', async () => {
    const { pres, slide } = await templateSlide();
    const title = findSlidePlaceholder(slide, 'title')!;
    setShapeText(title, '日本語のタイトル');
    const untouched = getSlideShapes(slide).map(getShapeXmlString);

    addSlidePlaceholder(slide, 'sldNum');

    expect(getSlideShapes(slide).slice(0, untouched.length).map(getShapeXmlString)).toEqual(
      untouched,
    );
    const ids = getSlideShapes(slide).map(getShapeId);
    expect(new Set(ids).size).toBe(ids.length);

    const reloaded = await loadPresentation(await savePresentation(pres));
    const saved = findSlidePlaceholder(getSlides(reloaded)[0]!, 'sldNum');
    expect(saved).not.toBeNull();
  });
});
