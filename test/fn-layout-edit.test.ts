// Editing a layout, not just applying one — the theme-builder slice.
// A layout handle now carries its own document, so a change written
// through it lands in the package part and every slide on that layout
// picks it up.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { expectSchemaValid, isSchemaValidationAvailable } from './lib/expect-schema-valid.ts';
import {
  addSlide,
  clearSlideLayoutBackground,
  getShapeBoundsResolved,
  getSlideLayout,
  getSlideLayoutBackground,
  getSlideLayoutName,
  getSlideLayoutPartName,
  getSlideLayoutPlaceholders,
  getSlideLayouts,
  getSlides,
  getSlideShapes,
  inches,
  loadPresentation,
  type PresentationData,
  readPackagePart,
  savePresentation,
  setShapeBounds,
  setSlideLayoutBackground,
  setSlideLayoutName,
  setSlideLayoutPlaceholderBounds,
  type SlideLayoutData,
} from '../src/api/index.ts';

const skipIfNoXmllint = isSchemaValidationAvailable() ? it : it.skip;
const fixturePath = fileURLToPath(new URL('./fixtures/minimal/blank.pptx', import.meta.url));

const template = async (): Promise<PresentationData> =>
  loadPresentation(await readFile(fixturePath));

const layoutNamed = (pres: PresentationData, name: string): SlideLayoutData =>
  getSlideLayouts(pres).find((l) => getSlideLayoutName(l) === name)!;

const layoutXml = (pres: PresentationData, layout: SlideLayoutData): string =>
  new TextDecoder().decode(readPackagePart(pres, getSlideLayoutPartName(layout))!);

describe('slide layout editing', () => {
  it('renames a layout, and the name survives a save', async () => {
    const pres = await template();
    const layout = layoutNamed(pres, 'Title and Content');

    setSlideLayoutName(layout, '本文レイアウト');

    expect(getSlideLayoutName(layout)).toBe('本文レイアウト');
    const reloaded = await loadPresentation(await savePresentation(pres));
    expect(getSlideLayouts(reloaded).map(getSlideLayoutName)).toContain('本文レイアウト');
  });

  it('gives the layout a background of its own', async () => {
    const pres = await template();
    const layout = layoutNamed(pres, 'Title and Content');
    expect(getSlideLayoutBackground(layout)).toEqual({ kind: 'inherit' });

    setSlideLayoutBackground(layout, '#1F3864');

    expect(getSlideLayoutBackground(layout)).toEqual({ kind: 'solid', color: '#1F3864' });
    expect(layoutXml(pres, layout)).toContain('<p:bg>');
  });

  it('accepts a scheme color, like the slide-level setter', async () => {
    const pres = await template();
    const layout = layoutNamed(pres, 'Title and Content');

    setSlideLayoutBackground(layout, 'scheme:accent2');

    expect(getSlideLayoutBackground(layout)).toEqual({ kind: 'solid', color: 'scheme:accent2' });
  });

  it('gives the background back to the master', async () => {
    const pres = await template();
    const layout = layoutNamed(pres, 'Title and Content');
    setSlideLayoutBackground(layout, '#1F3864');

    clearSlideLayoutBackground(layout);

    expect(getSlideLayoutBackground(layout)).toEqual({ kind: 'inherit' });
    expect(layoutXml(pres, layout)).not.toContain('<p:bg>');
  });

  it('a reloaded deck reads the background back off the layout part', async () => {
    const pres = await template();
    setSlideLayoutBackground(layoutNamed(pres, 'Title and Content'), '#1F3864');

    const reloaded = await loadPresentation(await savePresentation(pres));

    expect(getSlideLayoutBackground(layoutNamed(reloaded, 'Title and Content'))).toEqual({
      kind: 'solid',
      color: '#1F3864',
    });
  });

  it('moves a placeholder slot, and a slide that inherits follows it', async () => {
    const pres = await template();
    const layout = layoutNamed(pres, 'Title and Content');
    const slide = addSlide(pres, { layout });
    const title = getSlideShapes(slide)[0]!;
    const moved = { x: inches(1), y: inches(0.25), w: inches(6), h: inches(1.5) };

    setSlideLayoutPlaceholderBounds(layout, 0, moved);

    expect(getSlideLayoutPlaceholders(layout)[0]!.bounds).toEqual(moved);
    // The slide's own placeholder carries no `<a:xfrm>`, so its resolved
    // bounds are the layout's — which is the point of editing the layout.
    expect(getShapeBoundsResolved(pres, title)).toEqual(moved);
  });

  it('a slide that was nudged keeps its own box', async () => {
    const pres = await template();
    const layout = layoutNamed(pres, 'Title and Content');
    const slide = addSlide(pres, { layout });
    const title = getSlideShapes(slide)[0]!;
    const own = { x: inches(2), y: inches(2), w: inches(3), h: inches(1) };
    setShapeBounds(title, own);

    setSlideLayoutPlaceholderBounds(layout, 0, {
      x: inches(1),
      y: inches(0.25),
      w: inches(6),
      h: inches(1.5),
    });

    expect(getShapeBoundsResolved(pres, title)).toEqual(own);
  });

  it('refuses a slot the layout does not have', async () => {
    const pres = await template();
    const layout = layoutNamed(pres, 'Title and Content');
    const count = getSlideLayoutPlaceholders(layout).length;

    expect(() =>
      setSlideLayoutPlaceholderBounds(layout, count, {
        x: inches(0),
        y: inches(0),
        w: inches(1),
        h: inches(1),
      }),
    ).toThrow(/no index/);
  });

  it('edits the layout the slide is bound to, reached through the slide', async () => {
    const pres = await template();
    const slide = addSlide(pres, { layout: layoutNamed(pres, 'Title and Content') });
    const layout = getSlideLayout(slide)!;

    setSlideLayoutName(layout, 'Reached via the slide');

    const reloaded = await loadPresentation(await savePresentation(pres));
    expect(getSlideLayoutName(getSlideLayout(getSlides(reloaded)[0]!)!)).toBe(
      'Reached via the slide',
    );
  });

  skipIfNoXmllint('the edited layout part validates', async () => {
    const pres = await template();
    const layout = layoutNamed(pres, 'Title and Content');
    setSlideLayoutName(layout, 'Edited');
    setSlideLayoutBackground(layout, '#1F3864');
    setSlideLayoutPlaceholderBounds(layout, 0, {
      x: inches(1),
      y: inches(0.25),
      w: inches(6),
      h: inches(1.5),
    });
    const saved = await loadPresentation(await savePresentation(pres));

    expectSchemaValid(layoutXml(saved, layoutNamed(saved, 'Edited')), 'pml');
  });
});
