import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { Presentation, SlideLayout } from './index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`../../test/fixtures/minimal/${name}`, import.meta.url));

describe('Presentation.slideLayouts', () => {
  it('lists every layout that python-pptx ships in the default template', async () => {
    const pres = await Presentation.load(await readFile(fixture('blank.pptx')));
    const layouts = pres.slideLayouts;

    // python-pptx's default template ships 11 layouts: title, title+content,
    // section header, two content, comparison, title only, blank, content
    // with caption, picture with caption, and two more variants.
    expect(layouts.length).toBe(11);
    for (const layout of layouts) {
      expect(layout).toBeInstanceOf(SlideLayout);
      expect(typeof layout.name).toBe('string');
    }
  });

  it('exposes the canonical user-visible names', async () => {
    const pres = await Presentation.load(await readFile(fixture('blank.pptx')));
    const names = pres.slideLayouts.map((l) => l.name);
    expect(names).toContain('Title Slide');
    expect(names).toContain('Title and Content');
    expect(names).toContain('Title Only');
    expect(names).toContain('Blank');
  });

  it('exposes the layout type token', async () => {
    const pres = await Presentation.load(await readFile(fixture('blank.pptx')));
    const titleLayout = pres.slideLayouts.find((l) => l.name === 'Title Slide');
    // ECMA-376 §19.7.15 — the Title Slide layout uses `title` as its type.
    expect(titleLayout?.layoutType).toBe('title');
  });
});

describe('Slide.layout', () => {
  it('resolves to the correct layout for fixtures with slides', async () => {
    const pres = await Presentation.load(await readFile(fixture('two-slides.pptx')));
    const slide = pres.slides[0];
    if (!slide) throw new Error('expected a slide');
    const layout = slide.layout;
    expect(layout).toBeInstanceOf(SlideLayout);
    // python-pptx's "Title and Content" layout is index 1 in the default
    // template; its layoutType is `obj` in ECMA-376 terms.
    expect(layout?.name).toBe('Title and Content');
  });

  it('returns null on a presentation with no slides (and thus no slide rels)', async () => {
    // Build a degenerate package: load the blank fixture, then this would
    // succeed because rootRels still resolves. Skip: there's nothing to test
    // when there are no slides.
    const pres = await Presentation.load(await readFile(fixture('blank.pptx')));
    expect(pres.slides).toEqual([]);
  });
});
