// Level-3 first feature: add a new slide from a chosen layout.
//
// Scenario:
//   1. Load `blank.pptx` (no slides yet).
//   2. Pick the "Title and Content" layout.
//   3. Call `pres.addSlide({ layout })`.
//   4. Fill the title placeholder.
//   5. Save → reload → assert the deck now has one slide with our title.
//
// Also exercises:
//   - sldIdLst creation when presentation.xml didn't have one.
//   - sldId allocation in [256, 2³¹−1024].
//   - Multiple sequential adds.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { Presentation } from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('L3: addSlide from a layout', () => {
  it('adds a single slide bound to the chosen layout and persists', async () => {
    const pres = await Presentation.load(await readFile(fixture('blank.pptx')));
    expect(pres.slides.length).toBe(0);

    const layout = pres.slideLayouts.find((l) => l.name === 'Title and Content');
    if (!layout) throw new Error('expected Title and Content layout');

    const slide = pres.addSlide({ layout });
    expect(slide).toBeDefined();
    expect(pres.slides.length).toBe(1);

    // The new slide carries placeholder stubs from the layout.
    const title = slide.findPlaceholder('title');
    expect(title).not.toBeNull();
    title?.setText('Brand new slide');

    const reloaded = await Presentation.load(await pres.save());
    expect(reloaded.slides.length).toBe(1);
    expect(reloaded.slides[0]?.findPlaceholder('title')?.text).toBe('Brand new slide');
    expect(reloaded.slides[0]?.layout?.name).toBe('Title and Content');
  });

  it('preserves existing slides and appends in order', async () => {
    const pres = await Presentation.load(await readFile(fixture('two-slides.pptx')));
    const baseline = pres.slides.length;
    const layout = pres.slideLayouts.find((l) => l.name === 'Title Only');
    if (!layout) throw new Error('expected Title Only layout');
    pres.addSlide({ layout });
    pres.addSlide({ layout });

    expect(pres.slides.length).toBe(baseline + 2);
    const reloaded = await Presentation.load(await pres.save());
    expect(reloaded.slides.length).toBe(baseline + 2);
  });

  it('builds blank.pptx into a usable single-slide deck end-to-end', async () => {
    // Round-trip the "create from blank → add slide → fill text → save"
    // pipeline through a second load so we know the output is structurally
    // sound.
    const pres = await Presentation.load(await readFile(fixture('blank.pptx')));
    const layout = pres.slideLayouts.find((l) => l.name === 'Title Slide');
    if (!layout) throw new Error('expected Title Slide layout');
    const slide = pres.addSlide({ layout });
    slide.findPlaceholder('ctrTitle')?.setText('pptx-kit');
    slide.findPlaceholder('subTitle')?.setText('an OOXML library for TypeScript');

    const bytes = await pres.save();
    const reloaded = await Presentation.load(bytes);
    const reSlide = reloaded.slides[0];
    expect(reSlide?.findPlaceholder('ctrTitle')?.text).toBe('pptx-kit');
    expect(reSlide?.findPlaceholder('subTitle')?.text).toBe('an OOXML library for TypeScript');
  });
});
