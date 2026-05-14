// Level-3: moveSlide + addTextBox.
//
// moveSlide: reorder via `<p:sldIdLst>` child reordering. No part name
// or rId change; the slide identity is preserved.
//
// addTextBox: free-form `<p:sp>` with `txBox="1"`, geometry in EMU, and
// a single run of text. Appended to the slide's shape tree.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { Presentation, inches, pt } from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('L3: moveSlide', () => {
  it('reorders two slides and persists the new order', async () => {
    const pres = await Presentation.load(await readFile(fixture('two-slides.pptx')));
    const slide1Text = pres.slides[0]?.findPlaceholder('title')?.text;
    const slide2Text = pres.slides[1]?.findPlaceholder('title')?.text;
    expect(slide1Text).toBe('Slide 1');
    expect(slide2Text).toBe('Slide 2');

    const slide1 = pres.slides[0];
    if (!slide1) throw new Error('expected first slide');
    pres.moveSlide(slide1, 1);

    const reloaded = await Presentation.load(await pres.save());
    expect(reloaded.slides[0]?.findPlaceholder('title')?.text).toBe('Slide 2');
    expect(reloaded.slides[1]?.findPlaceholder('title')?.text).toBe('Slide 1');
  });

  it('clamps out-of-range indices', async () => {
    const pres = await Presentation.load(await readFile(fixture('two-slides.pptx')));
    const slide1 = pres.slides[0];
    if (!slide1) throw new Error('expected slide');
    // Moving to a huge index should clamp to last position.
    pres.moveSlide(slide1, 99);
    const reloaded = await Presentation.load(await pres.save());
    expect(reloaded.slides[1]?.findPlaceholder('title')?.text).toBe('Slide 1');
  });

  it('move + remove + add interleave round-trips cleanly', async () => {
    const pres = await Presentation.load(await readFile(fixture('two-slides.pptx')));
    const layout = pres.slideLayouts.find((l) => l.name === 'Title Only');
    if (!layout) throw new Error('expected layout');
    pres.addSlide({ layout });
    pres.moveSlide(pres.slides[2] as never, 0); // move the new slide to position 0
    const removed = pres.slides[2];
    if (!removed) throw new Error('expected slide');
    pres.removeSlide(removed);
    const reloaded = await Presentation.load(await pres.save());
    expect(reloaded.slides.length).toBe(2);
  });
});

describe('L3: Slide.addTextBox', () => {
  it('appends a free-form text box to a slide', async () => {
    const pres = await Presentation.load(await readFile(fixture('blank.pptx')));
    const layout = pres.slideLayouts.find((l) => l.name === 'Blank');
    if (!layout) throw new Error('expected Blank layout');
    const slide = pres.addSlide({ layout });
    const initialShapeCount = slide.shapes.length;

    const box = slide.addTextBox({
      x: inches(1),
      y: inches(1),
      w: inches(4),
      h: inches(1),
      text: 'Hello from addTextBox',
    });

    expect(box.text).toBe('Hello from addTextBox');
    expect(box.kind).toBe('shape');
    expect(slide.shapes.length).toBe(initialShapeCount + 1);

    // Geometry should match what we asked for.
    expect(box.position).toEqual({ x: inches(1), y: inches(1) });
    expect(box.size).toEqual({ w: inches(4), h: inches(1) });

    const reloaded = await Presentation.load(await pres.save());
    const reSlide = reloaded.slides[0];
    expect(reSlide).toBeDefined();
    expect(reSlide?.text).toContain('Hello from addTextBox');
  });

  it('supports a custom shape name', async () => {
    const pres = await Presentation.load(await readFile(fixture('blank.pptx')));
    const layout = pres.slideLayouts.find((l) => l.name === 'Blank');
    if (!layout) throw new Error('expected Blank layout');
    const slide = pres.addSlide({ layout });
    const box = slide.addTextBox({
      x: inches(2),
      y: pt(50),
      w: inches(3),
      h: pt(40),
      text: 'named',
      name: 'My Custom Box',
    });
    expect(box.name).toBe('My Custom Box');
  });
});
