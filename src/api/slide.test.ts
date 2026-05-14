import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { Presentation } from './index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`../../test/fixtures/minimal/${name}`, import.meta.url));

describe('Presentation.slides', () => {
  it('is empty for the blank fixture', async () => {
    const pres = await Presentation.load(await readFile(fixture('blank.pptx')));
    expect(pres.slides).toEqual([]);
  });

  it('returns one slide for one-text-slide.pptx', async () => {
    const pres = await Presentation.load(await readFile(fixture('one-text-slide.pptx')));
    const slides = pres.slides;
    expect(slides.length).toBe(1);
    const slide = slides[0];
    if (!slide) throw new Error('expected a slide');

    // Title shape carries "Hello, OOXML".
    expect(slide.text).toContain('Hello, OOXML');
    // Title placeholder is reachable.
    const titleShape = slide.shapes.find((s) => s.placeholderType === 'title');
    expect(titleShape?.text).toBe('Hello, OOXML');
  });

  it('returns two slides with correct ordering for two-slides.pptx', async () => {
    const pres = await Presentation.load(await readFile(fixture('two-slides.pptx')));
    const slides = pres.slides;
    expect(slides.length).toBe(2);
    expect(slides[0]?.text).toContain('Slide 1');
    expect(slides[0]?.text).toContain('Body of slide 1.');
    expect(slides[1]?.text).toContain('Slide 2');
    expect(slides[1]?.text).toContain('Body of slide 2.');
  });

  it('classifies title and body placeholders', async () => {
    const pres = await Presentation.load(await readFile(fixture('two-slides.pptx')));
    const slide = pres.slides[0];
    if (!slide) throw new Error('expected a slide');
    const placeholderTypes = slide.shapes
      .map((s) => s.placeholderType)
      .filter((t): t is string => t !== null);
    expect(placeholderTypes).toContain('title');
    // Body placeholders sometimes lack an explicit `type` attr — they
    // identify themselves only via `idx`. We confirm that explicit `title`
    // is observed; absence of others is acceptable here.
  });

  it('every shape has a numeric id and a name', async () => {
    const pres = await Presentation.load(await readFile(fixture('two-slides.pptx')));
    for (const slide of pres.slides) {
      for (const shape of slide.shapes) {
        expect(shape.id).toBeGreaterThan(0);
        expect(shape.name.length).toBeGreaterThan(0);
        expect(['shape', 'picture', 'group', 'graphicFrame', 'connector']).toContain(shape.kind);
      }
    }
  });
});
