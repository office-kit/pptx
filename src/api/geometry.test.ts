import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { Presentation, inches } from './index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`../../test/fixtures/minimal/${name}`, import.meta.url));

describe('SlideShape geometry', () => {
  it('reads position and size of a programmatically placed picture', async () => {
    const pres = await Presentation.load(await readFile(fixture('one-image-slide.pptx')));
    const picture = pres.slides[0]?.shapes.find((s) => s.kind === 'picture');
    if (!picture) throw new Error('expected a picture shape');

    // The fixture script calls add_picture(buf, Inches(2), Inches(2), Inches(3), Inches(3))
    expect(picture.position).toEqual({ x: inches(2), y: inches(2) });
    expect(picture.size).toEqual({ w: inches(3), h: inches(3) });
  });

  it('returns null position for placeholders inheriting from layout', async () => {
    const pres = await Presentation.load(await readFile(fixture('two-slides.pptx')));
    const slide = pres.slides[0];
    if (!slide) throw new Error('expected a slide');
    // python-pptx's "Title and Content" layout: slide-level placeholder
    // shapes don't carry their own xfrm; they inherit from the layout.
    const title = slide.findPlaceholder('title');
    expect(title?.position).toBeNull();
    expect(title?.size).toBeNull();
  });
});
