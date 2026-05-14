// L3: shape-level mutations (setPosition / setSize / remove) and
// an end-to-end "build a deck from scratch" smoke test that exercises
// every primitive we ship: addSlide, addTextBox, addImage, setText,
// setPosition, setSize, remove.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { Presentation, inches, pt } from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

// prettier-ignore
const PNG_1X1 = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
  0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
  0x89, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x63, 0x00, 0x01, 0x00, 0x00,
  0x05, 0x00, 0x01, 0x0d, 0x0a, 0x2d, 0xb4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae,
  0x42, 0x60, 0x82,
]);

describe('L3: setPosition / setSize', () => {
  it('moves and resizes an existing picture shape', async () => {
    const pres = await Presentation.load(await readFile(fixture('one-image-slide.pptx')));
    const slide = pres.slides[0];
    if (!slide) throw new Error('expected slide');
    const picture = slide.shapes.find((s) => s.kind === 'picture');
    if (!picture) throw new Error('expected picture');

    picture.setPosition(inches(4), inches(2));
    picture.setSize(inches(1), inches(1));

    const reloaded = await Presentation.load(await pres.save());
    const repic = reloaded.slides[0]?.shapes.find((s) => s.kind === 'picture');
    expect(repic?.position).toEqual({ x: inches(4), y: inches(2) });
    expect(repic?.size).toEqual({ w: inches(1), h: inches(1) });
  });

  it('creates a transform on a layout-inheriting placeholder when set', async () => {
    const pres = await Presentation.load(await readFile(fixture('two-slides.pptx')));
    const title = pres.slides[0]?.findPlaceholder('title');
    if (!title) throw new Error('expected title');
    expect(title.position).toBeNull(); // inheriting from layout
    title.setPosition(inches(1), inches(5));
    title.setSize(inches(6), inches(0.5));

    const reloaded = await Presentation.load(await pres.save());
    const reTitle = reloaded.slides[0]?.findPlaceholder('title');
    expect(reTitle?.position).toEqual({ x: inches(1), y: inches(5) });
    expect(reTitle?.size).toEqual({ w: inches(6), h: pt(36) /* 0.5in */ });
  });
});

describe('L3: SlideShape.remove', () => {
  it('removes a shape and shrinks the shape tree by one', async () => {
    const pres = await Presentation.load(await readFile(fixture('two-slides.pptx')));
    const slide = pres.slides[0];
    if (!slide) throw new Error('expected slide');
    const initialCount = slide.shapes.length;
    const body = slide.shapes.find((s) => s.placeholderIdx === 1);
    if (!body) throw new Error('expected body placeholder');
    body.remove();

    expect(slide.shapes.length).toBe(initialCount - 1);

    const reloaded = await Presentation.load(await pres.save());
    expect(reloaded.slides[0]?.shapes.length).toBe(initialCount - 1);
  });
});

describe('L3: end-to-end deck build from blank', () => {
  it('combines every authoring primitive into a single working deck', async () => {
    // Start from python-pptx's blank template (no slides, full layout set).
    const pres = await Presentation.load(await readFile(fixture('blank.pptx')));

    // Slide 1: Title Slide layout with both title and subtitle text.
    const titleLayout = pres.slideLayouts.find((l) => l.name === 'Title Slide');
    if (!titleLayout) throw new Error('expected Title Slide layout');
    const slide1 = pres.addSlide({ layout: titleLayout });
    slide1.findPlaceholder('ctrTitle')?.setText('pptx-kit demo');
    slide1.findPlaceholder('subTitle')?.setText('end-to-end deck from a blank template');

    // Slide 2: Blank layout with two free-form text boxes and an image.
    const blank = pres.slideLayouts.find((l) => l.name === 'Blank');
    if (!blank) throw new Error('expected Blank layout');
    const slide2 = pres.addSlide({ layout: blank });
    const heading = slide2.addTextBox({
      x: inches(1),
      y: inches(1),
      w: inches(8),
      h: inches(0.7),
      text: 'Body section heading',
    });
    heading.setSize(inches(8), inches(0.8));
    const body = slide2.addTextBox({
      x: inches(1),
      y: inches(2),
      w: inches(8),
      h: inches(3),
      text: 'Line 1\nLine 2\nLine 3',
    });
    body.setText('Replaced body text');
    slide2.addImage(PNG_1X1, { x: inches(1), y: inches(5), w: inches(2), h: inches(2) });

    // Slide 3 (added then removed to exercise removeSlide).
    const sectionLayout = pres.slideLayouts.find((l) => l.name === 'Section Header');
    if (!sectionLayout) throw new Error('expected Section Header layout');
    const slide3 = pres.addSlide({ layout: sectionLayout });
    pres.removeSlide(slide3);

    // Reorder: put slide 2 before slide 1.
    pres.moveSlide(slide2, 0);

    // Replace a token across the deck.
    slide1.findPlaceholder('ctrTitle')?.setText('Hello, {{name}}');
    const n = pres.replaceTokens({ name: 'Alice' });
    expect(n).toBe(1);

    const bytes = await pres.save();
    const reloaded = await Presentation.load(bytes);

    expect(reloaded.slides.length).toBe(2);

    // After the move, slide 0 is the blank one, slide 1 is the title.
    const slideA = reloaded.slides[0];
    expect(slideA?.text).toContain('Body section heading');
    expect(slideA?.text).toContain('Replaced body text');
    expect(slideA?.shapes.some((s) => s.kind === 'picture')).toBe(true);

    const slideB = reloaded.slides[1];
    expect(slideB?.findPlaceholder('ctrTitle')?.text).toBe('Hello, Alice');
  });
});
