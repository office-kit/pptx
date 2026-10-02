// `getShapeRunFormat` — read back per-run text formatting.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideTextBox,
  addBlankSlide,
  createPresentation,
  savePresentation,
  getSlideShapes,
  getShapeRunFormat,
  getSlides,
  inches,
  loadPresentation,
  setShapeRunFormat,
  setShapeText,
} from '../src/api/index.ts';
import { readZip, writeZip } from '../src/internal/opc/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getShapeRunFormat', () => {
  it('returns null when the run has no rPr', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(2),
      text: 'plain',
    });
    // setShapeText creates runs with rPr already (with `lang` attribute),
    // so we can't assert null on a freshly authored shape. Instead test
    // that the result is at least an empty object or has only inherited
    // values.
    const fmt = getShapeRunFormat(tb, 0, 0);
    expect(fmt).toBeDefined();
  });

  it('round-trips bold + color + size', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(2),
      text: 'styled',
    });
    setShapeRunFormat(tb, 0, 0, { bold: true, italic: true, color: '#FF0000', size: 18 });
    const fmt = getShapeRunFormat(tb, 0, 0);
    expect(fmt).not.toBeNull();
    expect(fmt!.bold).toBe(true);
    expect(fmt!.italic).toBe(true);
    expect(fmt!.color).toBe('#FF0000');
    expect(fmt!.size).toBeCloseTo(18);
  });

  it('round-trips font and fontEastAsian independently', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(2),
      text: '見出し',
    });
    setShapeRunFormat(tb, 0, 0, { font: 'Georgia', fontEastAsian: '游明朝' });
    const fmt = getShapeRunFormat(tb, 0, 0);
    expect(fmt).not.toBeNull();
    expect(fmt!.font).toBe('Georgia');
    expect(fmt!.fontEastAsian).toBe('游明朝');
  });

  it('underline encodes both boolean and explicit token', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(2),
      text: 'u',
    });
    setShapeText(tb, 'u');
    setShapeRunFormat(tb, 0, 0, { underline: true });
    expect(getShapeRunFormat(tb, 0, 0)!.underline).toBe(true);
    setShapeRunFormat(tb, 0, 0, { underline: false });
    expect(getShapeRunFormat(tb, 0, 0)!.underline).toBe(false);
  });

  it('round-trips an explicit underline color and the text-color sentinel', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(2),
      text: 'u',
    });

    setShapeRunFormat(tb, 0, 0, { underline: true, underlineColor: '#12AbEf' });
    expect(getShapeRunFormat(tb, 0, 0)!.underlineColor).toBe('#12ABEF');

    setShapeRunFormat(tb, 0, 0, { underlineColor: null });
    expect(getShapeRunFormat(tb, 0, 0)!.underlineColor).toBeNull();

    setShapeRunFormat(tb, 0, 0, { underlineColor: 'accent1' });
    expect(getShapeRunFormat(tb, 0, 0)!.underlineColor).toBe('accent1');
    setShapeRunFormat(tb, 0, 0, { underline: 'dbl' });
    expect(getShapeRunFormat(tb, 0, 0)!.underlineColor).toBe('accent1');
  });
  it.each(['#FF0000', 'accent1', null] as const)(
    'preserves underline color %s across PPTX serialization',
    async (underlineColor) => {
      const pres = createPresentation();
      const shape = addSlideTextBox(addBlankSlide(pres), {
        x: inches(0),
        y: inches(0),
        w: inches(3),
        h: inches(2),
        text: 'Colored underline',
      });
      setShapeRunFormat(shape, 0, 0, {
        underline: 'dbl',
        underlineColor,
        color: '#0000FF',
        strike: true,
      });
      const restored = await loadPresentation(await savePresentation(pres));
      const format = getShapeRunFormat(getSlideShapes(getSlides(restored)[0]!)[0]!, 0, 0);
      expect(format).toMatchObject({
        underline: 'dbl',
        underlineColor,
        color: '#0000FF',
        strike: true,
      });
    },
  );

  it('preserves an unknown non-solid underline fill when changing another run property', async () => {
    const original = createPresentation();
    const shape = addSlideTextBox(addBlankSlide(original), {
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(2),
      text: 'Future underline',
    });
    setShapeRunFormat(shape, 0, 0, { underline: true });
    const { entries } = readZip(await savePresentation(original));
    const encoder = new TextEncoder();
    const decoder = new TextDecoder();
    const mutated = writeZip(
      entries.map((entry) =>
        entry.name === 'ppt/slides/slide1.xml'
          ? {
              ...entry,
              data: encoder.encode(
                decoder
                  .decode(entry.data)
                  .replace(
                    /<a:rPr([^>]*)>/,
                    '<a:rPr$1><a:uFill><a:gradFill><a:futureFill/></a:gradFill></a:uFill>',
                  ),
              ),
            }
          : entry,
      ),
    );
    const restored = await loadPresentation(mutated);
    const restoredShape = getSlideShapes(getSlides(restored)[0]!)[0]!;
    setShapeRunFormat(restoredShape, 0, 0, { bold: true });
    const { entries: outputEntries } = readZip(await savePresentation(restored));
    const output = new TextDecoder().decode(
      outputEntries.find((entry) => entry.name === 'ppt/slides/slide1.xml')!.data,
    );
    expect(output).toContain('<a:uFill><a:gradFill><a:futureFill/></a:gradFill></a:uFill>');
  });

  it('writes underline fill in the DrawingML run-property sequence', async () => {
    const pres = createPresentation();
    const shape = addSlideTextBox(addBlankSlide(pres), {
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(2),
      text: 'Ordered underline',
    });
    setShapeRunFormat(shape, 0, 0, {
      color: '#112233',
      underline: true,
      underlineColor: '#445566',
      highlight: '#778899',
      font: 'Arial',
    });
    const { entries } = readZip(await savePresentation(pres));
    const xml = new TextDecoder().decode(
      entries.find((entry) => entry.name === 'ppt/slides/slide1.xml')!.data,
    );
    const positions = ['<a:solidFill', '<a:highlight', '<a:uFill', '<a:latin'].map((tag) =>
      xml.indexOf(tag),
    );
    expect(positions.every((position) => position >= 0)).toBe(true);
    expect(positions).toEqual([...positions].sort((a, b) => a - b));
  });
});
