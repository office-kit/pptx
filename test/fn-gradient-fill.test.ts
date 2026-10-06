// Linear gradient fill on a shape.

import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { expectSchemaValid, isSchemaValidationAvailable } from './lib/expect-schema-valid.ts';
import {
  getShapeGradientFill,
  setShapeFill,
  setShapeRotation,
  getSlideShapes,
  getSlideXmlString,
  getSlides,
  loadPresentation,
  savePresentation,
  setShapeGradientFill,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const slideXml = async (bytes: Uint8Array, slideIndex: number): Promise<string> => {
  const pres = await loadPresentation(bytes);
  return getSlideXmlString(getSlides(pres)[slideIndex]!);
};

describe('fn API: setShapeGradientFill', () => {
  it('writes a gradFill with the configured stops and angle', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = getSlideShapes(slide)[0]!;
    setShapeGradientFill(shape, {
      stops: [
        { offset: 0, color: '#FF0000' },
        { offset: 1, color: '#0000FF' },
      ],
      angleDeg: 90,
    });
    const xml = await slideXml(await savePresentation(pres), 0);
    expect(xml).toContain('<a:gradFill');
    expect(xml).toContain('<a:gs pos="0">');
    expect(xml).toContain('<a:gs pos="100000">');
    expect(xml).toContain('FF0000');
    expect(xml).toContain('0000FF');
    // 90 degrees × 60000 = 5400000
    expect(xml).toContain('ang="5400000"');
  });

  it('default angle is 90 degrees when omitted', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = getSlideShapes(slide)[0]!;
    setShapeGradientFill(shape, {
      stops: [
        { offset: 0, color: '#FFFFFF' },
        { offset: 1, color: '#000000' },
      ],
    });
    expect(await slideXml(await savePresentation(pres), 0)).toContain('ang="5400000"');
  });

  it('supports multi-stop gradients', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = getSlideShapes(slide)[0]!;
    setShapeGradientFill(shape, {
      stops: [
        { offset: 0, color: '#FF0000' },
        { offset: 0.5, color: '#00FF00' },
        { offset: 1, color: '#0000FF' },
      ],
    });
    const xml = await slideXml(await savePresentation(pres), 0);
    expect(xml).toContain('<a:gs pos="0">');
    expect(xml).toContain('<a:gs pos="50000">');
    expect(xml).toContain('<a:gs pos="100000">');
  });

  it('rejects fewer than two stops', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = getSlideShapes(slide)[0]!;
    expect(() => setShapeGradientFill(shape, { stops: [{ offset: 0, color: '#FFFFFF' }] })).toThrow(
      /two stops/,
    );
  });

  it('rejects offsets outside [0, 1]', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = getSlideShapes(slide)[0]!;
    expect(() =>
      setShapeGradientFill(shape, {
        stops: [
          { offset: 0, color: '#FFFFFF' },
          { offset: 1.5, color: '#000000' },
        ],
      }),
    ).toThrow(RangeError);
  });
});

describe('gradient stop editing', () => {
  it('round-trips stop brightness, opacity, and direction options', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    const shape = getSlideShapes(getSlides(pres)[0]!)[0]!;
    setShapeGradientFill(shape, {
      stops: [
        { offset: 0, color: 'accent1', brightness: 0.95, opacity: 0.25 },
        { offset: 1, color: '#FF0000', brightness: -0.5 },
      ],
      angleDeg: 45,
      rotateWithShape: false,
      scaled: true,
    });
    if (isSchemaValidationAvailable())
      expectSchemaValid(getSlideXmlString(getSlides(pres)[0]!), 'pml');
    const reloaded = await loadPresentation(await savePresentation(pres));
    expect(getShapeGradientFill(getSlideShapes(getSlides(reloaded)[0]!)[0]!)).toMatchObject({
      stops: [
        { offset: 0, color: 'scheme:accent1', brightness: 0.95, opacity: 0.25 },
        { offset: 1, color: '#FF0000', brightness: -0.5 },
      ],
      angleDeg: 45,
      rotateWithShape: false,
      scaled: true,
    });
  });

  it.each([
    { brightness: 1.1 },
    { brightness: Number.NaN },
    { opacity: -0.1 },
    { opacity: Number.NaN },
  ])('rejects invalid stop values atomically: %j', async (invalid) => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = getSlideShapes(slide)[0]!;
    setShapeFill(shape, '#123456');
    expect(() =>
      setShapeGradientFill(shape, {
        stops: [
          { offset: 0, color: '#000000' },
          { offset: 1, color: '#FFFFFF', ...invalid },
        ],
      }),
    ).toThrow();
    setShapeRotation(shape, 10);
    expect(getSlideXmlString(slide)).toContain('123456');
    expect(getSlideXmlString(slide)).not.toContain('<a:gradFill');
  });
});

describe('imported path gradient focus', () => {
  it.each([
    ['l="100000" t="100000"', { left: 1, top: 1, right: 0, bottom: 0 }],
    ['', { left: 0, top: 0, right: 0, bottom: 0 }],
    ['l="1" t="-1" r="50%" b="-25%"', { left: 0.00001, top: -0.00001, right: 0.5, bottom: -0.25 }],
  ])('reads OOXML inset percentages and omitted zero defaults: %s', async (attributes, focus) => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    setShapeGradientFill(getSlideShapes(getSlides(pres)[0]!)[0]!, {
      stops: [
        { offset: 0, color: '#FF0000' },
        { offset: 1, color: '#0000FF' },
      ],
      path: 'circle',
      focus: { left: 0.5, top: 0.5, right: 0.5, bottom: 0.5 },
    });
    const parts = unzipSync(await savePresentation(pres));
    const name = 'ppt/slides/slide1.xml';
    parts[name] = strToU8(
      strFromU8(parts[name]!).replace(/<a:fillToRect[^>]*\/>/, `<a:fillToRect ${attributes}/>`),
    );
    const loaded = await loadPresentation(zipSync(parts));
    const shape = getSlideShapes(getSlides(loaded)[0]!)[0]!;
    expect(getShapeGradientFill(shape)?.focus).toEqual(focus);
    setShapeGradientFill(shape, {
      stops: [
        { offset: 0, color: '#FF0000' },
        { offset: 1, color: '#0000FF' },
      ],
      path: 'circle',
      focus,
    });
    const reloaded = await loadPresentation(await savePresentation(loaded));
    expect(getShapeGradientFill(getSlideShapes(getSlides(reloaded)[0]!)[0]!)?.focus).toEqual(focus);
  });
});

describe('gradient tile rectangle', () => {
  it.each([Number.NaN, Number.POSITIVE_INFINITY, 21474.83648, -21474.83649])(
    'rejects invalid tile insets atomically: %s',
    async (right) => {
      const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
      const slide = getSlides(pres)[0]!;
      const shape = getSlideShapes(slide)[0]!;
      setShapeFill(shape, '#123456');
      expect(() =>
        setShapeGradientFill(shape, {
          stops: [
            { offset: 0, color: '#000000' },
            { offset: 1, color: '#FFFFFF' },
          ],
          tileRect: { left: 0, top: 0, right, bottom: 0 },
        }),
      ).toThrow(RangeError);
      expect(getSlideXmlString(slide)).toContain('123456');
      expect(getSlideXmlString(slide)).not.toContain('<a:gradFill');
    },
  );

  it.each(['circle', 'linear'] as const)(
    'retains native tile insets when editing a %s gradient',
    async (path) => {
      const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
      const stops = [
        { offset: 0, color: '#FF0000' },
        { offset: 1, color: '#0000FF' },
      ] as const;
      setShapeGradientFill(getSlideShapes(getSlides(pres)[0]!)[0]!, { stops, path });
      const parts = unzipSync(await savePresentation(pres));
      const name = 'ppt/slides/slide1.xml';
      // Mac PowerPoint's Radial > From Bottom Right Corner uses these insets.
      parts[name] = strToU8(
        strFromU8(parts[name]!).replace(
          '</a:gradFill>',
          '<a:tileRect r="-100000" b="-100000"/></a:gradFill>',
        ),
      );
      const loaded = await loadPresentation(zipSync(parts));
      const shape = getSlideShapes(getSlides(loaded)[0]!)[0]!;
      const gradient = getShapeGradientFill(shape)!;
      const tileRect = { left: 0, top: 0, right: -1, bottom: -1 };
      expect(gradient).toMatchObject({ tileRect });
      setShapeGradientFill(shape, { ...gradient, stops, rotateWithShape: false });
      const xml = getSlideXmlString(getSlides(loaded)[0]!);
      expect(xml).toContain('<a:tileRect l="0" t="0" r="-100000" b="-100000"/>');
      if (isSchemaValidationAvailable()) expectSchemaValid(xml, 'pml');
      const reloaded = await loadPresentation(await savePresentation(loaded));
      expect(getShapeGradientFill(getSlideShapes(getSlides(reloaded)[0]!)[0]!)).toMatchObject({
        tileRect,
        rotateWithShape: false,
      });
    },
  );
});
