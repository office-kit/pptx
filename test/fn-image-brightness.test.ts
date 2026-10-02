// Picture brightness via the shared <a:blip><a:lum bright="…"/>.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import {
  getShapeImageBrightness,
  getShapeImageBiLevelThreshold,
  getShapeImageContrast,
  getShapeImageOpacity,
  getShapeKind,
  getSlideXmlString,
  getSlideShapes,
  getSlides,
  isShapeImageGrayscale,
  loadPresentation,
  setShapeImageBrightness,
  setShapeImageContrast,
  resetShapeImageColorEffects,
  setShapeImageRecolor,
  savePresentation,
  createPresentation,
  addBlankSlide,
  addSlideShape,
  inches,
  setShapeImageFill,
} from '../src/api/index.ts';
import { SHAPE_ELEMENT } from '../src/api/_internal-symbols.ts';
import { NS, attr, elem, firstChildElement, qname } from '../src/internal/xml/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: setShapeImageBrightness', () => {
  it('supports image-filled shapes, preserves one correction when clearing the other, and round-trips', async () => {
    const pres = createPresentation();
    const shape = addSlideShape(addBlankSlide(pres), {
      preset: 'rect',
      x: inches(1),
      y: inches(1),
      w: inches(2),
      h: inches(2),
    });
    setShapeImageFill(shape, new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]));
    setShapeImageBrightness(shape, 0.2);
    setShapeImageContrast(shape, -0.3);
    expect(getShapeImageBrightness(shape)).toBeCloseTo(0.2);
    expect(getShapeImageContrast(shape)).toBeCloseTo(-0.3);

    setShapeImageBrightness(shape, null);
    expect(getShapeImageBrightness(shape)).toBeNull();
    expect(getShapeImageContrast(shape)).toBeCloseTo(-0.3);

    const restoredPresentation = await loadPresentation(await savePresentation(pres));
    const restored = getSlideShapes(getSlides(restoredPresentation)[0]!)[0]!;
    expect(getShapeImageBrightness(restored)).toBeNull();
    expect(getShapeImageContrast(restored)).toBeCloseTo(-0.3);
  });

  it('rejects shapes without an image fill', () => {
    const pres = createPresentation();
    const shape = addSlideShape(addBlankSlide(pres), {
      preset: 'rect',
      x: inches(1),
      y: inches(1),
      w: inches(2),
      h: inches(2),
    });
    expect(() => setShapeImageBrightness(shape, 0.2)).toThrow(
      /picture or a shape with an image fill/,
    );
    expect(() => setShapeImageContrast(shape, 0.2)).toThrow(
      /picture or a shape with an image fill/,
    );
  });

  it('round-trips a brightness fraction', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-image-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    const picture = getSlideShapes(slide).find((s) => getShapeKind(s) === 'picture')!;
    expect(getShapeImageBrightness(picture)).toBeNull();
    setShapeImageBrightness(picture, 0.4);
    expect(getShapeImageBrightness(picture)).toBeCloseTo(0.4);
    setShapeImageBrightness(picture, -0.2);
    expect(getShapeImageBrightness(picture)).toBeCloseTo(-0.2);
    setShapeImageBrightness(picture, null);
    expect(getShapeImageBrightness(picture)).toBeNull();
  });

  it('rejects non-pictures and out-of-range values', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    const text = getSlideShapes(slide).find((s) => getShapeKind(s) === 'shape')!;
    expect(() => setShapeImageBrightness(text, 0.5)).toThrow(/picture/);

    const pres2 = await loadPresentation(await readFile(fixture('one-image-slide.pptx')));
    const picture = getSlideShapes(getSlides(pres2)[0]!).find(
      (s) => getShapeKind(s) === 'picture',
    )!;
    expect(() => setShapeImageBrightness(picture, 1.5)).toThrow(RangeError);
    expect(() => setShapeImageBrightness(picture, -1.5)).toThrow(RangeError);
  });

  it('clears PowerPoint color corrections while preserving unrelated blip effects', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-image-slide.pptx')));
    const picture = getSlideShapes(getSlides(pres)[0]!).find((s) => getShapeKind(s) === 'picture')!;
    setShapeImageBrightness(picture, 0.2);
    setShapeImageContrast(picture, 0.3);
    const blipFill = firstChildElement(picture[SHAPE_ELEMENT], qname('p', 'blipFill', NS.pml))!;
    const blip = firstChildElement(blipFill, qname('a', 'blip', NS.dml))!;
    blip.children.push(
      elem(qname('a', 'grayscl', NS.dml)),
      elem(qname('a', 'duotone', NS.dml), {
        children: [
          elem(qname('a', 'srgbClr', NS.dml), { attrs: [attr(qname('', 'val', ''), 'FF0000')] }),
          elem(qname('a', 'srgbClr', NS.dml), { attrs: [attr(qname('', 'val', ''), '0000FF')] }),
        ],
      }),
      elem(qname('a', 'biLevel', NS.dml), {
        attrs: [attr(qname('', 'thresh', ''), '50000')],
      }),
      elem(qname('a', 'alphaModFix', NS.dml), {
        attrs: [attr(qname('', 'amt', ''), '75000')],
      }),
      elem(qname('a', 'extLst', NS.dml), {
        children: [
          elem(qname('a', 'ext', NS.dml), {
            attrs: [attr(qname('', 'uri', ''), '{00000000-0000-0000-0000-000000000001}')],
          }),
        ],
      }),
    );

    resetShapeImageColorEffects(picture);

    expect(getShapeImageBrightness(picture)).toBeNull();
    expect(getShapeImageContrast(picture)).toBeNull();
    const xml = getSlideXmlString(getSlides(pres)[0]!);
    expect(xml).not.toContain('<a:grayscl');
    expect(xml).not.toContain('<a:duotone');
    expect(xml).not.toContain('<a:biLevel');
    expect(xml).toContain('<a:alphaModFix');
    expect(xml).toContain('<a:extLst');

    const restored = await loadPresentation(await savePresentation(pres));
    const restoredPicture = getSlideShapes(getSlides(restored)[0]!).find(
      (s) => getShapeKind(s) === 'picture',
    )!;
    expect(getShapeImageOpacity(restoredPicture)).toBeCloseTo(0.75);
    const restoredXml = getSlideXmlString(getSlides(restored)[0]!);
    expect(restoredXml).not.toContain('<a:biLevel');
    expect(restoredXml).toContain('<a:extLst');
  });
});

describe('fn API: setShapeImageRecolor', () => {
  it('reads fixed-point and percent lexical threshold values', async () => {
    for (const [lexical, expected] of [
      ['1', 0.001],
      ['1%', 1],
      ['50000', 50],
    ] as const) {
      const parts = unzipSync(await readFile(fixture('one-image-slide.pptx')));
      const slidePart = 'ppt/slides/slide1.xml';
      parts[slidePart] = strToU8(
        strFromU8(parts[slidePart]!).replace(
          '<a:blip r:embed="rId2"/>',
          `<a:blip r:embed="rId2"><a:biLevel thresh="${lexical}"/></a:blip>`,
        ),
      );
      const pres = await loadPresentation(zipSync(parts));
      const picture = getSlideShapes(getSlides(pres)[0]!).find(
        (s) => getShapeKind(s) === 'picture',
      )!;
      expect(getShapeImageBiLevelThreshold(picture)).toBeCloseTo(expected, 6);
    }
  });

  it('round-trips the minimum fixed-point threshold', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-image-slide.pptx')));
    const picture = getSlideShapes(getSlides(pres)[0]!).find((s) => getShapeKind(s) === 'picture')!;
    setShapeImageRecolor(picture, { kind: 'threshold', threshold: 0.001 });
    expect(getShapeImageBiLevelThreshold(picture)).toBeCloseTo(0.001, 6);
    expect(getSlideXmlString(getSlides(pres)[0]!)).toContain('<a:biLevel thresh="1"/>');

    const roundTripped = await loadPresentation(await savePresentation(pres));
    const roundTrippedPicture = getSlideShapes(getSlides(roundTripped)[0]!).find(
      (s) => getShapeKind(s) === 'picture',
    )!;
    expect(getShapeImageBiLevelThreshold(roundTrippedPicture)).toBeCloseTo(0.001, 6);
  });

  it('writes PowerPoint recolor effects and clears them without touching opacity', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-image-slide.pptx')));
    const picture = getSlideShapes(getSlides(pres)[0]!).find((s) => getShapeKind(s) === 'picture')!;
    const blipFill = firstChildElement(picture[SHAPE_ELEMENT], qname('p', 'blipFill', NS.pml))!;
    const blip = firstChildElement(blipFill, qname('a', 'blip', NS.dml))!;
    blip.children.push(
      elem(qname('a', 'alphaModFix', NS.dml), {
        attrs: [attr(qname('', 'amt', ''), '75000')],
      }),
      elem(qname('a', 'extLst', NS.dml), {
        children: [
          elem(qname('a', 'ext', NS.dml), { attrs: [attr(qname('', 'uri', ''), 'test')] }),
        ],
      }),
    );

    setShapeImageRecolor(picture, { kind: 'grayscale' });
    expect(getSlideXmlString(getSlides(pres)[0]!)).toContain('<a:grayscl');

    setShapeImageRecolor(picture, {
      kind: 'duotone',
      colors: [
        { color: 'accent1', colorTransforms: [{ kind: 'tint', value: 0.45 }] },
        { color: '#D9C3A5', colorTransforms: [{ kind: 'satMod', value: 1.8 }] },
      ],
    });
    let xml = getSlideXmlString(getSlides(pres)[0]!);
    expect(xml).toContain('<a:duotone>');
    expect(xml).toContain('<a:schemeClr val="accent1"><a:tint val="45000"/></a:schemeClr>');
    expect(xml).toContain('<a:srgbClr val="D9C3A5"><a:satMod val="180000"/></a:srgbClr>');
    expect(xml.indexOf('<a:duotone>')).toBeLessThan(xml.indexOf('<a:extLst>'));
    const roundTripped = await loadPresentation(await savePresentation(pres));
    expect(getSlideXmlString(getSlides(roundTripped)[0]!)).toContain(
      '<a:schemeClr val="accent1"><a:tint val="45000"/></a:schemeClr>',
    );

    setShapeImageRecolor(picture, { kind: 'threshold', threshold: 50 });
    xml = getSlideXmlString(getSlides(pres)[0]!);
    expect(xml).toContain('<a:biLevel thresh="50000"/>');

    setShapeImageRecolor(picture, { kind: 'washout' });
    xml = getSlideXmlString(getSlides(pres)[0]!);
    expect(xml).toContain('<a:lum bright="70000" contrast="-70000"/>');

    setShapeImageRecolor(picture, { kind: 'none' });
    xml = getSlideXmlString(getSlides(pres)[0]!);
    expect(xml).not.toContain('<a:grayscl');
    expect(xml).not.toContain('<a:duotone');
    expect(xml).not.toContain('<a:biLevel');
    expect(xml).not.toContain('<a:lum');
    expect(xml).toContain('<a:alphaModFix amt="75000"/>');
    expect(xml).toContain('<a:extLst>');
  });

  it('validates duotone colors before replacing an existing effect', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-image-slide.pptx')));
    const picture = getSlideShapes(getSlides(pres)[0]!).find((s) => getShapeKind(s) === 'picture')!;
    setShapeImageRecolor(picture, { kind: 'grayscale' });
    expect(() =>
      setShapeImageRecolor(picture, {
        kind: 'duotone',
        colors: ['#nothex', '#FFFFFF'],
      }),
    ).toThrow(/unrecognized color/);
    expect(isShapeImageGrayscale(picture)).toBe(true);
  });

  it('rejects thresholds outside PowerPoint percent range', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-image-slide.pptx')));
    const picture = getSlideShapes(getSlides(pres)[0]!).find((s) => getShapeKind(s) === 'picture')!;
    expect(() => setShapeImageRecolor(picture, { kind: 'threshold', threshold: -1 })).toThrow(
      RangeError,
    );
    expect(() => setShapeImageRecolor(picture, { kind: 'threshold', threshold: 101 })).toThrow(
      RangeError,
    );
  });
});
