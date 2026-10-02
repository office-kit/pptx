import { describe, expect, it } from 'vitest';
import {
  addBlankSlide,
  addSlideShape,
  addSlideTextBox,
  createPresentation,
  getShapeFillColorResolved,
  getShapeRunFormatEffective,
  getSlideShapes,
  getSlides,
  inches,
  loadPresentation,
  resolveDrawingColor,
  savePresentation,
  setShapeFill,
  setShapeTextFormat,
} from '../src/api/index.ts';
import { parseXml } from '../src/internal/xml/index.ts';
import { readZip, writeZip } from '../src/internal/opc/index.ts';

const parseColorEl = (xml: string) => parseXml(xml).root;

describe('fn API: resolveDrawingColor alternate color models', () => {
  it('resolves scrgbClr percentages using the DrawingML linear-light model', () => {
    const el = parseColorEl(
      '<a:scrgbClr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" r="50%" g="50%" b="50%"/>',
    );
    expect(resolveDrawingColor(el, null)).toBe('#BCBCBC');
  });

  it('resolves hslClr hue, saturation, and luminance attributes', () => {
    const el = parseColorEl(
      '<a:hslClr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" hue="0" sat="100%" lum="50%"/>',
    );
    expect(resolveDrawingColor(el, null)).toBe('#FF0000');
  });

  it('accepts fixed-point percentages and nonzero hue', () => {
    const scrgb = parseColorEl(
      '<a:scrgbClr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" r="10000" g="20000" b="30000"/>',
    );
    expect(resolveDrawingColor(scrgb, null)).toBe('#597C95');
    const hsl = parseColorEl(
      '<a:hslClr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" hue="7200000" sat="100%" lum="50%"/>',
    );
    expect(resolveDrawingColor(hsl, null)).toBe('#00FF00');
  });

  it('applies transforms to alternate color models and clips channels', () => {
    const transformed = parseColorEl(
      '<a:scrgbClr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" r="100%" g="0%" b="0%"><a:shade val="50000"/></a:scrgbClr>',
    );
    expect(resolveDrawingColor(transformed, null)).toBe('#BC0000');
    const clipped = parseColorEl(
      '<a:scrgbClr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" r="-10%" g="150%" b="50%"/>',
    );
    expect(resolveDrawingColor(clipped, null)).toBe('#00FFBC');
  });

  it('returns null when required alternate-model attributes are absent or invalid', () => {
    const missing = parseColorEl(
      '<a:hslClr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" hue="0" sat="100%"/>',
    );
    expect(resolveDrawingColor(missing, null)).toBeNull();
    const missingHue = parseColorEl(
      '<a:hslClr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" sat="100%" lum="50%"/>',
    );
    expect(resolveDrawingColor(missingHue, null)).toBeNull();
    const invalid = parseColorEl(
      '<a:scrgbClr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" r="wat" g="0%" b="0%"/>',
    );
    expect(resolveDrawingColor(invalid, null)).toBeNull();
  });

  it('resolves imported alternate colors through shape fill and text consumers after reload', async () => {
    const presentation = createPresentation();
    const slide = addBlankSlide(presentation);
    const shape = addSlideShape(slide, {
      preset: 'rect',
      x: inches(1),
      y: inches(1),
      w: inches(2),
      h: inches(1),
    });
    setShapeFill(shape, '#FF0000');
    const text = addSlideTextBox(slide, {
      x: inches(1),
      y: inches(2.5),
      w: inches(4),
      h: inches(1),
      text: 'green',
    });
    setShapeTextFormat(text, { color: '#00FF00' });
    const zip = readZip(await savePresentation(presentation));
    const slideEntry = zip.entries.find((entry) => entry.name === 'ppt/slides/slide1.xml');
    if (!slideEntry) throw new Error('slide XML missing');
    const xml = new TextDecoder()
      .decode(slideEntry.data)
      .replace('<a:srgbClr val="FF0000"/>', '<a:scrgbClr r="100%" g="0%" b="0%"/>')
      .replace('<a:srgbClr val="00FF00"/>', '<a:hslClr hue="7200000" sat="100%" lum="50%"/>');
    expect(xml).toContain('<a:scrgbClr r="100%" g="0%" b="0%"/>');
    expect(xml).toContain('<a:hslClr hue="7200000" sat="100%" lum="50%"/>');
    const loaded = await loadPresentation(
      writeZip(
        zip.entries.map((entry) =>
          entry.name === slideEntry.name
            ? { ...entry, data: new TextEncoder().encode(xml) }
            : entry,
        ),
      ),
    );
    const assertColors = (reopened: Awaited<ReturnType<typeof loadPresentation>>) => {
      const shapes = getSlideShapes(getSlides(reopened)[0]!);
      expect(getShapeFillColorResolved(reopened, shapes[0]!)).toBe('#FF0000');
      expect(getShapeRunFormatEffective(reopened, shapes[1]!, 0, 0).color).toBe('#00FF00');
    };
    assertColors(loaded);
    assertColors(await loadPresentation(await savePresentation(loaded)));
  });
});
