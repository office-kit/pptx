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
  savePresentation,
  setShapeFill,
  setShapeTextFormat,
} from '../src/api/index.ts';
import { readZip, writeZip } from '../src/internal/opc/index.ts';

describe('public preset-color readers', () => {
  it('preserves preset fill and text colors through save and reload', async () => {
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
      text: 'red',
    });
    setShapeTextFormat(text, { color: '#FF0000' });

    const authored = await savePresentation(presentation);
    const zip = readZip(authored);
    const slideEntry = zip.entries.find((entry) => entry.name === 'ppt/slides/slide1.xml');
    if (!slideEntry) throw new Error('slide XML missing');
    const xml = new TextDecoder().decode(slideEntry.data);
    const presetXml = xml.replaceAll('<a:srgbClr val="FF0000"/>', '<a:prstClr val="red"/>');
    expect(presetXml).toContain('<a:prstClr val="red"/>');
    const loaded = await loadPresentation(
      writeZip(
        zip.entries.map((entry) =>
          entry.name === slideEntry.name
            ? { ...entry, data: new TextEncoder().encode(presetXml) }
            : entry,
        ),
      ),
    );

    const assertPresetColors = (reopened: Awaited<ReturnType<typeof loadPresentation>>) => {
      const restoredShapes = getSlideShapes(getSlides(reopened)[0]!);
      expect(getShapeFillColorResolved(reopened, restoredShapes[0]!)).toBe('#FF0000');
      expect(getShapeRunFormatEffective(reopened, restoredShapes[1]!, 0, 0).color).toBe('#FF0000');
    };
    assertPresetColors(loaded);
    assertPresetColors(await loadPresentation(await savePresentation(loaded)));
  });
});
