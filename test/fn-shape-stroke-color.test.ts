// getShapeStrokeColor / getShapeStrokeWidth — sugar over getShapeStroke.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideShape,
  emu,
  getShapeStrokeColor,
  getShapeStrokeOpacity,
  getSlideShapes,
  savePresentation,
  getShapeStrokeWidth,
  getSlides,
  inches,
  loadPresentation,
  setShapeNoStroke,
  setShapeStroke,
} from '../src/api/index.ts';

import { readZip, writeZip } from '../src/internal/opc/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getShapeStrokeColor / getShapeStrokeWidth', () => {
  it('returns the color and width set via setShapeStroke', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const rect = addSlideShape(slide, {
      preset: 'rect',
      x: inches(0),
      y: inches(0),
      w: inches(1),
      h: inches(1),
    });
    setShapeStroke(rect, { color: '#112233', widthEmu: emu(25400) });
    expect(getShapeStrokeColor(rect)).toBe('#112233');
    expect(getShapeStrokeWidth(rect)).toBe(25400);
  });

  it('returns null after clearing the stroke', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const rect = addSlideShape(slide, {
      preset: 'rect',
      x: inches(0),
      y: inches(0),
      w: inches(1),
      h: inches(1),
    });
    setShapeNoStroke(rect);
    expect(getShapeStrokeColor(rect)).toBeNull();
    expect(getShapeStrokeWidth(rect)).toBeNull();
  });
});

it.each(['#112233', 'scheme:accent1'])(
  'width-only stroke edits preserve %s and opacity through save',
  async (color) => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const shape = addSlideShape(getSlides(pres)[0]!, {
      preset: 'rect',
      x: inches(0),
      y: inches(0),
      w: inches(1),
      h: inches(1),
    });
    setShapeStroke(shape, { color, widthEmu: 12700 });
    const { entries } = readZip(await savePresentation(pres));
    const colorXml = color.startsWith('#')
      ? '<a:srgbClr val="112233"/>'
      : '<a:schemeClr val="accent1"/>';
    const transformed = colorXml.replace(
      '/>',
      '><a:alpha val="40000"/></' + (color.startsWith('#') ? 'a:srgbClr' : 'a:schemeClr') + '>',
    );
    const loaded = await loadPresentation(
      writeZip(
        entries.map((entry) =>
          entry.name === 'ppt/slides/slide1.xml'
            ? {
                ...entry,
                data: new TextEncoder().encode(
                  new TextDecoder().decode(entry.data).replace(colorXml, transformed),
                ),
              }
            : entry,
        ),
      ),
    );
    const target = getSlideShapes(getSlides(loaded)[0]!).at(-1)!;
    expect(getShapeStrokeOpacity(target)).toBe(0.4);
    setShapeStroke(target, { widthEmu: 76200 });
    expect(getShapeStrokeColor(target)).toBe(color);
    expect(getShapeStrokeOpacity(target)).toBe(0.4);
    const restored = await loadPresentation(await savePresentation(loaded));
    const result = getSlideShapes(getSlides(restored)[0]!).at(-1)!;
    expect(getShapeStrokeColor(result)).toBe(color);
    expect(getShapeStrokeOpacity(result)).toBe(0.4);
    expect(getShapeStrokeWidth(result)).toBe(76200);
  },
);
