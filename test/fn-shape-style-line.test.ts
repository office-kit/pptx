import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideShape,
  getShapeStrokeCap,
  getShapeStrokeColorResolved,
  getShapeStrokeEffective,
  getShapeStrokeWidth,
  getSlides,
  getSlideShapes,
  inches,
  loadPresentation,
  savePresentation,
} from '../src/api/index.ts';
import { readShapeStyleLineElement } from '../src/api/fn/shape-style-read.ts';
import { serializeFragment } from '../src/internal/xml/index.ts';
import { readZip, writeZip } from '../src/internal/opc/index.ts';
import { expectSchemaValid, isSchemaValidationAvailable } from './lib/expect-schema-valid.ts';

const deckPath = fileURLToPath(new URL('./fixtures/minimal/two-slides.pptx', import.meta.url));
const coloredStylePath = fileURLToPath(
  new URL('./fixtures/native/quickstyle-colored-fill-accent1-style.xml', import.meta.url),
);
const intenseStylePath = fileURLToPath(
  new URL('./fixtures/native/quickstyle-intense-accent1-style.xml', import.meta.url),
);
const decoder = new TextDecoder();
const encoder = new TextEncoder();

const styleDeck = async (stylePath: string, index?: string) => {
  const pres = await loadPresentation(await readFile(deckPath));
  const slide = getSlides(pres)[0]!;
  addSlideShape(slide, {
    preset: 'rect',
    x: inches(1),
    y: inches(1),
    w: inches(2),
    h: inches(1),
  });
  const style = (await readFile(stylePath, 'utf8')).replace(
    /(<a:lnRef idx=")\d+(")/,
    `$1${index ?? '2'}$2`,
  );
  const saved = readZip(await savePresentation(pres));
  return loadPresentation(
    writeZip(
      saved.entries.map((entry) => {
        if (entry.name !== 'ppt/slides/slide1.xml') return entry;
        const xml = decoder.decode(entry.data);
        const offset = xml.lastIndexOf('</p:sp>');
        return {
          ...entry,
          data: encoder.encode(`${xml.slice(0, offset)}${style}${xml.slice(offset)}`),
        };
      }),
    ),
  );
};

describe('shape style-matrix line references', () => {
  it.skipIf(!isSchemaValidationAvailable())(
    'preserves schema-valid native line references',
    async () => {
      const pres = await styleDeck(coloredStylePath);
      const entries = readZip(await savePresentation(pres)).entries;
      const slide = entries.find((entry) => entry.name === 'ppt/slides/slide1.xml')!;
      expectSchemaValid(decoder.decode(slide.data), 'pml');
    },
  );

  it('does not interpret a line index as a background fill index', async () => {
    const pres = await styleDeck(coloredStylePath, '1001');
    const shape = getSlideShapes(getSlides(pres)[0]!).at(-1)!;
    expect(readShapeStyleLineElement(pres, shape)).toBeNull();
  });

  it('retains theme line color transforms after reference color transforms', async () => {
    const pres = await styleDeck(coloredStylePath, '1');
    const shape = getSlideShapes(getSlides(pres)[0]!).at(-1)!;
    const line = readShapeStyleLineElement(pres, shape)!;
    const xml = serializeFragment(line);
    expect(xml).toContain('<a:shade val="15000"');
    expect(xml).toContain('<a:shade val="95000"');
    expect(xml).toContain('<a:satMod val="105000"');
    expect(xml.indexOf('15000')).toBeLessThan(xml.indexOf('95000'));
  });

  it('resolves the native Colored Fill line style and preserves it through save/load', async () => {
    const pres = await styleDeck(coloredStylePath);
    const shape = getSlideShapes(getSlides(pres)[0]!).at(-1)!;
    expect(getShapeStrokeEffective(pres, shape)).toEqual({
      kind: 'solid',
      color: 'scheme:accent1',
      widthEmu: 25400,
    });
    expect(getShapeStrokeColorResolved(pres, shape)).toBe('#1C334E');
    expect(getShapeStrokeWidth(shape)).toBeNull();
    expect(getShapeStrokeCap(shape, pres)).toBe('flat');

    const restoredPres = await loadPresentation(await savePresentation(pres));
    const restored = getSlideShapes(getSlides(restoredPres)[0]!).at(-1)!;
    expect(getShapeStrokeEffective(restoredPres, restored)).toEqual(
      getShapeStrokeEffective(pres, shape),
    );
    expect(getShapeStrokeColorResolved(restoredPres, restored)).toBe('#1C334E');
  });

  it('treats the native Intense Effect idx=0 line reference as no line', async () => {
    const pres = await styleDeck(intenseStylePath, '0');
    const shape = getSlideShapes(getSlides(pres)[0]!).at(-1)!;
    expect(getShapeStrokeEffective(pres, shape)).toEqual({ kind: 'none' });
    expect(getShapeStrokeColorResolved(pres, shape)).toBeNull();
  });
});
