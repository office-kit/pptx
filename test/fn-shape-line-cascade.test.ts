import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideShape,
  getShapeStrokeArrow,
  getShapeStrokeColorResolved,
  getShapeStrokeDash,
  getShapeStrokeEffective,
  getSlideShapes,
  getSlides,
  inches,
  loadPresentation,
  savePresentation,
} from '../src/api/index.ts';
import { readZip, writeZip } from '../src/internal/opc/index.ts';

const deckPath = fileURLToPath(new URL('./fixtures/minimal/two-slides.pptx', import.meta.url));
const stylePath = fileURLToPath(
  new URL('./fixtures/native/quickstyle-colored-fill-accent1-style.xml', import.meta.url),
);
const decoder = new TextDecoder();
const encoder = new TextEncoder();

/** Adds a native lnRef and a deliberately partial direct line to one shape. */
const partialLineDeck = async (
  line = '<a:ln w="12700"/>',
  themeLineExtras = false,
  styleIndex = '2',
) => {
  const pres = await loadPresentation(await readFile(deckPath));
  const slide = getSlides(pres)[0]!;
  addSlideShape(slide, {
    preset: 'rect',
    x: inches(1),
    y: inches(1),
    w: inches(2),
    h: inches(1),
  });
  const style = (await readFile(stylePath, 'utf8')).replace('idx="2"', `idx="${styleIndex}"`);
  const saved = readZip(await savePresentation(pres));
  return loadPresentation(
    writeZip(
      saved.entries.map((entry) => {
        if (entry.name === 'ppt/theme/theme1.xml' && themeLineExtras) {
          const xml = decoder.decode(entry.data);
          let lineNumber = 0;
          const themed = xml.replace(/<a:ln\b[^>]*>[\s\S]*?<\/a:ln>/g, (lineXml) => {
            lineNumber += 1;
            if (lineNumber !== 2) return lineXml;
            return lineXml
              .replace('<a:prstDash val="solid"/>', '<a:prstDash val="dashDot"/>')
              .replace('</a:ln>', '<a:tailEnd type="triangle" w="med" len="lg"/></a:ln>');
          });
          return { ...entry, data: encoder.encode(themed) };
        }
        if (entry.name !== 'ppt/slides/slide1.xml') return entry;
        const xml = decoder.decode(entry.data);
        const spPrEnd = xml.lastIndexOf('</p:spPr>');
        const withLine = `${xml.slice(0, spPrEnd)}${line}${xml.slice(spPrEnd)}`;
        const shapeEnd = withLine.lastIndexOf('</p:sp>');
        return {
          ...entry,
          data: encoder.encode(`${withLine.slice(0, shapeEnd)}${style}${withLine.slice(shapeEnd)}`),
        };
      }),
    ),
  );
};

describe('shape direct line and style-reference cascade', () => {
  it('overlays a direct width while retaining style paint', async () => {
    const pres = await partialLineDeck();
    const shape = getSlideShapes(getSlides(pres)[0]!).at(-1)!;

    expect(getShapeStrokeEffective(pres, shape)).toMatchObject({
      kind: 'solid',
      widthEmu: 12700,
    });
    expect(getShapeStrokeColorResolved(pres, shape)).toBe('#1C334E');
  });

  it('lets a direct noFill override the style line', async () => {
    const pres = await partialLineDeck('<a:ln><a:noFill/></a:ln>');
    const shape = getSlideShapes(getSlides(pres)[0]!).at(-1)!;

    expect(getShapeStrokeEffective(pres, shape)).toEqual({ kind: 'none' });
    expect(getShapeStrokeColorResolved(pres, shape)).toBeNull();
  });

  it('keeps an idx=0 no-line style when a direct line only sets width', async () => {
    const pres = await partialLineDeck('<a:ln w="12700"/>', false, '0');
    const shape = getSlideShapes(getSlides(pres)[0]!).at(-1)!;
    expect(getShapeStrokeEffective(pres, shape)).toEqual({ kind: 'none' });
  });

  it('allows direct paint to replace an idx=0 no-line style', async () => {
    const pres = await partialLineDeck(
      '<a:ln><a:solidFill><a:srgbClr val="123456"/></a:solidFill></a:ln>',
      false,
      '0',
    );
    const shape = getSlideShapes(getSlides(pres)[0]!).at(-1)!;
    expect(getShapeStrokeEffective(pres, shape)).toMatchObject({ kind: 'solid', color: '#123456' });
  });

  it('inherits style width when direct paint omits width', async () => {
    const pres = await partialLineDeck(
      '<a:ln><a:solidFill><a:srgbClr val="123456"/></a:solidFill></a:ln>',
    );
    const shape = getSlideShapes(getSlides(pres)[0]!).at(-1)!;
    expect(getShapeStrokeEffective(pres, shape)).toMatchObject({ kind: 'solid', widthEmu: 25400 });
  });

  it('reads style arrow properties when a direct line only supplies width', async () => {
    const pres = await partialLineDeck(undefined, true);
    const shape = getSlideShapes(getSlides(pres)[0]!).at(-1)!;
    expect(getShapeStrokeArrow(shape, 'head', pres)).toBeNull();
    expect(getShapeStrokeDash(shape, pres)).toBe('dashDot');
    expect(getShapeStrokeArrow(shape, 'tail', pres)).toEqual({
      type: 'triangle',
      width: 'med',
      length: 'lg',
    });
  });
});
