import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideShape,
  addSlideTextBox,
  getShapeFill,
  getShapeFillColorResolved,
  getShapeFillEffective,
  getShapeFillOpacity,
  getShapeRunFormatEffective,
  getSlideShapes,
  getSlideXmlString,
  getSlides,
  inches,
  loadPresentation,
  savePresentation,
  setShapeFill,
  setShapeRunFormat,
} from '../src/api/index.ts';
import { readZip, writeZip } from '../src/internal/opc/index.ts';
import { renderSlideToSvg } from '../packages/preview/src/index.ts';
import { expectSchemaValid, isSchemaValidationAvailable } from './lib/expect-schema-valid.ts';

const deckPath = fileURLToPath(new URL('./fixtures/minimal/two-slides.pptx', import.meta.url));
const stylePath = fileURLToPath(
  new URL('./fixtures/native/quickstyle-colored-fill-accent1-style.xml', import.meta.url),
);
const decoder = new TextDecoder();
const encoder = new TextEncoder();

const styleDeck = async (
  styleIndex = '1',
  alpha = false,
  text = false,
  paragraphColor?: string,
) => {
  const pres = await loadPresentation(await readFile(deckPath));
  const slide = getSlides(pres)[0]!;
  if (text) {
    addSlideTextBox(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(2),
      h: inches(1),
      text: 'Styled text',
    });
  } else {
    addSlideShape(slide, {
      preset: 'rect',
      x: inches(1),
      y: inches(1),
      w: inches(2),
      h: inches(1),
    });
  }
  const style = (await readFile(stylePath, 'utf8'))
    .replace('idx="1"', `idx="${styleIndex}"`)
    .replace(
      '<a:schemeClr val="accent1"/></a:fillRef>',
      `<a:schemeClr val="accent1"${alpha ? '><a:alpha val="27000"/></a:schemeClr>' : '/'}></a:fillRef>`,
    );
  const saved = readZip(await savePresentation(pres));
  return loadPresentation(
    writeZip(
      saved.entries.map((entry) =>
        entry.name === 'ppt/slides/slide1.xml'
          ? {
              ...entry,
              data: (() => {
                const xml = decoder.decode(entry.data);
                const closeMarker = '</p:sp>';
                const closeOffset = xml.lastIndexOf(closeMarker);
                if (closeOffset < 0) return encoder.encode(xml);
                const txBodyOffset = text ? xml.lastIndexOf('<p:txBody>', closeOffset) : -1;
                const offset = txBodyOffset >= 0 ? txBodyOffset : closeOffset;
                let styled = `${xml.slice(0, offset)}${style}${xml.slice(offset)}`;
                if (paragraphColor) {
                  const paragraphOffset = styled.lastIndexOf('<a:p>');
                  if (paragraphOffset >= 0) {
                    const paragraph = `<a:p><a:pPr><a:defRPr><a:solidFill><a:srgbClr val="${paragraphColor.slice(1)}"/></a:solidFill></a:defRPr></a:pPr>`;
                    styled = `${styled.slice(0, paragraphOffset)}${paragraph}${styled.slice(paragraphOffset + '<a:p>'.length)}`;
                  }
                }
                return encoder.encode(styled);
              })(),
            }
          : entry,
      ),
    ),
  );
};

describe('shape style-matrix fill references', () => {
  it('renders the native Colored Fill outline from its theme line reference', async () => {
    const pres = await styleDeck();
    const svg = renderSlideToSvg(pres, getSlides(pres)[0]!);
    expect(svg).toContain('stroke="#1C334E"');
    expect(svg).toContain('stroke-width="2.67"');
    expect(svg).toContain('stroke-linecap="butt"');
  });

  it.runIf(isSchemaValidationAvailable())(
    'keeps the native style fixture schema-valid',
    async () => {
      const pres = await styleDeck();
      expectSchemaValid(getSlideXmlString(getSlides(pres)[0]!), 'pml');
      const textPres = await styleDeck('1', false, true);
      expectSchemaValid(getSlideXmlString(getSlides(textPres)[0]!), 'pml');
    },
  );

  it('resolves native Quick Styles fillRef through the theme and survives save/load', async () => {
    const pres = await styleDeck();
    const shape = getSlideShapes(getSlides(pres)[0]!).at(-1)!;

    expect(getShapeFill(shape)).toEqual({ kind: 'inherit' });
    expect(getShapeFillEffective(pres, shape)).toEqual({ kind: 'solid', color: '#4F81BD' });
    expect(getShapeFillColorResolved(pres, shape)).toBe('#4F81BD');

    const reloaded = await loadPresentation(await savePresentation(pres));
    const restored = getSlideShapes(getSlides(reloaded)[0]!).at(-1)!;
    expect(getShapeFillEffective(reloaded, restored)).toEqual({ kind: 'solid', color: '#4F81BD' });
  });

  it('resolves native Quick Styles fontRef color and preserves direct overrides', async () => {
    const pres = await styleDeck('1', false, true);
    const shape = getSlideShapes(getSlides(pres)[0]!).at(-1)!;

    expect(getShapeRunFormatEffective(pres, shape, 0, 0)).toMatchObject({
      color: '#FFFFFF',
      font: 'Calibri',
    });
    const nativeReload = await loadPresentation(await savePresentation(pres));
    const nativeRestored = getSlideShapes(getSlides(nativeReload)[0]!).at(-1)!;
    expect(getShapeRunFormatEffective(nativeReload, nativeRestored, 0, 0).color).toBe('#FFFFFF');
    const svg = renderSlideToSvg(nativeReload, getSlides(nativeReload)[0]!);
    expect(svg).toContain('Styled text');
    expect(svg).toContain('color:#FFFFFF">Styled text');

    const paragraphPres = await styleDeck('1', false, true, '#ABCDEF');
    const paragraphShape = getSlideShapes(getSlides(paragraphPres)[0]!).at(-1)!;
    expect(getShapeRunFormatEffective(paragraphPres, paragraphShape, 0, 0).color).toBe('#ABCDEF');

    const paragraphColor = getShapeRunFormatEffective(paragraphPres, paragraphShape, 0, 0);
    expect(paragraphColor.font).toBe('Calibri');

    setShapeRunFormat(shape, 0, 0, { color: '#123456' });
    expect(getShapeRunFormatEffective(pres, shape, 0, 0).color).toBe('#123456');

    const reloaded = await loadPresentation(await savePresentation(pres));
    const restored = getSlideShapes(getSlides(reloaded)[0]!).at(-1)!;
    expect(getShapeRunFormatEffective(reloaded, restored, 0, 0).color).toBe('#123456');
  });

  it('keeps a direct shape fill ahead of its style-matrix fillRef', async () => {
    const pres = await styleDeck();
    const shape = getSlideShapes(getSlides(pres)[0]!).at(-1)!;
    setShapeFill(shape, '#123456');
    expect(getShapeFillEffective(pres, shape)).toEqual({ kind: 'solid', color: '#123456' });
    expect(getShapeFillColorResolved(pres, shape)).toBe('#123456');
  });

  it('resolves background fill style references at index 1001', async () => {
    const pres = await styleDeck('1001');
    const shape = getSlideShapes(getSlides(pres)[0]!).at(-1)!;
    expect(getShapeFillEffective(pres, shape)).toEqual({ kind: 'solid', color: '#4F81BD' });
  });

  it.each(['0', '1000'])('treats fillRef index %s as no fill', async (styleIndex) => {
    const pres = await styleDeck(styleIndex);
    const shape = getSlideShapes(getSlides(pres)[0]!).at(-1)!;
    expect(getShapeFillEffective(pres, shape)).toEqual({ kind: 'none' });
  });

  it('uses style fill alpha for preview while direct fill alpha takes precedence', async () => {
    const pres = await styleDeck('1', true);
    const shape = getSlideShapes(getSlides(pres)[0]!).at(-1)!;
    expect(getShapeFillOpacity(shape, pres)).toBeCloseTo(0.27, 6);
    expect(renderSlideToSvg(pres, getSlides(pres)[0]!)).toContain('fill-opacity="0.270"');

    setShapeFill(shape, '#123456');
    expect(getShapeFillOpacity(shape, pres)).toBeNull();
  });
});
