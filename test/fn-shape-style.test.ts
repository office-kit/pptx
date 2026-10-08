import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideShape,
  getShapeBounds,
  getShapeText,
  getShapeRunFormat,
  getSlideShapes,
  getSlides,
  inches,
  loadPresentation,
  savePresentation,
  setShapeFill,
  setShapeRunFormat,
  setShapeShadow,
  setShapeSlideBackgroundFill,
  setShapeStyle,
  setShapeStroke,
} from '../src/api/index.ts';
import { readZip, writeZip } from '../src/internal/opc/index.ts';
import { expectSchemaValid, isSchemaValidationAvailable } from './lib/expect-schema-valid.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));
const nativeFixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/native/${name}`, import.meta.url));

const xmlOf = async (bytes: Uint8Array): Promise<string> => {
  const entry = readZip(bytes).entries.find((item) => item.name === 'ppt/slides/slide1.xml');
  if (!entry) throw new Error('slide1.xml missing');
  return new TextDecoder().decode(entry.data);
};

const style = {
  line: {
    idx: 2,
    color: 'accent1' as const,
    colorTransforms: [{ kind: 'shade' as const, value: 0.15 }],
  },
  fill: { idx: 1, color: 'accent1' as const },
  effect: { idx: 0, color: 'accent1' as const },
  font: { idx: 'minor' as const, color: 'lt1' as const },
};

describe('fn API: setShapeStyle', () => {
  it('writes native refs and preserves geometry and text across a round trip', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const shape = addSlideShape(getSlides(pres)[0]!, {
      preset: 'rect',
      x: inches(1),
      y: inches(2),
      w: inches(3),
      h: inches(1),
      text: 'Keep this text',
    });
    const bounds = getShapeBounds(shape);
    setShapeRunFormat(shape, 0, 0, { bold: true, italic: true, color: '#123456' });
    const runFormat = getShapeRunFormat(shape, 0, 0);
    setShapeFill(shape, '#FF0000');
    setShapeStroke(shape, { color: '#000000', widthEmu: 12700 });
    setShapeStyle(shape, style);
    const xml = await xmlOf(await savePresentation(pres));
    expect(xml).toContain(
      '<p:style><a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="15000"/></a:schemeClr></a:lnRef>',
    );
    expect(xml).toContain('<a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef>');
    expect(xml).toContain('<a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>');
    expect(xml).toContain('<a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style>');
    expect(xml).not.toContain('<a:solidFill><a:srgbClr val="FF0000"');

    const reopened = await loadPresentation(await savePresentation(pres));
    const reopenedShape = getSlideShapes(getSlides(reopened)[0]!).at(-1)!;
    expect(getShapeText(reopenedShape)).toBe('Keep this text');
    expect(getShapeBounds(reopenedShape)).toEqual(bounds);
    expect(getShapeRunFormat(reopenedShape, 0, 0)).toEqual(runFormat);
  });

  it('matches the reference and rejects an out-of-range index atomically', async () => {
    const native = await readFile(
      nativeFixture('quickstyle-transparent-accent1-shape.xml'),
      'utf8',
    );
    const nativeRefs = [
      ...native.matchAll(/<(?:a:)?(lnRef|fillRef|effectRef|fontRef) idx="([^"]+)"/g),
    ].map((match) => `${match[1]}:${match[2]}`);

    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const shape = addSlideShape(getSlides(pres)[0]!, {
      preset: 'rect',
      x: inches(0),
      y: inches(0),
      w: inches(1),
      h: inches(1),
      text: 'Atomic',
    });
    setShapeStyle(shape, {
      line: { idx: 0, color: '#000000' },
      fill: { idx: 0, color: '#000000' },
      effect: { idx: 0, color: '#000000' },
      font: { idx: 'minor', color: 'accent1' },
    });
    const before = await xmlOf(await savePresentation(pres));
    const generatedRefs = [
      ...before.matchAll(/<(?:a:)?(lnRef|fillRef|effectRef|fontRef) idx="([^"]+)"/g),
    ].map((match) => `${match[1]}:${match[2]}`);
    expect(generatedRefs).toEqual(nativeRefs);
    const bounds = getShapeBounds(shape);
    const text = getShapeText(shape);
    const runFormat = getShapeRunFormat(shape, 0, 0);
    expect(() =>
      setShapeStyle(shape, {
        ...style,
        effect: { ...style.effect, idx: 0x1_0000_0000 },
      }),
    ).toThrow(/invalid effectRef index/);
    expect(await xmlOf(await savePresentation(pres))).toBe(before);
    expect(getShapeBounds(shape)).toEqual(bounds);
    expect(getShapeText(shape)).toBe(text);
    expect(getShapeRunFormat(shape, 0, 0)).toEqual(runFormat);
  });

  it('removes direct effect lists and emits schema-valid shape XML', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const shape = addSlideShape(getSlides(pres)[0]!, {
      preset: 'rect',
      x: inches(0),
      y: inches(0),
      w: inches(1),
      h: inches(1),
      text: 'Effects',
    });
    setShapeShadow(shape);
    setShapeStyle(shape, { ...style, font: { idx: 'none' } });
    const saved = await savePresentation(pres);
    const xml = await xmlOf(saved);
    expect(xml).not.toContain('<a:effectLst>');
    if (isSchemaValidationAvailable()) expectSchemaValid(xml, 'pml');

    // Preserve the same setter behavior when a legacy producer leaves an
    // effectDag in spPr, even though the reference desktop app itself normally writes effectLst.
    const patched = readZip(saved).entries.map((entry) => {
      if (entry.name !== 'ppt/slides/slide1.xml') return entry;
      const source = new TextDecoder().decode(entry.data);
      const end = source.lastIndexOf('</p:spPr>');
      return {
        ...entry,
        data: new TextEncoder().encode(`${source.slice(0, end)}<a:effectDag/>${source.slice(end)}`),
      };
    });
    const legacy = await loadPresentation(writeZip(patched));
    const legacyShape = getSlideShapes(getSlides(legacy)[0]!).at(-1)!;
    setShapeStyle(legacyShape, { ...style, font: { idx: 'none' } });
    expect(await xmlOf(await savePresentation(legacy))).not.toContain('<a:effectDag/>');
  });

  it('rejects invalid theme reference indices before mutating the shape', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const shape = addSlideShape(getSlides(pres)[0]!, {
      preset: 'rect',
      x: inches(0),
      y: inches(0),
      w: inches(1),
      h: inches(1),
    });
    for (const idx of [-1, 1.5, Number.NaN]) {
      expect(() => setShapeStyle(shape, { ...style, line: { ...style.line, idx } })).toThrow(
        /invalid lnRef index/,
      );
    }
  });

  it('accepts the uint32 boundary and clears a prior slide-background fill', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const shape = addSlideShape(getSlides(pres)[0]!, {
      preset: 'rect',
      x: inches(0),
      y: inches(0),
      w: inches(1),
      h: inches(1),
      text: 'Boundary',
    });
    setShapeSlideBackgroundFill(shape);
    expect(() =>
      setShapeStyle(shape, {
        ...style,
        line: { ...style.line, idx: 0xffff_ffff },
      }),
    ).not.toThrow();
    const xml = await xmlOf(await savePresentation(pres));
    expect(xml).toContain('<a:lnRef idx="4294967295">');
    expect(xml).not.toContain('useBgFill="1"');
  });
});
