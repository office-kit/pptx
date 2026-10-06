// Character-level inner shadow: native WordArt fixture read and TextFormat mutation.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideShape,
  getShapeRunFormat,
  getShapeRunFormatEffective,
  getSlideXmlString,
  getSlideShapes,
  getSlides,
  inches,
  loadPresentation,
  savePresentation,
  setShapeText,
  setShapeTextFormat,
  type TextFormat,
} from '../src/api/index.ts';
import { parseRPrLikeElement } from '../src/api/fn/shape-color.ts';
import { toWritableTextFormat } from '../src/internal/drawingml/text-format.ts';
import { SHAPE_ELEMENT } from '../src/api/_internal-symbols.ts';
import {
  NS,
  firstChildElement,
  parseXml,
  qname,
  type XmlElement,
} from '../src/internal/xml/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/native/${name}`, import.meta.url));

const firstRunProperties = (xml: string) => {
  const root = parseXml(xml).root;
  const visit = (node: XmlElement): XmlElement | null => {
    if (node.name.localName === 'rPr') return node;
    for (const child of node.children) {
      if (child.kind !== 'element') continue;
      const result = visit(child);
      if (result) return result;
    }
    return null;
  };
  return visit(root)!;
};

describe('fn API: character inner shadow', () => {
  it('reads the native inner-shadow fixture without confusing it with outer shadow', async () => {
    const xml = await readFile(fixture('wordart-background2-inner-shadow-shape.xml'), 'utf8');
    const format = parseRPrLikeElement(firstRunProperties(xml));

    expect(format.innerShadow).toMatchObject({
      blurEmu: 63500,
      offsetEmu: 50800,
      angleDeg: 225,
      color: '#000000',
    });
    expect(format.shadow).toBeUndefined();
  });

  it('writes and round-trips an inner shadow on character properties', async () => {
    const pres = await loadPresentation(
      await readFile(fileURLToPath(new URL('./fixtures/minimal/two-slides.pptx', import.meta.url))),
    );
    const slide = getSlides(pres)[0]!;
    const shape = addSlideShape(slide, {
      preset: 'rect',
      x: inches(0),
      y: inches(0),
      w: inches(2),
      h: inches(1),
    });
    setShapeText(shape, 'WordArt');
    setShapeTextFormat(shape, {
      innerShadow: { color: '#000000', blurEmu: 63500, offsetEmu: 50800, angleDeg: 225 },
    } satisfies TextFormat);

    const xml = getSlideXmlString(slide);
    expect(xml).toMatch(/<a:rPr[^>]*>[\s\S]*<a:effectLst>[\s\S]*<a:innerShdw\b/);
    expect(xml).toContain('blurRad="63500"');
    expect(xml).toContain('dist="50800"');
    expect(xml).toContain('dir="13500000"');

    const restored = await loadPresentation(await savePresentation(pres));
    const restoredShape = getSlideShapes(getSlides(restored)[0]!).at(-1)!;
    expect(getShapeRunFormat(restoredShape, 0, 0)).toMatchObject({
      innerShadow: { blurEmu: 63500, offsetEmu: 50800, angleDeg: 225 },
    });
  });

  it('removes only inner shadow and preserves an independent outer shadow', async () => {
    const pres = await loadPresentation(
      await readFile(fileURLToPath(new URL('./fixtures/minimal/two-slides.pptx', import.meta.url))),
    );
    const shape = addSlideShape(getSlides(pres)[0]!, {
      preset: 'rect',
      x: inches(0),
      y: inches(0),
      w: inches(2),
      h: inches(1),
    });
    setShapeText(shape, 'WordArt');
    setShapeTextFormat(shape, {
      shadow: { color: '#112233', blurEmu: 1016, offsetEmu: 2032, angleDeg: 45 },
      innerShadow: { color: '#000000', blurEmu: 63500, offsetEmu: 50800, angleDeg: 225 },
    } satisfies TextFormat);
    setShapeTextFormat(shape, { innerShadow: null });

    const xml = getSlideXmlString(getSlides(pres)[0]!);
    expect(xml).not.toContain('<a:innerShdw');
    expect(xml).toContain('<a:outerShdw');

    setShapeTextFormat(shape, {
      innerShadow: { color: '#000000', blurEmu: 63500, offsetEmu: 50800, angleDeg: 225 },
    });
    setShapeTextFormat(shape, {}, { reset: true });
    expect(getShapeRunFormat(shape, 0, 0)?.innerShadow).toBeUndefined();
    expect(getShapeRunFormatEffective(pres, shape, 0, 0)?.innerShadow).toBeUndefined();
  });

  it('rejects invalid inner-shadow geometry and opacity', async () => {
    const pres = await loadPresentation(
      await readFile(fileURLToPath(new URL('./fixtures/minimal/two-slides.pptx', import.meta.url))),
    );
    const shape = addSlideShape(getSlides(pres)[0]!, {
      preset: 'rect',
      x: inches(0),
      y: inches(0),
      w: inches(2),
      h: inches(1),
    });
    setShapeText(shape, 'WordArt');

    expect(() => setShapeTextFormat(shape, { innerShadow: { blurEmu: Number.NaN } })).toThrow(
      RangeError,
    );
    expect(() =>
      setShapeTextFormat(shape, { innerShadow: { angleDeg: Number.POSITIVE_INFINITY } }),
    ).toThrow(RangeError);
    expect(() => setShapeTextFormat(shape, { innerShadow: { opacity: 1.01 } })).toThrow(RangeError);
  });

  it('keeps null removal intent and inherits inner shadow from text-body defaults', async () => {
    expect(toWritableTextFormat({ innerShadow: null })).toEqual({ innerShadow: null });

    const pres = await loadPresentation(
      await readFile(fileURLToPath(new URL('./fixtures/minimal/two-slides.pptx', import.meta.url))),
    );
    const shape = addSlideShape(getSlides(pres)[0]!, {
      preset: 'rect',
      x: inches(0),
      y: inches(0),
      w: inches(2),
      h: inches(1),
    });
    setShapeText(shape, 'WordArt');
    const txBody = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'txBody', NS.pml));
    const lstStyle = txBody && firstChildElement(txBody, qname('a', 'lstStyle', NS.dml));
    if (!lstStyle) throw new Error('text body list style missing');
    lstStyle.children.push(
      parseXml(
        '<a:defPPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:defRPr><a:effectLst><a:innerShdw blurRad="63500" dist="50800" dir="13500000"><a:srgbClr val="000000"/></a:innerShdw></a:effectLst></a:defRPr></a:defPPr>',
      ).root,
    );
    const effective = getShapeRunFormatEffective(pres, shape, 0, 0);
    expect(effective.innerShadow).toMatchObject({
      blurEmu: 63500,
      offsetEmu: 50800,
      angleDeg: 225,
    });
  });
});
