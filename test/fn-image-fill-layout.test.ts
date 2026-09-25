import { readFile } from 'node:fs/promises';
import { describe, expect, it } from 'vitest';
import { renderSlideToSvg } from '../packages/preview/src/index.ts';
import { SHAPE_ELEMENT } from '../src/api/_internal-symbols.ts';
import { attr, firstChildElement, NS, qname } from '../src/internal/xml/index.ts';
import {
  addSlideShape,
  getMediaParts,
  getShapeImageBytes,
  getShapeImageFillBytes,
  getShapeImageFillLayout,
  getShapeImageOpacity,
  getShapeKind,
  getSlideShapes,
  getSlides,
  getSlideXmlString,
  inches,
  loadPresentation,
  pt,
  savePresentation,
  setShapeImageFill,
  setShapeImageFillLayout,
  setShapeImageOpacity,
  setShapeImageCrop,
  getShapeImageCrop,
} from '../src/api/index.ts';
import { expectSchemaValid, isSchemaValidationAvailable } from './lib/expect-schema-valid.ts';

async function fixture(kind: 'shape' | 'picture') {
  const pres = await loadPresentation(
    await readFile(new URL('./fixtures/minimal/one-image-slide.pptx', import.meta.url)),
  );
  const slide = getSlides(pres)[0]!;
  const picture = getSlideShapes(slide).find((shape) => getShapeKind(shape) === 'picture')!;
  const shape =
    kind === 'picture'
      ? picture
      : addSlideShape(slide, {
          preset: 'rect',
          x: inches(1),
          y: inches(1),
          w: inches(2),
          h: inches(2),
        });
  if (kind === 'shape') setShapeImageFill(shape, getShapeImageBytes(picture)!);
  return { pres, slide, shape };
}

describe.each(['shape', 'picture'] as const)('image fill layout on %s', (kind) => {
  it('round-trips stretch offsets and rotation without replacing media or opacity', async () => {
    const { pres, slide, shape } = await fixture(kind);
    const count = getMediaParts(pres).length;
    const bytes = kind === 'picture' ? getShapeImageBytes(shape) : getShapeImageFillBytes(shape);
    setShapeImageOpacity(shape, 0.4);
    const layout = {
      mode: 'stretch' as const,
      left: 0.25,
      right: -0.1,
      top: 0.3,
      bottom: 0.05,
      rotateWithShape: false,
    };
    setShapeImageFillLayout(shape, layout);
    expect(getShapeImageFillLayout(shape)).toEqual(layout);
    expect(getSlideXmlString(slide)).toContain(
      '<a:fillRect l="25000" t="30000" r="-10000" b="5000"/>',
    );
    expect(getMediaParts(pres)).toHaveLength(count);
    expect(kind === 'picture' ? getShapeImageBytes(shape) : getShapeImageFillBytes(shape)).toEqual(
      bytes,
    );
    expect(getShapeImageOpacity(shape)).toBe(0.4);
    const saved = await loadPresentation(await savePresentation(pres));
    const restored = getSlideShapes(getSlides(saved)[0]!).at(-1)!;
    expect(getShapeImageFillLayout(restored)).toEqual(layout);
    expect(getShapeImageOpacity(restored)).toBe(0.4);
  });

  it('switches modes without stacking choices and preserves rotation unless provided', async () => {
    const { slide, shape } = await fixture(kind);
    setShapeImageFillLayout(shape, { mode: 'stretch', rotateWithShape: false });
    const layout = {
      mode: 'tile' as const,
      offsetX: pt(12),
      offsetY: pt(-6),
      scaleX: 0.25,
      scaleY: 0.75,
      alignment: 'br' as const,
      flip: 'xy' as const,
    };
    setShapeImageFillLayout(shape, layout);
    expect(getShapeImageFillLayout(shape)).toEqual({ ...layout, rotateWithShape: false });
    const xml = getSlideXmlString(slide);
    expect(xml).toContain(
      '<a:tile tx="152400" ty="-76200" sx="25000" sy="75000" flip="xy" algn="br"/>',
    );
    setShapeImageFillLayout(shape, { mode: 'stretch' });
    expect(getShapeImageFillLayout(shape)).toEqual({
      mode: 'stretch',
      rotateWithShape: false,
      left: 0,
      top: 0,
      right: 0,
      bottom: 0,
    });
    expect(getSlideXmlString(slide)).not.toContain('<a:tile');
  });

  it('rejects invalid inputs without changing the shape', async () => {
    const { slide, shape } = await fixture(kind);
    const before = getSlideXmlString(slide);
    for (const bad of [NaN, Infinity, -Infinity, 100000]) {
      expect(() => setShapeImageFillLayout(shape, { mode: 'stretch', left: bad })).toThrow(
        RangeError,
      );
      expect(() => setShapeImageFillLayout(shape, { mode: 'tile', scaleY: bad })).toThrow(
        RangeError,
      );
      expect(getSlideXmlString(slide)).toBe(before);
    }
    expect(() => setShapeImageFillLayout(shape, { mode: 'tile', offsetX: pt(1e12) })).toThrow(
      RangeError,
    );
    expect(getSlideXmlString(slide)).toBe(before);
  });

  it.skipIf(!isSchemaValidationAvailable())(
    'emits schema-valid tile and stretch fills',
    async () => {
      const { slide, shape } = await fixture(kind);
      setShapeImageFillLayout(shape, { mode: 'tile', scaleX: 0, alignment: 'ctr', flip: 'x' });
      expectSchemaValid(getSlideXmlString(slide), 'pml');
      setShapeImageFillLayout(shape, {
        mode: 'stretch',
        left: -1000,
        bottom: 1000,
        rotateWithShape: true,
      });
      expectSchemaValid(getSlideXmlString(slide), 'pml');
    },
  );
});

it('preserves picture source crop when switching fill placement', async () => {
  const { shape } = await fixture('picture');
  setShapeImageCrop(shape, { left: 0.1, top: 0.2 });
  const before = getShapeImageCrop(shape);
  setShapeImageFillLayout(shape, { mode: 'tile' });
  expect(getShapeImageCrop(shape)).toEqual(before);
});

it('rejects shapes without image fills without changing their XML', async () => {
  const { slide } = await fixture('shape');
  const shape = addSlideShape(slide, {
    preset: 'rect',
    x: inches(0),
    y: inches(0),
    w: inches(1),
    h: inches(1),
  });
  const before = getSlideXmlString(slide);
  expect(getShapeImageFillLayout(shape)).toBeNull();
  expect(() => setShapeImageFillLayout(shape, { mode: 'tile' })).toThrow(/image fill/);
  expect(getSlideXmlString(slide)).toBe(before);
});

it('renders positive stretch offsets and clips negative offsets to the shape', async () => {
  const { pres, slide, shape } = await fixture('shape');
  setShapeImageFillLayout(shape, { mode: 'stretch', left: 0.25, right: 0.125, top: 0.5 });
  expect(renderSlideToSvg(pres, slide)).toContain(
    '<image x="144.00" y="192.00" width="120.00" height="96.00"',
  );
  setShapeImageFillLayout(shape, { mode: 'stretch', left: -0.5 });
  const svg = renderSlideToSvg(pres, slide);
  expect(svg).toContain('<image x="0.00" y="96.00" width="288.00" height="192.00"');
  expect(svg).toMatch(/<clipPath[^>]*><rect x="96.00" y="96.00" width="192.00" height="192.00"/);
});

it('renders empty destination rectangles without negative SVG sizes', async () => {
  const { pres, slide, shape } = await fixture('shape');
  setShapeImageFillLayout(shape, { mode: 'stretch', left: 0.75, right: 0.5 });
  expect(renderSlideToSvg(pres, slide)).toContain(
    '<image x="240.00" y="96.00" width="0.00" height="192.00"',
  );
});

it('reads percentage and universal-measure attributes in imported tile fills', async () => {
  const { shape } = await fixture('picture');
  setShapeImageFillLayout(shape, { mode: 'tile' });
  const fill = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'blipFill', NS.pml))!;
  const tile = firstChildElement(fill, qname('a', 'tile', NS.dml))!;
  tile.attrs = [
    attr(qname('', 'tx', ''), '1in'),
    attr(qname('', 'ty', ''), '-12pt'),
    attr(qname('', 'sx', ''), '25%'),
  ];
  fill.attrs.push(attr(qname('', 'rotWithShape', ''), ' false '));
  expect(getShapeImageFillLayout(shape)).toEqual({
    mode: 'tile',
    offsetX: inches(1),
    offsetY: pt(-12),
    scaleX: 0.25,
    scaleY: 1,
    alignment: 'tl',
    flip: 'none',
    rotateWithShape: false,
  });
});
