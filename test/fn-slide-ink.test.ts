import { expect, it } from 'vitest';
import * as pptx from '../src/api/index.ts';
import { appendSlideInk } from '../src/api/fn/slide-ink.ts';
import { INTERNAL_PACKAGE, SLIDE_DOCUMENT, SLIDE_PART_NAME } from '../src/api/_internal-symbols.ts';
import {
  NS,
  allChildElements,
  firstChildElement,
  getAttrValue,
  qname,
  serializeXml,
} from '../src/internal/xml/index.ts';
import { partName, resolveTarget } from '../src/internal/opc/index.ts';
import { REL_TYPES } from '../src/internal/presentationml/relationship-types.ts';
import { inkTransform, inkFallbackPicture } from '../src/internal/drawingml/ink-content.ts';
import { SHAPE_ELEMENT } from '../src/api/_internal-symbols.ts';
import { renderSlideToRgba } from '../packages/preview/src/node.ts';
import { buildPng } from './lib/build-png.ts';

const png = Uint8Array.from(
  Buffer.from(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jRZkAAAAASUVORK5CYII=',
    'base64',
  ),
);
const strokes = [
  {
    points: [
      { x: 360000, y: 720000 },
      { x: 720000, y: 360000 },
    ],
    widthEmu: 3600,
    color: '#FF0000',
  },
];

it('packages native ink and its picture fallback with stable relations after ZIP round trip', async () => {
  const deck = pptx.createPresentation();
  const slide = pptx.addBlankSlide(deck);
  expect(appendSlideInk(slide, strokes, png)).toBe(2);
  expect(appendSlideInk(slide, strokes, png)).toBe(3);
  const text = pptx.addSlideTextBox(slide, {
    x: pptx.emu(0),
    y: pptx.emu(0),
    w: pptx.emu(100000),
    h: pptx.emu(100000),
    text: 'After ink',
  });
  expect(pptx.getShapeId(text)).toBe(4);
  const loaded = await pptx.loadPresentation(await pptx.savePresentation(deck));
  const restored = pptx.getSlides(loaded)[0]!;
  expect(serializeXml(restored[SLIDE_DOCUMENT])).toBe(serializeXml(slide[SLIDE_DOCUMENT]));
  const pkg = restored[INTERNAL_PACKAGE];
  const rels = pkg.getRels(restored[SLIDE_PART_NAME])!.items;
  const custom = rels.filter((rel) => rel.type === REL_TYPES.customXml);
  const images = rels.filter((rel) => rel.type === REL_TYPES.image);
  expect(custom).toHaveLength(2);
  expect(images).toHaveLength(2);
  for (const rel of custom) {
    const part = pkg.getPart(resolveTarget(restored[SLIDE_PART_NAME], rel.target))!;
    expect(part.contentType).toBe('application/inkml+xml');
    expect(new TextDecoder().decode(part.data)).toContain('brushRef="#br0"');
  }
  for (const rel of images)
    expect(pkg.getPart(resolveTarget(restored[SLIDE_PART_NAME], rel.target))!.data).toEqual(png);
  const tree = firstChildElement(
    firstChildElement(restored[SLIDE_DOCUMENT].root, qname('p', 'cSld', NS.pml))!,
    qname('p', 'spTree', NS.pml),
  )!;
  const wrappers = allChildElements(tree, qname('mc', 'AlternateContent', NS.mc));
  expect(wrappers).toHaveLength(2);
  wrappers.forEach((wrapper, index) => {
    const choice = firstChildElement(wrapper, qname('mc', 'Choice', NS.mc))!;
    expect(choice.prefixDecls.get('p14')).toBe(NS.p14);
    expect(getAttrValue(choice, qname('', 'Requires', ''))).toBe('p14');
    const content = firstChildElement(choice, qname('p', 'contentPart', NS.pml))!;
    expect(getAttrValue(content, qname('r', 'id', NS.officeDocRels))).toBe(custom[index]!.id);
    const picture = firstChildElement(
      firstChildElement(wrapper, qname('mc', 'Fallback', NS.mc))!,
      qname('p', 'pic', NS.pml),
    )!;
    const fill = firstChildElement(picture, qname('p', 'blipFill', NS.pml))!;
    expect(
      getAttrValue(
        firstChildElement(fill, qname('a', 'blip', NS.dml))!,
        qname('r', 'embed', NS.officeDocRels),
      ),
    ).toBe(images[index]!.id);
    const nativeTransform = firstChildElement(content, qname('p14', 'xfrm', NS.p14))!;
    const fallbackTransform = firstChildElement(
      firstChildElement(picture, qname('p', 'spPr', NS.pml))!,
      qname('a', 'xfrm', NS.dml),
    )!;
    expect(nativeTransform.children).toEqual(fallbackTransform.children);
  });
  expect(pkg.getPart(partName('/ppt/ink/ink1.xml'))).toBeDefined();
});

it('does not change the package when ink or fallback validation fails', async () => {
  const deck = pptx.createPresentation();
  const slide = pptx.addBlankSlide(deck);
  const before = await pptx.savePresentation(deck);
  expect(() => appendSlideInk(slide, [], png)).toThrow();
  expect(() => appendSlideInk(slide, strokes, new Uint8Array())).toThrow();
  expect(await pptx.savePresentation(deck)).toEqual(before);
});

it('reads and transforms ink as one shape, keeping native and fallback geometry together', async () => {
  const deck = pptx.createPresentation();
  const slide = pptx.addBlankSlide(deck);
  appendSlideInk(slide, strokes, png);
  const shape = pptx.getSlideShapes(slide)[0]!;
  expect(pptx.getShapeKind(shape)).toBe('ink');
  expect(pptx.getShapeImageBytes(shape)).toEqual(png);
  expect(pptx.getShapePosition(shape)).toEqual({ x: 358200, y: 358200 });
  pptx.setShapePosition(shape, pptx.inches(1), pptx.inches(2));
  pptx.setShapeSize(shape, pptx.inches(3), pptx.inches(4));
  pptx.setShapeRotation(shape, 25);
  expect(pptx.getShapeRotation(shape)).toBe(25);
  const native = inkTransform(shape[SHAPE_ELEMENT])!;
  const fallback = inkFallbackPicture(shape[SHAPE_ELEMENT])!;
  const transform = firstChildElement(
    firstChildElement(fallback, qname('p', 'spPr', NS.pml))!,
    qname('a', 'xfrm', NS.dml),
  )!;
  expect(transform.attrs).toEqual(native.attrs);
  expect(transform.children).toEqual(native.children);
  const loaded = await pptx.loadPresentation(await pptx.savePresentation(deck));
  const restored = pptx.getSlideShapes(pptx.getSlides(loaded)[0]!)[0]!;
  expect(pptx.getShapeBounds(restored)).toEqual(pptx.getShapeBounds(shape));
  expect(pptx.getShapeRotation(restored)).toBe(25);
});

it('renders ink at its bounds and respects stacking, copy and deletion', () => {
  const deck = pptx.createPresentation();
  const slide = pptx.addBlankSlide(deck);
  const blue = buildPng(2, 2, [0, 0, 255]);
  appendSlideInk(slide, strokes, blue);
  const ink = pptx.getSlideShapes(slide)[0]!;
  pptx.setShapePosition(ink, pptx.inches(1), pptx.inches(1));
  pptx.setShapeSize(ink, pptx.inches(2), pptx.inches(2));
  const front = pptx.addSlideShape(slide, {
    preset: 'rect',
    x: pptx.inches(1),
    y: pptx.inches(1),
    w: pptx.inches(2),
    h: pptx.inches(2),
  });
  pptx.setShapeFill(front, '#FF0000');
  const color = () => {
    const { image } = renderSlideToRgba(deck, slide);
    const offset = (192 * image.width + 192) * 4;
    return Array.from(image.data.slice(offset, offset + 3));
  };
  expect(color()).toEqual([255, 0, 0]);
  pptx.sendShapeToBack(front);
  expect(pptx.getSlideShapes(slide).map(pptx.getShapeKind)).toEqual(['shape', 'ink']);
  expect(color()).toEqual([0, 0, 255]);
  const copy = pptx.copyShape(slide, pptx.getSlideShapes(slide)[1]!);
  expect(pptx.getShapeKind(copy)).toBe('ink');
  expect(pptx.getShapeId(copy)).toBe(4);
  expect(pptx.getShapeImageBytes(copy)).toEqual(blue);
  const fallback = inkFallbackPicture(copy[SHAPE_ELEMENT])!;
  const nv = firstChildElement(fallback, qname('p', 'nvPicPr', NS.pml))!;
  expect(
    getAttrValue(firstChildElement(nv, qname('p', 'cNvPr', NS.pml))!, qname('', 'id', '')),
  ).toBe('4');
  pptx.removeShape(copy);
  expect(pptx.getSlideShapes(slide)).toHaveLength(2);
});

it('keeps ink fallback names, accessibility and visibility in sync after editing and reload', async () => {
  const deck = pptx.createPresentation();
  const slide = pptx.addBlankSlide(deck);
  const ink = pptx.addSlideInk(slide, strokes, png);
  pptx.renameShape(ink, 'Presenter annotation');
  pptx.setShapeDescription(ink, 'Underline the important result');
  pptx.setShapeAltTitle(ink, 'Emphasis');
  pptx.setShapeHidden(ink, true);
  const fallbackMetadata = (shape: pptx.SlideShapeData) => {
    const picture = inkFallbackPicture(shape[SHAPE_ELEMENT])!;
    return firstChildElement(
      firstChildElement(picture, qname('p', 'nvPicPr', NS.pml))!,
      qname('p', 'cNvPr', NS.pml),
    )!;
  };
  const loaded = await pptx.loadPresentation(await pptx.savePresentation(deck));
  const restored = pptx.getSlideShapes(pptx.getSlides(loaded)[0]!)[0]!;
  for (const [name, value] of Object.entries({
    name: 'Presenter annotation',
    descr: 'Underline the important result',
    title: 'Emphasis',
    hidden: '1',
  }))
    expect(getAttrValue(fallbackMetadata(restored), qname('', name, ''))).toBe(value);
  pptx.setShapeHidden(restored, false);
  pptx.setShapeDescription(restored, null);
  pptx.setShapeAltTitle(restored, null);
  for (const name of ['hidden', 'descr', 'title'])
    expect(getAttrValue(fallbackMetadata(restored), qname('', name, ''))).toBeNull();
});

it('retains native ink aspect locks and mirrors click/hover actions to the fallback', async () => {
  const deck = pptx.createPresentation();
  const slide = pptx.addBlankSlide(deck);
  const ink = pptx.addSlideInk(slide, strokes, png);
  expect(pptx.getShapeAspectRatioLocked(ink)).toBe(false);
  const bounds = pptx.getShapeSize(ink);
  pptx.setShapeAspectRatioLocked(ink, true);
  pptx.setShapeClickAction(ink, { kind: 'url', url: 'https://example.com/ink' }, 'Ink link');
  pptx.setShapeHoverAction(ink, { kind: 'nextSlide' });
  const loaded = await pptx.loadPresentation(await pptx.savePresentation(deck));
  const restored = pptx.getSlideShapes(pptx.getSlides(loaded)[0]!)[0]!;
  expect(pptx.getShapeAspectRatioLocked(restored)).toBe(true);
  expect(pptx.getShapeSize(restored)).toEqual(bounds);
  const fallback = inkFallbackPicture(restored[SHAPE_ELEMENT])!;
  const nv = firstChildElement(fallback, qname('p', 'nvPicPr', NS.pml))!;
  const properties = firstChildElement(nv, qname('p', 'cNvPicPr', NS.pml))!;
  const locks = firstChildElement(properties, qname('a', 'picLocks', NS.dml))!;
  expect(getAttrValue(locks, qname('', 'noChangeAspect', ''))).toBe('1');
  const metadata = firstChildElement(nv, qname('p', 'cNvPr', NS.pml))!;
  const link = firstChildElement(metadata, qname('a', 'hlinkClick', NS.dml))!;
  const relId = getAttrValue(link, qname('r', 'id', NS.officeDocRels));
  const restoredSlide = pptx.getSlides(loaded)[0]!;
  expect(
    restoredSlide[INTERNAL_PACKAGE]
      .getRels(restoredSlide[SLIDE_PART_NAME])!
      .items.find((rel) => rel.id === relId)?.target,
  ).toBe('https://example.com/ink');
  expect(firstChildElement(metadata, qname('a', 'hlinkHover', NS.dml))).not.toBeNull();
  pptx.setShapeAspectRatioLocked(restored, false);
  pptx.setShapeClickAction(restored, null);
  pptx.setShapeHoverAction(restored, null);
  expect(pptx.getShapeAspectRatioLocked(restored)).toBe(false);
  expect(getAttrValue(locks, qname('', 'noChangeAspect', ''))).toBe('0');
  expect(firstChildElement(metadata, qname('a', 'hlinkClick', NS.dml))).toBeNull();
  expect(firstChildElement(metadata, qname('a', 'hlinkHover', NS.dml))).toBeNull();
});
