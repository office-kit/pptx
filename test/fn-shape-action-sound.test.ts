import { expect, it } from 'vitest';
import * as pptx from '../src/api/index.ts';

const wav = (sample: number) => {
  const bytes = new Uint8Array(46);
  const view = new DataView(bytes.buffer);
  for (const [offset, text] of [
    [0, 'RIFF'],
    [8, 'WAVE'],
    [12, 'fmt '],
    [36, 'data'],
  ] as const)
    bytes.set(new TextEncoder().encode(text), offset);
  view.setUint32(4, 38, true);
  view.setUint32(16, 16, true);
  view.setUint16(20, 1, true);
  view.setUint16(22, 1, true);
  view.setUint32(24, 8000, true);
  view.setUint32(28, 16000, true);
  view.setUint16(32, 2, true);
  view.setUint16(34, 16, true);
  view.setUint32(40, 2, true);
  view.setInt16(44, sample, true);
  return bytes;
};
const fixture = () => {
  const p = pptx.createPresentation();
  const slide = pptx.addBlankSlide(p);
  const shape = pptx.addSlideShape(slide, {
    preset: 'rect',
    x: pptx.inches(1),
    y: pptx.inches(1),
    w: pptx.inches(2),
    h: pptx.inches(1),
  });
  return { p, slide, shape };
};

it('round-trips independent trigger sounds and links with valid package relationships', async () => {
  const { p, shape } = fixture();
  const click = { sound: { name: 'Click chime', bytes: wav(10) }, stopPrevious: true };
  const hover = { sound: { name: 'Hover chime', bytes: wav(20) }, stopPrevious: false };
  pptx.setShapeClickAction(shape, { kind: 'nextSlide' });
  pptx.setShapeActionSound(shape, 'click', click);
  pptx.setShapeActionSound(shape, 'hover', hover);
  pptx.setShapeClickAction(shape, { kind: 'lastSlide' });
  const loaded = await pptx.loadPresentation(await pptx.savePresentation(p));
  const copy = pptx.getSlideShapes(pptx.getSlides(loaded)[0]!)[0]!;
  expect(pptx.getShapeActionSound(copy, 'click')).toEqual(click);
  expect(pptx.getShapeActionSound(copy, 'hover')).toEqual(hover);
  expect(pptx.getShapeClickAction(copy)).toEqual({ kind: 'lastSlide' });
  expect(pptx.getShapeHoverAction(copy)).toBeNull();
  expect(pptx.validatePresentation(loaded).filter((issue) => issue.severity === 'error')).toEqual(
    [],
  );
  const result = pptx.getShapeActionSound(copy, 'click');
  result.sound!.bytes[0] = 0;
  expect(pptx.getShapeActionSound(copy, 'click').sound!.bytes[0]).toBe(82);
});

it('shares identical embedded bytes while replacement and removal stay isolated', async () => {
  const { p, slide, shape } = fixture();
  const second = pptx.addSlideShape(slide, {
    preset: 'ellipse',
    x: pptx.inches(4),
    y: pptx.inches(1),
    w: pptx.inches(2),
    h: pptx.inches(1),
  });
  const value = { sound: { name: 'Shared', bytes: wav(10) }, stopPrevious: false };
  pptx.setShapeActionSound(shape, 'click', value);
  const rels = pptx.readPackagePart(p, '/ppt/slides/_rels/slide1.xml.rels');
  pptx.setShapeActionSound(second, 'hover', value);
  expect(pptx.readPackagePart(p, '/ppt/slides/_rels/slide1.xml.rels')).toEqual(rels);
  pptx.setShapeActionSound(shape, 'click', {
    sound: { name: 'Replacement', bytes: wav(30) },
    stopPrevious: false,
  });
  expect(pptx.getShapeActionSound(second, 'hover')).toEqual(value);
  pptx.setShapeActionSound(shape, 'click', { sound: null, stopPrevious: true });
  const loaded = await pptx.loadPresentation(await pptx.savePresentation(p));
  const shapes = pptx.getSlideShapes(pptx.getSlides(loaded)[0]!);
  expect(pptx.getShapeActionSound(shapes[0]!, 'click')).toEqual({
    sound: null,
    stopPrevious: true,
  });
  expect(pptx.getShapeActionSound(shapes[1]!, 'hover')).toEqual(value);
  expect(pptx.validatePresentation(loaded).filter((issue) => issue.severity === 'error')).toEqual(
    [],
  );
});

it('rejects invalid sound input without changing XML or package bytes', async () => {
  const { p, slide, shape } = fixture();
  const before = await pptx.savePresentation(p);
  const xml = pptx.getSlideXmlString(slide);
  expect(() =>
    pptx.setShapeActionSound(shape, 'click', {
      sound: { name: 'Not WAV', bytes: new Uint8Array([1, 2, 3]) },
      stopPrevious: false,
    }),
  ).toThrow('WAV');
  expect(pptx.getSlideXmlString(slide)).toBe(xml);
  expect(await pptx.savePresentation(p)).toEqual(before);
  pptx.setShapeActionSound(shape, 'hover', { sound: null, stopPrevious: false });
  expect(pptx.getSlideXmlString(slide)).toBe(xml);
});

it('keeps action audio available when duplicating its slide', async () => {
  const { p, slide, shape } = fixture();
  const value = { sound: { name: 'Duplicate sound', bytes: wav(42) }, stopPrevious: true };
  pptx.setShapeActionSound(shape, 'hover', value);
  const duplicate = pptx.duplicateSlide(p, slide);
  expect(pptx.getShapeActionSound(pptx.getSlideShapes(duplicate)[0]!, 'hover')).toEqual(value);
  const loaded = await pptx.loadPresentation(await pptx.savePresentation(p));
  expect(
    pptx.getShapeActionSound(pptx.getSlideShapes(pptx.getSlides(loaded)[1]!)[0]!, 'hover'),
  ).toEqual(value);
  expect(pptx.validatePresentation(loaded).filter((issue) => issue.severity === 'error')).toEqual(
    [],
  );
});
