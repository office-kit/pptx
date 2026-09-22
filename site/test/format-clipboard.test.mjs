// The format painter's clipboard: what it picks up off a shape, what it
// writes onto another, and what it refuses to claim it copied.

import assert from 'node:assert/strict';
import test from 'node:test';
import {
  addBlankSlide,
  addSlideShape,
  addSlideTextBox,
  createPresentation,
  getParagraphPropertiesEffective,
  getShapeEffects,
  getShapeFill,
  getShapeGradientFill,
  getShapeParagraphCount,
  getShapePatternFill,
  getShapeRunFormatEffective,
  getShapeStroke,
  getShapeStrokeDash,
  getSlides,
  getSlideShapes,
  inches,
  loadPresentation,
  savePresentation,
  setParagraphAlignment,
  setParagraphLineSpacing,
  setParagraphSpacing,
  setShapeFill,
  setShapeGlow,
  setShapeGradientFill,
  setShapeImageFill,
  setShapeNoFill,
  setShapePatternFill,
  setShapeShadow,
  setShapeStroke,
  setShapeStrokeDash,
  setShapeTextFormat,
} from '@office-kit/pptx';
import {
  applyShapeFormat,
  applyTextFormat,
  copiedFormatLimits,
  readShapeFormat,
  readTextFormat,
} from '../src/lib/editor/core/format-clipboard.ts';

const deck = () => {
  const pres = createPresentation();
  return { pres, slide: addBlankSlide(pres) };
};

const box = (slide, x, text) =>
  addSlideTextBox(slide, { x: inches(x), y: inches(1), w: inches(3), h: inches(1), text });

// What the controller hands `readShapeFormat`: the first run's effective
// format, or nothing at all for a shape with no text body.
const characterOf = (pres, shape) =>
  getShapeParagraphCount(shape) > 0 ? getShapeRunFormatEffective(pres, shape, 0, 0) : null;

test('a copied format carries paint, character and paragraph formatting', () => {
  const { pres, slide } = deck();
  const source = box(slide, 1, 'Source');
  setShapeFill(source, '#FF0000');
  setShapeStroke(source, { color: '#00FF00', widthEmu: 38100 });
  setShapeStrokeDash(source, 'dash');
  setShapeShadow(source, { color: '#123456', blurEmu: 50800, offsetEmu: 38100, angleDeg: 45 });
  setShapeGlow(source, { color: '#ABCDEF', radiusEmu: 63500 });
  setShapeTextFormat(source, { bold: true, size: 30, color: '#0000FF' });
  setParagraphAlignment(source, 0, 'right');
  setParagraphLineSpacing(source, 0, { kind: 'pct', value: 1.5 });
  setParagraphSpacing(source, 0, { beforePts: 12, afterPts: 6 });

  const target = box(slide, 5, 'Target');
  applyShapeFormat(target, readShapeFormat(pres, source, characterOf(pres, source)));

  assert.deepEqual(getShapeFill(target), { kind: 'solid', color: '#FF0000' });
  assert.deepEqual(getShapeStroke(target), { kind: 'solid', color: '#00FF00', widthEmu: 38100 });
  assert.equal(getShapeStrokeDash(target), 'dash');
  const effects = getShapeEffects(pres, target);
  assert.equal(effects.find((e) => e.kind === 'outerShdw')?.color, '#123456');
  assert.equal(effects.find((e) => e.kind === 'glow')?.radiusEmu, 63500);
  const character = characterOf(pres, target);
  assert.equal(character.bold, true);
  assert.equal(character.size, 30);
  assert.equal(character.color, '#0000FF');
  const paragraph = getParagraphPropertiesEffective(pres, target, 0);
  assert.equal(paragraph.align, 'right');
  assert.deepEqual(paragraph.lineSpacing, { kind: 'pct', value: 1.5 });
  assert.equal(paragraph.spcBefPts, 12);
  assert.equal(paragraph.spcAftPts, 6);
  // The text itself is formatting's passenger, never its cargo.
  assert.equal(getSlideShapes(getSlides(pres)[0]).length, 2);
});

test('gradient and pattern fills survive the round trip', () => {
  const { pres, slide } = deck();
  const gradient = box(slide, 1, 'Gradient');
  setShapeGradientFill(gradient, {
    stops: [
      { offset: 0, color: '#FF0000' },
      { offset: 1, color: '#0000FF' },
    ],
    angleDeg: 45,
  });
  const pattern = box(slide, 5, 'Pattern');
  setShapePatternFill(pattern, {
    preset: 'dkUpDiag',
    foreground: '#112233',
    background: '#FFFFFF',
  });

  const a = box(slide, 1, 'A');
  const b = box(slide, 5, 'B');
  applyShapeFormat(a, readShapeFormat(pres, gradient, characterOf(pres, gradient)));
  applyShapeFormat(b, readShapeFormat(pres, pattern, characterOf(pres, pattern)));

  assert.deepEqual(getShapeGradientFill(a), getShapeGradientFill(gradient));
  assert.deepEqual(getShapePatternFill(pres, b), getShapePatternFill(pres, pattern));
});

test('an explicit no-fill is a format, not an absence', () => {
  const { pres, slide } = deck();
  const source = box(slide, 1, 'Source');
  setShapeNoFill(source);
  const target = box(slide, 5, 'Target');
  setShapeFill(target, '#FF0000');
  applyShapeFormat(target, readShapeFormat(pres, source, characterOf(pres, source)));
  assert.deepEqual(getShapeFill(target), { kind: 'none' });
});

test('pasting a format replaces the target’s effects rather than adding to them', () => {
  const { pres, slide } = deck();
  const source = box(slide, 1, 'Source');
  setShapeShadow(source, { color: '#000000' });
  const target = box(slide, 5, 'Target');
  setShapeGlow(target, { color: '#FF00FF', radiusEmu: 63500 });
  applyShapeFormat(target, readShapeFormat(pres, source, characterOf(pres, source)));
  const kinds = getShapeEffects(pres, target).map((effect) => effect.kind);
  assert.deepEqual(kinds, ['outerShdw']);
});

test('what cannot be copied is reported instead of silently dropped', () => {
  const { pres, slide } = deck();
  const plain = box(slide, 1, 'Plain');
  assert.deepEqual(copiedFormatLimits(pres, plain), []);

  const picture = addSlideShape(slide, {
    preset: 'rect',
    x: inches(5),
    y: inches(1),
    w: inches(2),
    h: inches(2),
  });
  // A 1×1 transparent GIF is enough to make the fill a picture fill.
  setShapeImageFill(
    picture,
    Uint8Array.from(atob('R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7'), (c) =>
      c.charCodeAt(0),
    ),
  );
  assert.deepEqual(copiedFormatLimits(pres, picture), ['picture fill']);

  // …and the paste leaves the target's own fill alone rather than clearing it.
  const target = box(slide, 8, 'Target');
  setShapeFill(target, '#00FF00');
  applyShapeFormat(target, readShapeFormat(pres, picture, characterOf(pres, picture)));
  assert.deepEqual(getShapeFill(target), { kind: 'solid', color: '#00FF00' });
});

test('a text format repaints one range and the paragraphs it touches', () => {
  const { pres, slide } = deck();
  const source = box(slide, 1, 'Source');
  setShapeTextFormat(source, { bold: true, size: 28 });
  setParagraphAlignment(source, 0, 'center');

  const target = addSlideTextBox(slide, {
    x: inches(5),
    y: inches(1),
    w: inches(3),
    h: inches(2),
    text: 'first\nsecond',
  });
  const format = readTextFormat(pres, source, 0, characterOf(pres, source));
  applyTextFormat(target, { start: 0, end: 5 }, [0], format, false);

  assert.equal(getShapeRunFormatEffective(pres, target, 0, 0).bold, true);
  assert.equal(getParagraphPropertiesEffective(pres, target, 0).align, 'center');
  // The untouched paragraph keeps its own alignment.
  assert.notEqual(getParagraphPropertiesEffective(pres, target, 1).align, 'center');
  // A copied text format has no paint to give.
  assert.equal(format.paint, null);
});

test('a pasted format survives the save/load round trip', async () => {
  const { pres, slide } = deck();
  const source = box(slide, 1, 'Source');
  setShapeFill(source, '#FF0000');
  setShapeTextFormat(source, { bold: true, size: 30 });
  setParagraphAlignment(source, 0, 'right');
  const target = box(slide, 5, 'Target');
  applyShapeFormat(target, readShapeFormat(pres, source, characterOf(pres, source)));

  const reloaded = await loadPresentation(await savePresentation(pres));
  const saved = getSlideShapes(getSlides(reloaded)[0])[1];
  assert.deepEqual(getShapeFill(saved), { kind: 'solid', color: '#FF0000' });
  assert.equal(getShapeRunFormatEffective(reloaded, saved, 0, 0).size, 30);
  assert.equal(getParagraphPropertiesEffective(reloaded, saved, 0).align, 'right');
});
