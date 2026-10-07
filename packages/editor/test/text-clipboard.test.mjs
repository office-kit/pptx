import assert from 'node:assert/strict';
import test from 'node:test';
import {
  addBlankSlide,
  addSlideTextBox,
  createPresentation,
  getShapeRunFormatEffective,
  getSlides,
  getSlideShapes,
  inches,
  loadPresentation,
  savePresentation,
  setShapeTextFormat,
} from '@office-kit/pptx';
import { copyTextRange, parseTextClipboard } from '../src/core/text-clipboard.ts';

const clipboard = (format) =>
  JSON.stringify({ version: 1, text: '日本語', formats: [{ start: 0, end: 3, format }] });

test('invalid character metadata falls back to plain text before paste', () => {
  for (const format of [
    { underline: 'invalid' },
    { strike: 'dotted' },
    { normalizeHeight: 'true' },
    { normalizeHeight: 1 },
    { size: 0.1 },
    { spc: 400001 },
    { kern: -1 },
    { kern: 400001 },
    { baseline: 30000 },
    { baseline: -30000 },
    { color: 'ACCENT1' },
    { underlineColor: 'ACCENT1' },
  ]) {
    assert.equal(parseTextClipboard(clipboard(format), '日本語'), null, JSON.stringify(format));
  }
});

test('supported character metadata preserves theme colors and detailed underline styles', () => {
  const format = {
    underline: 'wavyDbl',
    strike: 'dblStrike',
    size: 1,
    spc: -400000,
    kern: 400000,
    baseline: 0.3,
    color: 'scheme:accent1',
    underlineColor: '#123456',
  };
  assert.deepEqual(parseTextClipboard(clipboard(format), '日本語')?.formats[0].format, format);
});

test('underlineColor accepts null as follow-text metadata', () => {
  for (const underlineColor of [null, '#123456', 'scheme:accent1']) {
    const format = { underline: true, underlineColor };
    assert.deepEqual(parseTextClipboard(clipboard(format), '日本語')?.formats[0].format, format);
  }
});

test('accepts nested text effects with schema-safe boundaries', () => {
  const format = {
    outline: { color: '#123456', widthEmu: 12700 },
    shadow: {
      color: '#000000',
      blurEmu: 50800,
      offsetEmu: 38100,
      angleDeg: 45,
      opacity: 0.5,
      alignment: 'ctr',
      rotateWithShape: true,
    },
    innerShadow: {
      color: 'scheme:accent1',
      blurEmu: 63500,
      offsetEmu: 50800,
      angleDeg: 225,
      opacity: 0.4,
    },
    glow: { color: '#00FF00', radiusEmu: 63500, opacity: 0.7 },
    reflection: {
      blurEmu: 6350,
      offsetEmu: 0,
      angleDeg: 180,
      opacity: 0.2,
      startOpacity: 0.8,
      startPosition: 0,
      endPosition: 1,
      fadeDirection: 90,
      scaleX: 1,
      scaleY: -0.9,
      skewX: -20,
      skewY: 20,
      alignment: 'b',
      rotateWithShape: true,
    },
  };
  assert.deepEqual(parseTextClipboard(clipboard(format), '日本語')?.formats[0].format, format);
});

test('rejects unsafe nested text effect values', () => {
  for (const format of [
    { outline: { widthEmu: 20116801 } },
    { glow: {} },
    { innerShadow: { color: 'ACCENT1' } },
    { innerShadow: { blurEmu: -1 } },
    { innerShadow: { opacity: 1.1 } },
    { shadow: { alignment: 'middle' } },
    { glow: { color: '#00FF00', radiusEmu: Number.MAX_SAFE_INTEGER } },
    { reflection: { skewX: 90 } },
    { reflection: { scaleY: 30000 } },
  ]) {
    assert.equal(parseTextClipboard(clipboard(format), '日本語'), null, JSON.stringify(format));
  }
});

test('copies nested text effects through paste and a save/load round trip', async () => {
  const pres = createPresentation();
  const slide = addBlankSlide(pres);
  const source = addSlideTextBox(slide, {
    x: inches(1),
    y: inches(1),
    w: inches(3),
    h: inches(1),
    text: 'Source',
  });
  const effect = {
    color: '#123456',
    blurEmu: 50800,
    offsetEmu: 38100,
    angleDeg: 45,
    opacity: 0.5,
  };
  setShapeTextFormat(source, { innerShadow: effect });

  const copied = copyTextRange(source, 0, 6);
  const parsed = parseTextClipboard(JSON.stringify(copied), copied.text);
  assert.ok(parsed);

  const target = addSlideTextBox(slide, {
    x: inches(5),
    y: inches(1),
    w: inches(3),
    h: inches(1),
    text: 'Target',
  });
  setShapeTextFormat(target, parsed.formats[0].format, {
    range: { start: 0, end: 6 },
    reset: true,
  });

  const restored = await loadPresentation(await savePresentation(pres));
  const restoredTarget = getSlideShapes(getSlides(restored)[0])[1];
  assert.deepEqual(getShapeRunFormatEffective(restored, restoredTarget, 0, 0).innerShadow, effect);
});
