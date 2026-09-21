import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import {
  addBlankSlide,
  addSlideTextBox,
  copyShape,
  createPresentation,
  duplicateSlide,
  getParagraphIndent,
  getParagraphLevel,
  getSlideShapes,
  getSlides,
  inches,
  loadPresentation,
  savePresentation,
  setParagraphBullet,
  setParagraphLevel,
  setShapeBulletStyle,
  setShapeText,
} from '../src/api/index.ts';

const makeBox = () => {
  const pres = createPresentation();
  const box = addSlideTextBox(addBlankSlide(pres), {
    x: inches(1),
    y: inches(1),
    w: inches(6),
    h: inches(2),
    text: 'parent\nchild',
  });
  return { pres, box };
};

const rootIndent = { leftEmu: 342900, rightEmu: null, firstLineEmu: -342900 };
const nestedIndent = { leftEmu: 742950, rightEmu: null, firstLineEmu: -285750 };

describe('bullet indentation follows paragraph level', () => {
  it('produces the same indent regardless of setShapeBulletStyle / setParagraphLevel order', () => {
    const first = makeBox().box;
    const second = makeBox().box;
    setShapeBulletStyle(first, 'bullet');
    setParagraphLevel(first, 1, 1);
    setParagraphLevel(second, 1, 1);
    setShapeBulletStyle(second, 'bullet');
    expect(getParagraphIndent(first, 1)).toEqual(getParagraphIndent(second, 1));
    expect(getParagraphIndent(first, 1)).toEqual(nestedIndent);
    expect(getParagraphIndent(first, 0)).toEqual(rootIndent);
  });

  it.each(['bullet', 'number', { char: '★' }, { autoNum: 'romanLcPeriod' }] as const)(
    'updates per-paragraph %j bullets and resets level to zero',
    (style) => {
      const { box } = makeBox();
      setParagraphBullet(box, 1, style);
      setParagraphLevel(box, 1, 1);
      expect(getParagraphIndent(box, 1)).toEqual(nestedIndent);
      setParagraphLevel(box, 1, 8);
      expect(getParagraphIndent(box, 1)).toEqual({
        leftEmu: 2057400,
        rightEmu: null,
        firstLineEmu: -228600,
      });
      setParagraphLevel(box, 1, 0);
      expect(getParagraphIndent(box, 1)).toEqual(rootIndent);
      expect(getParagraphLevel(box, 1)).toBe(0);
    },
  );

  it('supports the bullets option and preserves automatic indents when replacing text', () => {
    const { box } = makeBox();
    setShapeText(box, 'parent\nchild', { bullets: 'bullet' });
    setParagraphLevel(box, 1, 1);
    expect(getParagraphIndent(box, 1)).toEqual(nestedIndent);
    setShapeText(box, 'replacement\nchild');
    setParagraphLevel(box, 1, 1);
    expect(getParagraphIndent(box, 1)).toEqual(nestedIndent);
  });

  it.each([undefined, 'none'] as const)('does not add indentation with bullets %j', (style) => {
    const { box } = makeBox();
    if (style !== undefined) setShapeBulletStyle(box, style);
    setParagraphLevel(box, 1, 1);
    expect(getParagraphIndent(box, 1)).toEqual({
      leftEmu: null,
      rightEmu: null,
      firstLineEmu: null,
    });
  });

  it('updates default indentation even while bullets are disabled', () => {
    const { box } = makeBox();
    setShapeBulletStyle(box, 'bullet');
    setShapeBulletStyle(box, 'none');
    setParagraphLevel(box, 1, 1);
    expect(getParagraphIndent(box, 1)).toEqual(nestedIndent);
    setShapeBulletStyle(box, 'number');
    expect(getParagraphIndent(box, 1)).toEqual(nestedIndent);
  });

  it('updates default indentation after saving and reloading', async () => {
    const { pres, box } = makeBox();
    setShapeBulletStyle(box, 'bullet');
    const loaded = await loadPresentation(await savePresentation(pres));
    const loadedBox = getSlideShapes(getSlides(loaded)[0]!)[0]!;
    setParagraphLevel(loadedBox, 1, 1);
    expect(getParagraphIndent(loadedBox, 1)).toEqual(nestedIndent);
  });

  it('updates default indentation on a duplicated slide', () => {
    const { pres, box } = makeBox();
    setShapeBulletStyle(box, 'bullet');
    const copy = duplicateSlide(pres, getSlides(pres)[0]!);
    const copiedBox = getSlideShapes(copy)[0]!;
    setParagraphLevel(copiedBox, 1, 1);
    expect(getParagraphIndent(copiedBox, 1)).toEqual(nestedIndent);
    expect(getParagraphIndent(box, 1)).toEqual(rootIndent);
  });

  it('updates default indentation on a copied shape', () => {
    const { pres, box } = makeBox();
    setShapeBulletStyle(box, 'bullet');
    const copiedBox = copyShape(addBlankSlide(pres), box);
    setParagraphLevel(copiedBox, 1, 1);
    expect(getParagraphIndent(copiedBox, 1)).toEqual(nestedIndent);
    expect(getParagraphIndent(box, 1)).toEqual(rootIndent);
  });

  it.each([
    { attribute: 'marL', value: 900000, expected: { ...rootIndent, leftEmu: 900000 } },
    { attribute: 'indent', value: -200000, expected: { ...rootIndent, firstLineEmu: -200000 } },
  ])(
    'preserves both indents when only $attribute differs from the default',
    async ({ attribute, value, expected }) => {
      const { pres, box } = makeBox();
      setShapeBulletStyle(box, 'bullet');
      const parts = unzipSync(await savePresentation(pres));
      const path = 'ppt/slides/slide1.xml';
      parts[path] = strToU8(
        strFromU8(parts[path]!).replaceAll(
          new RegExp(` ${attribute}="[^"]*"`, 'g'),
          ` ${attribute}="${value}"`,
        ),
      );
      const loaded = await loadPresentation(zipSync(parts));
      const templateBox = getSlideShapes(getSlides(loaded)[0]!)[0]!;
      setShapeBulletStyle(templateBox, 'number');
      setShapeText(templateBox, 'replacement\nchild');
      setParagraphLevel(templateBox, 1, 1);
      expect(getParagraphIndent(templateBox, 1)).toEqual(expected);
    },
  );
});
