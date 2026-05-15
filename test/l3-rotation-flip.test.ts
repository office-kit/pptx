// Shape rotation + flip.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  addSlideTextBox,
  findSlideLayout,
  getShapeFlip,
  getShapeRotation,
  getSlideXmlString,
  getSlides,
  inches,
  loadPresentation,
  setShapeFlip,
  setShapeRotation,
} from '../src/api/index.ts';
import { expectSchemaValid, isSchemaValidationAvailable } from './lib/expect-schema-valid.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const skipIfNoXmllint = isSchemaValidationAvailable() ? it : it.skip;

const newBoxedDeck = async () => {
  const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
  const layout = findSlideLayout(pres, 'Blank');
  if (!layout) throw new Error('expected layout');
  const slide = addSlide(pres, { layout });
  const box = addSlideTextBox(slide, {
    x: inches(1), y: inches(1), w: inches(2), h: inches(2), text: 'x',
  });
  return { pres, box };
};

describe('L3: Shape rotation', () => {
  it('writes rot="5400000" for setShapeRotation(90)', async () => {
    const { pres, box } = await newBoxedDeck();
    setShapeRotation(box, 90);
    const xml = getSlideXmlString(getSlides(pres).at(-1)!);
    expect(xml).toContain('rot="5400000"');
    expect(getShapeRotation(box)).toBe(90);
  });

  it('normalizes out-of-range values', async () => {
    const { box } = await newBoxedDeck();
    setShapeRotation(box, -90);
    expect(getShapeRotation(box)).toBe(270);
    setShapeRotation(box, 540); // 540 % 360 = 180
    expect(getShapeRotation(box)).toBe(180);
  });

  it('clears rotation when set to 0', async () => {
    const { pres, box } = await newBoxedDeck();
    setShapeRotation(box, 45);
    setShapeRotation(box, 0);
    const xml = getSlideXmlString(getSlides(pres).at(-1)!);
    expect(xml).not.toContain('rot=');
  });
});

describe('L3: Shape flip', () => {
  it('writes flipH="1" / flipV="1"', async () => {
    const { pres, box } = await newBoxedDeck();
    setShapeFlip(box, { horizontal: true, vertical: true });
    const xml = getSlideXmlString(getSlides(pres).at(-1)!);
    expect(xml).toMatch(/flipH="1"/);
    expect(xml).toMatch(/flipV="1"/);
    expect(getShapeFlip(box)).toEqual({ horizontal: true, vertical: true });
  });

  it('clears individual flip flags', async () => {
    const { box } = await newBoxedDeck();
    setShapeFlip(box, { horizontal: true, vertical: true });
    setShapeFlip(box, { horizontal: false });
    expect(getShapeFlip(box)).toEqual({ horizontal: false, vertical: true });
  });

  skipIfNoXmllint('rotated+flipped shape validates against pml.xsd', async () => {
    const { pres, box } = await newBoxedDeck();
    setShapeRotation(box, 45);
    setShapeFlip(box, { horizontal: true });
    expectSchemaValid(getSlideXmlString(getSlides(pres).at(-1)!), 'pml');
  });
});
