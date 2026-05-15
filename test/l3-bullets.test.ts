// Bullet and numbered list authoring.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  addSlideTextBox,
  findSlideLayout,
  getSlideXmlString,
  getSlides,
  inches,
  loadPresentation,
  setShapeBullets,
  setShapeText,
} from '../src/api/index.ts';
import { expectSchemaValid, isSchemaValidationAvailable } from './lib/expect-schema-valid.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const skipIfNoXmllint = isSchemaValidationAvailable() ? it : it.skip;

const makeBoxedDeck = async (
  fixtureName: string,
): Promise<{ pres: Awaited<ReturnType<typeof loadPresentation>>; box: ReturnType<typeof addSlideTextBox> }> => {
  const pres = await loadPresentation(await readFile(fixture(fixtureName)));
  const layout = findSlideLayout(pres, 'Blank');
  if (!layout) throw new Error('expected Blank layout');
  const slide = addSlide(pres, { layout });
  const box = addSlideTextBox(slide, {
    x: inches(1), y: inches(1), w: inches(6), h: inches(3), text: 'seed',
  });
  return { pres, box };
};

describe('L3: bullet / numbered lists', () => {
  it('emits <a:buChar char="•"/> for bullet style', async () => {
    const { pres, box } = await makeBoxedDeck('blank.pptx');
    setShapeText(box, 'Apples\nOranges\nBananas', { bullets: 'bullet' });
    const xml = getSlideXmlString(getSlides(pres).at(-1)!);
    expect((xml.match(/<a:buChar/g) ?? []).length).toBe(3);
    expect(xml).toContain('char="•"');
  });

  it('emits <a:buAutoNum type="arabicPeriod"/> for numbered style', async () => {
    const { pres, box } = await makeBoxedDeck('blank.pptx');
    setShapeText(box, 'a\nb', { bullets: 'number' });
    const xml = getSlideXmlString(getSlides(pres).at(-1)!);
    expect((xml.match(/<a:buAutoNum/g) ?? []).length).toBe(2);
    expect(xml).toContain('type="arabicPeriod"');
  });

  it('accepts a custom bullet character', async () => {
    const { pres, box } = await makeBoxedDeck('blank.pptx');
    setShapeText(box, 'one\ntwo', { bullets: { char: '◆' } });
    const xml = getSlideXmlString(getSlides(pres).at(-1)!);
    expect(xml).toContain('char="◆"');
  });

  it('setShapeBullets after setShapeText applies bullets without touching text', async () => {
    const { pres, box } = await makeBoxedDeck('blank.pptx');
    setShapeText(box, 'First\nSecond');
    setShapeBullets(box, 'bullet');
    const xml = getSlideXmlString(getSlides(pres).at(-1)!);
    expect(xml).toContain('First');
    expect(xml).toContain('Second');
    expect((xml.match(/<a:buChar/g) ?? []).length).toBe(2);
  });

  it("setShapeBullets('none') forces explicit bullet-free paragraphs", async () => {
    const { pres, box } = await makeBoxedDeck('blank.pptx');
    setShapeText(box, 'a\nb');
    setShapeBullets(box, 'bullet');
    setShapeBullets(box, 'none');
    const xml = getSlideXmlString(getSlides(pres).at(-1)!);
    expect(xml).not.toContain('<a:buChar');
    expect((xml.match(/<a:buNone\/>/g) ?? []).length).toBe(2);
  });

  skipIfNoXmllint('bullet output validates against pml.xsd', async () => {
    const { pres, box } = await makeBoxedDeck('blank.pptx');
    setShapeText(box, 'Apples\nOranges\nBananas', { bullets: 'bullet' });
    expectSchemaValid(getSlideXmlString(getSlides(pres).at(-1)!), 'pml');
  });
});
