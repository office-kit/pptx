import { describe, expect, it } from 'vitest';
import {
  addBlankSlide,
  addSlideTextBox,
  createPresentation,
  getParagraphBulletStyle,
  getSlideShapes,
  getSlides,
  inches,
  loadPresentation,
  savePresentation,
  setParagraphBullet,
} from '../src/api/index.ts';
import { readZip, writeZip } from '../src/internal/opc/index.ts';

describe('imported bullet size percentages', () => {
  it.each([
    ['25%', 0.25],
    ['50%', 0.5],
    ['100%', 1],
    ['400%', 4],
    ['50000', 0.5],
  ])('reads %s without changing its magnitude', async (value, expected) => {
    const pres = createPresentation();
    const shape = addSlideTextBox(addBlankSlide(pres), {
      x: inches(1),
      y: inches(1),
      w: inches(5),
      h: inches(2),
      text: 'Bullet',
    });
    setParagraphBullet(shape, 0, 'bullet');
    const zip = readZip(await savePresentation(pres));
    const entry = zip.entries.find((e) => e.name === 'ppt/slides/slide1.xml')!;
    const xml = new TextDecoder().decode(entry.data);
    expect(xml).toContain('<a:buChar');
    const data = new TextEncoder().encode(
      xml.replace('<a:buChar', `<a:buSzPct val="${value}"/><a:buChar`),
    );
    const imported = await loadPresentation(
      writeZip(zip.entries.map((e) => (e === entry ? { ...e, data } : e))),
    );
    const reloaded = await loadPresentation(await savePresentation(imported));
    for (const presentation of [imported, reloaded]) {
      const box = getSlideShapes(getSlides(presentation)[0]!)[0]!;
      expect(getParagraphBulletStyle(presentation, box, 0).sizePct).toBe(expected);
    }
  });
});
