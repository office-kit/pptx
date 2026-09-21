import { describe, expect, it } from 'vitest';
import {
  addBlankSlide,
  addSlideChart,
  addSlideTable,
  addSlideTextBox,
  createPresentation,
  getShapeText,
  getSlideShapes,
  getSlides,
  inches,
  loadPresentation,
  savePresentation,
  setSlideNotes,
} from '../src/api/index.ts';
import { _internalsForTest } from '../src/internal/xml/serialize.ts';

const bounds = { x: inches(0), y: inches(0), w: inches(4), h: inches(2) };
const forbidden = [
  ...Array.from({ length: 0x20 }, (_, c) => c).filter((c) => ![9, 10, 13].includes(c)),
  0xfffe,
  0xffff,
  0xd800,
  0xdbff,
  0xdc00,
  0xdfff,
];

describe('issue-307: XML 1.0 characters', () => {
  for (const [name, escape] of Object.entries(_internalsForTest)) {
    it.each(forbidden)(`${name} rejects U+%i`, (code) => {
      expect(() => escape(`before${String.fromCharCode(code)}after`)).toThrow(
        `U+${code.toString(16).toUpperCase().padStart(4, '0')}`,
      );
    });

    it.each(['\ud800', '\udc00', '\ud800\ud800', '\udc00\ud800', '😀\udfff'])(
      `${name} rejects unpaired surrogates in %j`,
      (text) => expect(() => escape(text)).toThrow(/XML-illegal/),
    );

    it(`${name} preserves legal character boundaries and supplementary noncharacters`, () => {
      const text = String.fromCodePoint(
        0x20,
        0xd7ff,
        0xe000,
        0xfffd,
        0x10000,
        0x1f600,
        0x1fffe,
        0x10ffff,
      );
      expect(escape(text)).toBe(text);
    });
  }

  const authors = {
    'run text': (slide: ReturnType<typeof addBlankSlide>, text: string) =>
      addSlideTextBox(slide, { ...bounds, text }),
    'table cell': (slide: ReturnType<typeof addBlankSlide>, text: string) =>
      addSlideTable(slide, { ...bounds, rows: [[text]] }),
    'chart category': (slide: ReturnType<typeof addBlankSlide>, text: string) =>
      addSlideChart(slide, {
        ...bounds,
        spec: { kind: 'column', categories: [text], series: [{ name: 'Series', values: [1] }] },
      }),
    'chart series': (slide: ReturnType<typeof addBlankSlide>, text: string) =>
      addSlideChart(slide, {
        ...bounds,
        spec: { kind: 'column', categories: ['Category'], series: [{ name: text, values: [1] }] },
      }),
    notes: (slide: ReturnType<typeof addBlankSlide>, text: string) => setSlideNotes(slide, text),
  };

  for (const [name, author] of Object.entries(authors)) {
    it.each([0xfffe, 0xffff, 0xd800, 0xdc00])(`${name} rejects U+%i`, (code) => {
      const pres = createPresentation();
      expect(() => author(addBlankSlide(pres), `before${String.fromCharCode(code)}after`)).toThrow(
        /XML-illegal/,
      );
    });
  }

  it('preserves supplementary characters through save and load', async () => {
    const pres = createPresentation();
    const text = `日本語 😀 ${String.fromCodePoint(0x10000, 0x1fffe, 0x10ffff)}`;
    addSlideTextBox(addBlankSlide(pres), { ...bounds, text });
    const loaded = await loadPresentation(await savePresentation(pres));
    expect(getShapeText(getSlideShapes(getSlides(loaded)[0]!)[0]!)).toBe(text);
  });
});
