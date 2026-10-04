import { describe, expect, it } from 'vitest';
import {
  addBlankSlide,
  addSlideTextBox,
  createPresentation,
  emu,
  getSlideXmlString,
  setShapeGradientFill,
  setShapeTextFormat,
  toWritableTextFormat,
} from '../src/api/index.ts';

const gradient = {
  kind: 'gradient' as const,
  stops: [
    { offset: 0, color: '#112233' as const },
    { offset: 1, color: '#FFFFFF' as const },
  ],
};

describe('text fill choice boundaries', () => {
  it('rejects color together with textFill before touching the run', () => {
    const pres = createPresentation();
    const slide = addBlankSlide(pres);
    const box = addSlideTextBox(slide, {
      x: emu(0),
      y: emu(0),
      w: emu(914400),
      h: emu(457200),
      text: 'Paint',
    });
    const before = getSlideXmlString(slide);
    expect(() => setShapeTextFormat(box, { color: '#FF0000', textFill: gradient })).toThrow(
      /mutually exclusive/,
    );
    expect(getSlideXmlString(slide)).toBe(before);
  });

  it('keeps the rest of a copied format when its fill uses an unwritable color token', () => {
    const pattern = {
      kind: 'pattern' as const,
      preset: 'pct5' as const,
      foreground: 'notASchemeToken',
      background: '#FFFFFF',
    };
    expect(toWritableTextFormat({ bold: true, textFill: pattern })).toEqual({ bold: true });
    expect(
      toWritableTextFormat({
        italic: true,
        textFill: {
          kind: 'gradient',
          stops: [
            { offset: 0, color: '#112233' },
            { offset: 1, color: 'notASchemeToken' },
          ],
        },
      }),
    ).toEqual({ italic: true });
    expect(toWritableTextFormat({ textFill: gradient })).toEqual({ textFill: gradient });
  });

  it('names the public caller when a gradient path is invalid and leaves the slide unchanged', () => {
    const pres = createPresentation();
    const slide = addBlankSlide(pres);
    const box = addSlideTextBox(slide, {
      x: emu(0),
      y: emu(0),
      w: emu(914400),
      h: emu(457200),
      text: 'Paint',
    });
    const before = getSlideXmlString(slide);
    // The cast exercises the JavaScript boundary with an out-of-schema path.
    const bogus = { ...gradient, path: 'bogus' as 'linear' };
    expect(() => setShapeGradientFill(box, bogus)).toThrow(
      /^setShapeGradientFill: path: .*is not one of:/,
    );
    expect(() => setShapeTextFormat(box, { bold: true, textFill: bogus })).toThrow(
      /^setShapeTextFormat: textFill: path: .*is not one of:/,
    );
    expect(getSlideXmlString(slide)).toBe(before);
  });
});
