// `setShapeTextField` writes `<a:fld>` — text PowerPoint fills in on open
// (the slide's number, today's date) rather than text the file states.

import { describe, expect, it } from 'vitest';
import {
  addBlankSlide,
  addSlideTextBox,
  createPresentation,
  getShapeParagraphElements,
  getShapeText,
  getShapeXmlString,
  getSlides,
  getSlideShapes,
  inches,
  loadPresentation,
  savePresentation,
  setShapeText,
  setShapeTextField,
  setShapeTextFormat,
} from '../src/api/index.ts';

const textBox = (text = '') => {
  const pres = createPresentation();
  const slide = addBlankSlide(pres);
  const shape = addSlideTextBox(slide, {
    x: inches(1),
    y: inches(1),
    w: inches(3),
    h: inches(1),
    text,
  });
  return { pres, shape };
};

describe('setShapeTextField', () => {
  it('writes one field with a braced uppercase GUID and the given type', () => {
    const { shape } = textBox('placeholder');
    setShapeTextField(shape, 'slidenum');
    const xml = getShapeXmlString(shape);
    expect(xml).toMatch(
      /<a:fld id="\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}" type="slidenum">/,
    );
    // ST_Guid is required on CT_TextField, so no field may go out without one.
    expect(xml.match(/<a:fld\b/g)).toHaveLength(1);
    const elements = getShapeParagraphElements(shape, 0);
    expect(elements).toHaveLength(1);
    expect(elements[0]).toMatchObject({ kind: 'fld', type: 'slidenum' });
  });

  it('replaces the whole body — a field placeholder holds nothing else', () => {
    const { shape } = textBox('first\nsecond');
    setShapeTextField(shape, 'slidenum');
    expect(getShapeXmlString(shape).match(/<a:p>/g)).toHaveLength(1);
    expect(getShapeXmlString(shape)).not.toContain('second');
  });

  it('carries the replaced text’s formatting onto the field', () => {
    const { shape } = textBox('1');
    setShapeTextFormat(shape, { bold: true, size: 18, color: '#FF0000' });
    setShapeTextField(shape, 'slidenum');
    const format = getShapeParagraphElements(shape, 0)[0]?.format;
    expect(format).toMatchObject({ bold: true, size: 18, color: '#FF0000' });
  });

  it('stores the cached value a field-blind reader shows', () => {
    const { shape } = textBox();
    setShapeTextField(shape, 'datetime1', { text: '2026-09-23' });
    expect(getShapeText(shape)).toBe('2026-09-23');
    // Without one the field is still valid; `<a:t>` is optional in the schema
    // and PowerPoint overwrites whatever is there.
    setShapeTextField(shape, 'slidenum');
    expect(getShapeText(shape)).toBe('');
  });

  it('keeps bodyPr and lstStyle, which describe the body rather than its text', () => {
    const { shape } = textBox('x');
    setShapeTextField(shape, 'slidenum');
    const xml = getShapeXmlString(shape);
    expect(xml).toContain('<a:bodyPr');
    expect(xml).toContain('<a:lstStyle');
  });

  it('refuses an empty type', () => {
    const { shape } = textBox();
    expect(() => setShapeTextField(shape, '')).toThrow(/type must not be empty/);
  });

  it('survives the save/load round trip', async () => {
    const { pres, shape } = textBox('x');
    setShapeTextField(shape, 'slidenum', { text: '3' });
    const reloaded = await loadPresentation(await savePresentation(pres));
    const saved = getSlideShapes(getSlides(reloaded)[0]!)[0]!;
    expect(getShapeParagraphElements(saved, 0)[0]).toMatchObject({
      kind: 'fld',
      type: 'slidenum',
      text: '3',
    });
  });

  it('is replaced by ordinary text, which is how a field is removed', () => {
    const { shape } = textBox('x');
    setShapeTextField(shape, 'slidenum');
    setShapeText(shape, 'literal');
    expect(getShapeXmlString(shape)).not.toContain('<a:fld');
    expect(getShapeText(shape)).toBe('literal');
  });
});
