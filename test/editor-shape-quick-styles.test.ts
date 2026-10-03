import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideShape,
  getSlides,
  inches,
  loadPresentation,
  savePresentation,
} from '../src/api/index.ts';
import { readZip } from '../src/internal/opc/index.ts';
import {
  applyShapeQuickStyle,
  quickStyleReferences,
  type ShapeQuickStyle,
} from '../site/src/lib/editor/core/shape-quick-styles.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/native/${name}`, import.meta.url));
const deckFixture = fileURLToPath(new URL('./fixtures/minimal/two-slides.pptx', import.meta.url));

const readGeneratedShape = async (style: ShapeQuickStyle, color: 'accent1' | 'dk1') => {
  const pres = await loadPresentation(await readFile(deckFixture));
  const shape = addSlideShape(getSlides(pres)[0]!, {
    preset: 'rect',
    x: inches(1),
    y: inches(1),
    w: inches(3),
    h: inches(2),
  });
  applyShapeQuickStyle(shape, style, color);
  const { entries } = readZip(await savePresentation(pres));
  const slide = new TextDecoder().decode(
    entries.find((entry) => entry.name === 'ppt/slides/slide1.xml')!.data,
  );
  return slide.match(/<p:sp>[\s\S]*?<\/p:sp>/g)!.at(-1)!;
};

const readNativeShape = async (name: string) => readFile(fixture(name), 'utf8');

/** Compare the paint/style payload while ignoring serializer defaults. */
const normalize = (xml: string): string =>
  xml
    .replace(/\s+xmlns(?::\w+)?="[^"]*"/g, '')
    .replace(/<a:scrgbClr r="0" g="0" b="0"\s*\/>/g, '<a:srgbClr val="000000"/>')
    .replace(/\s+algn="ctr"/g, '')
    .replace(/\s+flip="none"/g, '')
    .replace(/\s+rotWithShape="1"/g, '')
    .replace(/\s+scaled="1"/g, '')
    .replace(/\s+(?:left|l)="0"\s+(?:top|t)="0"\s+(?:right|r)="0"\s+(?:bottom|b)="0"/g, '')
    .replace(/\s+/g, '');

const paintPayload = (xml: string): string =>
  normalize(
    xml
      .match(/<p:spPr>[\s\S]*?<\/p:spPr>/)![0]
      .replace(/<a:xfrm>[\s\S]*?<\/a:xfrm>/, '')
      .replace(/<a:prstGeom[^>]*>[\s\S]*?<\/a:prstGeom>/, '') +
      xml.match(/<p:style>[\s\S]*?<\/p:style>/)![0],
  );

describe('editor shape quick styles', () => {
  it.each([
    ['Transparent', 'quickstyle-transparent-accent1-shape.xml'],
    ['Transparent, Colored Outline', 'quickstyle-transparent-outline-accent1-shape.xml'],
    ['Semitransparent', 'quickstyle-semitransparent-accent1-shape.xml'],
    ['Colored Fill, No Outline', 'quickstyle-fill-no-outline-accent1-shape.xml'],
    ['Gradient Fill, No Outline', 'quickstyle-gradient-no-outline-accent1-shape.xml'],
  ] as const)('matches native %s direct paint and style refs', async (style, nativeName) => {
    expect(paintPayload(await readGeneratedShape(style, 'accent1'))).toBe(
      paintPayload(await readNativeShape(nativeName)),
    );
  });

  it('matches the captured dark-theme gradient row as well as the accent row', async () => {
    expect(paintPayload(await readGeneratedShape('Gradient Fill, No Outline', 'dk1'))).toBe(
      paintPayload(await readNativeShape('quickstyle-gradient-no-outline-dk1-shape.xml')),
    );
  });

  it.each([
    ['Colored Fill', 'quickstyle-colored-fill-accent1-style.xml'],
    ['Intense Effect', 'quickstyle-intense-accent1-style.xml'],
  ] as const)('keeps the native theme reference row for %s', async (style, nativeName) => {
    const generated = quickStyleReferences(style, 'accent1');
    const native = await readNativeShape(nativeName);
    expect(generated.line).toMatchObject({
      idx: style === 'Colored Fill' ? 2 : 0,
      color: 'accent1',
    });
    expect(generated.fill).toMatchObject({
      idx: style === 'Colored Fill' ? 1 : 3,
      color: 'accent1',
    });
    expect(generated.effect).toMatchObject({
      idx: style === 'Colored Fill' ? 0 : 3,
      color: 'accent1',
    });
    expect(generated.font).toMatchObject({ idx: 'minor', color: 'lt1' });
    expect(native).toContain(
      `<a:effectRef idx="${style === 'Colored Fill' ? 0 : 3}"><a:schemeClr val="accent1"/></a:effectRef>`,
    );
  });
});
