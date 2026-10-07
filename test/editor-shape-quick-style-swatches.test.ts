import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  getShapeId,
  getSlideShapes,
  getSlideColorMapOverride,
  getSlides,
  loadPresentation,
  savePresentation,
} from '../src/api/index.ts';
import { SLIDE_DOCUMENT } from '../src/api/_internal-symbols.ts';
import { readZip, writeZip } from '../src/internal/opc/index.ts';
import { parseXml } from '../src/internal/xml/index.ts';
import { shapeQuickStyleSwatches } from '../packages/editor/src/core/shape-quick-style-swatches.ts';

const deckFixture = fileURLToPath(new URL('./fixtures/minimal/two-slides.pptx', import.meta.url));

const packageSnapshot = (bytes: Uint8Array): ReadonlyMap<string, Uint8Array> =>
  new Map(
    readZip(bytes)
      .entries.map((entry) => [entry.name, entry.data] as const)
      .sort(([a], [b]) => a.localeCompare(b)),
  );

const svgFor = (image: string): string =>
  decodeURIComponent(image.slice('data:image/svg+xml,'.length));

describe('editor shape quick style swatches', () => {
  it('does not mutate the source deck and returns every gallery color/style', async () => {
    const source = await loadPresentation(await readFile(deckFixture));
    const before = await savePresentation(source);
    const swatches = shapeQuickStyleSwatches(getSlides(source)[0]!);
    const after = await savePresentation(source);

    expect(packageSnapshot(after)).toEqual(packageSnapshot(before));
    expect(swatches.size).toBe(11 * 7);
    expect([...swatches.keys()]).toContain('Colored Fill:accent1');
    expect([...swatches.keys()]).toContain('Gradient Fill, No Outline:dk1');
    for (const image of swatches.values())
      expect(image.startsWith('data:image/svg+xml,')).toBe(true);
  });

  it('produces distinct theme and preset thumbnails', async () => {
    const source = await loadPresentation(await readFile(deckFixture));
    const swatches = shapeQuickStyleSwatches(getSlides(source)[0]!);
    expect(swatches.get('Colored Outline:accent1')).not.toBe(swatches.get('Colored Fill:accent1'));
    expect(swatches.get('Colored Fill:accent1')).not.toBe(swatches.get('Intense Effect:accent1'));
    expect(swatches.get('Transparent:accent1')).not.toBe(swatches.get('Semitransparent:accent1'));
    expect(swatches.get('Semitransparent:accent1')).not.toBe(
      swatches.get('Gradient Fill, No Outline:accent1'),
    );
    expect(swatches.get('Gradient Fill, No Outline:accent1')).not.toBe(
      swatches.get('Gradient Fill, No Outline:dk1'),
    );
  }, 30_000);

  it('resolves theme colors from the source deck owner', async () => {
    const original = await readFile(deckFixture);
    const decoder = new TextDecoder();
    const entries = readZip(original).entries.map((entry) => {
      if (entry.name !== 'ppt/theme/theme1.xml') return entry;
      const xml = decoder.decode(entry.data).replace('val="4F81BD"', 'val="123456"');
      return { ...entry, data: new TextEncoder().encode(xml) };
    });
    const source = await loadPresentation(writeZip(entries));
    const swatches = shapeQuickStyleSwatches(getSlides(source)[0]!);
    expect(svgFor(swatches.get('Colored Fill:accent1')!)).toContain('#123456');
    expect(svgFor(swatches.get('Gradient Fill, No Outline:accent1')!)).toMatch(
      /stop-color="#(?!5485BF|95B3D7)[0-9A-F]{6}"/,
    );
  }, 30_000);

  it('preserves a source slide color-map override for gallery colors', async () => {
    const original = await readFile(deckFixture);
    const decoder = new TextDecoder();
    const entries = readZip(original).entries.map((entry) => {
      if (entry.name !== 'ppt/slides/slide1.xml') return entry;
      const xml = decoder
        .decode(entry.data)
        .replace(
          '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>',
          '<p:clrMapOvr><a:overrideClrMapping accent1="accent2"/></p:clrMapOvr>',
        );
      return { ...entry, data: new TextEncoder().encode(xml) };
    });
    const source = await loadPresentation(writeZip(entries));
    expect(getSlideColorMapOverride(getSlides(source)[0]!)).toEqual({ accent1: 'accent2' });
    const svg = svgFor(shapeQuickStyleSwatches(getSlides(source)[0]!).get('Colored Fill:accent1')!);
    expect(svg).toContain('#C0504D');
    expect(svg).not.toContain('#4F81BD');
  }, 30_000);

  it('handles an animated source without removing its timing tree', async () => {
    const source = await loadPresentation(await readFile(deckFixture));
    const slide = getSlides(source)[0]!;
    const shapeId = getShapeId(getSlideShapes(slide)[0]!);
    const timing = parseXml(
      `<p:timing xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">` +
        `<p:tnLst><p:par><p:cTn id="1" dur="indefinite" nodeType="tmRoot">` +
        `<p:childTnLst><p:seq><p:cTn id="2" dur="indefinite" nodeType="mainSeq">` +
        `<p:childTnLst><p:par><p:cTn id="3" presetID="1" presetClass="entr">` +
        `<p:childTnLst><p:set><p:cBhvr><p:cTn id="4"/>` +
        `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl></p:cBhvr></p:set>` +
        `</p:childTnLst></p:cTn></p:par></p:childTnLst>` +
        `</p:cTn></p:seq></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>`,
    ).root;
    slide[SLIDE_DOCUMENT].root.children.push(timing);
    const before = await savePresentation(source);
    expect(() => shapeQuickStyleSwatches(slide)).not.toThrow();
    expect(packageSnapshot(await savePresentation(source))).toEqual(packageSnapshot(before));
  }, 30_000);
});
