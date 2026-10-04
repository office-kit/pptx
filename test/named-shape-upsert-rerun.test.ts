import { describe, expect, it } from 'vitest';
import {
  addSlideImage,
  addSlideTextBox,
  addTitleSlide,
  createPresentation,
  emu,
  findShapeByName,
  getMediaParts,
  getShapeId,
  getShapeName,
  getShapeText,
  getSlideShapes,
  getSlides,
  listPackageParts,
  loadPresentation,
  readPackagePart,
  savePresentation,
  setShapeText,
} from '../src/api/index.ts';
import { readZip, writeZip } from '../src/internal/opc/index.ts';

// Minimal 1×1 PNG (transparent).
const PNG = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
  0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
  0x89, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x62, 0x00, 0x01, 0x00, 0x00,
  0x05, 0x00, 0x01, 0x0d, 0x0a, 0x2d, 0xb4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae,
  0x42, 0x60, 0x82,
]);
const UNKNOWN_PART = '/customXml/office-kit-unknown.xml';
const UNKNOWN_XML = new TextEncoder().encode('<vendor xmlns="urn:example:vendor"><keep/></vendor>');

// A save script written only against the public API. Adding shapes is not
// idempotent, so the script finds its own shapes by name and updates them;
// running it again on its output must not add a second copy.
const updateDeck = async (input: Uint8Array, kpi: string): Promise<Uint8Array> => {
  const pres = await loadPresentation(input);
  const slide = getSlides(pres)[0]!;
  const box = findShapeByName(slide, 'kpi');
  if (box) setShapeText(box, kpi);
  else
    addSlideTextBox(slide, {
      x: emu(457200),
      y: emu(457200),
      w: emu(3657600),
      h: emu(457200),
      text: kpi,
      name: 'kpi',
    });
  if (!findShapeByName(slide, 'logo'))
    addSlideImage(slide, PNG, {
      x: emu(0),
      y: emu(0),
      w: emu(457200),
      h: emu(457200),
      name: 'logo',
    });
  return savePresentation(pres);
};

// Fixture only: a vendor part this library does not model, which every save must carry over.
const sourceDeck = async (): Promise<Uint8Array> => {
  const pres = createPresentation();
  addTitleSlide(pres, 'Quarterly review');
  const { entries } = readZip(await savePresentation(pres));
  return writeZip([...entries, { name: UNKNOWN_PART.slice(1), data: UNKNOWN_XML }]);
};

const inspect = async (bytes: Uint8Array) => {
  const pres = await loadPresentation(bytes);
  const shapes = getSlideShapes(getSlides(pres)[0]!);
  return {
    names: shapes.map(getShapeName),
    ids: shapes.map(getShapeId),
    kpi: getShapeText(shapes.find((shape) => getShapeName(shape) === 'kpi')!),
    parts: listPackageParts(pres).map((part) => part.name),
    media: getMediaParts(pres).length,
    unknown: readPackagePart(pres, UNKNOWN_PART),
  };
};

describe('re-running a named upsert save script', () => {
  it('updates the same shapes instead of adding copies, keeping IDs and unknown parts', async () => {
    const first = await inspect(await updateDeck(await sourceDeck(), 'Q1'));
    const second = await inspect(
      await updateDeck(await updateDeck(await sourceDeck(), 'Q1'), 'Q2'),
    );

    expect(first.names.filter((name) => name === 'kpi')).toHaveLength(1);
    expect(first.names.filter((name) => name === 'logo')).toHaveLength(1);
    expect(first.kpi).toBe('Q1');
    expect(second.kpi).toBe('Q2');
    expect(second.names).toEqual(first.names);
    expect(second.ids).toEqual(first.ids);
    expect(second.parts).toEqual(first.parts);
    expect(second.media).toBe(first.media);
    expect(second.unknown).toEqual(UNKNOWN_XML);
  });
});
