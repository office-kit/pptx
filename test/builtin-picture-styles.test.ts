import { readdirSync, readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import {
  BUILTIN_PICTURE_STYLES,
  getShapeImageCompressionState,
  getShapeImageCrop,
  getShapePictureStyle,
  getSlideShapes,
  getSlides,
  loadPresentation,
  savePresentation,
  setShapeImageCompressionState,
  setShapeImageCrop,
  setShapePictureStyle,
  setShapeStroke,
  type PresentationData,
} from '../src/api/index.ts';
import { readZip } from '../src/internal/opc/index.ts';
import { renderSlideToSvg } from '../packages/preview/src/index.ts';
import { expectSchemaValid, isSchemaValidationAvailable } from './lib/expect-schema-valid.ts';

const decoder = new TextDecoder();
const STYLE_DIR = new URL('./fixtures/native/picture-styles/', import.meta.url);
const COMPRESS_DIR = new URL('./fixtures/native/compress-pictures/', import.meta.url);
const captures = readdirSync(STYLE_DIR)
  .filter((file) => file.endsWith('.xml'))
  .sort()
  .map((file) => readFileSync(new URL(file, STYLE_DIR), 'utf8'));

const blank = readFileSync(new URL('./fixtures/minimal/one-image-slide.pptx', import.meta.url));
const picture = (pres: PresentationData) => getSlideShapes(getSlides(pres)[0]!)[1]!;
const slideXml = async (pres: PresentationData) => {
  const files = readZip(await savePresentation(pres)).entries;
  return decoder.decode(files.find((entry) => entry.name === 'ppt/slides/slide1.xml')!.data);
};
// `p:spPr` after `a:xfrm`: the part a picture style writes.
const styledPart = (xml: string) => /<p:spPr><a:xfrm>.*?<\/a:xfrm>(.*?)<\/p:spPr>/.exec(xml)![1];

describe('built-in picture styles', () => {
  it('lists the 28 styles in the order of the gallery captures', () => {
    expect(BUILTIN_PICTURE_STYLES).toEqual(
      captures.map((xml) => /Picture Style "([^"]+)"/.exec(xml)![1]),
    );
  });

  it.each(BUILTIN_PICTURE_STYLES.map((name, index) => [name, index] as const))(
    'writes exactly what the reference desktop app writes for %s',
    async (name, index) => {
      const pres = await loadPresentation(blank);
      const before = await slideXml(pres);
      setShapePictureStyle(picture(pres), name);
      const xml = await slideXml(pres);
      expect(styledPart(xml)).toBe(styledPart(captures[index]!));
      // The picture itself and its frame stay as they were.
      expect(/<p:blipFill>.*?<\/p:blipFill>/.exec(xml)![0]).toBe(
        /<p:blipFill>.*?<\/p:blipFill>/.exec(before)![0],
      );
      expect(/<a:xfrm>.*?<\/a:xfrm>/.exec(xml)![0]).toBe(/<a:xfrm>.*?<\/a:xfrm>/.exec(before)![0]);
      if (isSchemaValidationAvailable()) expectSchemaValid(xml, 'pml');
      const reloaded = await loadPresentation(await savePresentation(pres));
      expect(getShapePictureStyle(picture(reloaded))).toBe(name);
    },
  );

  it('replaces a previous style and keeps the crop', async () => {
    const pres = await loadPresentation(blank);
    setShapeImageCrop(picture(pres), { left: 0.1, top: 0, right: 0, bottom: 0 });
    setShapePictureStyle(picture(pres), 'Metal Oval');
    setShapePictureStyle(picture(pres), 'Soft Edge Rectangle');
    expect(styledPart(await slideXml(pres))).toBe(styledPart(captures[5]!));
    expect(getShapeImageCrop(picture(pres))?.left).toBeCloseTo(0.1);
  });

  it('reads no style for a plain picture or one edited after styling', async () => {
    const pres = await loadPresentation(blank);
    expect(getShapePictureStyle(picture(pres))).toBeNull();
    setShapePictureStyle(picture(pres), 'Simple Frame, Black');
    setShapeStroke(picture(pres), { widthEmu: 12700 });
    expect(getShapePictureStyle(picture(pres))).toBeNull();
  });

  it('rejects unknown names and non-pictures', async () => {
    const pres = await loadPresentation(blank);
    // @ts-expect-error not a built-in style
    expect(() => setShapePictureStyle(picture(pres), 'Gold Frame')).toThrow(/built-in/);
    const title = getSlideShapes(getSlides(pres)[0]!)[0]!;
    expect(() => setShapePictureStyle(title, 'Metal Oval')).toThrow(/picture/);
  });

  it('renders every style without failing', async () => {
    const pres = await loadPresentation(blank);
    for (const name of BUILTIN_PICTURE_STYLES) {
      setShapePictureStyle(picture(pres), name);
      expect(renderSlideToSvg(pres, getSlides(pres)[0]!)).toContain('<image');
    }
  });
});

describe('picture compression state', () => {
  const capturedBlip = (file: string) =>
    /<a:blip .*?<\/a:blip>/.exec(readFileSync(new URL(file, COMPRESS_DIR), 'utf8'))![0];

  it.each([
    ['print', '4-print-220-ppi.xml'],
    ['screen', '5-on-screen-150-ppi.xml'],
    ['email', '6-email-96-ppi.xml'],
  ] as const)('writes the reference desktop app’s %s blip', async (state, file) => {
    const pres = await loadPresentation(blank);
    setShapeImageCompressionState(picture(pres), state);
    const xml = await slideXml(pres);
    const embed = /<a:blip r:embed="(rId\d+)"/.exec(xml)![1];
    expect(/<a:blip .*?<\/a:blip>/.exec(xml)![0]).toBe(capturedBlip(file).replace('rId2', embed!));
    if (isSchemaValidationAvailable()) expectSchemaValid(xml, 'pml');
    const reloaded = await loadPresentation(await savePresentation(pres));
    expect(getShapeImageCompressionState(picture(reloaded))).toBe(state);
  });

  it('replaces the state and removes it with its extension', async () => {
    const pres = await loadPresentation(blank);
    const before = await slideXml(pres);
    setShapeImageCompressionState(picture(pres), 'print');
    setShapeImageCompressionState(picture(pres), 'email');
    const xml = await slideXml(pres);
    expect(xml.match(/useLocalDpi/g)).toHaveLength(1);
    expect(xml).toContain('cstate="email"');
    setShapeImageCompressionState(picture(pres), null);
    expect(getShapeImageCompressionState(picture(pres))).toBeNull();
    const pic = (slide: string) => /<p:pic>.*<\/p:pic>/.exec(slide)![0];
    expect(pic(await slideXml(pres))).toBe(pic(before));
  });

  it('rejects values outside ST_BlipCompression', async () => {
    const pres = await loadPresentation(blank);
    // @ts-expect-error not an ST_BlipCompression value
    expect(() => setShapeImageCompressionState(picture(pres), 'web')).toThrow(/ST_BlipCompression/);
  });
});
