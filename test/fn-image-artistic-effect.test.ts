import { readFile } from 'node:fs/promises';
import { describe, expect, it } from 'vitest';
import {
  duplicateSlide,
  getShapeImageArtisticEffect,
  getShapeImageBytes,
  getSlideShapes,
  getSlides,
  loadPresentation,
  savePresentation,
  setShapeImageBrightness,
  setShapeImageOpacity,
  setShapeImageRecolor,
} from '../src/api/index.ts';
import { readZip, writeZip } from '../src/internal/opc/index.ts';
import { renderSlideToSvg } from '../packages/preview/src/index.ts';
import { expectSchemaValid, isSchemaValidationAvailable } from './lib/expect-schema-valid.ts';

const decoder = new TextDecoder();
const encoder = new TextEncoder();
const HDPHOTO_REL = 'http://schemas.microsoft.com/office/2007/relationships/hdphoto';
// Opaque stand-in for PowerPoint's JPEG XR original: the library only has to
// keep it, never decode it. Starts with the JPEG XR signature.
const ORIGINAL = new Uint8Array([0x49, 0x49, 0xbc, 0x01, 1, 2, 3, 4, 5, 6, 7, 8]);

// The structure PowerPoint writes ([MS-ODRAWXML] "Pictures", and the
// Wood Type theme shipped with Mac PowerPoint): the blip embeds the processed
// picture; the a14:imgLayer points at the original in JPEG XR.
const blipWithEffect = (effect: string) =>
  `<a:blip r:embed="rId2"><a:extLst><a:ext uri="{BEBA8EAE-BF5A-486C-A8C5-ECC9F3942E4B}"><a14:imgProps xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"><a14:imgLayer r:embed="rId3"><a14:imgEffect>${effect}</a14:imgEffect><a14:imgEffect><a14:saturation sat="200000"/></a14:imgEffect></a14:imgLayer></a14:imgProps></a:ext><a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}"><a14:useLocalDpi xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" val="0"/></a:ext></a:extLst></a:blip>`;

const deckWithEffect = async (effect: string): Promise<Uint8Array> => {
  const fixture = await readFile(
    new URL('./fixtures/minimal/one-image-slide.pptx', import.meta.url),
  );
  const entries = readZip(fixture).entries.map((entry) => {
    const text = () => decoder.decode(entry.data);
    if (entry.name === 'ppt/slides/slide1.xml')
      return {
        ...entry,
        data: encoder.encode(text().replace('<a:blip r:embed="rId2"/>', blipWithEffect(effect))),
      };
    if (entry.name === 'ppt/slides/_rels/slide1.xml.rels')
      return {
        ...entry,
        data: encoder.encode(
          text().replace(
            '</Relationships>',
            `<Relationship Id="rId3" Type="${HDPHOTO_REL}" Target="../media/hdphoto1.wdp"/></Relationships>`,
          ),
        ),
      };
    if (entry.name === '[Content_Types].xml')
      return {
        ...entry,
        data: encoder.encode(
          text().replace(
            '<Default ',
            '<Default Extension="wdp" ContentType="image/vnd.ms-photo"/><Default ',
          ),
        ),
      };
    return entry;
  });
  return writeZip([...entries, { name: 'ppt/media/hdphoto1.wdp', data: ORIGINAL }]);
};

const picture = (pres: Awaited<ReturnType<typeof loadPresentation>>, index = 0) =>
  getSlideShapes(getSlides(pres)[index]!)[1]!;

describe('Artistic Effects (a14:imgProps)', () => {
  it.each([
    ['<a14:artisticPencilSketch trans="0" pressure="50"/>', 'pencilSketch'],
    ['<a14:artisticBlur radius="20"/>', 'blur'],
    ['<a14:artisticMosiaicBubbles/>', 'mosaicBubbles'],
    ['<a14:artisticGlowEdges smoothness="5"/>', 'glowEdges'],
  ])('reads %s', async (effect, kind) => {
    const pres = await loadPresentation(await deckWithEffect(effect));
    expect(getShapeImageArtisticEffect(picture(pres))).toBe(kind);
  });

  it('reports no artistic effect for corrections-only image properties or plain pictures', async () => {
    const corrected = await loadPresentation(
      await deckWithEffect('<a14:sharpenSoften amount="25000"/>'),
    );
    expect(getShapeImageArtisticEffect(picture(corrected))).toBeNull();
    const plain = await loadPresentation(
      await readFile(new URL('./fixtures/minimal/one-image-slide.pptx', import.meta.url)),
    );
    expect(getShapeImageArtisticEffect(picture(plain))).toBeNull();
  });

  it('keeps the effect, the original picture and its relationship through edits and saves', async () => {
    const pres = await loadPresentation(
      await deckWithEffect('<a14:artisticCement trans="10000" crackSpacing="40"/>'),
    );
    const rendered = getShapeImageBytes(picture(pres));
    // Every blip effect must land before the blip's extLst to stay schema-valid.
    setShapeImageOpacity(picture(pres), 0.5);
    setShapeImageBrightness(picture(pres), 0.2);
    setShapeImageRecolor(picture(pres), { kind: 'grayscale' });
    duplicateSlide(pres, getSlides(pres)[0]!);
    const saved = await savePresentation(pres);
    const reloaded = await loadPresentation(saved);
    for (const index of [0, 1]) {
      expect(getShapeImageArtisticEffect(picture(reloaded, index))).toBe('cement');
      expect(getShapeImageBytes(picture(reloaded, index))).toEqual(rendered);
    }
    const files = new Map(readZip(saved).entries.map((entry) => [entry.name, entry.data]));
    expect(decoder.decode(files.get('[Content_Types].xml'))).toContain(
      'Extension="wdp" ContentType="image/vnd.ms-photo"',
    );
    for (const slide of ['slide1', 'slide2']) {
      const xml = decoder.decode(files.get(`ppt/slides/${slide}.xml`));
      const layer = /<a14:imgLayer r:embed="(rId\d+)">/.exec(xml)?.[1];
      expect(layer).toBeDefined();
      expect(xml).toContain('<a14:artisticCement trans="10000" crackSpacing="40"/>');
      const rels = decoder.decode(files.get(`ppt/slides/_rels/${slide}.xml.rels`));
      const target = new RegExp(`Id="${layer}" Type="${HDPHOTO_REL}" Target="([^"]+)"`).exec(
        rels,
      )?.[1];
      expect(target).toMatch(/\.wdp$/);
      const part = target!.startsWith('/') ? target!.slice(1) : `ppt/${target!.slice(3)}`;
      expect(files.get(part)).toEqual(ORIGINAL);
      if (isSchemaValidationAvailable()) expectSchemaValid(xml, 'pml');
    }
  });

  it('renders the embedded (already processed) picture, not the original', async () => {
    const pres = await loadPresentation(
      await deckWithEffect('<a14:artisticPaintBrush trans="0" brushSize="3"/>'),
    );
    const svg = renderSlideToSvg(pres, getSlides(pres)[0]!);
    const rendered = Buffer.from(getShapeImageBytes(picture(pres))!).toString('base64');
    expect(svg).toContain(rendered);
    expect(svg).not.toContain(Buffer.from(ORIGINAL).toString('base64'));
    // No filter approximates the effect on top of the result PowerPoint saved.
    expect(svg).not.toMatch(/<filter/);
  });
});
