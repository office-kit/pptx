import { readFile } from 'node:fs/promises';
import { describe, expect, it } from 'vitest';
import { expectSchemaValid, isSchemaValidationAvailable } from './lib/expect-schema-valid.ts';
import {
  addBlankSlide,
  addSlideMedia,
  createPresentation,
  getShapeMediaPlayback,
  getShapeId,
  getSlides,
  getSlideShapes,
  inches,
  loadPresentation,
  readPackagePart,
  savePresentation,
  setShapeMediaPlayback,
  _internalPackageOf,
} from '../src/api/index.ts';
import { partName } from '../src/internal/opc/index.ts';

const mediaData = (): Uint8Array =>
  new Uint8Array([0x49, 0x44, 0x33, 3, 0, 0, 0, 0, 0, 0, 0xff, 0xfb]);
const box = { x: inches(1), y: inches(1), w: inches(4), h: inches(2.25) };
const skipIfNoXmllint = isSchemaValidationAvailable() ? it : it.skip;
const slideXml = (pres: Parameters<typeof readPackagePart>[0]): string =>
  new TextDecoder().decode(readPackagePart(pres, '/ppt/slides/slide1.xml')!);

type Condition = { attrs?: string; child?: string };

const nestedTiming = (
  shapeId: number,
  parent: Condition,
  own: Condition,
  container: 'par' | 'seq' = 'par',
) => {
  const condition = ({ attrs = 'delay="0"', child = '' }: Condition): string =>
    `<p:cond ${attrs}>${child}</p:cond>`;
  const ownMedia = `<p:audio><p:cMediaNode><p:cTn id="4"><p:stCondLst>${condition(own)}</p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl></p:cMediaNode></p:audio>`;
  return `<p:timing xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:tnLst><p:par><p:cTn id="1" nodeType="tmRoot"><p:childTnLst><p:${container}><p:cTn id="2"><p:stCondLst>${condition(parent)}</p:stCondLst><p:childTnLst>${ownMedia}</p:childTnLst></p:cTn></p:${container}></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>`;
};

const nestedDeck = async (
  parent: Condition,
  own: Condition,
  container: 'par' | 'seq' = 'par',
  transform: (xml: string) => string = (xml) => xml,
) => {
  const pres = createPresentation();
  const slide = addBlankSlide(pres);
  const shape = addSlideMedia(slide, { kind: 'audio', data: mediaData(), format: 'mp3', ...box });
  const shapeId = getShapeId(shape);
  const part = _internalPackageOf(pres).getPart(partName('/ppt/slides/slide1.xml'))!;
  part.data = new TextEncoder().encode(
    new TextDecoder()
      .decode(part.data)
      .replace(
        /<p:timing[\s\S]*?<\/p:timing>/,
        transform(nestedTiming(shapeId, parent, own, container)),
      ),
  );
  const loaded = await loadPresentation(await savePresentation(pres));
  return { pres: loaded, shape: getSlideShapes(getSlides(loaded)[0]!)[0]! };
};

describe('nested media playback timing', () => {
  it('sums the parent and media-own cTn delays and survives round-trip', async () => {
    const { pres, shape } = await nestedDeck({ attrs: 'delay="100"' }, { attrs: 'delay="250"' });
    expect(getShapeMediaPlayback(shape)).toMatchObject({ autoplay: true, delayMs: 350 });

    const reloaded = await loadPresentation(await savePresentation(pres));
    expect(getShapeMediaPlayback(getSlideShapes(getSlides(reloaded)[0]!)[0]!)).toMatchObject({
      autoplay: true,
      delayMs: 350,
    });
  });

  it.each([
    ['indefinite', { attrs: 'delay="indefinite"' }],
    ['click', { attrs: 'evt="onClick" delay="0"' }],
  ] as const)('treats a media-own %s condition as manual playback', async (_name, own) => {
    const { shape } = await nestedDeck({ attrs: 'delay="100"' }, own);
    expect(getShapeMediaPlayback(shape)?.autoplay).toBe(false);
  });

  it('refuses an atomic autoplay conversion when a parent is click-triggered', async () => {
    const { pres, shape } = await nestedDeck(
      { attrs: 'evt="onClick" delay="0"' },
      { attrs: 'delay="250"' },
    );
    const before = slideXml(pres);
    expect(() => setShapeMediaPlayback(shape, { autoplay: true, volume: 0.2 })).toThrow();
    expect(slideXml(pres)).toBe(before);
  });

  it('changes only the media-own delay when setting a nested total delay', async () => {
    const { pres, shape } = await nestedDeck({ attrs: 'delay="100"' }, { attrs: 'delay="250"' });
    setShapeMediaPlayback(shape, { autoplay: true, delayMs: 500 });
    expect(slideXml(pres)).toContain('<p:cond delay="100"/>');
    expect(slideXml(pres)).toContain('<p:cond delay="400"/>');
    expect(getShapeMediaPlayback(shape)).toMatchObject({ autoplay: true, delayMs: 500 });
  });

  it('does not flatten sequence timing into slide autoplay', async () => {
    const { shape } = await nestedDeck({ attrs: 'delay="100"' }, { attrs: 'delay="250"' }, 'seq');
    expect(getShapeMediaPlayback(shape)?.autoplay).toBe(false);
  });

  it.each(['<p:tn val="2"/>', '<p:rtn val="all"/>', '<p:tgtEl><p:sldTgt/></p:tgtEl>'])(
    'does not treat a referenced start %s as a plain delay',
    async (child) => {
      const { pres, shape } = await nestedDeck(
        { attrs: 'delay="100"' },
        { attrs: 'delay="250"', child },
      );
      expect(getShapeMediaPlayback(shape)?.autoplay).toBe(false);
      const before = slideXml(pres);
      expect(() => setShapeMediaPlayback(shape, { delayMs: 500, volume: 0.2 })).toThrow();
      expect(slideXml(pres)).toBe(before);
    },
  );

  it('edits non-start properties without losing a click ancestor', async () => {
    const { pres, shape } = await nestedDeck(
      { attrs: 'evt="onClick" delay="100"' },
      { attrs: 'delay="250"' },
    );
    setShapeMediaPlayback(shape, { volume: 0.2, muted: true, loop: true, hideWhenStopped: true });
    expect(slideXml(pres)).toContain('evt="onClick" delay="100"');
    const saved = await loadPresentation(await savePresentation(pres));
    expect(getShapeMediaPlayback(getSlideShapes(getSlides(saved)[0]!)[0]!)).toMatchObject({
      autoplay: false,
      volume: 0.2,
      muted: true,
      loop: true,
      hideWhenStopped: true,
    });
  });

  it('changes only the media timing fill when toggling rewind-after-playing', async () => {
    const { pres, shape } = await nestedDeck({ attrs: 'delay="100"' }, { attrs: 'delay="250"' });
    setShapeMediaPlayback(shape, { rewindAfterPlaying: true });
    expect(slideXml(pres)).toContain('delay="100"');
    expect(slideXml(pres)).toContain('delay="250"');
    expect(slideXml(pres)).toContain('fill="remove"');
    expect(getShapeMediaPlayback(shape)).toMatchObject({ rewindAfterPlaying: true });
  });

  it('recognizes native Play in Background command timing without rewriting it', async () => {
    const { pres } = await nestedDeck({}, {});
    const part = _internalPackageOf(pres).getPart(partName('/ppt/slides/slide1.xml'))!;
    const timing = await readFile(
      new URL('./fixtures/native-media-background-timing.xml', import.meta.url),
      'utf8',
    );
    part.data = new TextEncoder().encode(
      new TextDecoder().decode(part.data).replace(/<p:timing[\s\S]*?<\/p:timing>/, timing),
    );
    const loaded = await loadPresentation(await savePresentation(pres));
    const loadedShape = getSlideShapes(getSlides(loaded)[0]!)[0]!;

    expect(getShapeMediaPlayback(loadedShape)).toMatchObject({
      autoplay: true,
      loop: true,
      slideCount: 999,
      hideWhenStopped: true,
    });
    setShapeMediaPlayback(loadedShape, { autoplay: true, volume: 0.25 });
    expect(getShapeMediaPlayback(loadedShape)).toMatchObject({ autoplay: true, volume: 0.25 });
    const before = slideXml(loaded);
    expect(() => setShapeMediaPlayback(loadedShape, { autoplay: false })).toThrow(
      /background media timing/,
    );
    expect(slideXml(loaded)).toBe(before);
    expect(() => setShapeMediaPlayback(loadedShape, { delayMs: 250 })).toThrow(
      /background media timing/,
    );
    expect(slideXml(loaded)).toBe(before);
  });

  it.each([
    ['a nonzero playFrom offset', (xml: string) => xml.replace('playFrom(0.0)', 'playFrom(5.0)')],
    [
      'duplicate playFrom commands',
      (xml: string) => xml.replace(/(<p:cmd type="call"[\s\S]*?<\/p:cmd>)/, '$1$1'),
    ],
    [
      'a delayed intermediary group',
      (xml: string) =>
        xml.replace('<p:cond delay="0"/></p:stCondLst>', '<p:cond delay="100"/></p:stCondLst>'),
    ],
    [
      'a referenced intermediary condition',
      (xml: string) =>
        xml.replace(
          '<p:cond delay="0"/></p:stCondLst>',
          '<p:cond delay="0"><p:tn val="9"/></p:cond></p:stCondLst>',
        ),
    ],
    [
      'a later main-sequence group',
      (xml: string) =>
        xml.replace(
          '<p:childTnLst>\n                <p:par>',
          '<p:childTnLst>\n                <p:par><p:cTn id="8"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn></p:par>\n                <p:par>',
        ),
    ],
  ])('keeps %s out of automatic playback recognition', async (_name, transform) => {
    const { pres } = await nestedDeck({}, {});
    const part = _internalPackageOf(pres).getPart(partName('/ppt/slides/slide1.xml'))!;
    const timing = await readFile(
      new URL('./fixtures/native-media-background-timing.xml', import.meta.url),
      'utf8',
    );
    part.data = new TextEncoder().encode(
      new TextDecoder()
        .decode(part.data)
        .replace(/<p:timing[\s\S]*?<\/p:timing>/, transform(timing)),
    );
    const loaded = await loadPresentation(await savePresentation(pres));
    const loadedShape = getSlideShapes(getSlides(loaded)[0]!)[0]!;
    expect(getShapeMediaPlayback(loadedShape)?.autoplay).toBe(false);
    expect(() => setShapeMediaPlayback(loadedShape, { autoplay: true })).toThrow(
      /command timing is unsupported/,
    );
  });

  it('converts its own manual start while retaining an automatic parent delay', async () => {
    const { pres, shape } = await nestedDeck(
      { attrs: 'delay="100"' },
      { attrs: 'delay="indefinite"' },
    );
    setShapeMediaPlayback(shape, { autoplay: true, delayMs: 350 });
    expect(getShapeMediaPlayback(shape)).toMatchObject({ autoplay: true, delayMs: 350 });
    expect(slideXml(pres)).toContain('<p:cond delay="100"/>');
    expect(slideXml(pres)).toContain('<p:cond delay="250"/>');
  });

  it.each(['2', '4'])('does not flatten masterRel on cTn %s', async (id) => {
    const { pres, shape } = await nestedDeck({}, {}, 'par', (xml) =>
      xml.replace(`id="${id}"`, `id="${id}" masterRel="nextClick"`),
    );
    expect(getShapeMediaPlayback(shape)?.autoplay).toBe(false);
    const before = slideXml(pres);
    expect(() => setShapeMediaPlayback(shape, { autoplay: true, volume: 0.2 })).toThrow();
    expect(slideXml(pres)).toBe(before);
  });

  it('does not interpret a referenced root-level media condition as autoplay', async () => {
    const { pres, shape } = await nestedDeck({}, { child: '<p:tn val="1"/>' }, 'par', (xml) =>
      xml
        .replace(/<p:par><p:cTn id="2">[\s\S]*?<p:childTnLst>/, '')
        .replace('</p:audio></p:childTnLst></p:cTn></p:par>', '</p:audio>'),
    );
    expect(getShapeMediaPlayback(shape)?.autoplay).toBe(false);
    const before = slideXml(pres);
    expect(() => setShapeMediaPlayback(shape, { autoplay: true, volume: 0.2 })).toThrow();
    expect(slideXml(pres)).toBe(before);
  });

  it('preserves subordinate timing instead of flattening it into autoplay', async () => {
    const { pres, shape } = await nestedDeck({}, {}, 'par', (xml) =>
      xml
        .replace('<p:childTnLst><p:audio>', '<p:subTnLst><p:audio>')
        .replace('</p:audio></p:childTnLst>', '</p:audio></p:subTnLst>')
        .replace('id="4"', 'id="4" masterRel="nextClick"'),
    );
    expect(getShapeMediaPlayback(shape)?.autoplay).toBe(false);
    const before = slideXml(pres);
    expect(() => setShapeMediaPlayback(shape, { autoplay: true, volume: 0.2 })).toThrow();
    expect(slideXml(pres)).toBe(before);
    setShapeMediaPlayback(shape, { volume: 0.2 });
    expect(getShapeMediaPlayback(shape)?.volume).toBe(0.2);
    expect(slideXml(pres)).toContain('masterRel="nextClick"');
  });

  it('rejects ambiguous duplicate media edits without changing either node', async () => {
    const { pres, shape } = await nestedDeck({}, {}, 'par', (xml) =>
      xml.replace(
        /(<p:audio>[\s\S]*?<\/p:audio>)/,
        (node) => node + node.replace('id="4"', 'id="5"'),
      ),
    );
    expect(getShapeMediaPlayback(shape)?.autoplay).toBe(false);
    const before = slideXml(pres);
    expect(() => setShapeMediaPlayback(shape, { volume: 0.2, muted: true })).toThrow(
      /multiple media/,
    );
    expect(slideXml(pres)).toBe(before);
  });

  it('rejects a total delay earlier than its parent without applying other edits', async () => {
    const { pres, shape } = await nestedDeck({ attrs: 'delay="100"' }, { attrs: 'delay="250"' });
    const before = slideXml(pres);
    expect(() => setShapeMediaPlayback(shape, { delayMs: 50, volume: 0.2 })).toThrow(/delayMs/);
    expect(slideXml(pres)).toBe(before);
  });

  skipIfNoXmllint('writes schema-valid nested timing after changing the total delay', async () => {
    const { pres, shape } = await nestedDeck({ attrs: 'delay="100"' }, { attrs: 'delay="250"' });
    setShapeMediaPlayback(shape, { autoplay: true, delayMs: 500 });
    const saved = await loadPresentation(await savePresentation(pres));
    expectSchemaValid(slideXml(saved), 'pml');
  });
});
