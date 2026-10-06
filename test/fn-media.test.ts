// addSlideMedia / getShapeMedia — video, audio and online-video authoring,
// read-back, and how the clip survives copy / duplicate / import / remove.

import { describe, expect, it } from 'vitest';
import { expectSchemaValid, isSchemaValidationAvailable } from './lib/expect-schema-valid.ts';
import { buildPng } from './lib/build-png.ts';
import { serializeContentTypes } from '../src/internal/opc/index.ts';
import {
  type PresentationData,
  type SlideData,
  _internalPackageOf,
  addBlankSlide,
  addSlideMedia,
  addSlideShape,
  clearSlideAnimations,
  clearSlideShapes,
  compactPackage,
  copyShape,
  createPresentation,
  duplicateSlide,
  findShapesWithMedia,
  getMediaParts,
  getShapeAnimation,
  getShapeId,
  getShapeImageBytes,
  getShapeKind,
  getShapeMedia,
  getShapeName,
  getSlideLayouts,
  getSlideShapes,
  getSlides,
  importSlide,
  inches,
  loadPresentation,
  readPackagePart,
  removeShape,
  savePresentation,
  setShapeAnimation,
  setShapeImage,
  validatePresentation,
} from '../src/api/index.ts';

const REL_VIDEO = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/video';
const REL_AUDIO = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio';
const REL_MEDIA = 'http://schemas.microsoft.com/office/2007/relationships/media';
const REL_IMAGE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';

const decoder = new TextDecoder();
const encoder = new TextEncoder();
const skipIfNoXmllint = isSchemaValidationAvailable() ? it : it.skip;

const ascii = (s: string): number[] => Array.from(s, (ch) => ch.charCodeAt(0));
// Container headers only — enough for signature detection, not playable.
const mp4 = (tag = 0): Uint8Array =>
  new Uint8Array([0, 0, 0, 0x18, ...ascii('ftypmp42'), 0, 0, 0, 0, ...ascii('mp42isom'), tag]);
const mov = (): Uint8Array => new Uint8Array([0, 0, 0, 0x14, ...ascii('ftypqt  '), 0, 0, 0, 0]);
const webm = (): Uint8Array =>
  new Uint8Array([0x1a, 0x45, 0xdf, 0xa3, 0x9f, 0x42, 0x82, 0x84, ...ascii('webm'), 0x42, 0x87]);
const mkv = (): Uint8Array =>
  new Uint8Array([0x1a, 0x45, 0xdf, 0xa3, 0x9f, 0x42, 0x82, 0x88, ...ascii('matroska')]);
const mp3 = (): Uint8Array => new Uint8Array([...ascii('ID3'), 3, 0, 0, 0, 0, 0, 0, 0xff, 0xfb]);
const wav = (): Uint8Array =>
  new Uint8Array([...ascii('RIFF'), 0x24, 0, 0, 0, ...ascii('WAVEfmt ')]);

const box = { x: inches(1), y: inches(1), w: inches(4), h: inches(2.25) };

const newDeck = (): { pres: PresentationData; slide: SlideData } => {
  const pres = createPresentation();
  return { pres, slide: addBlankSlide(pres) };
};

const partText = (pres: PresentationData, name: string): string => {
  const bytes = readPackagePart(pres, name);
  if (bytes === null) throw new Error(`missing part ${name}`);
  return decoder.decode(bytes);
};

interface Rel {
  id: string;
  type: string;
  target: string;
  external: boolean;
}
const readRels = (pres: PresentationData, relsPart: string): Rel[] => {
  const xml = partText(pres, relsPart);
  return [...xml.matchAll(/<Relationship\b[^>]*>/g)].map(([tag]) => ({
    id: /\bId="([^"]*)"/.exec(tag)![1]!,
    type: /\bType="([^"]*)"/.exec(tag)![1]!,
    target: /\bTarget="([^"]*)"/.exec(tag)![1]!,
    external: /TargetMode="External"/.test(tag),
  }));
};

// `[Content_Types].xml` only materialises on save; read the live model instead.
const contentTypesXml = (pres: PresentationData): string =>
  serializeContentTypes(_internalPackageOf(pres).contentTypes);

const SLIDE1 = '/ppt/slides/slide1.xml';
const SLIDE1_RELS = '/ppt/slides/_rels/slide1.xml.rels';

describe('addSlideMedia: embedded video', () => {
  it('writes the PowerPoint wire shape: pic + two rels to one part + poster + time node', () => {
    const { pres, slide } = newDeck();
    const shape = addSlideMedia(slide, { kind: 'video', data: mp4(), ...box });

    expect(getShapeKind(shape)).toBe('picture');
    expect(getShapeName(shape)).toBe(`Video ${getShapeId(shape)}`);

    const rels = readRels(pres, SLIDE1_RELS);
    const media = rels.filter((r) => r.type === REL_MEDIA);
    const video = rels.filter((r) => r.type === REL_VIDEO);
    const image = rels.filter((r) => r.type === REL_IMAGE);
    expect(media).toHaveLength(1);
    expect(video).toHaveLength(1);
    expect(image).toHaveLength(1);
    expect(media[0]!.target).toBe('../media/media1.mp4');
    expect(video[0]!.target).toBe('../media/media1.mp4');
    expect(image[0]!.target).toBe('../media/image1.png');
    expect(new Set(rels.map((r) => r.id)).size).toBe(rels.length);

    const xml = partText(pres, SLIDE1);
    expect(xml).toContain('<a:hlinkClick r:id="" action="ppaction://media"/>');
    expect(xml).toContain(`<a:videoFile r:link="${video[0]!.id}"/>`);
    expect(xml).toContain('<p:ext uri="{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}">');
    expect(xml).toMatch(
      new RegExp(
        `<p14:media xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" r:embed="${media[0]!.id}"/>`,
      ),
    );
    expect(xml).toContain(`<a:blip r:embed="${image[0]!.id}"/>`);
    expect(xml).toMatch(
      new RegExp(
        `<p:video><p:cMediaNode vol="80000"><p:cTn id="2" fill="hold" display="0">.*<p:spTgt spid="${getShapeId(shape)}"/>`,
      ),
    );

    const contentTypes = contentTypesXml(pres);
    expect(contentTypes).toContain('<Default Extension="mp4" ContentType="video/mp4"/>');
    expect(contentTypes).toContain('<Default Extension="png" ContentType="image/png"/>');
    expect(validatePresentation(pres).filter((i) => i.severity === 'error')).toEqual([]);
  });

  skipIfNoXmllint('emits schema-valid slide XML for every kind', () => {
    const { pres, slide } = newDeck();
    addSlideMedia(slide, { kind: 'video', data: mp4(), ...box });
    addSlideMedia(slide, { kind: 'audio', data: mp3(), ...box });
    addSlideMedia(slide, { kind: 'online', url: 'https://youtu.be/dQw4w9WgXcQ', ...box });
    expectSchemaValid(partText(pres, SLIDE1), 'pml');
    expectSchemaValid(partText(pres, SLIDE1_RELS), 'rels');
    expectSchemaValid(contentTypesXml(pres), 'contentTypes');
  });

  it('round-trips through save → load with identical read-back', async () => {
    const { pres, slide } = newDeck();
    const poster = buildPng(4, 4, [200, 30, 30]);
    addSlideMedia(slide, { kind: 'video', data: mp4(), poster, name: 'Demo clip', ...box });

    const reloaded = await loadPresentation(await savePresentation(pres));
    const [shape] = findShapesWithMedia(getSlides(reloaded)[0]!);
    expect(getShapeName(shape!)).toBe('Demo clip');
    const media = getShapeMedia(shape!);
    expect(media).toMatchObject({
      kind: 'video',
      partName: '/ppt/media/media1.mp4',
      contentType: 'video/mp4',
    });
    expect(media?.kind === 'video' && Array.from(media.bytes)).toEqual(Array.from(mp4()));
    expect(Array.from(getShapeImageBytes(shape!)!)).toEqual(Array.from(poster));
  });

  it('stores identical clip bytes once and gives distinct clips distinct parts', () => {
    const { pres, slide } = newDeck();
    const second = addBlankSlide(pres);
    addSlideMedia(slide, { kind: 'video', data: mp4(), ...box });
    addSlideMedia(slide, { kind: 'video', data: mp4(), ...box });
    addSlideMedia(second, { kind: 'video', data: mp4(), ...box });
    addSlideMedia(second, { kind: 'video', data: mp4(1), ...box });

    const clips = getMediaParts(pres)
      .map((p) => p.name)
      .filter((n) => n.includes('/media/media'));
    expect(clips.sort()).toEqual(['/ppt/media/media1.mp4', '/ppt/media/media2.mp4']);

    // Same clip twice on one slide shares its two relationships as well.
    const rels = readRels(pres, SLIDE1_RELS);
    expect(rels.filter((r) => r.type === REL_MEDIA)).toHaveLength(1);
    expect(rels.filter((r) => r.type === REL_VIDEO)).toHaveLength(1);
    // Posters stay per-shape so replacing one does not repaint the others.
    expect(rels.filter((r) => r.type === REL_IMAGE)).toHaveLength(2);
  });

  it('replacing one poster leaves the other clip poster untouched', () => {
    const { slide } = newDeck();
    const a = addSlideMedia(slide, { kind: 'video', data: mp4(), ...box });
    const b = addSlideMedia(slide, { kind: 'video', data: mp4(), ...box });
    const before = Array.from(getShapeImageBytes(b)!);
    setShapeImage(a, buildPng(2, 2, [0, 0, 255]));
    expect(Array.from(getShapeImageBytes(getSlideShapes(slide)[1]!)!)).toEqual(before);
    expect(getShapeMedia(getSlideShapes(slide)[0]!)?.kind).toBe('video');
  });

  it('detects containers and honours an explicit format', () => {
    const { pres, slide } = newDeck();
    addSlideMedia(slide, { kind: 'video', data: mov(), ...box });
    addSlideMedia(slide, { kind: 'video', data: webm(), ...box });
    addSlideMedia(slide, { kind: 'video', data: new Uint8Array([1, 2, 3]), format: 'wmv', ...box });
    const names = getMediaParts(pres).map((p) => p.name);
    expect(names).toEqual(
      expect.arrayContaining([
        '/ppt/media/media1.mov',
        '/ppt/media/media2.webm',
        '/ppt/media/media3.wmv',
      ]),
    );
    const contentTypes = contentTypesXml(pres);
    expect(contentTypes).toContain('<Default Extension="mov" ContentType="video/quicktime"/>');
    expect(contentTypes).toContain('<Default Extension="webm" ContentType="video/webm"/>');
    expect(contentTypes).toContain('<Default Extension="wmv" ContentType="video/x-ms-wmv"/>');
  });
});

describe('addSlideMedia: embedded audio', () => {
  it('uses audioFile, the audio rel type and a <p:audio> time node', async () => {
    const { pres, slide } = newDeck();
    const shape = addSlideMedia(slide, { kind: 'audio', data: mp3(), ...box });

    const rels = readRels(pres, SLIDE1_RELS);
    const audio = rels.filter((r) => r.type === REL_AUDIO);
    expect(audio).toHaveLength(1);
    expect(audio[0]!.target).toBe('../media/media1.mp3');
    expect(rels.filter((r) => r.type === REL_MEDIA)).toHaveLength(1);
    expect(rels.filter((r) => r.type === REL_VIDEO)).toHaveLength(0);

    const xml = partText(pres, SLIDE1);
    expect(xml).toContain(`<a:audioFile r:link="${audio[0]!.id}"/>`);
    expect(xml).toContain('<p:audio><p:cMediaNode');
    expect(xml).not.toContain('<p:video>');
    expect(getShapeName(shape)).toBe(`Audio ${getShapeId(shape)}`);
    expect(contentTypesXml(pres)).toContain('<Default Extension="mp3" ContentType="audio/mpeg"/>');

    const reloaded = await loadPresentation(await savePresentation(pres));
    const media = getShapeMedia(getSlideShapes(getSlides(reloaded)[0]!)[0]!);
    expect(media).toMatchObject({ kind: 'audio', contentType: 'audio/mpeg' });
  });

  it('detects wav', () => {
    const { pres, slide } = newDeck();
    addSlideMedia(slide, { kind: 'audio', data: wav(), ...box });
    expect(getMediaParts(pres).map((p) => p.name)).toContain('/ppt/media/media1.wav');
  });
});

describe('addSlideMedia: online video', () => {
  it.each([
    ['https://www.youtube.com/watch?v=dQw4w9WgXcQ&t=42s'],
    ['https://youtu.be/dQw4w9WgXcQ?si=abc'],
    ['https://m.youtube.com/shorts/dQw4w9WgXcQ'],
  ])('rewrites the YouTube page URL %s to the embed form', (url) => {
    const { slide } = newDeck();
    const shape = addSlideMedia(slide, { kind: 'online', url, ...box });
    expect(getShapeMedia(shape)).toEqual({
      kind: 'online',
      url: 'https://www.youtube.com/embed/dQw4w9WgXcQ?feature=oembed',
    });
  });

  it('keeps embed and non-YouTube URLs as given, as an external video rel', async () => {
    const { pres, slide } = newDeck();
    const embed = 'https://www.youtube.com/embed/dQw4w9WgXcQ?start=30';
    const vimeo = 'https://player.vimeo.com/video/76979871';
    addSlideMedia(slide, { kind: 'online', url: embed, ...box });
    const shape = addSlideMedia(slide, { kind: 'online', url: vimeo, ...box });
    expect(getShapeName(shape)).toBe(`Online Media ${getShapeId(shape)}`);

    const rels = readRels(pres, SLIDE1_RELS);
    const video = rels.filter((r) => r.type === REL_VIDEO);
    expect(video.map((r) => [r.target, r.external])).toEqual([
      [embed, true],
      [vimeo, true],
    ]);
    expect(rels.filter((r) => r.type === REL_MEDIA)).toHaveLength(0);
    expect(partText(pres, SLIDE1)).not.toContain('p14:media');
    // No clip part — only the two posters.
    expect(getMediaParts(pres).map((p) => p.name)).toEqual([
      '/ppt/media/image1.png',
      '/ppt/media/image2.png',
    ]);

    const reloaded = await loadPresentation(await savePresentation(pres));
    expect(findShapesWithMedia(getSlides(reloaded)[0]!).map((s) => getShapeMedia(s))).toEqual([
      { kind: 'online', url: embed },
      { kind: 'online', url: vimeo },
    ]);
  });
});

describe('addSlideMedia: rejected input leaves the package untouched', () => {
  const cases: Array<[string, (slide: SlideData) => void, RegExp]> = [
    [
      'unknown video container',
      (s) => addSlideMedia(s, { kind: 'video', data: new Uint8Array([1, 2, 3, 4]), ...box }),
      /could not detect the video format/,
    ],
    [
      'matroska is not webm',
      (s) => addSlideMedia(s, { kind: 'video', data: mkv(), ...box }),
      /could not detect the video format/,
    ],
    [
      'unknown audio container',
      (s) => addSlideMedia(s, { kind: 'audio', data: new Uint8Array([1, 2, 3, 4]), ...box }),
      /could not detect the audio format/,
    ],
    [
      'unreadable poster',
      (s) =>
        addSlideMedia(s, { kind: 'video', data: mp4(), poster: new Uint8Array([9, 9]), ...box }),
      /poster image format/,
    ],
    [
      'relative url',
      (s) => addSlideMedia(s, { kind: 'online', url: '/watch?v=abc', ...box }),
      /not an absolute URL/,
    ],
    [
      'non-http url',
      (s) => addSlideMedia(s, { kind: 'online', url: 'file:///C:/clip.mp4', ...box }),
      /must be http\(s\)/,
    ],
    [
      'negative width',
      (s) => addSlideMedia(s, { kind: 'video', data: mp4(), ...box, w: inches(-1) }),
      /addSlideMedia: w/,
    ],
  ];
  it.each(cases)('%s', (_label, run, message) => {
    const { pres, slide } = newDeck();
    const partsBefore = _internalPackageOf(pres).parts.map((p) => p.name);
    const xmlBefore = partText(pres, SLIDE1);
    expect(() => run(slide)).toThrow(message);
    expect(_internalPackageOf(pres).parts.map((p) => p.name)).toEqual(partsBefore);
    expect(partText(pres, SLIDE1)).toBe(xmlBefore);
  });
});

describe('getShapeMedia', () => {
  it('returns null for plain shapes and pictures', () => {
    const { slide } = newDeck();
    const rect = addSlideShape(slide, { preset: 'rect', ...box });
    expect(getShapeMedia(rect)).toBeNull();
    expect(findShapesWithMedia(slide)).toEqual([]);
  });

  // PptxGenJS 4.x writes no time node, no hlinkClick on online videos, and no
  // <p14:media> for them; the embedded case is only reachable via r:link when
  // the extension list is absent (PowerPoint 2007-era files).
  it('reads a foreign-authored deck: r:link only, no p14:media, no timing', async () => {
    const { pres, slide } = newDeck();
    addSlideMedia(slide, { kind: 'video', data: mp4(), ...box });
    addSlideMedia(slide, { kind: 'online', url: 'https://example.com/v', ...box });
    const pkg = _internalPackageOf(pres);
    const part = pkg.parts.find((p) => p.name === SLIDE1)!;
    const stripped = decoder
      .decode(part.data)
      .replace(/<p:timing>.*<\/p:timing>/s, '')
      .replace(/<p:extLst><p:ext uri="\{DAA4B4D4[^]*?<\/p:extLst>/, '')
      .replaceAll(/<a:hlinkClick[^>]*\/>/g, '');
    expect(stripped).not.toContain('p14:media');
    part.data = encoder.encode(stripped);

    const reloaded = await loadPresentation(await savePresentation(pres));
    const media = findShapesWithMedia(getSlides(reloaded)[0]!).map((s) => getShapeMedia(s));
    expect(media[0]).toMatchObject({ kind: 'video', partName: '/ppt/media/media1.mp4' });
    expect(media[1]).toEqual({ kind: 'online', url: 'https://example.com/v' });
  });
});

const timingOf = (pres: PresentationData, slidePart: string): string =>
  /<p:timing>.*<\/p:timing>/s.exec(partText(pres, slidePart))?.[0] ?? '';

describe('media shapes through shape / slide operations', () => {
  it('removeShape drops the clip time node, and the timing element with the last one', () => {
    const { pres, slide } = newDeck();
    const video = addSlideMedia(slide, { kind: 'video', data: mp4(), ...box });
    addSlideMedia(slide, { kind: 'audio', data: mp3(), ...box });
    const videoId = getShapeId(video);

    removeShape(video);
    expect(timingOf(pres, SLIDE1)).not.toContain(`spid="${videoId}"`);
    expect(timingOf(pres, SLIDE1)).toContain('<p:audio>');
    expect(validatePresentation(pres).filter((i) => i.severity === 'error')).toEqual([]);

    removeShape(getSlideShapes(slide)[0]!);
    expect(partText(pres, SLIDE1)).not.toContain('<p:timing');
  });

  it('clearSlideShapes drops every clip time node', () => {
    const { pres, slide } = newDeck();
    addSlideMedia(slide, { kind: 'video', data: mp4(), ...box });
    clearSlideShapes(slide);
    expect(partText(pres, SLIDE1)).not.toContain('<p:timing');
  });

  it('compactPackage keeps a referenced clip and poster', () => {
    const { pres, slide } = newDeck();
    addSlideMedia(slide, { kind: 'video', data: mp4(), ...box });
    expect(compactPackage(pres).removed).toEqual([]);
    expect(getMediaParts(pres).map((p) => p.name)).toEqual([
      '/ppt/media/media1.mp4',
      '/ppt/media/image1.png',
    ]);
  });

  skipIfNoXmllint('copyShape carries both clip rels, the poster and a time node', () => {
    const { pres, slide } = newDeck();
    const target = addBlankSlide(pres);
    const source = addSlideMedia(slide, { kind: 'video', data: mp4(), ...box });
    const online = addSlideMedia(slide, { kind: 'online', url: 'https://example.com/v', ...box });

    const copy = copyShape(target, source);
    const onlineCopy = copyShape(target, online);

    expect(getShapeMedia(copy)).toMatchObject({ kind: 'video', partName: '/ppt/media/media1.mp4' });
    expect(getShapeMedia(onlineCopy)).toEqual({ kind: 'online', url: 'https://example.com/v' });
    const rels = readRels(pres, '/ppt/slides/_rels/slide2.xml.rels');
    expect(rels.filter((r) => r.type === REL_MEDIA)).toHaveLength(1);
    expect(rels.filter((r) => r.type === REL_VIDEO)).toHaveLength(2);
    expect(rels.find((r) => r.external)?.target).toBe('https://example.com/v');
    // No second copy of the clip bytes.
    expect(getMediaParts(pres).filter((p) => p.name.includes('/media/media'))).toHaveLength(1);

    const timing = timingOf(pres, '/ppt/slides/slide2.xml');
    expect(timing).toContain(`<p:spTgt spid="${getShapeId(copy)}"/>`);
    expect(timing).toContain(`<p:spTgt spid="${getShapeId(onlineCopy)}"/>`);
    expectSchemaValid(partText(pres, '/ppt/slides/slide2.xml'), 'pml');
    expect(validatePresentation(pres).filter((i) => i.severity === 'error')).toEqual([]);
  });

  it('duplicateSlide shares the clip part and keeps the external rel', async () => {
    const { pres, slide } = newDeck();
    addSlideMedia(slide, { kind: 'video', data: mp4(), ...box });
    addSlideMedia(slide, { kind: 'online', url: 'https://example.com/v', ...box });
    const partsBefore = getMediaParts(pres).length;

    const dup = duplicateSlide(pres, slide);
    expect(getMediaParts(pres)).toHaveLength(partsBefore);
    expect(findShapesWithMedia(dup).map((s) => getShapeMedia(s))).toEqual(
      findShapesWithMedia(slide).map((s) => getShapeMedia(s)),
    );
    expect(timingOf(pres, '/ppt/slides/slide2.xml')).toBe(timingOf(pres, SLIDE1));

    const reloaded = await loadPresentation(await savePresentation(pres));
    expect(findShapesWithMedia(getSlides(reloaded)[1]!)).toHaveLength(2);
    expect(validatePresentation(reloaded).filter((i) => i.severity === 'error')).toEqual([]);
  });

  it('importSlide copies the clip once and keeps both rels pointing at it', async () => {
    const { slide } = newDeck();
    addSlideMedia(slide, { kind: 'audio', data: mp3(), ...box });
    addSlideMedia(slide, { kind: 'online', url: 'https://example.com/v', ...box });

    const targetPres = createPresentation();
    const imported = importSlide(targetPres, slide, getSlideLayouts(targetPres)[0]!);

    const media = findShapesWithMedia(imported).map((s) => getShapeMedia(s));
    expect(media[0]).toMatchObject({ kind: 'audio', partName: '/ppt/media/media1-copy1.mp3' });
    expect(media[1]).toEqual({ kind: 'online', url: 'https://example.com/v' });
    const rels = readRels(targetPres, SLIDE1_RELS);
    expect(rels.filter((r) => r.type === REL_MEDIA).map((r) => r.target)).toEqual([
      '/ppt/media/media1-copy1.mp3',
    ]);
    expect(rels.filter((r) => r.type === REL_AUDIO).map((r) => r.target)).toEqual([
      '/ppt/media/media1-copy1.mp3',
    ]);
    expect(getMediaParts(targetPres).filter((p) => p.name.endsWith('.mp3'))).toHaveLength(1);
    expect(contentTypesXml(targetPres)).toContain('ContentType="audio/mpeg"');

    const reloaded = await loadPresentation(await savePresentation(targetPres));
    expect(validatePresentation(reloaded).filter((i) => i.severity === 'error')).toEqual([]);
  });
});

describe('media time nodes and animations share <p:timing>', () => {
  skipIfNoXmllint('animating a shape on a slide that already holds a clip', () => {
    const { pres, slide } = newDeck();
    addSlideMedia(slide, { kind: 'video', data: mp4(), ...box });
    const rect = addSlideShape(slide, { preset: 'rect', ...box });
    setShapeAnimation(rect, { effect: 'fadeIn' });

    expect(getShapeAnimation(getSlideShapes(slide)[1]!)).toBe('fadeIn');
    const timing = timingOf(pres, SLIDE1);
    expect(timing).toContain('<p:video>');
    expect(timing.indexOf('nodeType="mainSeq"')).toBeLessThan(timing.indexOf('<p:video>'));
    const ids = [...timing.matchAll(/<p:cTn id="(\d+)"/g)].map((m) => m[1]);
    expect(new Set(ids).size).toBe(ids.length);
    expectSchemaValid(partText(pres, SLIDE1), 'pml');
  });

  skipIfNoXmllint('adding a clip to a slide that already animates', () => {
    const { pres, slide } = newDeck();
    const rect = addSlideShape(slide, { preset: 'rect', ...box });
    setShapeAnimation(rect, { effect: 'appear' });
    const clip = addSlideMedia(slide, { kind: 'audio', data: mp3(), ...box });

    const timing = timingOf(pres, SLIDE1);
    expect(timing).toContain(`<p:spTgt spid="${getShapeId(clip)}"/>`);
    expect(getShapeAnimation(getSlideShapes(slide)[0]!)).toBe('appear');
    const ids = [...timing.matchAll(/<p:cTn id="(\d+)"/g)].map((m) => m[1]);
    expect(new Set(ids).size).toBe(ids.length);
    expectSchemaValid(partText(pres, SLIDE1), 'pml');
  });

  skipIfNoXmllint('clearSlideAnimations keeps the clip playable', () => {
    const { pres, slide } = newDeck();
    const clip = addSlideMedia(slide, { kind: 'video', data: mp4(), ...box });
    const rect = addSlideShape(slide, { preset: 'rect', ...box });
    setShapeAnimation(rect, { effect: 'fadeOut' });

    clearSlideAnimations(slide);
    const timing = timingOf(pres, SLIDE1);
    expect(timing).not.toContain('mainSeq');
    expect(timing).not.toContain('<p:bldLst>');
    expect(timing).toContain(`<p:spTgt spid="${getShapeId(clip)}"/>`);
    expect(getShapeAnimation(getSlideShapes(slide)[1]!)).toBeNull();
    expectSchemaValid(partText(pres, SLIDE1), 'pml');
  });
});

describe('validatePresentation: media', () => {
  it('flags a dangling clip rel and a time node without its shape', () => {
    const { pres, slide } = newDeck();
    addSlideMedia(slide, { kind: 'video', data: mp4(), ...box });
    const pkg = _internalPackageOf(pres);
    pkg.removePart(pkg.parts.find((p) => p.name === '/ppt/media/media1.mp4')!.name);
    const part = pkg.parts.find((p) => p.name === SLIDE1)!;
    part.data = encoder.encode(decoder.decode(part.data).replace(/<p:pic>.*<\/p:pic>/s, ''));

    const messages = validatePresentation(pres).map((i) => i.message);
    expect(messages.some((m) => /dangling media rel/.test(m))).toBe(true);
    expect(messages.some((m) => /dangling video rel/.test(m))).toBe(true);
    expect(messages.some((m) => /media time node .* targets missing shape/.test(m))).toBe(true);
  });
});
