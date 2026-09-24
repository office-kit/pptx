import { expect, it } from 'vitest';
import * as pptx from '../src/api/index.ts';
import { INTERNAL_PACKAGE, SLIDE_DOCUMENT, SLIDE_PART_NAME } from '../src/api/_internal-symbols.ts';
import { NS, parseXml } from '../src/internal/xml/index.ts';
const wav = (sample: number) => {
  const bytes = new Uint8Array(46);
  const view = new DataView(bytes.buffer);
  for (const [offset, text] of [
    [0, 'RIFF'],
    [8, 'WAVE'],
    [12, 'fmt '],
    [36, 'data'],
  ] as const)
    bytes.set(new TextEncoder().encode(text), offset);
  view.setUint32(4, 38, true);
  view.setUint32(16, 16, true);
  view.setUint16(20, 1, true);
  view.setUint16(22, 1, true);
  view.setUint32(24, 8000, true);
  view.setUint32(28, 16000, true);
  view.setUint16(32, 2, true);
  view.setUint16(34, 16, true);
  view.setUint32(40, 2, true);
  view.setInt16(44, sample, true);
  return bytes;
};

it('round-trips transition sound, looping, stop and removal while retaining visual settings', async () => {
  const p = pptx.createPresentation();
  const slide = pptx.addBlankSlide(p);
  const other = pptx.addBlankSlide(p);
  pptx.setSlideTransition(slide, {
    effect: 'push',
    direction: 'r',
    durationMs: 1350,
    advanceAfterMs: 4000,
  });
  const original = pptx.getSlideTransition(slide);
  const sound = { kind: 'play' as const, name: 'Chime', bytes: wav(20), loop: true };
  pptx.setSlideTransitionSound(slide, sound);
  expect(pptx.getSlideTransition(slide)).toEqual(original);
  expect(pptx.getSlideTransitionSound(slide)).toEqual(sound);
  pptx.setSlideTransitionSound(slide, sound);
  expect(
    p[INTERNAL_PACKAGE]
      .getRels(slide[SLIDE_PART_NAME])!
      .items.filter((r) => r.type.endsWith('/audio')),
  ).toHaveLength(1);
  pptx.applySlideTransitionToAll(p, slide);
  expect(pptx.getSlideTransitionSound(other)).toEqual(sound);
  const loaded = await pptx.loadPresentation(await pptx.savePresentation(p));
  const copy = pptx.getSlides(loaded)[0]!;
  expect(pptx.getSlideTransitionSound(copy)).toEqual(sound);
  const detached = pptx.getSlideTransitionSound(copy);
  if (detached.kind === 'play') detached.bytes.fill(0);
  expect(pptx.getSlideTransitionSound(copy)).toEqual(sound);
  pptx.setSlideTransitionSound(copy, { kind: 'stop' });
  expect(pptx.getSlideTransitionSound(copy)).toEqual({ kind: 'stop' });
  pptx.setSlideTransitionSound(copy, { kind: 'none' });
  expect(pptx.getSlideTransitionSound(copy)).toEqual({ kind: 'none' });
  expect(pptx.getSlideTransition(copy)).toEqual(original);
});

it('rejects invalid sound before mutation and updates compatibility branches', async () => {
  const p = pptx.createPresentation();
  const slide = pptx.addBlankSlide(p);
  const before = await pptx.savePresentation(p);
  expect(() =>
    pptx.setSlideTransitionSound(slide, {
      kind: 'play',
      name: 'Bad',
      bytes: new Uint8Array([1]),
      loop: false,
    }),
  ).toThrow(/WAV/);
  expect(await pptx.savePresentation(p)).toEqual(before);
  slide[SLIDE_DOCUMENT].root.children.push(
    parseXml(
      `<mc:AlternateContent xmlns:mc="${NS.mc}" xmlns:p="${NS.pml}" xmlns:p14="${NS.p14}"><mc:Choice Requires="p14"><p:transition><p:fade/><p:extLst/></p:transition></mc:Choice><mc:Fallback><p:transition><p:fade/></p:transition></mc:Fallback></mc:AlternateContent>`,
    ).root,
  );
  pptx.setSlideTransitionSound(slide, { kind: 'play', name: 'Loop', bytes: wav(5), loop: false });
  const xml = new TextDecoder().decode(p[INTERNAL_PACKAGE].getPart(slide[SLIDE_PART_NAME])!.data);
  expect(xml.match(/<p:stSnd/g)).toHaveLength(2);
  expect(xml).toMatch(/<p:sndAc>.*<\/p:sndAc><p:extLst/s);
  pptx.setSlideTransitionSound(slide, { kind: 'none' });
  const removed = new TextDecoder().decode(
    p[INTERNAL_PACKAGE].getPart(slide[SLIDE_PART_NAME])!.data,
  );
  expect(removed).not.toContain('sndAc');
  expect(removed).toContain('p:extLst');
});
