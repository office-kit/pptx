// Every effect of the reference desktop app's Entrance, Emphasis and Exit galleries, written
// at its defaults and compared byte for byte with what the reference desktop app (Mac 16.113)
// itself saved for the same gallery entry (test/fixtures/native/animations/,
// provenance in test/fixtures/SOURCES.md).
//
// Each capture is one effect on a rectangle with a fill and a line of text
// (shape 2), so its build entry animates the background too; the text captures
// are Fly In on a three-paragraph text box with no fill, in each of the
// Effect Options' three Sequence choices.

import { readFile, readdir } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { expectSchemaValid, isSchemaValidationAvailable } from './lib/expect-schema-valid.ts';
import { partName } from '../src/internal/opc/index.ts';
import {
  type AnimationEffect,
  type AnimationOptions,
  _internalPackageOf,
  addSlideTextBox,
  emu,
  getShapeAnimation,
  getSlideAnimations,
  getSlideShapes,
  getSlideXmlString,
  getSlides,
  loadPresentation,
  savePresentation,
  setShapeAnimation,
  setShapeFill,
  updateSlideAnimation,
} from '../src/api/index.ts';
import {
  ANIMATION_EFFECTS,
  defaultAnimationDurationMs,
} from '../src/internal/presentationml/index.ts';

const NATIVE = new URL('./fixtures/native/animations/', import.meta.url);

// Capture number -> the token it stands for, spelled out here rather than read
// from the generator, so a slip in the generator's table cannot agree with
// itself.
const EFFECTS: readonly AnimationEffect[] = [
  'appear',
  'blindsIn',
  'checkerboardIn',
  'dissolveIn',
  'flyIn',
  'peekIn',
  'randomBarsIn',
  'shapeIn',
  'splitIn',
  'stripsIn',
  'wedgeIn',
  'wheelIn',
  'wipeIn',
  'expandIn',
  'fadeIn',
  'swivelIn',
  'zoomIn',
  'centerRevolveIn',
  'floatIn',
  'growTurnIn',
  'riseUpIn',
  'spinnerIn',
  'basicZoomIn',
  'stretchIn',
  'boomerangIn',
  'bounceIn',
  'creditsIn',
  'curveUpIn',
  'dropIn',
  'flipIn',
  'floatingIn',
  'pinwheelIn',
  'spiralIn',
  'basicSwivelIn',
  'whipIn',
  'blindsOut',
  'checkerboardOut',
  'disappear',
  'dissolveOut',
  'flyOut',
  'peekOut',
  'randomBarsOut',
  'shapeOut',
  'splitOut',
  'stripsOut',
  'wedgeOut',
  'wheelOut',
  'wipeOut',
  'contractOut',
  'fadeOut',
  'swivelOut',
  'zoomOut',
  'centerRevolveOut',
  'collapseOut',
  'floatOut',
  'shrinkTurnOut',
  'sinkDownOut',
  'spinnerOut',
  'basicZoomOut',
  'stretchyOut',
  'boomerangOut',
  'bounceOut',
  'creditsOut',
  'curveDownOut',
  'dropOut',
  'flipOut',
  'floatingOut',
  'pinwheelOut',
  'spiralOut',
  'basicSwivelOut',
  'whipOut',
  'fillColor',
  'fontColor',
  'growShrink',
  'lineColor',
  'spin',
  'transparency',
  'boldFlash',
  'brushColor',
  'complementaryColor',
  'complementaryColor2',
  'contrastingColor',
  'darken',
  'desaturate',
  'lighten',
  'objectColor',
  'pulse',
  'underline',
  'colorPulse',
  'growWithColor',
  'shimmer',
  'teeter',
  'blink',
  'boldReveal',
  'wave',
];

interface Capture {
  readonly file: string;
  readonly effect: AnimationEffect;
  readonly timing: string;
}

const timingOf = (xml: string): string => {
  const match = /<p:timing>[\s\S]*<\/p:timing>/.exec(xml);
  if (match === null) throw new Error('no <p:timing>');
  return match[0];
};

const captures = async (): Promise<Capture[]> => {
  const out: Capture[] = [];
  for (const group of ['entrance', 'exit', 'emphasis']) {
    const dir = new URL(`${group}/`, NATIVE);
    for (const file of (await readdir(dir)).sort()) {
      const number = Number.parseInt(file, 10);
      out.push({
        file: `${group}/${file}`,
        effect: EFFECTS[number - 1]!,
        timing: timingOf(await readFile(new URL(file, dir), 'utf8')),
      });
    }
  }
  return out;
};

const ALL = await captures();

const fixture = fileURLToPath(new URL('./fixtures/minimal/one-text-slide.pptx', import.meta.url));

/** A slide whose first shape draws a fill, as the captured rectangle does. */
const filledShape = async () => {
  const pres = await loadPresentation(await readFile(fixture));
  const slide = getSlides(pres)[0]!;
  const shape = getSlideShapes(slide)[0]!;
  setShapeFill(shape, '#4472C4');
  return { pres, slide, shape };
};

/** The slide's timing with the shape's id written as the capture's `2`. */
const writtenTiming = (xml: string, spid: number): string =>
  timingOf(xml).replaceAll(`spid="${spid}"`, 'spid="2"');

describe('animations: every gallery effect as the reference desktop app writes it', () => {
  it('covers the 95 effects of the three galleries, one capture each', () => {
    expect(ALL).toHaveLength(95);
    expect(new Set(ALL.map((c) => c.effect)).size).toBe(95);
    expect([...ANIMATION_EFFECTS].sort()).toEqual([...EFFECTS].sort());
  });

  it.each(ALL.map((c) => [c.file, c] as const))(
    '%s is written byte for byte',
    async (_, capture) => {
      const { slide, shape } = await filledShape();
      setShapeAnimation(shape, { effect: capture.effect });
      const spid = Number(
        /spid="(\d+)"/.exec(getSlideXmlString(slide).split('<p:timing>')[1]!)![1],
      );
      expect(writtenTiming(getSlideXmlString(slide), spid)).toBe(capture.timing);
    },
  );

  it.each(ALL.map((c) => [c.file, c] as const))(
    '%s reads back as the effect it is, with the reference desktop app’s duration',
    async (_, capture) => {
      const { pres, shape } = await filledShape();
      setShapeAnimation(shape, { effect: capture.effect });
      const reopened = getSlides(await loadPresentation(await savePresentation(pres)))[0]!;
      const [step] = getSlideAnimations(reopened);
      expect(step!.effect).toBe(capture.effect);
      expect(step!.playable).toBe(true);
      expect(step!.editable).toBe(true);
      expect(step!.build).toBe('asOneObject');
      expect(step!.durationMs).toBe(defaultAnimationDurationMs(capture.effect));
    },
  );

  it.each(ALL.map((c) => [c.file, c] as const))(
    '%s, as the reference desktop app saved it, reads as the effect it is',
    async (_, capture) => {
      const { pres } = await filledShape();
      const part = _internalPackageOf(pres).getPart(partName('/ppt/slides/slide1.xml'))!;
      const id = /<p:cNvPr id="(\d+)"/.exec(new TextDecoder().decode(part.data))![1];
      const native = capture.timing.replaceAll('spid="2"', `spid="${id}"`);
      const xml = new TextDecoder()
        .decode(part.data)
        .replace(/<p:timing>[\s\S]*<\/p:timing>/, '')
        .replace('</p:sld>', `${native}</p:sld>`);
      part.data = new TextEncoder().encode(xml);
      const reopened = getSlides(await loadPresentation(await savePresentation(pres)))[0]!;
      const steps = getSlideAnimations(reopened);
      expect(steps.map((s) => s.effect)).toEqual([capture.effect]);
      expect(steps[0]!.durationMs).toBe(defaultAnimationDurationMs(capture.effect));
      expect(steps[0]!.editable).toBe(true);
    },
  );

  it('scales every behaviour of a composite effect with its duration', async () => {
    const { slide, shape } = await filledShape();
    setShapeAnimation(shape, { effect: 'bounceIn', durationMs: 1000 });
    const xml = getSlideXmlString(slide);
    const durs = [...timingOf(xml).matchAll(/ dur="(\d+)"/g)].map((m) => Number(m[1]));
    // Bounce's own 2000ms halved: 580 -> 290, 1822 -> 911 … and the 1ms kick kept.
    expect(durs).toEqual([1, 290, 911, 332, 332, 166, 82, 13, 83, 13, 83, 13, 83, 13, 83]);
    expect(getSlideAnimations(slide)[0]!.durationMs).toBe(1000);
  });

  it('retimes an exit and keeps its hide at the end', async () => {
    const { slide, shape } = await filledShape();
    setShapeAnimation(shape, { effect: 'centerRevolveOut' });
    const id = getSlideAnimations(slide)[0]!.id!;
    updateSlideAnimation(slide, id, { durationMs: 2000 });
    const xml = timingOf(getSlideXmlString(slide));
    expect(xml).toContain('<p:cond delay="1999"/>');
    expect(xml).toContain('<p:cond delay="1800"/>');
    expect(getSlideAnimations(slide)[0]!.durationMs).toBe(2000);
  });

  it('refuses a duration for the effects that hold until the slide ends', async () => {
    const { shape } = await filledShape();
    expect(() => setShapeAnimation(shape, { effect: 'transparency', durationMs: 500 })).toThrow(
      /holds until the end of the slide/,
    );
    expect(defaultAnimationDurationMs('transparency')).toBeNull();
    expect(defaultAnimationDurationMs('boldReveal')).toBeNull();
  });

  it('measures a reversing behaviour twice over when chaining after it', async () => {
    const { slide, shape } = await filledShape();
    // Pulse grows for 250ms and shrinks back for another 250 (`autoRev`).
    setShapeAnimation(shape, { effect: 'pulse' });
    setShapeAnimation(shape, { effect: 'fadeOut', start: 'afterPrevious' });
    expect(timingOf(getSlideXmlString(slide))).toContain(
      '<p:par><p:cTn id="8" fill="hold"><p:stCondLst><p:cond delay="500"/>',
    );
  });

  it('refuses to chain after an effect that runs letter by letter', async () => {
    const { shape } = await filledShape();
    setShapeAnimation(shape, { effect: 'dropIn' });
    expect(() => setShapeAnimation(shape, { effect: 'fadeOut', start: 'afterPrevious' })).toThrow(
      /runs letter by letter/,
    );
  });

  it('Fill Color joins no build group, and is still the shape’s effect', async () => {
    const { pres, slide, shape } = await filledShape();
    setShapeAnimation(shape, { effect: 'fadeIn' });
    setShapeAnimation(shape, { effect: 'fillColor' });
    const xml = timingOf(getSlideXmlString(slide));
    expect(xml.match(/<p:bldP /g)).toHaveLength(1);
    const reopened = getSlides(await loadPresentation(await savePresentation(pres)))[0]!;
    expect(getSlideAnimations(reopened).map((s) => s.effect)).toEqual(['fadeIn', 'fillColor']);

    // Turning it into an effect that does build opens a group of its own, and
    // back again leaves no entry behind.
    const { slide: other, shape: lone } = await filledShape();
    setShapeAnimation(lone, { effect: 'fillColor' });
    expect(getShapeAnimation(lone)).toBe('fillColor');
    const id = getSlideAnimations(other)[0]!.id!;
    updateSlideAnimation(other, id, { effect: 'pulse' });
    expect(timingOf(getSlideXmlString(other))).toContain('grpId="0" animBg="1"/>');
    updateSlideAnimation(other, id, { effect: 'lineColor' });
    expect(timingOf(getSlideXmlString(other))).not.toContain('<p:bldLst>');
    expect(getSlideAnimations(other).map((s) => s.effect)).toEqual(['lineColor']);
  });

  it('a text effect builds the text alone', async () => {
    const { slide, shape } = await filledShape();
    setShapeAnimation(shape, { effect: 'fontColor' });
    expect(timingOf(getSlideXmlString(slide))).toContain('<p:bldP spid="2" grpId="0"/>');
  });

  it.skipIf(!isSchemaValidationAvailable())(
    'every effect is schema-valid, all on one slide',
    async () => {
      const { slide, shape } = await filledShape();
      for (const effect of EFFECTS) {
        const opts: AnimationOptions = { effect, start: 'click' };
        setShapeAnimation(shape, opts);
      }
      expectSchemaValid(getSlideXmlString(slide), 'pml');
      expect(getSlideAnimations(slide).map((s) => s.effect)).toEqual(EFFECTS);
    },
  );
});

describe('animations: text build (Effect Options › Sequence)', () => {
  const TEXT = {
    asOneObject: '96-fly-in-on-3-paragraph-text-box-gallery-default-sequence-as-one-object.xml',
    allAtOnce: '98-fly-in-on-3-paragraph-text-box-sequence-all-at-once.xml',
    byParagraph: '99-fly-in-on-3-paragraph-text-box-sequence-by-paragraph.xml',
  } as const;

  const textBox = async () => {
    const pres = await loadPresentation(await readFile(fixture));
    const slide = getSlides(pres)[0]!;
    const shape = addSlideTextBox(slide, {
      x: emu(0),
      y: emu(0),
      w: emu(3_000_000),
      h: emu(1_000_000),
      text: 'First\nSecond\nThird',
    });
    return { pres, slide, shape };
  };

  it.each(Object.entries(TEXT))('%s is written byte for byte', async (build, file) => {
    const { slide, shape } = await textBox();
    setShapeAnimation(shape, { effect: 'flyIn', build: build as keyof typeof TEXT });
    const capture = timingOf(await readFile(new URL(`text/${file}`, NATIVE), 'utf8'));
    const xml = getSlideXmlString(slide);
    const spid = Number(/spid="(\d+)"/.exec(xml.split('<p:timing>')[1]!)![1]);
    expect(writtenTiming(xml, spid)).toBe(capture);
  });

  it.each(Object.keys(TEXT))('%s reads back and round-trips', async (build) => {
    const { pres, shape } = await textBox();
    setShapeAnimation(shape, { effect: 'flyIn', build: build as keyof typeof TEXT });
    const steps = getSlideAnimations(
      getSlides(await loadPresentation(await savePresentation(pres)))[0]!,
    );
    expect(steps.map((s) => s.build)).toEqual(
      build === 'asOneObject' ? ['asOneObject'] : [build, build, build],
    );
    expect(steps.map((s) => s.start)).toEqual(
      build === 'asOneObject'
        ? ['click']
        : build === 'allAtOnce'
          ? ['click', 'withPrevious', 'withPrevious']
          : ['click', 'click', 'click'],
    );
  });

  it('switches between the three sequences in place', async () => {
    const { slide, shape } = await textBox();
    setShapeAnimation(shape, { effect: 'flyIn' });
    const id = getSlideAnimations(slide)[0]!.id!;
    for (const build of ['allAtOnce', 'byParagraph', 'allAtOnce', 'asOneObject'] as const) {
      updateSlideAnimation(slide, id, { build });
      const capture = timingOf(await readFile(new URL(`text/${TEXT[build]}`, NATIVE), 'utf8'));
      const steps = getSlideAnimations(slide);
      expect(steps[0]!.id).toBe(id);
      expect(steps.every((s) => s.build === build)).toBe(true);
      // The same tree the reference desktop app writes, ids aside.
      const xml = getSlideXmlString(slide);
      const spid = Number(/spid="(\d+)"/.exec(xml.split('<p:timing>')[1]!)![1]);
      const strip = (s: string): string => s.replaceAll(/ id="\d+"/g, '');
      expect(strip(writtenTiming(xml, spid))).toBe(strip(capture));
    }
  });

  it('refuses a text build for an effect that animates the shape itself', async () => {
    const { shape } = await textBox();
    expect(() => setShapeAnimation(shape, { effect: 'fillColor', build: 'byParagraph' })).toThrow(
      /has no text build/,
    );
  });
});
