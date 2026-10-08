// The emphasis effects' own Effect Options — Spin's direction and angle,
// Grow/Shrink's axes and size, Transparency's amount and the colour of the
// colour effects — and the duration a changed preset takes.
//
// None of these options is in the preset numbers, so each is checked where it
// is written: the behaviour the reference desktop app's capture of the gallery default shows
// it in (test/fixtures/native/animations/emphasis/). The defaults themselves
// are compared byte for byte in fn-animation-native.test.ts.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { expectSchemaValid, isSchemaValidationAvailable } from './lib/expect-schema-valid.ts';
import { partName } from '../src/internal/opc/index.ts';
import {
  type AnimationEffect,
  type AnimationOptions,
  type SlideAnimationStep,
  _internalPackageOf,
  addSlideTextBox,
  emu,
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

const fixture = fileURLToPath(new URL('./fixtures/minimal/one-text-slide.pptx', import.meta.url));

const openDeck = async () => {
  const pres = await loadPresentation(await readFile(fixture));
  const slide = getSlides(pres)[0]!;
  const shape = getSlideShapes(slide)[0]!;
  setShapeFill(shape, '#4472C4');
  return { pres, slide, shape };
};

type Deck = Awaited<ReturnType<typeof openDeck>>;

/** The slide as saved and loaded again, so every check is about what reaches disk. */
const reopened = async (pres: Deck['pres']) => {
  const slide = getSlides(await loadPresentation(await savePresentation(pres)))[0]!;
  return { xml: getSlideXmlString(slide), steps: getSlideAnimations(slide) };
};

const timingOf = (xml: string): string => /<p:timing>[\s\S]*<\/p:timing>/.exec(xml)![0];

const OPTION_FIELDS = [
  'spinDirection',
  'spinDegrees',
  'scaleDirection',
  'scalePercent',
  'transparencyPercent',
  'color',
] as const;

const optionsOf = (step: SlideAnimationStep) =>
  Object.fromEntries(OPTION_FIELDS.map((name) => [name, step[name]]));

/** Every option an effect may carry, with the ones it does not take left `null`. */
const options = (values: Partial<Record<(typeof OPTION_FIELDS)[number], unknown>>) => ({
  spinDirection: null,
  spinDegrees: null,
  scaleDirection: null,
  scalePercent: null,
  transparencyPercent: null,
  color: null,
  ...values,
});

interface Case {
  readonly name: string;
  readonly opts: AnimationOptions;
  /** What the saved timing must contain. */
  readonly xml: readonly string[];
  /** What it must no longer contain. */
  readonly not?: readonly string[];
  /** What `getSlideAnimations` reads back. */
  readonly read: ReturnType<typeof options>;
}

const CASES: readonly Case[] = [
  {
    name: 'Spin: counterclockwise quarter turn',
    opts: { effect: 'spin', spinDirection: 'counterclockwise', spinDegrees: 90 },
    xml: ['<p:animRot by="-5400000">'],
    read: options({ spinDirection: 'counterclockwise', spinDegrees: 90 }),
  },
  {
    name: 'Spin: two clockwise turns',
    opts: { effect: 'spin', spinDegrees: 720 },
    xml: ['<p:animRot by="43200000">'],
    read: options({ spinDirection: 'clockwise', spinDegrees: 720 }),
  },
  {
    name: 'Spin: half a turn back',
    opts: { effect: 'spin', spinDirection: 'counterclockwise', spinDegrees: 180 },
    xml: ['<p:animRot by="-10800000">'],
    read: options({ spinDirection: 'counterclockwise', spinDegrees: 180 }),
  },
  {
    name: 'Grow/Shrink: huge, both axes',
    opts: { effect: 'growShrink', scalePercent: 400 },
    xml: ['<p:by x="400000" y="400000"/>'],
    read: options({ scaleDirection: 'both', scalePercent: 400 }),
  },
  {
    name: 'Grow/Shrink: tiny, width only',
    opts: { effect: 'growShrink', scaleDirection: 'horizontal', scalePercent: 25 },
    xml: ['<p:by x="25000" y="100000"/>'],
    read: options({ scaleDirection: 'horizontal', scalePercent: 25 }),
  },
  {
    name: 'Grow/Shrink: larger, height only',
    opts: { effect: 'growShrink', scaleDirection: 'vertical' },
    xml: ['<p:by x="100000" y="150000"/>'],
    read: options({ scaleDirection: 'vertical', scalePercent: 150 }),
  },
  {
    name: 'Transparency: 75 %',
    opts: { effect: 'transparency', transparencyPercent: 75 },
    xml: ['<p:strVal val="0.25"/>', 'prLst="opacity: 0.25"'],
    not: ['0.5'],
    read: options({ transparencyPercent: 75 }),
  },
  {
    name: 'Transparency: 100 %',
    opts: { effect: 'transparency', transparencyPercent: 100 },
    xml: ['<p:strVal val="0"/>', 'prLst="opacity: 0"'],
    read: options({ transparencyPercent: 100 }),
  },
  ...(
    [
      ['fillColor', 1],
      ['fontColor', 1],
      ['lineColor', 1],
      ['objectColor', 2],
      ['growWithColor', 2],
      ['colorPulse', 2],
    ] as const
  ).map(
    ([effect, count]): Case => ({
      name: `${effect}: an sRGB colour on each of its ${count} colour behaviours`,
      opts: { effect, color: '#FF0000' },
      xml: Array.from({ length: count }, () => '<p:to><a:srgbClr val="FF0000"/></p:to>'),
      not: ['accent2', 'bg1'],
      read: options({ color: '#FF0000' }),
    }),
  ),
  {
    name: 'Brush Color: a theme colour in both of its <p:clrVal>',
    opts: { effect: 'brushColor', color: 'accent6' },
    xml: [
      '<p:clrVal><a:schemeClr val="accent6"/></p:clrVal>',
      '<p:clrVal><a:schemeClr val="accent6"/></p:clrVal>',
    ],
    not: ['accent2'],
    read: options({ color: 'scheme:accent6' }),
  },
  {
    name: 'Fill Color: a lighter theme shade',
    opts: {
      effect: 'fillColor',
      color: {
        color: 'scheme:accent1',
        colorTransforms: [
          { kind: 'lumMod', value: 0.2 },
          { kind: 'lumOff', value: 0.8 },
        ],
      },
    },
    xml: [
      '<p:to><a:schemeClr val="accent1"><a:lumMod val="20000"/><a:lumOff val="80000"/></a:schemeClr></p:to>',
    ],
    read: options({
      color: {
        color: 'scheme:accent1',
        colorTransforms: [
          { kind: 'lumMod', value: 0.2 },
          { kind: 'lumOff', value: 0.8 },
        ],
      },
    }),
  },
];

/** Each occurrence of `needle` counted, so "on both behaviours" means both. */
const count = (haystack: string, needle: string): number => haystack.split(needle).length - 1;

describe('animations: the emphasis effects’ own Effect Options', () => {
  it.each(CASES.map((c) => [c.name, c] as const))('%s', async (_, c) => {
    const { pres, shape } = await openDeck();
    setShapeAnimation(shape, c.opts);
    const { xml, steps } = await reopened(pres);
    const timing = timingOf(xml);
    const wanted = new Map<string, number>();
    for (const needle of c.xml) wanted.set(needle, (wanted.get(needle) ?? 0) + 1);
    for (const [needle, times] of wanted) expect(count(timing, needle), needle).toBe(times);
    for (const needle of c.not ?? []) expect(timing, needle).not.toContain(needle);
    expect(steps).toHaveLength(1);
    expect(steps[0]!.effect).toBe(c.opts.effect);
    expect(steps[0]!.editable).toBe(true);
    expect(optionsOf(steps[0]!)).toEqual(c.read);
  });

  it.skipIf(!isSchemaValidationAvailable())(
    'every option writes a schema-valid slide',
    async () => {
      for (const c of CASES) {
        const { slide, shape } = await openDeck();
        setShapeAnimation(shape, c.opts);
        expectSchemaValid(getSlideXmlString(slide), 'pml');
      }
    },
  );

  it('reads the gallery defaults back as options', async () => {
    const defaults: ReadonlyArray<readonly [AnimationEffect, ReturnType<typeof options>]> = [
      ['spin', options({ spinDirection: 'clockwise', spinDegrees: 360 })],
      ['growShrink', options({ scaleDirection: 'both', scalePercent: 150 })],
      ['transparency', options({ transparencyPercent: 50 })],
      ['fillColor', options({ color: 'scheme:accent2' })],
      ['colorPulse', options({ color: 'scheme:bg1' })],
      ['pulse', options({})],
      ['flyIn', options({})],
    ];
    for (const [effect, read] of defaults) {
      const { pres, shape } = await openDeck();
      setShapeAnimation(shape, { effect });
      expect(optionsOf((await reopened(pres)).steps[0]!), effect).toEqual(read);
    }
  });

  it('round-trips: writing back what was read changes nothing', async () => {
    for (const c of CASES) {
      const { pres, shape } = await openDeck();
      setShapeAnimation(shape, c.opts);
      const first = await reopened(pres);
      const slide = getSlides(await loadPresentation(await savePresentation(pres)))[0]!;
      const step = getSlideAnimations(slide)[0]!;
      const patch = Object.fromEntries(
        OPTION_FIELDS.filter((name) => step[name] !== null).map((name) => [name, step[name]]),
      );
      updateSlideAnimation(slide, step.id!, patch);
      // The click stop around the effect is laid out again on any edit, so its
      // ids move; the effect itself must come out exactly as it went in.
      const withoutIds = (xml: string): string => timingOf(xml).replaceAll(/ id="\d+"/g, '');
      expect(withoutIds(getSlideXmlString(slide)), c.name).toBe(withoutIds(first.xml));
    }
  });

  it('changes an option on an existing effect, keeping its duration', async () => {
    const { pres, slide, shape } = await openDeck();
    setShapeAnimation(shape, { effect: 'spin', durationMs: 3000 });
    const id = getSlideAnimations(slide)[0]!.id!;
    updateSlideAnimation(slide, id, { spinDegrees: 90 });
    updateSlideAnimation(slide, id, { spinDirection: 'counterclockwise' });
    const step = (await reopened(pres)).steps[0]!;
    expect([step.id, step.spinDirection, step.spinDegrees, step.durationMs]).toEqual([
      id,
      'counterclockwise',
      90,
      3000,
    ]);
  });

  it('keeps a colour when the effect changes to another colour effect', async () => {
    const { pres, slide, shape } = await openDeck();
    setShapeAnimation(shape, { effect: 'fillColor', color: '#00B050' });
    const id = getSlideAnimations(slide)[0]!.id!;
    updateSlideAnimation(slide, id, { effect: 'lineColor' });
    expect((await reopened(pres)).steps[0]!.color).toBe('#00B050');
    updateSlideAnimation(slide, id, { effect: 'spin' });
    const step = (await reopened(pres)).steps[0]!;
    expect([step.effect, step.color, step.spinDegrees]).toEqual(['spin', null, 360]);
  });

  it('applies the options to every paragraph of a text build', async () => {
    const { pres, slide } = await openDeck();
    const box = addSlideTextBox(slide, {
      x: emu(0),
      y: emu(0),
      w: emu(3_000_000),
      h: emu(1_000_000),
      text: 'One\nTwo\nThree',
    });
    setShapeAnimation(box, { effect: 'growShrink', scalePercent: 50, build: 'byParagraph' });
    const built = (await reopened(pres)).steps.filter((s) => s.target.kind === 'paragraphs');
    expect(built.map((s) => s.scalePercent)).toEqual([50, 50, 50]);
    // Collapsing the build keeps the size.
    updateSlideAnimation(slide, getSlideAnimations(slide).at(-1)!.id!, { build: 'asOneObject' });
    const whole = (await reopened(pres)).steps.filter((s) => s.target.kind === 'shape');
    expect(whole.at(-1)!.scalePercent).toBe(50);
  });

  it('refuses an option the effect does not take, or a value it cannot write', async () => {
    const { shape } = await openDeck();
    const refused: readonly AnimationOptions[] = [
      { effect: 'fadeIn', spinDegrees: 90 },
      { effect: 'spin', color: 'accent1' },
      { effect: 'growShrink', transparencyPercent: 20 },
      { effect: 'darken', color: '#000000' },
      { effect: 'spin', spinDegrees: 0 },
      { effect: 'spin', spinDegrees: Number.POSITIVE_INFINITY },
      { effect: 'spin', spinDirection: 'left' as never },
      { effect: 'growShrink', scalePercent: -10 },
      { effect: 'growShrink', scaleDirection: 'diagonal' as never },
      { effect: 'transparency', transparencyPercent: 120 },
      { effect: 'fillColor', color: 'red' as never },
      {
        effect: 'fillColor',
        color: { color: '#FF0000', colorTransforms: [{ kind: 'alpha', value: 2 }] },
      },
    ];
    for (const opts of refused) {
      expect(() => setShapeAnimation(shape, opts), JSON.stringify(opts)).toThrow(RangeError);
    }
  });
});

describe('animations: emphasis options the reader cannot name', () => {
  /**
   * The effect as written, with its saved slide edited by `edit` and loaded
   * again — a tree the reference desktop app could hand us that says more than the options do.
   */
  const edited = async (opts: AnimationOptions, edit: (xml: string) => string) => {
    const { pres } = await openDeck();
    setShapeAnimation(getSlideShapes(getSlides(pres)[0]!)[0]!, opts);
    const saved = await loadPresentation(await savePresentation(pres));
    const part = _internalPackageOf(saved).getPart(partName('/ppt/slides/slide1.xml'))!;
    const xml = edit(new TextDecoder().decode(part.data));
    if (isSchemaValidationAvailable()) expectSchemaValid(xml, 'pml');
    part.data = new TextEncoder().encode(xml);
    const slide = getSlides(await loadPresentation(await savePresentation(saved)))[0]!;
    return getSlideAnimations(slide)[0]!;
  };

  it.each([
    [
      'a scale that stretches the two axes by different amounts',
      { effect: 'growShrink' },
      (xml: string) =>
        xml.replace('<p:by x="150000" y="150000"/>', '<p:by x="150000" y="200000"/>'),
    ],
    [
      'a scale fixed by <p:from>',
      { effect: 'growShrink' },
      (xml: string) =>
        xml.replace(
          '<p:by x="150000" y="150000"/>',
          '<p:from x="50000" y="50000"/><p:to x="100000" y="100000"/>',
        ),
    ],
    [
      'two colours on one effect',
      { effect: 'objectColor' },
      (xml: string) => xml.replace('<a:schemeClr val="accent2"/>', '<a:srgbClr val="FF0000"/>'),
    ],
    [
      'a preset colour this library does not write',
      { effect: 'fontColor' },
      (xml: string) => xml.replace('<a:schemeClr val="accent2"/>', '<a:prstClr val="red"/>'),
    ],
    [
      'an opacity that is not a number',
      { effect: 'transparency' },
      (xml: string) => xml.replace('<p:strVal val="0.5"/>', '<p:strVal val="half"/>'),
    ],
  ] as const)('%s is read, not named', async (_, opts, edit) => {
    const step = await edited(opts as AnimationOptions, edit);
    expect(step.effect).toBeNull();
    expect(step.editable).toBe(false);
  });

  it('a colour with transforms the reference desktop app wrote is named with them', async () => {
    const step = await edited({ effect: 'fontColor' }, (xml) =>
      xml.replace(
        '<a:schemeClr val="accent2"/>',
        '<a:schemeClr val="accent2"><a:lumMod val="75000"/></a:schemeClr>',
      ),
    );
    expect(step.effect).toBe('fontColor');
    expect(step.color).toEqual({
      color: 'scheme:accent2',
      colorTransforms: [{ kind: 'lumMod', value: 0.75 }],
    });
  });
});

describe('animations: changing the preset takes its default duration', () => {
  it('a new preset starts from its own length, a new option keeps the old one', async () => {
    const { pres, slide, shape } = await openDeck();
    setShapeAnimation(shape, { effect: 'fadeIn', durationMs: 3000 });
    const id = getSlideAnimations(slide)[0]!.id!;
    updateSlideAnimation(slide, id, { effect: 'wipeIn' });
    expect((await reopened(pres)).steps[0]!.durationMs).toBe(500);
    updateSlideAnimation(slide, id, { durationMs: 2500 });
    updateSlideAnimation(slide, id, { direction: 'left' });
    expect((await reopened(pres)).steps[0]!.durationMs).toBe(2500);
    updateSlideAnimation(slide, id, { effect: 'shapeIn' });
    expect((await reopened(pres)).steps[0]!.durationMs).toBe(2000);
    updateSlideAnimation(slide, id, { effect: 'bounceIn', durationMs: 1000 });
    expect((await reopened(pres)).steps[0]!.durationMs).toBe(1000);
    // Transparency holds until the slide ends, and so has none.
    updateSlideAnimation(slide, id, { effect: 'transparency' });
    expect((await reopened(pres)).steps[0]!.durationMs).toBeNull();
    updateSlideAnimation(slide, id, { effect: 'spin' });
    expect((await reopened(pres)).steps[0]!.durationMs).toBe(2000);
  });
});
