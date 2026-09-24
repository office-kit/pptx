// getShapeAnimation — read back the v1 single-effect animation state.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  clearSlideAnimations,
  getSlideAnimationSequence,
  setSlideAnimationDuration,
  setSlideAnimationDelay,
  setSlideAnimationStart,
  setSlideAnimationEffect,
  setSlideAnimationsSettings,
  removeSlideAnimation,
  removeSlideAnimations,
  moveSlideAnimation,
  moveSlideAnimations,
  reorderSlideAnimations,
  getSlideXmlString,
  getShapeId,
  savePresentation,
  getShapeAnimation,
  getSlideShapes,
  getSlides,
  loadPresentation,
  setShapeAnimation,
} from '../src/api/index.ts';

import { SLIDE_DOCUMENT } from '../src/api/_internal-symbols.ts';
import { parseXml } from '../src/internal/xml/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

it.each([
  ['earlier', [1, 2], [1, 2, 0, 3, 4]],
  ['later', [1, 2], [0, 3, 1, 2, 4]],
  ['earlier', [0, 2, 4], [0, 2, 1, 4, 3]],
  ['later', [0, 2, 4], [1, 0, 3, 2, 4]],
  ['earlier', [0, 1, 2, 3, 4], [0, 1, 2, 3, 4]],
] as const)(
  'moves selected blocks %s without changing effect settings',
  async (direction, selection, expected) => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    for (let i = 0; i < 5; i++)
      setShapeAnimation(getSlideShapes(slide)[0]!, { effect: 'fadeIn', durationMs: 100 + i });
    const ids = getSlideAnimationSequence(slide)
      .flat()
      .map((effect) => effect.timingId!);
    setSlideAnimationStart(slide, ids[2]!, 'afterEffect');
    setSlideAnimationStart(slide, ids[3]!, 'withEffect');
    const before = getSlideAnimationSequence(slide).flat();
    moveSlideAnimations(
      slide,
      [...selection].reverse().map((i) => ids[i]!),
      direction,
    );
    expect(getSlideAnimationSequence(slide).flat()).toEqual(expected.map((i) => before[i]));
    expect(
      getSlideAnimationSequence(
        getSlides(await loadPresentation(await savePresentation(pres)))[0]!,
      ),
    ).toEqual(getSlideAnimationSequence(slide));
  },
);

it('rolls back earlier moves when another selected effect has a timing dependency', async () => {
  const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
  const slide = getSlides(pres)[0]!;
  for (let i = 0; i < 4; i++) setShapeAnimation(getSlideShapes(slide)[0]!, { effect: 'appear' });
  const ids = getSlideAnimationSequence(slide)
    .flat()
    .map((effect) => effect.timingId!);
  slide[SLIDE_DOCUMENT] = parseXml(
    getSlideXmlString(slide).replace(
      '<p:tnLst>',
      `<p:tnLst><p:par><p:cTn id="1000"><p:stCondLst><p:cond evt="onEnd"><p:tn val="${ids[3]}"/></p:cond></p:stCondLst></p:cTn></p:par>`,
    ),
  );
  const before = getSlideXmlString(slide);
  expect(() => moveSlideAnimations(slide, [ids[1]!, ids[3]!], 'earlier')).toThrow(/depends on/);
  expect(getSlideXmlString(slide)).toBe(before);
  expect(
    getSlideXmlString(getSlides(await loadPresentation(await savePresentation(pres)))[0]!),
  ).toBe(before);
  expect(() => moveSlideAnimations(slide, [ids[1]!, 'missing'], 'earlier')).toThrow(
    /no longer exists/,
  );
});

it.each([
  [[3, 1], 0, [1, 3, 0, 2, 4]],
  [[0, 2], null, [1, 3, 4, 0, 2]],
  [[0, 2], 4, [1, 3, 0, 2, 4]],
  [[0, 2], 2, [0, 1, 2, 3, 4]],
] as const)(
  'places a selection at a drag destination (%j before %s)',
  async (selection, target, expected) => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    for (let i = 0; i < 5; i++)
      setShapeAnimation(getSlideShapes(slide)[0]!, { effect: 'fadeIn', durationMs: 100 + i });
    const ids = getSlideAnimationSequence(slide)
      .flat()
      .map((effect) => effect.timingId!);
    setSlideAnimationStart(slide, ids[2]!, 'afterEffect');
    const before = getSlideAnimationSequence(slide).flat();
    reorderSlideAnimations(
      slide,
      selection.map((i) => ids[i]!),
      target === null ? null : ids[target]!,
    );
    expect(getSlideAnimationSequence(slide).flat()).toEqual(expected.map((i) => before[i]));
    expect(
      getSlideAnimationSequence(
        getSlides(await loadPresentation(await savePresentation(pres)))[0]!,
      ),
    ).toEqual(getSlideAnimationSequence(slide));
  },
);

it('restores a drag reorder if a later crossed effect has an external dependency', async () => {
  const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
  const slide = getSlides(pres)[0]!;
  for (let i = 0; i < 4; i++) setShapeAnimation(getSlideShapes(slide)[0]!, { effect: 'appear' });
  const ids = getSlideAnimationSequence(slide)
    .flat()
    .map((effect) => effect.timingId!);
  slide[SLIDE_DOCUMENT] = parseXml(
    getSlideXmlString(slide).replace(
      '<p:tnLst>',
      `<p:tnLst><p:par><p:cTn id="1000"><p:stCondLst><p:cond evt="onEnd"><p:tn val="${ids[2]}"/></p:cond></p:stCondLst></p:cTn></p:par>`,
    ),
  );
  const before = getSlideXmlString(slide);
  expect(() => reorderSlideAnimations(slide, [ids[0]!], null)).toThrow(/depends on/);
  expect(getSlideXmlString(slide)).toBe(before);
  expect(
    getSlideXmlString(getSlides(await loadPresentation(await savePresentation(pres)))[0]!),
  ).toBe(before);
  expect(() => reorderSlideAnimations(slide, [ids[0]!], 'missing')).toThrow(/no longer exists/);
});

it('applies batch settings to selected effects and persists them', async () => {
  const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
  const slide = getSlides(pres)[0]!;
  for (let i = 0; i < 3; i++) setShapeAnimation(getSlideShapes(slide)[0]!, { effect: 'fadeIn' });
  const before = getSlideAnimationSequence(slide).flat();
  const ids = [before[0]!.timingId!, before[2]!.timingId!];
  setSlideAnimationsSettings(slide, [...ids, ids[0]!], {
    effect: 'fadeOut',
    start: 'afterEffect',
    durationMs: 1250,
    delayMs: 375,
  });
  const after = getSlideAnimationSequence(slide).flat();
  expect(after[1]).toEqual(before[1]);
  for (const index of [0, 2])
    expect(after[index]).toMatchObject({
      timingId: before[index]!.timingId,
      effect: 'fadeOut',
      trigger: 'afterEffect',
      durationMs: 1250,
      delayMs: 375,
    });
  expect(
    getSlideAnimationSequence(getSlides(await loadPresentation(await savePresentation(pres)))[0]!),
  ).toEqual(getSlideAnimationSequence(slide));
});

it('restores all settings if a later batch duration is unsupported', async () => {
  const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
  const slide = getSlides(pres)[0]!;
  setShapeAnimation(getSlideShapes(slide)[0]!, { effect: 'fadeIn' });
  setShapeAnimation(getSlideShapes(slide)[0]!, { effect: 'appear' });
  const ids = getSlideAnimationSequence(slide)
    .flat()
    .map((item) => item.timingId!);
  const before = getSlideXmlString(slide);
  expect(() =>
    setSlideAnimationsSettings(slide, ids, {
      start: 'afterEffect',
      durationMs: 1200,
      delayMs: 200,
    }),
  ).toThrow();
  expect(getSlideXmlString(slide)).toBe(before);
  expect(
    getSlideXmlString(getSlides(await loadPresentation(await savePresentation(pres)))[0]!),
  ).toBe(before);
  expect(() =>
    setSlideAnimationsSettings(slide, [ids[0]!, 'missing'], { effect: 'fadeOut' }),
  ).toThrow(/no longer exists/);
  expect(getSlideXmlString(slide)).toBe(before);
});

it('removes multiple grouped effects together and saves the remaining sequence', async () => {
  const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
  const slide = getSlides(pres)[0]!;
  for (let i = 0; i < 4; i++) setShapeAnimation(getSlideShapes(slide)[0]!, { effect: 'fadeIn' });
  const ids = getSlideAnimationSequence(slide)
    .flat()
    .map((item) => item.timingId!);
  setSlideAnimationStart(slide, ids[1]!, 'afterEffect');
  setSlideAnimationStart(slide, ids[2]!, 'withEffect');
  const before = getSlideAnimationSequence(slide).flat();
  removeSlideAnimations(slide, [ids[0]!, ids[2]!, ids[0]!]);
  expect(getSlideAnimationSequence(slide).flat()).toEqual([before[1], before[3]]);
  const saved = getSlides(await loadPresentation(await savePresentation(pres)))[0]!;
  expect(getSlideAnimationSequence(saved)).toEqual(getSlideAnimationSequence(slide));
  removeSlideAnimations(slide, [ids[1]!, ids[3]!]);
  expect(getSlideAnimationSequence(slide)).toEqual([]);
  expect(getSlideXmlString(slide)).not.toContain('<p:bldLst>');
});

it('rolls back earlier removals when a later selected effect has a dependency', async () => {
  const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
  const slide = getSlides(pres)[0]!;
  for (let i = 0; i < 3; i++) setShapeAnimation(getSlideShapes(slide)[0]!, { effect: 'appear' });
  const ids = getSlideAnimationSequence(slide)
    .flat()
    .map((item) => item.timingId!);
  slide[SLIDE_DOCUMENT] = parseXml(
    getSlideXmlString(slide).replace(
      '<p:tnLst>',
      `<p:tnLst><p:par><p:cTn id="1000"><p:stCondLst><p:cond evt="onEnd"><p:tn val="${ids[1]}"/></p:cond></p:stCondLst></p:cTn></p:par>`,
    ),
  );
  const before = getSlideXmlString(slide);
  expect(() => removeSlideAnimations(slide, [ids[0]!, ids[1]!])).toThrow(/depends on/);
  expect(getSlideXmlString(slide)).toBe(before);
  const saved = getSlides(await loadPresentation(await savePresentation(pres)))[0]!;
  expect(getSlideXmlString(saved)).toBe(before);
  expect(() => removeSlideAnimations(slide, [ids[0]!, 'missing'])).toThrow(/no longer exists/);
  expect(getSlideXmlString(slide)).toBe(before);
});

it.each(['appear', 'disappear', 'fadeIn', 'fadeOut'] as const)(
  'replaces selected animation with %s without adding or regrouping',
  async (effect) => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    for (let i = 0; i < 3; i++)
      setShapeAnimation(getSlideShapes(slide)[0]!, { effect: 'fadeIn', durationMs: 1250 });
    const id = getSlideAnimationSequence(slide).flat()[1]!.timingId!;
    setSlideAnimationStart(slide, id, 'afterEffect');
    setSlideAnimationDelay(slide, id, 375);
    const before = getSlideAnimationSequence(slide);
    setSlideAnimationEffect(slide, id, effect);
    const after = getSlideAnimationSequence(slide);
    expect(after[0]![0]).toEqual(before[0]![0]);
    expect(after[1]).toEqual(before[1]);
    expect(after[0]![1]).toMatchObject({
      timingId: id,
      effect,
      trigger: 'afterEffect',
      delayMs: 375,
      durationMs: effect === 'fadeIn' || effect === 'fadeOut' ? 1250 : 0,
    });
    const saved = getSlides(await loadPresentation(await savePresentation(pres)))[0]!;
    expect(getSlideAnimationSequence(saved)).toEqual(after);
    const xml = getSlideXmlString(slide);
    expect(() => setSlideAnimationEffect(slide, 'missing', effect)).toThrow(/no longer exists/);
    expect(() => setSlideAnimationEffect(slide, id, 'invalid' as 'appear')).toThrow();
    expect(getSlideXmlString(slide)).toBe(xml);
  },
);

it('preserves zero fade duration and refuses referenced behavior replacement atomically', async () => {
  const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
  const slide = getSlides(pres)[0]!;
  setShapeAnimation(getSlideShapes(slide)[0]!, { effect: 'fadeIn', durationMs: 0 });
  const id = getSlideAnimationSequence(slide).flat()[0]!.timingId!;
  setSlideAnimationEffect(slide, id, 'fadeOut');
  expect(getSlideAnimationSequence(slide).flat()[0]!.durationMs).toBe(0);
  const xml = getSlideXmlString(slide);
  const behaviorId = /<p:cBhvr[^>]*>\s*<p:cTn id="(\d+)"/.exec(xml)![1];
  slide[SLIDE_DOCUMENT] = parseXml(
    xml.replace(
      '<p:tnLst>',
      `<p:tnLst><p:par><p:cTn id="1000"><p:stCondLst><p:cond evt="onEnd"><p:tn val="${behaviorId}"/></p:cond></p:stCondLst></p:cTn></p:par>`,
    ),
  );
  const before = getSlideXmlString(slide);
  expect(() => setSlideAnimationEffect(slide, id, 'appear')).toThrow(/depends on/);
  expect(getSlideXmlString(slide)).toBe(before);
});

describe('fn API: getShapeAnimation', () => {
  it('regroups start modes, retains effect timing and serializes end dependencies', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    for (const effect of ['fadeIn', 'fadeOut', 'appear'] as const)
      setShapeAnimation(getSlideShapes(slide)[0]!, { effect });
    const before = getSlideAnimationSequence(slide).flat();
    const id = before[1]!.timingId!;
    setSlideAnimationDelay(slide, id, 350);
    setSlideAnimationStart(slide, id, 'withEffect');
    expect(getSlideAnimationSequence(slide).map((g) => g.length)).toEqual([2, 1]);
    expect(getSlideAnimationSequence(slide)[0]![1]).toMatchObject({
      timingId: id,
      trigger: 'withEffect',
      delayMs: 350,
    });
    setSlideAnimationStart(slide, id, 'afterEffect');
    expect(getSlideXmlString(slide)).toContain('evt="onEnd"');
    setSlideAnimationDuration(slide, before[0]!.timingId!, 2750);
    expect(getSlideAnimationSequence(slide)[0]![0]!.durationMs).toBe(2750);
    const loaded = getSlides(await loadPresentation(await savePresentation(pres)))[0]!;
    expect(getSlideAnimationSequence(loaded)).toEqual(getSlideAnimationSequence(slide));
    setSlideAnimationStart(slide, id, 'clickEffect');
    expect(getSlideAnimationSequence(slide).map((g) => g.length)).toEqual([1, 1, 1]);
    expect(getSlideXmlString(slide)).not.toContain('evt="onEnd"');
    setSlideAnimationStart(slide, before[0]!.timingId!, 'withEffect');
    expect(getSlideAnimationSequence(slide)[0]![0]!.trigger).toBe('withEffect');
    const xml = getSlideXmlString(slide);
    expect(() => setSlideAnimationStart(slide, 'missing', 'clickEffect')).toThrow();
    expect(() => setSlideAnimationStart(slide, id, 'bad' as 'clickEffect')).toThrow();
    expect(getSlideXmlString(slide)).toBe(xml);
  });
  it('rejects timing wrapper settings that cannot be retained without mutating the slide', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    setShapeAnimation(getSlideShapes(slide)[0]!, { effect: 'fadeIn' });
    const id = getSlideAnimationSequence(slide)[0]![0]!.timingId!;
    const xml = getSlideXmlString(slide).replace('delay="indefinite"', 'delay="500"');
    slide[SLIDE_DOCUMENT] = parseXml(xml);
    expect(() => setSlideAnimationStart(slide, id, 'afterEffect')).toThrow(/simple main/);
    expect(getSlideXmlString(slide)).toBe(xml);
  });
  it.each([0, 1, 2, 3, 4, 5])(
    'removes grouped effect %i while retaining remaining starts and valid timing references',
    async (position) => {
      const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
      const slide = getSlides(pres)[0]!;
      for (let n = 0; n < 6; n++)
        setShapeAnimation(getSlideShapes(slide)[0]!, { effect: n % 2 ? 'fadeOut' : 'fadeIn' });
      const ids = getSlideAnimationSequence(slide)
        .flat()
        .map((effect) => effect.timingId!);
      const starts = [
        'clickEffect',
        'withEffect',
        'afterEffect',
        'withEffect',
        'afterEffect',
        'clickEffect',
      ] as const;
      for (let n = 0; n < ids.length; n++) {
        setSlideAnimationStart(slide, ids[n]!, starts[n]!);
        setSlideAnimationDuration(slide, ids[n]!, 100 + n * 200);
        setSlideAnimationDelay(slide, ids[n]!, n * 75);
      }
      const before = getSlideAnimationSequence(slide).flat();
      removeSlideAnimation(slide, ids[position]!);
      const remaining = before.filter((_, n) => n !== position);
      expect(getSlideAnimationSequence(slide).flat()).toEqual(remaining);
      const xml = getSlideXmlString(slide);
      const timeIds = [...xml.matchAll(/<p:cTn[^>]* id="([^"]+)"/g)].map((match) => match[1]);
      for (const match of xml.matchAll(/<p:tn val="([^"]+)"/g)) expect(timeIds).toContain(match[1]);
      expect(xml).not.toContain(`grpId="${position}"`);
      expect(
        getSlideAnimationSequence(
          getSlides(await loadPresentation(await savePresentation(pres)))[0]!,
        ),
      ).toEqual(getSlideAnimationSequence(slide));
      for (const effect of remaining) removeSlideAnimation(slide, effect.timingId!);
      expect(getSlideAnimationSequence(slide)).toEqual([]);
      expect(getSlideXmlString(slide)).not.toContain('<p:bldLst');
      setShapeAnimation(getSlideShapes(slide)[0]!, { effect: 'appear' });
      expect(getSlideAnimationSequence(slide).flat()).toHaveLength(1);
    },
  );
  it('removes only a selected effect, cleans its build and can animate again after removing all', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    for (const effect of ['fadeIn', 'fadeOut', 'appear'] as const)
      setShapeAnimation(getSlideShapes(slide)[0]!, { effect });
    const before = getSlideAnimationSequence(slide);
    removeSlideAnimation(slide, before[1]![0]!.timingId!);
    expect(getSlideAnimationSequence(slide)).toEqual([before[0], before[2]]);
    expect(getSlideXmlString(slide)).not.toContain('grpId="1"');
    const saved = getSlides(await loadPresentation(await savePresentation(pres)))[0]!;
    expect(getSlideAnimationSequence(saved)).toEqual([before[0], before[2]]);
    for (const group of getSlideAnimationSequence(slide))
      removeSlideAnimation(slide, group[0]!.timingId!);
    expect(getSlideAnimationSequence(slide)).toEqual([]);
    expect(getSlideXmlString(slide)).not.toContain('<p:bldLst');
    setShapeAnimation(getSlideShapes(slide)[0]!, { effect: 'fadeOut' });
    expect(
      getSlideAnimationSequence(slide)
        .flat()
        .map((effect) => effect.effect),
    ).toEqual(['fadeOut']);
    const xml = getSlideXmlString(slide);
    expect(() => removeSlideAnimation(slide, 'missing')).toThrow(/no longer exists/);
    expect(getSlideXmlString(slide)).toBe(xml);
  });
  it('returns null when no animation is set', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = getSlideShapes(slide)[0]!;
    expect(getShapeAnimation(shape)).toBeNull();
  });

  it('returns the configured preset for every supported effect', async () => {
    for (const effect of ['appear', 'fadeIn', 'disappear', 'fadeOut'] as const) {
      const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
      const slide = getSlides(pres)[0]!;
      const shape = getSlideShapes(slide)[0]!;
      setShapeAnimation(shape, { effect });
      expect(getShapeAnimation(shape)).toBe(effect);
    }
  });

  it('returns null after clearSlideAnimations', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = getSlideShapes(slide)[0]!;
    setShapeAnimation(shape, { effect: 'fadeIn' });
    clearSlideAnimations(slide);
    expect(getShapeAnimation(shape)).toBeNull();
  });
});

describe('getSlideAnimationSequence', () => {
  it('preserves repeated effects on a shape across save/reload as separate click groups', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    expect(getSlideAnimationSequence(slide)).toEqual([]);
    const shape = getSlideShapes(slide)[0]!;
    setShapeAnimation(shape, { effect: 'fadeIn', durationMs: 1234 });
    setShapeAnimation(shape, { effect: 'fadeOut', durationMs: 789 });
    setShapeAnimation(shape, { effect: 'appear' });
    const restored = getSlides(await loadPresentation(await savePresentation(pres)))[0]!;
    const sequence = getSlideAnimationSequence(restored);
    expect(
      sequence.map((group) =>
        group.map(({ effect, durationMs, shapeIds, trigger, delayMs }) => ({
          effect,
          durationMs,
          shapeIds,
          trigger,
          delayMs,
        })),
      ),
    ).toEqual([
      [
        {
          effect: 'fadeIn',
          durationMs: 1234,
          shapeIds: [String(getShapeId(shape))],
          trigger: 'clickEffect',
          delayMs: 0,
        },
      ],
      [
        {
          effect: 'fadeOut',
          durationMs: 789,
          shapeIds: [String(getShapeId(shape))],
          trigger: 'clickEffect',
          delayMs: 0,
        },
      ],
      [
        {
          effect: 'appear',
          durationMs: 0,
          shapeIds: [String(getShapeId(shape))],
          trigger: 'clickEffect',
          delayMs: 0,
        },
      ],
    ]);
    expect(new Set(sequence.flat().map((effect) => effect.timingId)).size).toBe(3);
    clearSlideAnimations(restored);
    expect(getSlideAnimationSequence(restored)).toEqual([]);
  });
});

it('retains unknown presets and simultaneous/after triggers without including interactive effects', async () => {
  const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
  const slide = getSlides(pres)[0]!;
  slide[SLIDE_DOCUMENT].root.children.push(
    parseXml(`
    <p:timing xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
      <p:tnLst><p:par><p:cTn nodeType="tmRoot"><p:childTnLst>
        <p:seq><p:cTn nodeType="mainSeq"><p:childTnLst>
          <p:par><p:cTn><p:childTnLst>
            <p:par><p:cTn id="8" presetID="999" presetClass="entr" nodeType="clickEffect">
              <p:stCondLst><p:cond delay="indefinite"/></p:stCondLst>
              <p:childTnLst><p:set><p:cBhvr><p:tgtEl><p:spTgt spid="42"/></p:tgtEl></p:cBhvr></p:set></p:childTnLst>
            </p:cTn></p:par>
            <p:par><p:cTn id="9" presetID="1" presetClass="entr" nodeType="withEffect">
              <p:stCondLst><p:cond delay="200"/></p:stCondLst>
              <p:childTnLst><p:set><p:cBhvr><p:tgtEl><p:spTgt spid="2"/></p:tgtEl></p:cBhvr></p:set></p:childTnLst>
            </p:cTn></p:par>
            <p:par><p:cTn id="10" presetID="1" presetClass="exit" nodeType="afterEffect"/></p:par>
          </p:childTnLst></p:cTn></p:par>
          <p:par><p:cTn/></p:par>
        </p:childTnLst></p:cTn></p:seq>
        <p:seq><p:cTn nodeType="interactiveSeq"><p:childTnLst>
          <p:par><p:cTn presetID="1" presetClass="entr"/></p:par>
        </p:childTnLst></p:cTn></p:seq>
      </p:childTnLst></p:cTn></p:par></p:tnLst>
    </p:timing>`).root,
  );
  const sequence = getSlideAnimationSequence(slide);
  expect(sequence).toHaveLength(2);
  expect(sequence[1]).toEqual([]);
  expect(sequence[0]).toHaveLength(3);
  expect(sequence[0]![0]).toMatchObject({
    shapeIds: ['42'],
    effect: null,
    presetId: '999',
    delayMs: null,
    durationMs: null,
  });
  expect(sequence[0]![1]).toMatchObject({
    shapeIds: ['2'],
    effect: 'appear',
    trigger: 'withEffect',
    delayMs: 200,
    durationMs: 0,
  });
  expect(sequence[0]![2]).toMatchObject({
    shapeIds: [],
    effect: 'disappear',
    trigger: 'afterEffect',
    delayMs: null,
  });
  removeSlideAnimation(slide, '9');
  expect(getSlideAnimationSequence(slide)).toEqual([[sequence[0]![0], sequence[0]![2]], []]);
  expect(getSlideXmlString(slide)).toContain('interactiveSeq');
});

it('rejects removal with incoming timing references without mutating the document', async () => {
  const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
  const slide = getSlides(pres)[0]!;
  setShapeAnimation(getSlideShapes(slide)[0]!, { effect: 'appear' });
  setShapeAnimation(getSlideShapes(slide)[0]!, { effect: 'disappear' });
  const sequence = getSlideAnimationSequence(slide);
  const id = sequence[0]![0]!.timingId!;
  const xml = getSlideXmlString(slide);
  // An external root condition references the first effect's timing node.
  slide[SLIDE_DOCUMENT] = parseXml(
    xml.replace(
      '<p:tnLst>',
      `<p:tnLst><p:par><p:cTn id="1000"><p:stCondLst><p:cond evt="onEnd"><p:tn val="${id}"/></p:cond></p:stCondLst></p:cTn></p:par>`,
    ),
  );
  const before = getSlideXmlString(slide);
  expect(() => removeSlideAnimation(slide, id)).toThrow(/depends/);
  expect(getSlideXmlString(slide)).toBe(before);
  expect(getSlideAnimationSequence(slide)).toEqual(sequence);
});

it('edits a fade duration in place without duplicating effects or changing their identities', async () => {
  const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
  const slide = getSlides(pres)[0]!;
  const shape = getSlideShapes(slide)[0]!;
  setShapeAnimation(shape, { effect: 'fadeIn', durationMs: 700 });
  setShapeAnimation(shape, { effect: 'fadeOut', durationMs: 1200 });
  const before = getSlideAnimationSequence(slide).flat();
  setSlideAnimationDuration(slide, before[1]!.timingId!, 2400);
  const restored = getSlides(await loadPresentation(await savePresentation(pres)))[0]!;
  expect(getSlideAnimationSequence(restored).flat()).toEqual([
    before[0],
    { ...before[1], durationMs: 2400 },
  ]);
  expect(getSlideXmlString(restored)).toContain('delay="2400"');
  setSlideAnimationDuration(restored, before[1]!.timingId!, 0);
  expect(getSlideAnimationSequence(restored).flat()[1]!.durationMs).toBe(0);
  expect(getSlideXmlString(restored)).not.toContain('delay="2400"');
});

it('rejects invalid durations, missing effects and unsupported effects without changing the slide', async () => {
  const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
  const slide = getSlides(pres)[0]!;
  setShapeAnimation(getSlideShapes(slide)[0]!, { effect: 'appear' });
  const id = getSlideAnimationSequence(slide)[0]![0]!.timingId!;
  const before = getSlideXmlString(slide);
  for (const duration of [-1, 0.5, NaN, Infinity, 4294967296])
    expect(() => setSlideAnimationDuration(slide, id, duration)).toThrow();
  expect(() => setSlideAnimationDuration(slide, 'missing', 500)).toThrow('no longer exists');
  expect(() => setSlideAnimationDuration(slide, id, 500)).toThrow('not supported');
  expect(getSlideXmlString(slide)).toBe(before);
});

const nativeFadeTiming = (direction: 'in' | 'out', behavior = '') => `
<p:timing xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:tnLst><p:par><p:cTn id="1" nodeType="tmRoot"><p:childTnLst>
    <p:seq><p:cTn id="2" nodeType="mainSeq"><p:childTnLst>
      <p:par><p:cTn id="3"><p:childTnLst>
        <p:par><p:cTn id="4" presetID="10" presetClass="${direction === 'in' ? 'entr' : 'exit'}" nodeType="clickEffect" grpId="0">
          <p:stCondLst><p:cond delay="100"/></p:stCondLst>
          <p:childTnLst><p:animEffect transition="${direction}" filter="fade">
            <p:cBhvr><p:cTn id="5" dur="1350" fill="hold"/><p:tgtEl><p:spTgt spid="2"/></p:tgtEl></p:cBhvr>${behavior}
          </p:animEffect></p:childTnLst>
        </p:cTn></p:par>
      </p:childTnLst></p:cTn></p:par>
    </p:childTnLst></p:cTn></p:seq>
  </p:childTnLst></p:cTn></p:par></p:tnLst>
  <p:bldLst><p:bldP spid="2" grpId="0"/></p:bldLst>
</p:timing>`;

it('changes only the selected delay and persists a newly added start condition', async () => {
  const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
  const slide = getSlides(pres)[0]!;
  for (const effect of ['appear', 'fadeOut'] as const)
    setShapeAnimation(getSlideShapes(slide)[0]!, { effect });
  const before = getSlideAnimationSequence(slide);
  setSlideAnimationDelay(slide, before[1]![0]!.timingId!, 1750);
  const restored = getSlides(await loadPresentation(await savePresentation(pres)))[0]!;
  expect(getSlideAnimationSequence(restored)).toEqual([
    before[0],
    [{ ...before[1]![0], delayMs: 1750 }],
  ]);
  const other = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
  const otherSlide = getSlides(other)[0]!;
  otherSlide[SLIDE_DOCUMENT].root.children.push(
    parseXml(nativeFadeTiming('in').replace('<p:stCondLst><p:cond delay="100"/></p:stCondLst>', ''))
      .root,
  );
  setSlideAnimationDelay(otherSlide, '4', 0);
  expect(
    getSlideAnimationSequence(
      getSlides(await loadPresentation(await savePresentation(other)))[0]!,
    )[0]![0],
  ).toMatchObject({ delayMs: 0, durationMs: 1350, trigger: 'clickEffect' });
});

it('rejects invalid delays and complex start conditions without mutating timing', async () => {
  const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
  const slide = getSlides(pres)[0]!;
  setShapeAnimation(getSlideShapes(slide)[0]!, { effect: 'appear' });
  const id = getSlideAnimationSequence(slide)[0]![0]!.timingId!;
  const before = getSlideXmlString(slide);
  for (const delay of [-1, NaN, Infinity, 4294967296])
    expect(() => setSlideAnimationDelay(slide, id, delay)).toThrow();
  expect(() => setSlideAnimationDelay(slide, 'missing', 200)).toThrow();
  expect(getSlideXmlString(slide)).toBe(before);
  setSlideAnimationDelay(slide, id, 123.6);
  expect(getSlideAnimationSequence(slide)[0]![0]!.delayMs).toBe(124);
  for (const condition of [
    '<p:cond delay="indefinite"/>',
    '<p:cond delay="100"/><p:cond delay="200"/>',
    '<p:cond evt="onClick" delay="0"/>',
    '<p:cond delay="0"><p:tn val="3"/></p:cond>',
  ]) {
    const p = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    const s = getSlides(p)[0]!;
    s[SLIDE_DOCUMENT].root.children.push(
      parseXml(nativeFadeTiming('in').replace('<p:cond delay="100"/>', condition)).root,
    );
    const xml = getSlideXmlString(s);
    expect(() => setSlideAnimationDelay(s, '4', 200)).toThrow(/simple numeric/);
    expect(getSlideXmlString(s)).toBe(xml);
  }
});

it.each(['in', 'out'] as const)(
  'reads and edits native %s filter fades without converting their XML',
  async (direction) => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    slide[SLIDE_DOCUMENT].root.children.push(parseXml(nativeFadeTiming(direction)).root);
    const before = getSlideAnimationSequence(slide)[0]![0]!;
    expect(before).toMatchObject({
      effect: direction === 'in' ? 'fadeIn' : 'fadeOut',
      durationMs: 1350,
      delayMs: 100,
      shapeIds: ['2'],
    });
    setSlideAnimationDuration(slide, '4', 2700);
    const restored = getSlides(await loadPresentation(await savePresentation(pres)))[0]!;
    expect(getSlideAnimationSequence(restored)[0]![0]).toEqual({ ...before, durationMs: 2700 });
    expect(getSlideXmlString(restored)).toContain(`transition="${direction}" filter="fade"`);
    expect(getSlideXmlString(restored)).not.toContain('style.opacity');
  },
);

it('does not treat custom-progress or different filter effects as editable standard fades', async () => {
  for (const xml of [
    nativeFadeTiming('in', '<p:progress><p:fltVal val="0.5"/></p:progress>'),
    nativeFadeTiming('in').replace('filter="fade"', 'filter="blinds(horizontal)"'),
    nativeFadeTiming('in').replace('transition="in"', 'transition="out"'),
  ]) {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    slide[SLIDE_DOCUMENT].root.children.push(parseXml(xml).root);
    expect(getSlideAnimationSequence(slide)[0]![0]!.durationMs).toBeNull();
    const before = getSlideXmlString(slide);
    expect(() => setSlideAnimationDuration(slide, '4', 2000)).toThrow(/single supported fade/);
    expect(getSlideXmlString(slide)).toBe(before);
  }
});

it('reorders independent effects while preserving timing, targets and serialized behaviors', async () => {
  const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
  const slide = getSlides(pres)[0]!;
  for (const effect of ['appear', 'fadeOut', 'fadeIn'] as const)
    setShapeAnimation(getSlideShapes(slide)[0]!, { effect, durationMs: 750 });
  const id = getSlideAnimationSequence(slide)[1]![0]!.timingId!;
  setSlideAnimationDelay(slide, id, 325);
  const before = getSlideAnimationSequence(slide);
  const xml = getSlideXmlString(slide);
  moveSlideAnimation(slide, id, 'earlier');
  expect(getSlideAnimationSequence(slide)).toEqual([before[1], before[0], before[2]]);
  const saved = getSlides(await loadPresentation(await savePresentation(pres)))[0]!;
  expect(getSlideAnimationSequence(saved)).toEqual([before[1], before[0], before[2]]);
  moveSlideAnimation(slide, id, 'earlier'); // Already at the boundary.
  expect(getSlideAnimationSequence(slide)).toEqual([before[1], before[0], before[2]]);
  moveSlideAnimation(slide, id, 'later');
  expect(getSlideXmlString(slide)).toBe(xml);
  moveSlideAnimation(slide, id, 'later');
  expect(getSlideAnimationSequence(slide)).toEqual([before[0], before[2], before[1]]);
  const end = getSlideXmlString(slide);
  moveSlideAnimation(slide, id, 'later');
  expect(() => moveSlideAnimation(slide, 'missing', 'earlier')).toThrow(/no longer exists/);
  expect(() => moveSlideAnimation(slide, id, 'invalid' as 'earlier')).toThrow(/direction/);
  expect(getSlideXmlString(slide)).toBe(end);
});

it('rejects reordering across timing references without mutation', async () => {
  const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
  const slide = getSlides(pres)[0]!;
  for (const effect of ['appear', 'disappear'] as const)
    setShapeAnimation(getSlideShapes(slide)[0]!, { effect });
  const id = getSlideAnimationSequence(slide)[0]![0]!.timingId!;
  slide[SLIDE_DOCUMENT] = parseXml(
    getSlideXmlString(slide).replace(
      '<p:tnLst>',
      `<p:tnLst><p:par><p:cTn id="1000"><p:stCondLst><p:cond evt="onEnd"><p:tn val="${id}"/></p:cond></p:stCondLst></p:cTn></p:par>`,
    ),
  );
  const before = getSlideXmlString(slide);
  expect(() => moveSlideAnimation(slide, id, 'later')).toThrow(/depends on/);
  expect(getSlideXmlString(slide)).toBe(before);
});

it.each([0, 1, 2, 3, 4])(
  'reorders grouped effect at position %i and preserves start settings',
  async (position) => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    for (let i = 0; i < 6; i++)
      setShapeAnimation(getSlideShapes(slide)[0]!, {
        effect: i % 2 ? 'fadeOut' : 'fadeIn',
        durationMs: 250 + i * 100,
      });
    const ids = getSlideAnimationSequence(slide)
      .flat()
      .map((effect) => effect.timingId!);
    const modes = [
      'clickEffect',
      'withEffect',
      'afterEffect',
      'withEffect',
      'afterEffect',
      'clickEffect',
    ] as const;
    ids.forEach((id, i) => {
      setSlideAnimationStart(slide, id, modes[i]!);
      setSlideAnimationDelay(slide, id, i * 75);
    });
    const before = getSlideAnimationSequence(slide).flat();
    const expected = [...before];
    [expected[position], expected[position + 1]] = [expected[position + 1]!, expected[position]!];
    moveSlideAnimation(slide, ids[position]!, 'later');
    expect(getSlideAnimationSequence(slide).flat()).toEqual(expected);
    const saved = getSlides(await loadPresentation(await savePresentation(pres)))[0]!;
    expect(getSlideAnimationSequence(saved).flat()).toEqual(expected);
    moveSlideAnimation(slide, ids[position]!, 'earlier');
    expect(getSlideAnimationSequence(slide).flat()).toEqual(before);
  },
);

it('rejects grouped reorder with external effect dependencies atomically', async () => {
  const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
  const slide = getSlides(pres)[0]!;
  for (const effect of ['appear', 'disappear'] as const)
    setShapeAnimation(getSlideShapes(slide)[0]!, { effect });
  const ids = getSlideAnimationSequence(slide)
    .flat()
    .map((effect) => effect.timingId!);
  setSlideAnimationStart(slide, ids[1]!, 'afterEffect');
  slide[SLIDE_DOCUMENT] = parseXml(
    getSlideXmlString(slide).replace(
      '<p:tnLst>',
      `<p:tnLst><p:par><p:cTn id="1000"><p:stCondLst><p:cond evt="onEnd"><p:tn val="${ids[0]}"/></p:cond></p:stCondLst></p:cTn></p:par>`,
    ),
  );
  const before = getSlideXmlString(slide);
  expect(() => moveSlideAnimation(slide, ids[0]!, 'later')).toThrow(/depends on/);
  expect(getSlideXmlString(slide)).toBe(before);
});
