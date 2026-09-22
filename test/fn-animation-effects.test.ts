// The effects beyond fade and appear: flying in and out from an edge, zooming
// in and out about the centre, and spinning.
//
// What is checked here is the behaviours, not only the preset numbers. The
// numbers name a gallery entry; it is `<p:anim>` on `ppt_x` / `ppt_y`, on
// `ppt_w` / `ppt_h`, and `<p:animRot by>` that decide what a renderer does.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  type AnimationDirection,
  type AnimationEffect,
  type SlideAnimationStep,
  addBlankSlide,
  addSlideTextBox,
  createPresentation,
  getShapeId,
  getSlideAnimations,
  getSlideShapes,
  getSlideXmlString,
  getSlides,
  inches,
  loadPresentation,
  savePresentation,
  setShapeAnimation,
  updateSlideAnimation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const openDeck = async () => {
  const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
  const slide = getSlides(pres)[0]!;
  return { pres, slide, shape: getSlideShapes(slide)[0]! };
};

/** The saved slide XML, so every assertion is about what actually reaches disk. */
const savedXml = async (pres: Awaited<ReturnType<typeof openDeck>>['pres']): Promise<string> => {
  const reopened = await loadPresentation(await savePresentation(pres));
  return getSlideXmlString(getSlides(reopened)[0]!);
};

/** The steps the reader gets back from the saved bytes, not from the tree in hand. */
const savedSteps = async (
  pres: Awaited<ReturnType<typeof openDeck>>['pres'],
): Promise<readonly SlideAnimationStep[]> => {
  const reopened = await loadPresentation(await savePresentation(pres));
  return getSlideAnimations(getSlides(reopened)[0]!);
};

/** The `<p:anim>` element that drives one attribute, as written. */
const animFor = (xml: string, attrName: string): string => {
  const anims = xml.match(/<p:anim\b[\s\S]*?<\/p:anim>/g) ?? [];
  const found = anims.filter((a) => a.includes(`<p:attrName>${attrName}</p:attrName>`));
  expect(found, `one <p:anim> on ${attrName}`).toHaveLength(1);
  return found[0]!;
};

/** Its two keyframe values, in order. */
const keyframes = (anim: string): string[] =>
  [...anim.matchAll(/<p:(?:str|flt)Val val="([^"]*)"\s*\/>/g)].map((m) => m[1]!);

describe('fn API: fly, zoom and spin', () => {
  it('a fly in states both axes, and only the named one moves', async () => {
    const { pres, shape } = await openDeck();
    setShapeAnimation(shape, { effect: 'flyIn', direction: 'bottom', durationMs: 700 });
    const xml = await savedXml(pres);

    expect(xml).toContain('presetID="2"');
    expect(xml).toContain('presetClass="entr"');
    expect(xml).toContain('presetSubtype="4"');
    // The centre comes up from one half-height below the slide's bottom edge.
    expect(keyframes(animFor(xml, 'ppt_y'))).toEqual(['1+#ppt_h/2', '#ppt_y']);
    expect(keyframes(animFor(xml, 'ppt_x'))).toEqual(['#ppt_x', '#ppt_x']);
    expect(animFor(xml, 'ppt_y')).toContain('dur="700"');
  });

  it('each direction writes its own bit and its own edge', async () => {
    const cases: [AnimationDirection, string, string, string][] = [
      ['top', '1', 'ppt_y', '0-#ppt_h/2'],
      ['right', '2', 'ppt_x', '1+#ppt_w/2'],
      ['bottom', '4', 'ppt_y', '1+#ppt_h/2'],
      ['left', '8', 'ppt_x', '0-#ppt_w/2'],
    ];
    for (const [direction, subtype, moving, from] of cases) {
      const { pres, shape } = await openDeck();
      setShapeAnimation(shape, { effect: 'flyIn', direction });
      const xml = await savedXml(pres);
      expect(xml, direction).toContain(`presetSubtype="${subtype}"`);
      expect(keyframes(animFor(xml, moving))[0], direction).toEqual(from);
      const steps = await savedSteps(pres);
      expect(steps[0]!.direction, direction).toBe(direction);
      expect(steps[0]!.effect, direction).toBe('flyIn');
    }
  });

  it('a fly out leaves through the edge it names and hides only at the end', async () => {
    const { pres, shape } = await openDeck();
    setShapeAnimation(shape, { effect: 'flyOut', direction: 'right', durationMs: 900 });
    const xml = await savedXml(pres);

    expect(xml).toContain('presetClass="exit"');
    expect(xml).toContain('presetSubtype="2"');
    expect(keyframes(animFor(xml, 'ppt_x'))).toEqual(['ppt_x', '1+ppt_w/2']);
    expect(keyframes(animFor(xml, 'ppt_y'))).toEqual(['ppt_y', 'ppt_y']);
    // The shape stays on the slide for the whole flight.
    expect(xml).toMatch(/<p:cond delay="899"\s*\/>/);
    expect(xml).toContain('<p:strVal val="hidden"');
  });

  it('a zoom scales about the centre and states no direction', async () => {
    const { pres, shape } = await openDeck();
    setShapeAnimation(shape, { effect: 'zoomIn' });
    const inXml = await savedXml(pres);
    expect(inXml).toContain('presetID="23"');
    expect(inXml).toContain('presetSubtype="16"');
    expect(keyframes(animFor(inXml, 'ppt_w'))).toEqual(['0', '#ppt_w']);
    expect(keyframes(animFor(inXml, 'ppt_h'))).toEqual(['0', '#ppt_h']);
    expect((await savedSteps(pres))[0]!.direction).toBeNull();

    const out = await openDeck();
    setShapeAnimation(out.shape, { effect: 'zoomOut' });
    const outXml = await savedXml(out.pres);
    expect(outXml).toContain('presetSubtype="32"');
    expect(keyframes(animFor(outXml, 'ppt_w'))).toEqual(['ppt_w', '0']);
  });

  it('a spin turns once clockwise and never touches visibility', async () => {
    const { pres, shape } = await openDeck();
    setShapeAnimation(shape, { effect: 'spin', durationMs: 2000 });
    const xml = await savedXml(pres);

    expect(xml).toContain('presetClass="emph"');
    expect(xml).toContain('presetID="8"');
    expect(xml).toContain('<p:animRot by="21600000">');
    expect(xml).toContain('<p:attrName>r</p:attrName>');
    // An emphasis effect animates a shape that is already on the slide.
    expect(xml).not.toContain('style.visibility');

    const step = (await savedSteps(pres))[0]!;
    expect(step.effect).toBe('spin');
    expect(step.durationMs).toBe(2000);
    expect(step.playable).toBe(true);
  });

  it('every effect survives the round trip', async () => {
    const effects: AnimationEffect[] = [
      'appear',
      'fadeIn',
      'flyIn',
      'zoomIn',
      'disappear',
      'fadeOut',
      'flyOut',
      'zoomOut',
      'spin',
    ];
    for (const effect of effects) {
      const { pres, shape } = await openDeck();
      setShapeAnimation(shape, { effect });
      const step = (await savedSteps(pres))[0]!;
      expect(step.effect, effect).toBe(effect);
      expect(step.playable, effect).toBe(true);
      expect(step.editable, effect).toBe(true);
      expect(step.valueAfterEnd, effect).toBe('held');
    }
  });

  it('any effect can be changed into any other', async () => {
    // Changing a preset replaces the effect node, which the editing path only
    // does when everything on that node is something this library wrote. A new
    // behaviour — `<p:animRot>`, for the spin — has to be on that list, or the
    // effect it writes comes back refusing to be edited again.
    const effects: AnimationEffect[] = [
      'appear',
      'fadeIn',
      'flyIn',
      'zoomIn',
      'disappear',
      'fadeOut',
      'flyOut',
      'zoomOut',
      'spin',
    ];
    for (const from of effects) {
      for (const to of effects) {
        if (from === to) continue;
        const { pres, slide, shape } = await openDeck();
        setShapeAnimation(shape, { effect: from });
        const id = getSlideAnimations(slide)[0]!.id!;
        updateSlideAnimation(slide, id, {
          effect: to,
          ...(to === 'flyIn' || to === 'flyOut' ? { direction: 'top' as const } : {}),
        });
        const step = (await savedSteps(pres))[0]!;
        expect(step.effect, `${from} -> ${to}`).toBe(to);
        expect(step.editable, `${from} -> ${to}`).toBe(true);
      }
    }
  });

  it('a spin splits into a paragraph build and collapses back', async () => {
    // Three paragraphs, so collapsing the build really has several effects to
    // fold back into one rather than the single-paragraph case that would pass
    // without the join ever running.
    const pres = createPresentation();
    const slide = addBlankSlide(pres);
    const shape = addSlideTextBox(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(4),
      h: inches(2),
      text: 'one\ntwo\nthree',
    });
    const spid = getShapeId(shape);
    setShapeAnimation(shape, { effect: 'spin', durationMs: 1500 });
    const id = getSlideAnimations(slide)[0]!.id!;

    updateSlideAnimation(slide, id, { byParagraph: true });
    const built = await savedSteps(pres);
    expect(built).toHaveLength(3);
    expect(built.map((step) => step.effect)).toEqual(['spin', 'spin', 'spin']);
    expect(
      built.map((step) =>
        step.target.kind === 'paragraphs'
          ? [step.target.firstParagraph, step.target.lastParagraph]
          : step.target.kind,
      ),
    ).toEqual([
      [0, 0],
      [1, 1],
      [2, 2],
    ]);
    expect(built.every((step) => step.buildByParagraph)).toBe(true);
    expect(built.every((step) => step.target.shapeId === spid)).toBe(true);

    updateSlideAnimation(slide, getSlideAnimations(slide)[0]!.id!, { byParagraph: false });
    const whole = await savedSteps(pres);
    expect(whole).toHaveLength(1);
    expect(whole[0]!.effect).toBe('spin');
    expect(whole[0]!.target.kind).toBe('shape');
    expect(whole[0]!.buildByParagraph).toBe(false);
    expect(whole[0]!.durationMs).toBe(1500);
  });

  it('a direction is refused for an effect that does not fly', async () => {
    const { pres, slide, shape } = await openDeck();
    expect(() => setShapeAnimation(shape, { effect: 'fadeIn', direction: 'top' })).toThrow(
      /direction only applies/,
    );
    expect(getSlideAnimations(slide)).toHaveLength(0);

    setShapeAnimation(shape, { effect: 'spin' });
    const id = getSlideAnimations(slide)[0]!.id!;
    expect(() => updateSlideAnimation(slide, id, { direction: 'left' })).toThrow(
      /direction only applies/,
    );
    expect(await savedXml(pres)).toContain('<p:animRot by="21600000">');
  });
});

describe('fn API: retiming an effect that hides at the end', () => {
  const hideDelays = (xml: string): string[] =>
    (
      xml.match(/<p:cTn id="\d+" dur="1" fill="hold">\s*<p:stCondLst>\s*<p:cond delay="(\d+)"/g) ??
      []
    ).map((m) => /delay="(\d+)"/.exec(m)![1]!);

  it('the trailing hide moves with the duration, including from one millisecond', async () => {
    const { pres, slide, shape } = await openDeck();
    // Written short enough that the hide already sits at zero, which is where a
    // reveal sits too — so "not at zero" would not have found it again.
    setShapeAnimation(shape, { effect: 'fadeOut', durationMs: 1 });
    expect(hideDelays(await savedXml(pres))).toEqual(['0']);

    const id = getSlideAnimations(slide)[0]!.id!;
    updateSlideAnimation(slide, id, { durationMs: 1200 });
    expect(hideDelays(await savedXml(pres))).toEqual(['1199']);
    expect((await savedSteps(pres))[0]!.durationMs).toBe(1200);
  });

  it('it survives being shortened to nothing and stretched again', async () => {
    const { pres, slide, shape } = await openDeck();
    setShapeAnimation(shape, { effect: 'flyOut', direction: 'left', durationMs: 500 });
    expect(hideDelays(await savedXml(pres))).toEqual(['499']);

    const id = getSlideAnimations(slide)[0]!.id!;
    updateSlideAnimation(slide, id, { durationMs: 0 });
    expect(hideDelays(await savedXml(pres))).toEqual(['0']);
    updateSlideAnimation(slide, id, { durationMs: 1200 });
    expect(hideDelays(await savedXml(pres))).toEqual(['1199']);
    // Both axes are retimed together, not just the first one found.
    const xml = await savedXml(pres);
    expect(animFor(xml, 'ppt_x')).toContain('dur="1200"');
    expect(animFor(xml, 'ppt_y')).toContain('dur="1200"');
  });

  it('an instant exit keeps hiding at once', async () => {
    const { pres, slide, shape } = await openDeck();
    setShapeAnimation(shape, { effect: 'disappear' });
    const id = getSlideAnimations(slide)[0]!.id!;
    updateSlideAnimation(slide, id, { durationMs: 900 });
    expect(hideDelays(await savedXml(pres))).toEqual(['0']);
  });

  it('an afterPrevious effect still starts when the one before it ends', async () => {
    const { pres, shape } = await openDeck();
    // The trailing hide is one millisecond long and ends with the motion, so
    // moving it must not have changed how long the effect before it runs.
    setShapeAnimation(shape, { effect: 'fadeOut', durationMs: 500 });
    setShapeAnimation(shape, { effect: 'fadeIn', start: 'afterPrevious' });
    const xml = await savedXml(pres);
    // The second group starts at the first one's end, not at 501 or 1.
    expect(xml).toMatch(/<p:cond delay="500"\s*\/>/);

    const steps = await savedSteps(pres);
    expect(steps.map((step) => step.start)).toEqual(['click', 'afterPrevious']);
  });
});
