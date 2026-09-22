// updateSlideAnimation / removeSlideAnimation / moveSlideAnimation, and what
// removing or copying a shape does to the effects bound to it.
//
// The assertions are about the saved XML as much as the read model: reordering
// rewrites the click stops and the offsets inside them, and the point of these
// calls is that what PowerPoint plays changes with them.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { expectSchemaValid, isSchemaValidationAvailable } from './lib/expect-schema-valid.ts';
import { partName } from '../src/internal/opc/index.ts';
import {
  type PresentationData,
  type SlideData,
  type SlideShapeData,
  _internalPackageOf,
  addBlankSlide,
  addSlideImage,
  addSlideShape,
  addSlideTextBox,
  copyShape,
  createPresentation,
  getShapeId,
  getSlideAnimations,
  getSlideShapes,
  getSlides,
  groupShapes,
  inches,
  loadPresentation,
  moveSlideAnimation,
  removeShape,
  removeSlideAnimation,
  savePresentation,
  setShapeAnimation,
  slideHasAnimations,
  updateSlideAnimation,
} from '../src/api/index.ts';

const SLIDE1 = '/ppt/slides/slide1.xml';
const decoder = new TextDecoder();
const encoder = new TextEncoder();
const box = { x: inches(1), y: inches(1), w: inches(2), h: inches(1) };
const skipIfNoXmllint = isSchemaValidationAvailable() ? it : it.skip;

const slideXml = (pres: PresentationData): string =>
  decoder.decode(_internalPackageOf(pres).getPart(partName(SLIDE1))!.data);

/**
 * Start offsets of the click stops and the groups inside them, in document
 * order. Effect nodes carry `presetID` and are dropped, so what is left is the
 * timeline the wrappers impose — the thing a reorder has to get right.
 */
const wrapperDelays = (xml: string): string[] =>
  [...xml.matchAll(/<p:par><p:cTn\b([^>]*)><p:stCondLst><p:cond delay="([^"]*)"\/>/g)]
    .filter((m) => !m[1]!.includes('presetID'))
    .map((m) => m[2]!);

const deck = (): { pres: PresentationData; slide: SlideData } => {
  const pres = createPresentation();
  return { pres, slide: addBlankSlide(pres) };
};
const rect = (slide: SlideData): SlideShapeData => addSlideShape(slide, { preset: 'rect', ...box });

/** Three click effects on three shapes, the ordinary case every operation starts from. */
const threeClicks = (): {
  pres: PresentationData;
  slide: SlideData;
  shapes: SlideShapeData[];
  ids: number[];
} => {
  const { pres, slide } = deck();
  const shapes = ['rect', 'ellipse', 'triangle'].map((preset) =>
    addSlideShape(slide, { preset, ...box }),
  );
  for (const shape of shapes) setShapeAnimation(shape, { effect: 'fadeIn', durationMs: 400 });
  return { pres, slide, shapes, ids: getSlideAnimations(slide).map((s) => s.id!) };
};

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

/** Splices hand-written timing into a real slide and reloads it. */
const withTiming = async (
  timing: string,
): Promise<{ pres: PresentationData; slide: SlideData; spid: number }> => {
  const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
  const part = _internalPackageOf(pres).getPart(partName(SLIDE1))!;
  const spid = getShapeId(getSlideShapes(getSlides(pres)[0]!)[0]!);
  const xml = decoder.decode(part.data).replace(/<p:timing>[\s\S]*<\/p:timing>/, '');
  const spliced = xml.replace('</p:sld>', `${timing.replaceAll('{spid}', String(spid))}</p:sld>`);
  if (isSchemaValidationAvailable()) expectSchemaValid(spliced, 'pml');
  part.data = encoder.encode(spliced);
  const reloaded = await loadPresentation(await savePresentation(pres));
  return { pres: reloaded, slide: getSlides(reloaded)[0]!, spid };
};

const timingRoot = (seqChildren: string, bldLst = ''): string =>
  `<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">` +
  `<p:childTnLst>${seqChildren}</p:childTnLst></p:cTn></p:par></p:tnLst>${bldLst}</p:timing>`;

const mainSeq = (childTnLst: string): string =>
  `<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq">` +
  `<p:childTnLst>${childTnLst}</p:childTnLst></p:cTn></p:seq>`;

/** A click stop holding one preset effect, with the wrapper attributes given. */
const stop = (
  opts: {
    stopDelay?: string;
    groupDelay?: string;
    stopFill?: string;
    effectId?: number;
    firstId?: number;
    grpId?: string;
    presetId?: string;
  } = {},
): string => {
  const first = opts.firstId ?? 3;
  const id = opts.effectId ?? first + 2;
  return (
    `<p:par><p:cTn id="${first}" fill="${opts.stopFill ?? 'hold'}"><p:stCondLst>` +
    `<p:cond delay="${opts.stopDelay ?? 'indefinite'}"/></p:stCondLst><p:childTnLst>` +
    `<p:par><p:cTn id="${first + 1}" fill="hold"><p:stCondLst>` +
    `<p:cond delay="${opts.groupDelay ?? '0'}"/></p:stCondLst><p:childTnLst>` +
    `<p:par><p:cTn id="${id}" presetID="${opts.presetId ?? '10'}" presetClass="entr" ` +
    `presetSubtype="0" fill="hold" grpId="${opts.grpId ?? '0'}" nodeType="clickEffect">` +
    `<p:stCondLst><p:cond delay="0"/></p:stCondLst>` +
    `<p:childTnLst><p:anim calcmode="lin" valueType="num"><p:cBhvr additive="base">` +
    `<p:cTn id="${id + 1}" dur="500" fill="hold"/><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>` +
    `<p:attrNameLst><p:attrName>style.opacity</p:attrName></p:attrNameLst></p:cBhvr>` +
    `<p:tavLst/></p:anim></p:childTnLst></p:cTn></p:par>` +
    `</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>`
  );
};

const BLD = `<p:bldLst><p:bldP spid="{spid}" grpId="0"/></p:bldLst>`;

/** The same click stop, against a shape id this test already knows. */
const stopOn = (spid: number, opts: Parameters<typeof stop>[0] = {}): string =>
  stop(opts).replaceAll('{spid}', String(spid));

/** A sequence that runs on its own trigger rather than on the slide's clicks. */
const interactiveSeq = (spid: number, childTnLst: string): string =>
  `<p:seq concurrent="1" nextAc="seek"><p:cTn id="20" restart="whenNotActive" fill="hold" ` +
  `nodeType="interactiveSeq"><p:stCondLst><p:cond evt="onClick" delay="0"><p:tgtEl>` +
  `<p:spTgt spid="${spid}"/></p:tgtEl></p:cond></p:stCondLst>` +
  `<p:childTnLst>${childTnLst}</p:childTnLst></p:cTn></p:seq>`;

/** Puts hand-written timing on the first slide of a deck built in the test. */
const withSplicedTiming = async (
  pres: PresentationData,
  timing: string,
): Promise<PresentationData> => {
  const part = _internalPackageOf(pres).getPart(partName(SLIDE1))!;
  const spliced = decoder
    .decode(part.data)
    .replace(/<p:timing>[\s\S]*<\/p:timing>/, '')
    .replace('</p:sld>', `${timing}</p:sld>`);
  if (isSchemaValidationAvailable()) expectSchemaValid(spliced, 'pml');
  part.data = encoder.encode(spliced);
  return loadPresentation(await savePresentation(pres));
};

const PNG = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
  0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
  0x89, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x63, 0x00, 0x01, 0x00, 0x00,
  0x05, 0x00, 0x01, 0x0d, 0x0a, 0x2d, 0xb4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae,
  0x42, 0x60, 0x82,
]);

describe('fn API: moveSlideAnimation', () => {
  it('reorders the click order and keeps every handle', () => {
    const { pres, slide, shapes, ids } = threeClicks();
    moveSlideAnimation(slide, ids[2]!, 0);

    const steps = getSlideAnimations(slide);
    expect(steps.map((s) => s.id)).toEqual([ids[2], ids[0], ids[1]]);
    expect(steps.map((s) => s.target.shapeId)).toEqual([
      getShapeId(shapes[2]!),
      getShapeId(shapes[0]!),
      getShapeId(shapes[1]!),
    ]);
    // Three click stops, still one group each.
    expect(wrapperDelays(slideXml(pres))).toEqual([
      'indefinite',
      '0',
      'indefinite',
      '0',
      'indefinite',
      '0',
    ]);
  });

  // The start conditions are what a move has to preserve: a "with previous"
  // effect dragged to the front is not suddenly a click, it is the first thing
  // that runs as the slide appears.
  it('keeps a with-previous effect auto-starting when it moves to the front', () => {
    const { pres, slide } = deck();
    const a = rect(slide);
    const b = rect(slide);
    setShapeAnimation(a, { effect: 'fadeIn', durationMs: 400 });
    setShapeAnimation(b, { effect: 'fadeIn', durationMs: 400, start: 'withPrevious' });
    const ids = getSlideAnimations(slide).map((s) => s.id!);

    moveSlideAnimation(slide, ids[1]!, 0);
    const steps = getSlideAnimations(slide);
    expect(steps.map((s) => s.start)).toEqual(['withPrevious', 'click']);
    // The stop the leading effect opens starts with the slide, and the click
    // effect after it gets a stop of its own.
    expect(wrapperDelays(slideXml(pres))).toEqual(['0', '0', 'indefinite', '0']);
  });

  it('recomputes the offsets of an after-previous chain', () => {
    const { pres, slide } = deck();
    const a = rect(slide);
    const b = rect(slide);
    const c = rect(slide);
    setShapeAnimation(a, { effect: 'fadeIn', durationMs: 300 });
    setShapeAnimation(b, { effect: 'fadeIn', durationMs: 900, start: 'afterPrevious' });
    setShapeAnimation(c, { effect: 'fadeIn', durationMs: 100, start: 'afterPrevious' });
    expect(wrapperDelays(slideXml(pres))).toEqual(['indefinite', '0', '300', '1200']);

    const ids = getSlideAnimations(slide).map((s) => s.id!);
    moveSlideAnimation(slide, ids[2]!, 1);
    // The chain runs 300 then 100 now, so the last effect waits 400 instead.
    expect(getSlideAnimations(slide).map((s) => s.durationMs)).toEqual([300, 100, 900]);
    expect(wrapperDelays(slideXml(pres))).toEqual(['indefinite', '0', '300', '400']);
  });

  // An "after previous" effect dragged to the front has nothing to follow, so
  // it runs as the slide appears — the start it was given is not rewritten.
  it('auto-starts an after-previous effect moved to the front', () => {
    const { pres, slide } = deck();
    const a = rect(slide);
    const b = rect(slide);
    setShapeAnimation(a, { effect: 'fadeIn', durationMs: 300 });
    setShapeAnimation(b, { effect: 'fadeIn', durationMs: 900, start: 'afterPrevious' });
    const ids = getSlideAnimations(slide).map((s) => s.id!);

    moveSlideAnimation(slide, ids[1]!, 0);
    expect(getSlideAnimations(slide).map((s) => s.start)).toEqual(['afterPrevious', 'click']);
    expect(wrapperDelays(slideXml(pres))).toEqual(['0', '0', 'indefinite', '0']);
  });

  it('refuses an index outside the sequence and changes nothing', () => {
    const { pres, slide, ids } = threeClicks();
    const before = slideXml(pres);
    expect(() => moveSlideAnimation(slide, ids[0]!, 3)).toThrow(/outside the slide's 3/);
    expect(() => moveSlideAnimation(slide, ids[0]!, -1)).toThrow(/outside the slide's 3/);
    expect(slideXml(pres)).toBe(before);
  });

  it('refuses an id the slide does not have', () => {
    const { slide, ids } = threeClicks();
    expect(() => moveSlideAnimation(slide, Math.max(...ids) + 100, 0)).toThrow(
      /no animation with id/,
    );
  });

  skipIfNoXmllint('emits a schema-valid tree after reordering', () => {
    const { pres, slide, ids } = threeClicks();
    moveSlideAnimation(slide, ids[0]!, 2);
    expectSchemaValid(slideXml(pres), 'pml');
  });

  it('survives save and reload', async () => {
    const { pres, slide, ids } = threeClicks();
    moveSlideAnimation(slide, ids[2]!, 0);
    const before = slideXml(pres);
    const steps = getSlideAnimations(slide);

    const reloaded = await loadPresentation(await savePresentation(pres));
    expect(slideXml(reloaded)).toBe(before);
    expect(getSlideAnimations(getSlides(reloaded)[0]!)).toEqual(steps);
  });
});

describe('fn API: removeSlideAnimation', () => {
  it('drops one effect and closes the gap in the click order', () => {
    const { pres, slide, ids } = threeClicks();
    removeSlideAnimation(slide, ids[1]!);

    expect(getSlideAnimations(slide).map((s) => s.id)).toEqual([ids[0], ids[2]]);
    expect(wrapperDelays(slideXml(pres))).toEqual(['indefinite', '0', 'indefinite', '0']);
  });

  it('moves an after-previous effect up when what it followed is gone', () => {
    const { pres, slide } = deck();
    const a = rect(slide);
    const b = rect(slide);
    const c = rect(slide);
    setShapeAnimation(a, { effect: 'fadeIn', durationMs: 300 });
    setShapeAnimation(b, { effect: 'fadeIn', durationMs: 700, start: 'afterPrevious' });
    setShapeAnimation(c, { effect: 'fadeIn', durationMs: 100, start: 'afterPrevious' });
    const ids = getSlideAnimations(slide).map((s) => s.id!);
    expect(wrapperDelays(slideXml(pres))).toEqual(['indefinite', '0', '300', '1000']);

    removeSlideAnimation(slide, ids[1]!);
    expect(wrapperDelays(slideXml(pres))).toEqual(['indefinite', '0', '300']);
    expect(getSlideAnimations(slide).map((s) => s.durationMs)).toEqual([300, 100]);
  });

  it('takes the timing off the slide once the last effect goes', () => {
    const { pres, slide } = deck();
    const shape = rect(slide);
    setShapeAnimation(shape, { effect: 'fadeIn' });
    removeSlideAnimation(slide, getSlideAnimations(slide)[0]!.id!);

    expect(getSlideAnimations(slide)).toEqual([]);
    expect(slideHasAnimations(slide)).toBe(false);
    expect(slideXml(pres)).not.toContain('<p:timing>');
  });

  it('refuses an effect it cannot edit and leaves the slide alone', async () => {
    // A composite effect that drives two shapes at once: safe to read, not to
    // take out from under whatever else depends on it.
    const { pres, slide, spid } = await withTiming(
      timingRoot(
        mainSeq(
          `<p:par><p:cTn id="3" fill="hold"><p:stCondLst><p:cond delay="indefinite"/>` +
            `</p:stCondLst><p:childTnLst><p:par><p:cTn id="4" fill="hold"><p:stCondLst>` +
            `<p:cond delay="0"/></p:stCondLst><p:childTnLst>` +
            `<p:par><p:cTn id="5" presetID="10" presetClass="entr" presetSubtype="0" fill="hold" ` +
            `grpId="0" nodeType="clickEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst>` +
            `<p:childTnLst><p:anim calcmode="lin" valueType="num"><p:cBhvr additive="base">` +
            `<p:cTn id="6" dur="500" fill="hold"/><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>` +
            `<p:attrNameLst><p:attrName>style.opacity</p:attrName></p:attrNameLst></p:cBhvr>` +
            `<p:tavLst/></p:anim><p:anim calcmode="lin" valueType="num"><p:cBhvr additive="base">` +
            `<p:cTn id="7" dur="500" fill="hold"/><p:tgtEl><p:spTgt spid="999"/></p:tgtEl>` +
            `<p:attrNameLst><p:attrName>style.opacity</p:attrName></p:attrNameLst></p:cBhvr>` +
            `<p:tavLst/></p:anim></p:childTnLst></p:cTn></p:par>` +
            `</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>`,
        ),
        BLD,
      ),
    );
    void spid;
    const before = slideXml(pres);
    expect(() => removeSlideAnimation(slide, 5)).toThrow(/cannot edit safely/);
    expect(slideXml(pres)).toBe(before);
  });

  it('refuses when another time node starts from the effect being removed', async () => {
    const { pres, slide } = await withTiming(
      timingRoot(
        mainSeq(stop()) +
          `<p:seq concurrent="1" nextAc="seek"><p:cTn id="10" nodeType="interactiveSeq">` +
          `<p:stCondLst><p:cond evt="onEnd" delay="0"><p:tn val="5"/></p:cond></p:stCondLst>` +
          `<p:childTnLst><p:par><p:cTn id="11" fill="hold"><p:stCondLst>` +
          `<p:cond delay="0"/></p:stCondLst><p:childTnLst>` +
          `<p:par><p:cTn id="12" presetID="10" presetClass="exit" presetSubtype="0" fill="hold" ` +
          `grpId="1" nodeType="clickEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst>` +
          `<p:childTnLst><p:anim calcmode="lin" valueType="num"><p:cBhvr additive="base">` +
          `<p:cTn id="13" dur="500" fill="hold"/><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>` +
          `<p:attrNameLst><p:attrName>style.opacity</p:attrName></p:attrNameLst></p:cBhvr>` +
          `<p:tavLst/></p:anim></p:childTnLst></p:cTn></p:par>` +
          `</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:seq>`,
        BLD,
      ),
    );
    const before = slideXml(pres);
    expect(() => removeSlideAnimation(slide, 5)).toThrow(/waiting on nothing/);
    expect(slideXml(pres)).toBe(before);
  });
});

describe('fn API: updateSlideAnimation', () => {
  it('changes how long an effect runs and retimes what follows it', () => {
    const { pres, slide } = deck();
    const a = rect(slide);
    const b = rect(slide);
    setShapeAnimation(a, { effect: 'fadeIn', durationMs: 300 });
    setShapeAnimation(b, { effect: 'fadeIn', durationMs: 200, start: 'afterPrevious' });
    const ids = getSlideAnimations(slide).map((s) => s.id!);

    updateSlideAnimation(slide, ids[0]!, { durationMs: 1200 });
    expect(getSlideAnimations(slide).map((s) => s.durationMs)).toEqual([1200, 200]);
    expect(wrapperDelays(slideXml(pres))).toEqual(['indefinite', '0', '1200']);
  });

  it('changes a start condition and folds the effect into the stop before it', () => {
    const { pres, slide, ids } = threeClicks();
    updateSlideAnimation(slide, ids[1]!, { start: 'withPrevious' });

    expect(getSlideAnimations(slide).map((s) => s.start)).toEqual([
      'click',
      'withPrevious',
      'click',
    ]);
    // Two stops now, and the first holds both of the effects that run together.
    expect(wrapperDelays(slideXml(pres))).toEqual(['indefinite', '0', 'indefinite', '0']);
  });

  it('changes the preset while keeping the handle', () => {
    const { slide, ids } = threeClicks();
    updateSlideAnimation(slide, ids[0]!, { effect: 'fadeOut' });
    const steps = getSlideAnimations(slide);
    expect(steps.map((s) => s.id)).toEqual(ids);
    expect(steps[0]!.effect).toBe('fadeOut');
    expect(steps[0]!.presetClass).toBe('exit');
  });

  it('sets a delay without touching anything else about the effect', () => {
    const { slide, ids } = threeClicks();
    updateSlideAnimation(slide, ids[0]!, { delayMs: 250 });
    const steps = getSlideAnimations(slide);
    expect(steps[0]!.delayMs).toBe(250);
    expect(steps[0]!.durationMs).toBe(400);
    expect(steps[0]!.effect).toBe('fadeIn');
  });

  // An effect this library did not author can carry behaviours it does not
  // model. A timing change must not quietly replace it with a stock preset.
  it('keeps an authored effect intact when only its delay changes', async () => {
    const { pres, slide } = await withTiming(
      timingRoot(
        mainSeq(
          `<p:par><p:cTn id="3" fill="hold"><p:stCondLst><p:cond delay="indefinite"/>` +
            `</p:stCondLst><p:childTnLst><p:par><p:cTn id="4" fill="hold"><p:stCondLst>` +
            `<p:cond delay="0"/></p:stCondLst><p:childTnLst>` +
            `<p:par><p:cTn id="5" presetID="10" presetClass="entr" presetSubtype="0" fill="hold" ` +
            `grpId="0" nodeType="clickEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst>` +
            `<p:childTnLst><p:animRot by="900000"><p:cBhvr>` +
            `<p:cTn id="6" dur="500" fill="hold"/><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>` +
            `</p:cBhvr></p:animRot></p:childTnLst></p:cTn></p:par>` +
            `</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>`,
        ),
        BLD,
      ),
    );
    updateSlideAnimation(slide, 5, { delayMs: 120 });
    const xml = slideXml(pres);
    // The rotation the file authored is still there, and only the delay moved.
    expect(xml).toContain('<p:animRot by="900000">');
    expect(getSlideAnimations(slide)[0]!.delayMs).toBe(120);
  });

  it('refuses to replace an authored effect when the preset changes', async () => {
    const { pres, slide } = await withTiming(
      timingRoot(
        mainSeq(
          `<p:par><p:cTn id="3" fill="hold"><p:stCondLst><p:cond delay="indefinite"/>` +
            `</p:stCondLst><p:childTnLst><p:par><p:cTn id="4" fill="hold"><p:stCondLst>` +
            `<p:cond delay="0"/></p:stCondLst><p:childTnLst>` +
            `<p:par><p:cTn id="5" presetID="10" presetClass="entr" presetSubtype="0" fill="hold" ` +
            `grpId="0" nodeType="clickEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst>` +
            `<p:childTnLst><p:animRot by="900000"><p:cBhvr>` +
            `<p:cTn id="6" dur="500" fill="hold"/><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>` +
            `</p:cBhvr></p:animRot></p:childTnLst></p:cTn></p:par>` +
            `</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>`,
        ),
        BLD,
      ),
    );
    const before = slideXml(pres);
    expect(() => updateSlideAnimation(slide, 5, { effect: 'fadeOut' })).toThrow(
      /did not author|lose that/,
    );
    expect(slideXml(pres)).toBe(before);
  });

  it('refuses a negative delay and leaves the slide alone', () => {
    const { pres, slide, ids } = threeClicks();
    const before = slideXml(pres);
    expect(() => updateSlideAnimation(slide, ids[0]!, { delayMs: -5 })).toThrow(
      /updateSlideAnimation: delayMs/,
    );
    expect(slideXml(pres)).toBe(before);
  });

  skipIfNoXmllint('emits a schema-valid tree after a change of start', () => {
    const { pres, slide, ids } = threeClicks();
    updateSlideAnimation(slide, ids[2]!, { start: 'afterPrevious', delayMs: 100 });
    expectSchemaValid(slideXml(pres), 'pml');
  });
});

describe('fn API: updateSlideAnimation — turning a paragraph build on and off', () => {
  const textDeck = (
    lines: readonly string[],
  ): { pres: PresentationData; slide: SlideData; shape: SlideShapeData } => {
    const pres = createPresentation();
    const slide = addBlankSlide(pres);
    const shape = addSlideTextBox(slide, { ...box, text: lines.join('\n') });
    return { pres, slide, shape };
  };

  it('turns one effect into one per paragraph, keeping the addressed handle', () => {
    const { pres, slide, shape } = textDeck(['One', 'Two', 'Three']);
    setShapeAnimation(shape, { effect: 'fadeIn' });
    const id = getSlideAnimations(slide)[0]!.id!;

    updateSlideAnimation(slide, id, { byParagraph: true });
    const steps = getSlideAnimations(slide);
    expect(steps).toHaveLength(3);
    expect(steps[0]!.id).toBe(id);
    expect(steps.map((s) => s.target)).toEqual([
      { kind: 'paragraphs', shapeId: getShapeId(shape), firstParagraph: 0, lastParagraph: 0 },
      { kind: 'paragraphs', shapeId: getShapeId(shape), firstParagraph: 1, lastParagraph: 1 },
      { kind: 'paragraphs', shapeId: getShapeId(shape), firstParagraph: 2, lastParagraph: 2 },
    ]);
    expect(steps.every((s) => s.buildByParagraph)).toBe(true);
    expect(new Set(steps.map((s) => s.id)).size).toBe(3);
    // Still one build entry, now saying the body builds a paragraph at a time.
    expect(slideXml(pres).match(/<p:bldP\b/g)).toHaveLength(1);
    expect(slideXml(pres)).toContain(`<p:bldP spid="${getShapeId(shape)}" grpId="0" build="p"/>`);
  });

  it('collapses a build back into one effect on the shape', () => {
    const { pres, slide, shape } = textDeck(['One', 'Two', 'Three']);
    setShapeAnimation(shape, { effect: 'fadeIn', byParagraph: true });
    const ids = getSlideAnimations(slide).map((s) => s.id!);

    updateSlideAnimation(slide, ids[1]!, { byParagraph: false });
    const steps = getSlideAnimations(slide);
    expect(steps).toHaveLength(1);
    // The handle passed in survives; the other paragraphs' handles are gone.
    expect(steps[0]!.id).toBe(ids[1]);
    expect(steps[0]!.target).toEqual({ kind: 'shape', shapeId: getShapeId(shape) });
    expect(steps[0]!.buildByParagraph).toBe(false);
    expect(slideXml(pres)).toContain(`<p:bldP spid="${getShapeId(shape)}" grpId="0"/>`);
    expect(slideXml(pres)).not.toContain('build="p"');
  });

  skipIfNoXmllint('emits a schema-valid tree either way', () => {
    const { pres, slide, shape } = textDeck(['One', 'Two']);
    setShapeAnimation(shape, { effect: 'fadeIn' });
    const id = getSlideAnimations(slide)[0]!.id!;
    updateSlideAnimation(slide, id, { byParagraph: true });
    expectSchemaValid(slideXml(pres), 'pml');
    updateSlideAnimation(slide, id, { byParagraph: false });
    expectSchemaValid(slideXml(pres), 'pml');
  });

  it('refuses to build a shape with no text and changes nothing', () => {
    const { pres, slide } = deck();
    const shape = rect(slide);
    setShapeAnimation(shape, { effect: 'fadeIn' });
    const id = getSlideAnimations(slide)[0]!.id!;
    const before = slideXml(pres);
    expect(() => updateSlideAnimation(slide, id, { byParagraph: true })).toThrow(
      /byParagraph needs a shape with text/,
    );
    expect(slideXml(pres)).toBe(before);
  });

  // The paragraphs of one build share a single `<p:bldP>`: taking one paragraph
  // out must leave the entry the others still need.
  it('keeps the build entry while other paragraphs still use it', () => {
    const { pres, slide, shape } = textDeck(['One', 'Two', 'Three']);
    setShapeAnimation(shape, { effect: 'fadeIn', byParagraph: true });
    const ids = getSlideAnimations(slide).map((s) => s.id!);

    removeSlideAnimation(slide, ids[1]!);
    expect(getSlideAnimations(slide).map((s) => s.id)).toEqual([ids[0], ids[2]]);
    expect(slideXml(pres).match(/<p:bldP\b/g)).toHaveLength(1);
    expect(getSlideAnimations(slide).every((s) => s.buildByParagraph)).toBe(true);

    removeSlideAnimation(slide, ids[0]!);
    expect(slideXml(pres).match(/<p:bldP\b/g)).toHaveLength(1);
  });

  it('drops the build entry with the last paragraph that used it', () => {
    const { pres, slide, shape } = textDeck(['One', 'Two']);
    const other = rect(slide);
    setShapeAnimation(shape, { effect: 'fadeIn', byParagraph: true });
    setShapeAnimation(other, { effect: 'fadeIn' });
    const ids = getSlideAnimations(slide).map((s) => s.id!);
    expect(slideXml(pres).match(/<p:bldP\b/g)).toHaveLength(2);

    removeSlideAnimation(slide, ids[0]!);
    removeSlideAnimation(slide, ids[1]!);
    // The other shape's entry is untouched; the build's is gone with it.
    const xml = slideXml(pres);
    expect(xml.match(/<p:bldP\b/g)).toHaveLength(1);
    expect(xml).toContain(`<p:bldP spid="${getShapeId(other)}"`);
    expect(xml).not.toContain('build="p"');
  });
});

describe('fn API: animation edits this library refuses to lay out again', () => {
  // Reordering rewrites every stop and group offset, so a sequence whose
  // offsets are not the ones we would write is one we cannot reproduce.
  it('refuses a stop that waits a fixed time rather than for a click', async () => {
    const { pres, slide } = await withTiming(timingRoot(mainSeq(stop({ stopDelay: '250' })), BLD));
    const before = slideXml(pres);
    expect(() => removeSlideAnimation(slide, 5)).toThrow(/cannot lay out again/);
    expect(slideXml(pres)).toBe(before);
  });

  it('refuses a group placed somewhere other than where we would put it', async () => {
    const { pres, slide } = await withTiming(timingRoot(mainSeq(stop({ groupDelay: '900' })), BLD));
    const before = slideXml(pres);
    expect(() => removeSlideAnimation(slide, 5)).toThrow(/cannot lay out again/);
    expect(slideXml(pres)).toBe(before);
  });

  it('refuses a wrapper whose fill is not the one we write', async () => {
    const { pres, slide } = await withTiming(
      timingRoot(mainSeq(stop({ stopFill: 'remove' })), BLD),
    );
    const before = slideXml(pres);
    expect(() => removeSlideAnimation(slide, 5)).toThrow(/cannot lay out again/);
    expect(slideXml(pres)).toBe(before);
  });

  it('refuses when something starts from a wrapper the layout would replace', async () => {
    const { pres, slide } = await withTiming(
      timingRoot(
        mainSeq(stop()) +
          `<p:seq concurrent="1" nextAc="seek"><p:cTn id="10" nodeType="interactiveSeq">` +
          `<p:stCondLst><p:cond evt="onEnd" delay="0"><p:tn val="4"/></p:cond></p:stCondLst>` +
          `<p:childTnLst><p:par><p:cTn id="11" fill="hold"><p:stCondLst>` +
          `<p:cond delay="0"/></p:stCondLst><p:childTnLst>` +
          `<p:par><p:cTn id="12" presetID="10" presetClass="exit" presetSubtype="0" fill="hold" ` +
          `grpId="1" nodeType="clickEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst>` +
          `<p:childTnLst><p:anim calcmode="lin" valueType="num"><p:cBhvr additive="base">` +
          `<p:cTn id="13" dur="500" fill="hold"/><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>` +
          `<p:attrNameLst><p:attrName>style.opacity</p:attrName></p:attrNameLst></p:cBhvr>` +
          `<p:tavLst/></p:anim></p:childTnLst></p:cTn></p:par>` +
          `</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:seq>`,
        BLD,
      ),
    );
    const before = slideXml(pres);
    expect(() => moveSlideAnimation(slide, 5, 0)).toThrow(/cannot lay out again/);
    expect(slideXml(pres)).toBe(before);
  });
});

describe('fn API: removing a shape takes its animations with it', () => {
  it('drops the effects bound to the shape and retimes the rest', () => {
    const { pres, slide, shapes, ids } = threeClicks();
    removeShape(shapes[1]!);

    const steps = getSlideAnimations(slide);
    expect(steps.map((s) => s.id)).toEqual([ids[0], ids[2]]);
    expect(steps.map((s) => s.target.shapeId)).toEqual([
      getShapeId(shapes[0]!),
      getShapeId(shapes[2]!),
    ]);
    expect(slideXml(pres).match(/<p:bldP\b/g)).toHaveLength(2);
  });

  it('leaves no timing behind when the last animated shape goes', () => {
    const { pres, slide } = deck();
    const shape = rect(slide);
    setShapeAnimation(shape, { effect: 'fadeIn' });
    removeShape(shape);
    expect(slideXml(pres)).not.toContain('<p:timing>');
    expect(slideHasAnimations(slide)).toBe(false);
  });

  it('takes a whole paragraph build with the shape that held the text', () => {
    const pres = createPresentation();
    const slide = addBlankSlide(pres);
    const text = addSlideTextBox(slide, { ...box, text: 'One\nTwo\nThree' });
    const other = rect(slide);
    setShapeAnimation(text, { effect: 'fadeIn', byParagraph: true });
    setShapeAnimation(other, { effect: 'fadeIn' });

    removeShape(text);
    const steps = getSlideAnimations(slide);
    expect(steps).toHaveLength(1);
    expect(steps[0]!.target.shapeId).toBe(getShapeId(other));
    const xml = slideXml(pres);
    expect(xml.match(/<p:bldP\b/g)).toHaveLength(1);
    expect(xml).not.toContain(`spid="${getShapeId(text)}"`);
  });

  it('moves an after-previous effect up when the shape it followed goes', () => {
    const { pres, slide } = deck();
    const a = rect(slide);
    const b = rect(slide);
    const c = rect(slide);
    setShapeAnimation(a, { effect: 'fadeIn', durationMs: 300 });
    setShapeAnimation(b, { effect: 'fadeIn', durationMs: 700, start: 'afterPrevious' });
    setShapeAnimation(c, { effect: 'fadeIn', durationMs: 100, start: 'afterPrevious' });
    expect(wrapperDelays(slideXml(pres))).toEqual(['indefinite', '0', '300', '1000']);

    removeShape(b);
    expect(wrapperDelays(slideXml(pres))).toEqual(['indefinite', '0', '300']);
    expect(getSlideAnimations(slide).map((s) => s.durationMs)).toEqual([300, 100]);
  });

  it('takes the animations of a group’s children with the group', () => {
    const { pres, slide } = deck();
    const a = rect(slide);
    const b = rect(slide);
    const outside = rect(slide);
    setShapeAnimation(a, { effect: 'fadeIn' });
    setShapeAnimation(b, { effect: 'fadeIn' });
    setShapeAnimation(outside, { effect: 'fadeIn' });
    const group = groupShapes([a, b]);

    removeShape(group);
    const steps = getSlideAnimations(slide);
    expect(steps).toHaveLength(1);
    expect(steps[0]!.target.shapeId).toBe(getShapeId(outside));
    expect(slideXml(pres).match(/<p:bldP\b/g)).toHaveLength(1);
  });

  skipIfNoXmllint('leaves a schema-valid tree behind', () => {
    const { pres, shapes } = threeClicks();
    removeShape(shapes[0]!);
    expectSchemaValid(slideXml(pres), 'pml');
  });
});

describe('fn API: removals the slide’s timing could not survive', () => {
  /** Two shapes, then hand-written timing that names them both. */
  const twoShapes = async (
    timing: (a: number, b: number) => string,
  ): Promise<{ pres: PresentationData; slide: SlideData; shapes: SlideShapeData[] }> => {
    const pres = createPresentation();
    const slide = addBlankSlide(pres);
    const [a, b] = [rect(slide), rect(slide)];
    const reloaded = await withSplicedTiming(pres, timing(getShapeId(a), getShapeId(b)));
    const only = getSlides(reloaded)[0]!;
    return { pres: reloaded, slide: only, shapes: [...getSlideShapes(only)] };
  };

  it('refuses a shape that only triggers an interactive sequence', async () => {
    const { pres, slide, shapes } = await twoShapes((a, b) =>
      timingRoot(
        mainSeq(stopOn(b)) + interactiveSeq(a, stopOn(b, { firstId: 21, grpId: '1' })),
        `<p:bldLst><p:bldP spid="${b}" grpId="0"/><p:bldP spid="${b}" grpId="1"/></p:bldLst>`,
      ),
    );
    const before = slideXml(pres);

    expect(() => removeShape(shapes[0]!)).toThrow(/does not model/);
    expect(slideXml(pres)).toBe(before);
    expect(getSlideShapes(slide)).toHaveLength(2);
  });

  it('refuses a shape whose effect also drives one that is staying', async () => {
    const { pres, slide, shapes } = await twoShapes((a, b) =>
      timingRoot(
        mainSeq(
          `<p:par><p:cTn id="3" fill="hold"><p:stCondLst><p:cond delay="indefinite"/>` +
            `</p:stCondLst><p:childTnLst><p:par><p:cTn id="4" fill="hold"><p:stCondLst>` +
            `<p:cond delay="0"/></p:stCondLst><p:childTnLst>` +
            `<p:par><p:cTn id="5" presetID="10" presetClass="entr" presetSubtype="0" fill="hold" ` +
            `grpId="0" nodeType="clickEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst>` +
            `<p:childTnLst><p:anim calcmode="lin" valueType="num"><p:cBhvr additive="base">` +
            `<p:cTn id="6" dur="500" fill="hold"/><p:tgtEl><p:spTgt spid="${a}"/></p:tgtEl>` +
            `<p:attrNameLst><p:attrName>style.opacity</p:attrName></p:attrNameLst></p:cBhvr>` +
            `<p:tavLst/></p:anim><p:anim calcmode="lin" valueType="num"><p:cBhvr additive="base">` +
            `<p:cTn id="7" dur="500" fill="hold"/><p:tgtEl><p:spTgt spid="${b}"/></p:tgtEl>` +
            `<p:attrNameLst><p:attrName>style.opacity</p:attrName></p:attrNameLst></p:cBhvr>` +
            `<p:tavLst/></p:anim></p:childTnLst></p:cTn></p:par>` +
            `</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>`,
        ),
        `<p:bldLst><p:bldP spid="${a}" grpId="0"/></p:bldLst>`,
      ),
    );
    const before = slideXml(pres);

    expect(() => removeShape(shapes[0]!)).toThrow(/drives both this shape/);
    expect(slideXml(pres)).toBe(before);
    expect(getSlideShapes(slide)).toHaveLength(2);
    expect(getSlideAnimations(slide)).toHaveLength(1);
  });

  it('keeps the whole slide when a removal is refused, package included', async () => {
    const { pres, shapes } = await twoShapes((a, b) =>
      timingRoot(
        mainSeq(stopOn(b)) + interactiveSeq(a, stopOn(b, { firstId: 21, grpId: '1' })),
        `<p:bldLst><p:bldP spid="${b}" grpId="0"/><p:bldP spid="${b}" grpId="1"/></p:bldLst>`,
      ),
    );
    const before = await savePresentation(pres);

    expect(() => removeShape(shapes[0]!)).toThrow();
    expect(await savePresentation(pres)).toEqual(before);
  });
});

describe('fn API: copying a shape copies its animations', () => {
  it('gives the copy its own effect and its own build group', () => {
    const { pres, slide } = deck();
    const shape = rect(slide);
    setShapeAnimation(shape, { effect: 'fadeIn', durationMs: 750, delayMs: 100 });

    const copy = copyShape(slide, shape);
    const steps = getSlideAnimations(slide);
    expect(steps.map((s) => s.target.shapeId)).toEqual([getShapeId(shape), getShapeId(copy)]);
    expect(steps.map((s) => s.durationMs)).toEqual([750, 750]);
    expect(steps.map((s) => s.delayMs)).toEqual([100, 100]);
    expect(new Set(steps.map((s) => s.id)).size).toBe(2);

    const xml = slideXml(pres);
    expect(xml.match(/<p:bldP\b/g)).toHaveLength(2);
    expect(xml).toContain(`<p:bldP spid="${getShapeId(copy)}" grpId="1"/>`);
  });

  it('copies a paragraph build as a build over the copy’s own paragraphs', () => {
    const pres = createPresentation();
    const slide = addBlankSlide(pres);
    const text = addSlideTextBox(slide, { ...box, text: 'One\nTwo\nThree' });
    setShapeAnimation(text, { effect: 'fadeIn', byParagraph: true });

    const copy = copyShape(slide, text);
    const steps = getSlideAnimations(slide);
    expect(steps).toHaveLength(6);
    expect(steps.slice(3).map((s) => s.target)).toEqual([
      { kind: 'paragraphs', shapeId: getShapeId(copy), firstParagraph: 0, lastParagraph: 0 },
      { kind: 'paragraphs', shapeId: getShapeId(copy), firstParagraph: 1, lastParagraph: 1 },
      { kind: 'paragraphs', shapeId: getShapeId(copy), firstParagraph: 2, lastParagraph: 2 },
    ]);
    // One entry per build, not one per paragraph.
    expect(slideXml(pres).match(/<p:bldP\b/g)).toHaveLength(2);
    expect(new Set(steps.map((s) => s.id)).size).toBe(6);
  });

  it('copies the animations of a group’s children under their new ids', () => {
    const { pres, slide } = deck();
    const a = rect(slide);
    const b = rect(slide);
    setShapeAnimation(a, { effect: 'fadeIn' });
    setShapeAnimation(b, { effect: 'fadeOut' });
    const group = groupShapes([a, b]);

    const copy = copyShape(slide, group);
    void copy;

    const steps = getSlideAnimations(slide);
    expect(steps).toHaveLength(4);
    expect(steps.map((s) => s.effect)).toEqual(['fadeIn', 'fadeOut', 'fadeIn', 'fadeOut']);
    // Every effect names a shape that is really on the slide, and no id twice.
    const ids = steps.map((s) => s.target.shapeId!);
    expect(new Set(ids).size).toBe(4);
    const xml = slideXml(pres);
    for (const id of ids) expect(xml).toContain(`spid="${id}"`);
    expect(xml.match(/<p:bldP\b/g)).toHaveLength(4);
  });

  it('copies across slides without touching the source', async () => {
    const pres = createPresentation();
    const source = addBlankSlide(pres);
    const target = addBlankSlide(pres);
    const shape = addSlideShape(source, { preset: 'rect', ...box });
    setShapeAnimation(shape, { effect: 'fadeIn', durationMs: 600 });
    const before = getSlideAnimations(source);

    copyShape(target, shape);
    expect(getSlideAnimations(source)).toEqual(before);
    const copied = getSlideAnimations(target);
    expect(copied).toHaveLength(1);
    expect(copied[0]!.durationMs).toBe(600);

    const reloaded = await loadPresentation(await savePresentation(pres));
    expect(getSlideAnimations(getSlides(reloaded)[1]!)).toEqual(copied);
  });

  it('copies a build paragraph by paragraph, edits and all', () => {
    const pres = createPresentation();
    const slide = addBlankSlide(pres);
    const text = addSlideTextBox(slide, { ...box, text: 'One\nTwo\nThree' });
    setShapeAnimation(text, { effect: 'fadeIn', byParagraph: true });
    const ids = getSlideAnimations(slide).map((s) => s.id!);
    updateSlideAnimation(slide, ids[1]!, { durationMs: 900, delayMs: 120 });
    updateSlideAnimation(slide, ids[2]!, { effect: 'appear' });

    const copy = copyShape(slide, text);
    const steps = getSlideAnimations(slide);
    const shape = (s: (typeof steps)[number]): number | null => s.target.shapeId;
    const source = steps.filter((s) => shape(s) === getShapeId(text));
    const copied = steps.filter((s) => shape(s) === getShapeId(copy));

    expect(copied.map((s) => [s.effect, s.durationMs, s.delayMs])).toEqual(
      source.map((s) => [s.effect, s.durationMs, s.delayMs]),
    );
    expect(copied.map((s) => s.effect)).toEqual(['fadeIn', 'fadeIn', 'appear']);
    // `appear` is instantaneous, so it states no duration of its own.
    expect(copied.map((s) => s.durationMs)).toEqual([500, 900, null]);
    expect(copied.map((s) => s.delayMs)).toEqual([0, 120, 0]);
    // One entry for the whole build, carrying the paragraph build flag.
    expect(slideXml(pres)).toContain(`<p:bldP spid="${getShapeId(copy)}" grpId="1" build="p"/>`);
  });

  it('does not bring back a paragraph taken out of the build', () => {
    const pres = createPresentation();
    const slide = addBlankSlide(pres);
    const text = addSlideTextBox(slide, { ...box, text: 'One\nTwo\nThree' });
    setShapeAnimation(text, { effect: 'fadeIn', byParagraph: true });
    removeSlideAnimation(slide, getSlideAnimations(slide)[1]!.id!);

    const copy = copyShape(slide, text);
    const copied = getSlideAnimations(slide).filter((s) => s.target.shapeId === getShapeId(copy));
    expect(copied.map((s) => s.target)).toEqual([
      { kind: 'paragraphs', shapeId: getShapeId(copy), firstParagraph: 0, lastParagraph: 0 },
      { kind: 'paragraphs', shapeId: getShapeId(copy), firstParagraph: 2, lastParagraph: 2 },
    ]);
    expect(slideXml(pres).match(/<p:bldP\b/g)).toHaveLength(2);
  });

  it('copies an effect it can only read, rather than dropping it', async () => {
    const { pres, slide, spid } = await withTiming(
      timingRoot(mainSeq(stop({ presetId: '22' })), BLD),
    );
    const shape = getSlideShapes(slide)[0]!;
    expect(getSlideAnimations(slide)[0]!.effect).toBe(null);

    const copy = copyShape(slide, shape);
    const xml = slideXml(pres);
    expect(xml.match(/presetID="22"/g)).toHaveLength(2);
    expect(getSlideAnimations(slide).map((s) => s.target.shapeId)).toEqual([
      spid,
      getShapeId(copy),
    ]);
    expect(getSlideAnimations(slide)[1]!.presetId).toBe(22);
  });

  it('adds nothing to the target when the animations cannot come along', async () => {
    const source = createPresentation();
    const sourceSlide = addBlankSlide(source);
    const picture = addSlideImage(sourceSlide, PNG, box);
    const other = addSlideShape(sourceSlide, { preset: 'rect', ...box });
    const reloaded = await withSplicedTiming(
      source,
      timingRoot(
        mainSeq(stopOn(getShapeId(other))) +
          interactiveSeq(
            getShapeId(other),
            stopOn(getShapeId(picture), { firstId: 21, grpId: '1' }),
          ),
        `<p:bldLst><p:bldP spid="${getShapeId(other)}" grpId="0"/>` +
          `<p:bldP spid="${getShapeId(picture)}" grpId="1"/></p:bldLst>`,
      ),
    );
    const pic = getSlideShapes(getSlides(reloaded)[0]!)[0]!;

    const target = createPresentation();
    const targetSlide = addBlankSlide(target);
    const before = await savePresentation(target);

    expect(() => copyShape(targetSlide, pic)).toThrow(/interactive sequence/);
    expect(getSlideShapes(targetSlide)).toHaveLength(0);
    // No shape, no relationship, and no media part: the whole package is as it
    // was before the copy was asked for.
    expect(await savePresentation(target)).toEqual(before);
  });

  skipIfNoXmllint('emits a schema-valid tree for the copy', () => {
    const { pres, slide } = deck();
    const shape = rect(slide);
    setShapeAnimation(shape, { effect: 'fadeIn' });
    copyShape(slide, shape);
    expectSchemaValid(slideXml(pres), 'pml');
  });
});
