// getSlideAnimations — the ordered read of a slide's <p:timing> tree.
//
// Several cases splice hand-written timing XML into a real slide part: the
// authoring API cannot produce a composite effect, a duplicated `<p:cTn id>`
// or a paragraph build, and those are exactly the trees a template drops on us.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { partName } from '../src/internal/opc/index.ts';
import {
  type PresentationData,
  type SlideData,
  _internalPackageOf,
  addBlankSlide,
  addSlideMedia,
  addSlideShape,
  createPresentation,
  findShapesWithAnimation,
  getShapeAnimation,
  getShapeId,
  getSlideAnimations,
  getSlideShapes,
  getSlides,
  inches,
  loadPresentation,
  savePresentation,
  setShapeAnimation,
  slideHasAnimations,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const SLIDE1 = '/ppt/slides/slide1.xml';
const decoder = new TextDecoder();
const encoder = new TextEncoder();

/** Replaces slide 1's `<p:timing>` with `timing` and reloads the deck. */
const withTiming = async (
  timing: string,
  source?: Uint8Array,
): Promise<{ pres: PresentationData; slide: SlideData }> => {
  const pres = await loadPresentation(source ?? (await readFile(fixture('one-text-slide.pptx'))));
  const part = _internalPackageOf(pres).getPart(partName(SLIDE1));
  if (part === null) throw new Error('missing slide part');
  const xml = decoder.decode(part.data).replace(/<p:timing>[\s\S]*<\/p:timing>/, '');
  part.data = encoder.encode(xml.replace('</p:sld>', `${timing}</p:sld>`));
  const reloaded = await loadPresentation(await savePresentation(pres));
  return { pres: reloaded, slide: getSlides(reloaded)[0]! };
};

/** The spid of slide 1's first shape in `one-text-slide.pptx`. */
const firstShapeId = async (): Promise<number> => {
  const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
  return getShapeId(getSlideShapes(getSlides(pres)[0]!)[0]!);
};

const timingRoot = (seqChildren: string, bldLst = ''): string =>
  `<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">` +
  `<p:childTnLst>${seqChildren}</p:childTnLst></p:cTn></p:par></p:tnLst>${bldLst}</p:timing>`;

const mainSeq = (childTnLst: string): string =>
  `<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq">` +
  `<p:childTnLst>${childTnLst}</p:childTnLst></p:cTn></p:seq>`;

const box = { x: inches(1), y: inches(1), w: inches(2), h: inches(1) };

describe('fn API: getSlideAnimations', () => {
  it('lists click effects in the order they were added, with unique handles', async () => {
    const pres = createPresentation();
    const slide = addBlankSlide(pres);
    const shapes = ['rect', 'ellipse', 'triangle'].map((preset) =>
      addSlideShape(slide, { preset, ...box }),
    );
    setShapeAnimation(shapes[0]!, { effect: 'fadeIn' });
    setShapeAnimation(shapes[1]!, { effect: 'appear' });
    setShapeAnimation(shapes[2]!, { effect: 'fadeOut', durationMs: 1200 });

    const steps = getSlideAnimations(slide);
    expect(steps.map((s) => s.effect)).toEqual(['fadeIn', 'appear', 'fadeOut']);
    expect(steps.map((s) => s.target.shapeId)).toEqual(shapes.map(getShapeId));
    expect(steps.map((s) => s.start)).toEqual(['click', 'click', 'click']);
    expect(steps.map((s) => s.sequence)).toEqual(['mainSeq', 'mainSeq', 'mainSeq']);
    expect(steps.every((s) => s.playable)).toBe(true);
    expect(steps.every((s) => s.editable)).toBe(true);
    expect(new Set(steps.map((s) => s.id)).size).toBe(3);
    expect(steps.every((s) => s.id !== null)).toBe(true);
  });

  it('reports the authored duration and leaves an unstated one null', async () => {
    const pres = createPresentation();
    const slide = addBlankSlide(pres);
    const fade = addSlideShape(slide, { preset: 'rect', ...box });
    const instant = addSlideShape(slide, { preset: 'rect', ...box });
    setShapeAnimation(fade, { effect: 'fadeIn', durationMs: 900 });
    setShapeAnimation(instant, { effect: 'appear', durationMs: 900 });

    const steps = getSlideAnimations(slide);
    expect(steps[0]!.durationMs).toBe(900);
    // `appear` is instantaneous: the builder writes no animated behaviour, so
    // there is no duration to report and we must not invent the 500ms default.
    expect(steps[1]!.durationMs).toBeNull();
    expect(steps.map((s) => s.delayMs)).toEqual([0, 0]);
  });

  it('keeps both effects when one shape carries an entrance and an exit', async () => {
    const pres = createPresentation();
    const slide = addBlankSlide(pres);
    const shape = addSlideShape(slide, { preset: 'rect', ...box });
    setShapeAnimation(shape, { effect: 'fadeIn' });
    setShapeAnimation(shape, { effect: 'fadeOut' });

    const steps = getSlideAnimations(slide);
    expect(steps.map((s) => s.effect)).toEqual(['fadeIn', 'fadeOut']);
    expect(steps[0]!.id).not.toBe(steps[1]!.id);
    // The single-effect read stays on its documented contract: first one wins.
    expect(getShapeAnimation(shape)).toBe('fadeIn');
  });

  it('survives a save / reload round trip unchanged', async () => {
    const pres = createPresentation();
    const slide = addBlankSlide(pres);
    const a = addSlideShape(slide, { preset: 'rect', ...box });
    const b = addSlideShape(slide, { preset: 'ellipse', ...box });
    setShapeAnimation(a, { effect: 'fadeIn', durationMs: 750 });
    setShapeAnimation(b, { effect: 'disappear' });
    const before = getSlideAnimations(slide);

    const reloaded = await loadPresentation(await savePresentation(pres));
    expect(getSlideAnimations(getSlides(reloaded)[0]!)).toEqual(before);
  });

  it('is empty for a slide with no timing at all', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    expect(getSlideAnimations(slide)).toEqual([]);
    expect(slideHasAnimations(slide)).toBe(false);
  });

  it('is empty for a media-only timing tree', () => {
    const pres = createPresentation();
    const slide = addBlankSlide(pres);
    const ascii = (s: string): number[] => Array.from(s, (ch) => ch.charCodeAt(0));
    addSlideMedia(slide, {
      kind: 'video',
      data: new Uint8Array([0, 0, 0, 0x18, ...ascii('ftypmp42'), 0, 0, 0, 0, ...ascii('mp42isom')]),
      ...box,
    });
    expect(getSlideAnimations(slide)).toEqual([]);
    expect(slideHasAnimations(slide)).toBe(false);
  });

  it('is empty for an empty main sequence', async () => {
    const { slide } = await withTiming(timingRoot(mainSeq('')));
    expect(getSlideAnimations(slide)).toEqual([]);
    expect(slideHasAnimations(slide)).toBe(false);
  });

  it('is empty for a main sequence holding only media play controls', async () => {
    const spid = await firstShapeId();
    const media =
      `<p:audio><p:cMediaNode><p:cTn id="3" fill="hold" display="0"><p:stCondLst>` +
      `<p:cond delay="indefinite"/></p:stCondLst></p:cTn>` +
      `<p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl></p:cMediaNode></p:audio>`;
    const { slide } = await withTiming(timingRoot(mainSeq(media)));
    expect(getSlideAnimations(slide)).toEqual([]);
    expect(slideHasAnimations(slide)).toBe(false);
  });
});

// A `<p:cTn>` carrying an animation behaviour, with a click node type and no
// preset attributes — a custom effect PowerPoint or another tool authored.
const customEffect = (id: number, spid: number): string =>
  `<p:par><p:cTn id="${id}" nodeType="clickEffect" fill="hold">` +
  `<p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>` +
  `<p:anim calcmode="lin" valueType="num"><p:cBhvr><p:cTn id="${id + 1}" dur="640"/>` +
  `<p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl>` +
  `<p:attrNameLst><p:attrName>ppt_x</p:attrName></p:attrNameLst></p:cBhvr>` +
  `<p:tavLst/></p:anim></p:childTnLst></p:cTn></p:par>`;

const presetEffect = (
  id: number,
  spid: number,
  opts: { presetId?: string; presetClass?: string; nodeType?: string; target?: string } = {},
): string => {
  const presetId = opts.presetId ?? '10';
  const presetClass = opts.presetClass ?? 'entr';
  const nodeType = opts.nodeType ?? 'clickEffect';
  const target = opts.target ?? `<p:spTgt spid="${spid}"/>`;
  return (
    `<p:par><p:cTn id="${id}" presetID="${presetId}" presetClass="${presetClass}" ` +
    `presetSubtype="0" fill="hold" grpId="0" nodeType="${nodeType}">` +
    `<p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>` +
    `<p:anim calcmode="lin" valueType="num"><p:cBhvr additive="base">` +
    `<p:cTn id="${id + 1}" dur="500" fill="hold"/><p:tgtEl>${target}</p:tgtEl>` +
    `<p:attrNameLst><p:attrName>style.opacity</p:attrName></p:attrNameLst></p:cBhvr>` +
    `<p:tavLst/></p:anim></p:childTnLst></p:cTn></p:par>`
  );
};

describe('fn API: getSlideAnimations — trees this library did not author', () => {
  it('keeps a custom effect that carries no preset attributes, as read-only', async () => {
    const spid = await firstShapeId();
    const { slide } = await withTiming(
      timingRoot(mainSeq(customEffect(3, spid) + presetEffect(10, spid))),
    );
    const steps = getSlideAnimations(slide);
    expect(steps).toHaveLength(2);
    expect(steps[0]!.effect).toBeNull();
    expect(steps[0]!.presetId).toBeNull();
    expect(steps[0]!.presetClass).toBeNull();
    expect(steps[0]!.start).toBe('click');
    expect(steps[0]!.target).toEqual({ kind: 'shape', shapeId: spid });
    // We cannot re-author an effect we cannot name, so it stays read-only —
    // but the recognised effect next to it is still editable.
    expect(steps[0]!.editable).toBe(false);
    expect(steps[1]!.effect).toBe('fadeIn');
    expect(steps[1]!.editable).toBe(true);
  });

  it('keeps an effect whose preset attributes are half-written, as read-only', async () => {
    const spid = await firstShapeId();
    const { slide } = await withTiming(
      timingRoot(mainSeq(presetEffect(3, spid, { presetClass: '' }) + presetEffect(10, spid))),
    );
    const steps = getSlideAnimations(slide);
    expect(steps).toHaveLength(2);
    expect(steps[0]!.editable).toBe(false);
    expect(steps[1]!.editable).toBe(true);
  });

  it('reports an unmodelled node type as unknown without touching its neighbour', async () => {
    const spid = await firstShapeId();
    const { slide } = await withTiming(
      timingRoot(mainSeq(presetEffect(3, spid, { nodeType: 'mainSeq' }) + presetEffect(10, spid))),
    );
    const steps = getSlideAnimations(slide);
    expect(steps).toHaveLength(2);
    expect(steps[0]!.start).toBe('unknown');
    expect(steps[0]!.editable).toBe(false);
    expect(steps[1]!.start).toBe('click');
    expect(steps[1]!.editable).toBe(true);
  });

  it('refuses to claim a composite effect that drives two shapes', async () => {
    const spid = await firstShapeId();
    const composite =
      `<p:par><p:cTn id="3" presetID="10" presetClass="entr" presetSubtype="0" fill="hold" ` +
      `grpId="0" nodeType="clickEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst>` +
      `<p:childTnLst>` +
      `<p:set><p:cBhvr><p:cTn id="4" dur="1" fill="hold"/>` +
      `<p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl>` +
      `<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr>` +
      `<p:to><p:strVal val="visible"/></p:to></p:set>` +
      `<p:anim calcmode="lin" valueType="num"><p:cBhvr><p:cTn id="5" dur="500"/>` +
      `<p:tgtEl><p:spTgt spid="${spid + 1000}"/></p:tgtEl>` +
      `<p:attrNameLst><p:attrName>style.opacity</p:attrName></p:attrNameLst></p:cBhvr>` +
      `<p:tavLst/></p:anim>` +
      `</p:childTnLst></p:cTn></p:par>`;
    const { slide } = await withTiming(timingRoot(mainSeq(composite + presetEffect(10, spid))));

    const steps = getSlideAnimations(slide);
    expect(steps).toHaveLength(2);
    expect(steps[0]!.target).toEqual({ kind: 'unsupported', shapeId: null });
    expect(steps[0]!.editable).toBe(false);
    expect(steps[1]!.editable).toBe(true);
  });

  it('withholds the handle when a cTn id repeats or is not an integer', async () => {
    const spid = await firstShapeId();
    const { slide } = await withTiming(
      timingRoot(
        mainSeq(
          // Two effects sharing id 3, plus one whose id has trailing garbage.
          presetEffect(3, spid) +
            presetEffect(3, spid) +
            presetEffect(20, spid).replace('id="20"', 'id="20abc"'),
        ),
      ),
    );
    const steps = getSlideAnimations(slide);
    expect(steps).toHaveLength(3);
    expect(steps.map((s) => s.id)).toEqual([null, null, null]);
    expect(steps.every((s) => s.editable)).toBe(false);
  });

  it('reads a paragraph build target and its bldP settings', async () => {
    const spid = await firstShapeId();
    const target = `<p:spTgt spid="${spid}"><p:txEl><p:pRg st="1" end="2"/></p:txEl></p:spTgt>`;
    const bldLst = `<p:bldLst><p:bldP spid="${spid}" grpId="0" build="p" bldLvl="2"/></p:bldLst>`;
    const { slide } = await withTiming(
      timingRoot(mainSeq(presetEffect(3, spid, { target })), bldLst),
    );

    const steps = getSlideAnimations(slide);
    expect(steps).toHaveLength(1);
    expect(steps[0]!.target).toEqual({
      kind: 'paragraphs',
      shapeId: spid,
      firstParagraph: 1,
      lastParagraph: 2,
    });
    expect(steps[0]!.buildByParagraph).toBe(true);
    expect(steps[0]!.buildLevel).toBe(2);
    expect(steps[0]!.editable).toBe(true);
  });

  it('lists a preset it cannot name, read-only, without disabling its neighbour', async () => {
    const spid = await firstShapeId();
    // presetID 2 / entr is PowerPoint's "fly in" — a real preset, not one of
    // the four this library authors.
    const { slide } = await withTiming(
      timingRoot(mainSeq(presetEffect(3, spid, { presetId: '2' }) + presetEffect(10, spid))),
    );
    const steps = getSlideAnimations(slide);
    expect(steps[0]!.presetId).toBe(2);
    expect(steps[0]!.presetClass).toBe('entr');
    expect(steps[0]!.effect).toBeNull();
    expect(steps[0]!.editable).toBe(false);
    expect(steps[1]!.editable).toBe(true);
  });

  it('keeps the shape-level readers working on a composite effect', async () => {
    const deck = createPresentation();
    const authored = addBlankSlide(deck);
    const first = addSlideShape(authored, { preset: 'rect', ...box });
    const second = addSlideShape(authored, { preset: 'ellipse', ...box });
    const [idA, idB] = [getShapeId(first), getShapeId(second)];
    const bytes = await savePresentation(deck);

    const composite =
      `<p:par><p:cTn id="3" presetID="10" presetClass="entr" presetSubtype="0" fill="hold" ` +
      `grpId="0" nodeType="clickEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst>` +
      `<p:childTnLst>` +
      `<p:set><p:cBhvr><p:cTn id="4" dur="1" fill="hold"/>` +
      `<p:tgtEl><p:spTgt spid="${idA}"/></p:tgtEl>` +
      `<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr>` +
      `<p:to><p:strVal val="visible"/></p:to></p:set>` +
      `<p:anim calcmode="lin" valueType="num"><p:cBhvr><p:cTn id="5" dur="500"/>` +
      `<p:tgtEl><p:spTgt spid="${idB}"/></p:tgtEl>` +
      `<p:attrNameLst><p:attrName>style.opacity</p:attrName></p:attrNameLst></p:cBhvr>` +
      `<p:tavLst/></p:anim></p:childTnLst></p:cTn></p:par>`;
    const bldLst = `<p:bldLst><p:bldP spid="${idA}" grpId="0"/><p:bldP spid="${idB}" grpId="0"/></p:bldLst>`;
    const { slide } = await withTiming(timingRoot(mainSeq(composite), bldLst), bytes);

    const steps = getSlideAnimations(slide);
    expect(steps).toHaveLength(1);
    // The step refuses to name one target, but still says which shapes move,
    // so the shape-level readers keep reporting both of them.
    expect(steps[0]!.target).toEqual({ kind: 'unsupported', shapeId: null });
    expect(steps[0]!.targetShapeIds).toEqual([idA, idB]);
    expect(steps[0]!.editable).toBe(false);

    const [shapeA, shapeB] = getSlideShapes(slide);
    expect(getShapeAnimation(shapeA!)).toBe('fadeIn');
    expect(getShapeAnimation(shapeB!)).toBe('fadeIn');
    expect(findShapesWithAnimation(slide).map(getShapeId)).toEqual([idA, idB]);
  });

  // The two shadowing cases, side by side: only a node that *names* a preset
  // was ever a candidate for the shape-level readers.
  it('lets a recognised effect through when an unnamed effect precedes it', async () => {
    const spid = await firstShapeId();
    const bldLst = `<p:bldLst><p:bldP spid="${spid}" grpId="0"/></p:bldLst>`;
    const { slide } = await withTiming(
      timingRoot(mainSeq(customEffect(3, spid) + presetEffect(10, spid)), bldLst),
    );
    expect(getSlideAnimations(slide).map((s) => s.effect)).toEqual([null, 'fadeIn']);
    expect(getShapeAnimation(getSlideShapes(slide)[0]!)).toBe('fadeIn');
    expect(findShapesWithAnimation(slide)).toHaveLength(1);
  });

  it('stops at a named preset it cannot map, even when a known effect follows', async () => {
    const spid = await firstShapeId();
    const bldLst = `<p:bldLst><p:bldP spid="${spid}" grpId="0"/></p:bldLst>`;
    const { slide } = await withTiming(
      timingRoot(
        mainSeq(presetEffect(3, spid, { presetId: '2' }) + presetEffect(10, spid)),
        bldLst,
      ),
    );
    expect(getSlideAnimations(slide).map((s) => s.effect)).toEqual([null, 'fadeIn']);
    expect(getShapeAnimation(getSlideShapes(slide)[0]!)).toBeNull();
    expect(findShapesWithAnimation(slide)).toEqual([]);
  });

  it('reports no animation for a shape the build list does not mention', async () => {
    const spid = await firstShapeId();
    const { slide } = await withTiming(timingRoot(mainSeq(presetEffect(3, spid))));
    expect(getSlideAnimations(slide)).toHaveLength(1);
    expect(getShapeAnimation(getSlideShapes(slide)[0]!)).toBeNull();
    expect(findShapesWithAnimation(slide)).toEqual([]);
  });

  // A shape-triggered sequence looks exactly like a main-sequence click effect
  // from `start` alone. A player that went by `start` would fire it on the
  // wrong click, so the sequence it lives in has to travel with the step.
  it('separates an interactive sequence effect from the slide click order', async () => {
    const spid = await firstShapeId();
    const interactive =
      `<p:seq concurrent="1" nextAc="seek"><p:cTn id="30" restart="whenNotActive" ` +
      `fill="hold" evtFilter="cancelBubble" nodeType="interactive">` +
      `<p:childTnLst>${presetEffect(31, spid)}</p:childTnLst></p:cTn></p:seq>`;
    const { slide } = await withTiming(timingRoot(mainSeq(presetEffect(3, spid)) + interactive));

    const steps = getSlideAnimations(slide);
    expect(steps).toHaveLength(2);
    expect(steps.map((s) => s.sequence)).toEqual(['mainSeq', 'interactive']);
    // Both read as a click effect; only the main-sequence one may be played.
    expect(steps.map((s) => s.start)).toEqual(['click', 'click']);
    expect(steps.map((s) => s.playable)).toEqual([true, false]);
    expect(steps.map((s) => s.editable)).toEqual([true, false]);
    expect(slideHasAnimations(slide)).toBe(true);
  });

  it('marks a step playable but not editable when its handle is unusable', async () => {
    const spid = await firstShapeId();
    const { slide } = await withTiming(
      timingRoot(mainSeq(presetEffect(3, spid).replace('id="3"', 'id="3abc"'))),
    );
    const steps = getSlideAnimations(slide);
    expect(steps[0]!.sequence).toBe('mainSeq');
    expect(steps[0]!.id).toBeNull();
    expect(steps[0]!.playable).toBe(true);
    expect(steps[0]!.editable).toBe(false);
  });
});
