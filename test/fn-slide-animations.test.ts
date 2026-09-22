// getSlideAnimations — the ordered read of a slide's <p:timing> tree.
//
// Several cases splice hand-written timing XML into a real slide part: the
// authoring API cannot produce a composite effect, a duplicated `<p:cTn id>`
// or a paragraph build, and those are exactly the trees a template drops on us.

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
  addSlideMedia,
  addSlideShape,
  addSlideTextBox,
  createPresentation,
  findShapesWithAnimation,
  getShapeAnimation,
  getShapeId,
  getShapeParagraphCount,
  getSlideAnimations,
  getSlideShapes,
  getSlides,
  inches,
  loadPresentation,
  savePresentation,
  setShapeAnimation,
  setShapeBullets,
  slideHasAnimations,
  updateSlideAnimation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const SLIDE1 = '/ppt/slides/slide1.xml';
const decoder = new TextDecoder();
const encoder = new TextEncoder();

/**
 * Replaces slide 1's `<p:timing>` with `timing` and reloads the deck. The
 * spliced slide is schema-checked first: a fixture PowerPoint could never
 * hand us would test the parser against nothing real. (`nodeType="interactive"`
 * looked right and is not in ST_TLTimeNodeType — this is what caught it.)
 * Pass `malformed` for the cases whose whole point is a tree that breaks the
 * schema, and say in the test why.
 */
const withTiming = async (
  timing: string,
  opts: { source?: Uint8Array; malformed?: boolean } = {},
): Promise<{ pres: PresentationData; slide: SlideData }> => {
  const pres = await loadPresentation(
    opts.source ?? (await readFile(fixture('one-text-slide.pptx'))),
  );
  const part = _internalPackageOf(pres).getPart(partName(SLIDE1));
  if (part === null) throw new Error('missing slide part');
  const xml = decoder.decode(part.data).replace(/<p:timing>[\s\S]*<\/p:timing>/, '');
  const spliced = xml.replace('</p:sld>', `${timing}</p:sld>`);
  if (isSchemaValidationAvailable()) {
    // A fixture flagged malformed has to actually break the schema, or the
    // flag is hiding a mistake in the fixture instead of exercising tolerance.
    if (opts.malformed) {
      expect(() => expectSchemaValid(spliced, 'pml')).toThrow(/schema validation failed/);
    } else {
      expectSchemaValid(spliced, 'pml');
    }
  }
  part.data = encoder.encode(spliced);
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

// CT_TimeNodeList needs at least one child, so a main sequence with nothing
// in it omits `<p:childTnLst>` rather than writing an empty one.
const mainSeq = (childTnLst: string): string =>
  `<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq">` +
  (childTnLst === '' ? '' : `<p:childTnLst>${childTnLst}</p:childTnLst>`) +
  `</p:cTn></p:seq>`;

const box = { x: inches(1), y: inches(1), w: inches(2), h: inches(1) };
const skipIfNoXmllint = isSchemaValidationAvailable() ? it : it.skip;

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

  skipIfNoXmllint('emits a schema-valid timing tree for several stacked effects', () => {
    const pres = createPresentation();
    const slide = addBlankSlide(pres);
    const a = addSlideShape(slide, { preset: 'rect', ...box });
    const b = addSlideShape(slide, { preset: 'ellipse', ...box });
    setShapeAnimation(a, { effect: 'fadeIn', durationMs: 750 });
    setShapeAnimation(b, { effect: 'disappear' });
    setShapeAnimation(a, { effect: 'fadeOut' });
    const part = _internalPackageOf(pres).getPart(partName(SLIDE1));
    expectSchemaValid(decoder.decode(part!.data), 'pml');
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
      // An empty presetClass is outside ST_TLPresetClassType: the tree is
      // broken on purpose, and the parser must survive it.
      { malformed: true },
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

  // A preset effect written the way PowerPoint writes one: a `<p:set>` that
  // puts the shape on the slide, then the fade itself. `fill` says what becomes
  // of each node's value once it has run, and `null` leaves the attribute out —
  // which CT_TLCommonTimeNodeData allows and gives no default for.
  const heldEffect = (
    id: number,
    spid: number,
    opts: {
      effect?: string | null;
      visibility?: string | null;
      fade?: string | null;
      /** False for an effect that only fades, with no visibility set at all. */
      flipsVisibility?: boolean;
      nodeType?: string;
      /** A paragraph range, for a build rather than a whole-shape effect. */
      paragraphs?: [number, number];
      delayMs?: number;
      durationMs?: number;
    } = {},
  ): string => {
    const fill = (value: string | null | undefined, fallback: string | null): string => {
      const use = value === undefined ? fallback : value;
      return use === null ? '' : ` fill="${use}"`;
    };
    const target =
      opts.paragraphs === undefined
        ? `<p:spTgt spid="${spid}"/>`
        : `<p:spTgt spid="${spid}"><p:txEl><p:pRg st="${opts.paragraphs[0]}" ` +
          `end="${opts.paragraphs[1]}"/></p:txEl></p:spTgt>`;
    const set =
      opts.flipsVisibility === false
        ? ''
        : `<p:set><p:cBhvr><p:cTn id="${id + 1}" dur="1"${fill(opts.visibility, 'hold')}/>` +
          `<p:tgtEl>${target}</p:tgtEl>` +
          `<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr>` +
          `<p:to><p:strVal val="visible"/></p:to></p:set>`;
    return (
      `<p:par><p:cTn id="${id}" presetID="10" presetClass="entr" presetSubtype="0"` +
      `${fill(opts.effect, 'hold')} grpId="0" nodeType="${opts.nodeType ?? 'clickEffect'}">` +
      `<p:stCondLst><p:cond delay="${opts.delayMs ?? 0}"/></p:stCondLst><p:childTnLst>${set}` +
      `<p:anim calcmode="lin" valueType="num"><p:cBhvr additive="base">` +
      `<p:cTn id="${id + 2}" dur="${opts.durationMs ?? 500}"${fill(opts.fade, null)}/>` +
      `<p:tgtEl>${target}</p:tgtEl>` +
      `<p:attrNameLst><p:attrName>style.opacity</p:attrName></p:attrNameLst></p:cBhvr>` +
      `<p:tavLst/></p:anim></p:childTnLst></p:cTn></p:par>`
    );
  };

  it('plays a preset whose visibility is held, whatever the fade says', async () => {
    const spid = await firstShapeId();
    // The fade carries no fill of its own. It ends on the opacity the
    // visibility set already implies, so nothing about where the shape ends up
    // is left unsaid — the ordinary shape of a PowerPoint entrance must not be
    // read as something this library cannot play.
    const { slide } = await withTiming(timingRoot(mainSeq(heldEffect(3, spid))));
    const steps = getSlideAnimations(slide);
    expect(steps).toHaveLength(1);
    expect(steps[0]!.effect).toBe('fadeIn');
    expect(steps[0]!.playable).toBe(true);
    expect(steps[0]!.editable).toBe(true);
  });

  it('keeps an effect the tree takes away again, as read-only', async () => {
    const spid = await firstShapeId();
    // `fill="remove"` on the effect node drops everything under it when the
    // effect ends, so the shape does not stay where the entrance put it. What
    // it does instead is not something this library models.
    const { slide } = await withTiming(
      timingRoot(mainSeq(heldEffect(3, spid, { effect: 'remove' }) + heldEffect(10, spid))),
    );
    const steps = getSlideAnimations(slide);
    expect(steps).toHaveLength(2);
    expect(steps[0]!.effect).toBe('fadeIn');
    expect(steps[0]!.playable).toBe(false);
    expect(steps[0]!.editable).toBe(false);
    expect(steps[1]!.playable).toBe(true);
  });

  it('keeps an effect whose visibility is taken away again, as read-only', async () => {
    const spid = await firstShapeId();
    // The effect node holds, but the behaviour that puts the shape on the
    // slide does not: reading only the node above it would play this as a
    // plain entrance and show what the deck ends up hiding.
    const { slide } = await withTiming(
      timingRoot(mainSeq(heldEffect(3, spid, { visibility: 'remove' }))),
    );
    const steps = getSlideAnimations(slide);
    expect(steps[0]!.effect).toBe('fadeIn');
    expect(steps[0]!.playable).toBe(false);
    expect(steps[0]!.editable).toBe(false);
  });

  it('asks the fade about its fill when nothing else says where the shape is', async () => {
    const spid = await firstShapeId();
    // No visibility set: the opacity the fade lands on is all there is to say
    // whether the shape can be seen, and `remove` puts it back where it began.
    // An effect that does flip visibility answers that question with the set,
    // so its fade is asked about something else — see the overlap case below.
    const { slide } = await withTiming(
      timingRoot(mainSeq(heldEffect(3, spid, { flipsVisibility: false, fade: 'remove' }))),
    );
    expect(getSlideAnimations(slide)[0]!.effect).toBe('fadeIn');
    expect(getSlideAnimations(slide)[0]!.playable).toBe(false);

    const held = await withTiming(
      timingRoot(mainSeq(heldEffect(3, spid, { flipsVisibility: false, fade: 'hold' }))),
    );
    expect(getSlideAnimations(held.slide)[0]!.playable).toBe(true);
  });

  it('reports what the tree says becomes of an effect\u2019s value when it ends', async () => {
    const spid = await firstShapeId();
    // `fill` on the behaviour that animates something. What turns on it is how
    // two effects over one object settle when the first of them ends, which is
    // the player's question, not this one's \u2014 so it is reported, not judged:
    // every one of these still plays, alone.
    const cases: [string | null, string][] = [
      ['hold', 'held'],
      ['freeze', 'held'],
      ['remove', 'removed'],
      ['transition', 'removed'],
      [null, 'unstated'],
    ];
    for (const [fade, expected] of cases) {
      const { slide } = await withTiming(timingRoot(mainSeq(heldEffect(3, spid, { fade }))));
      const step = getSlideAnimations(slide)[0]!;
      expect([fade, step.valueAfterEnd]).toEqual([fade, expected]);
      expect(step.playable).toBe(true);
    }

    // An instantaneous preset animates nothing, so there is nothing to take
    // away: its `<p:set>` is asked by `playable` instead.
    const instant = await withTiming(
      timingRoot(mainSeq(heldEffect(3, spid, { flipsVisibility: true, fade: null }))),
    );
    expect(getSlideAnimations(instant.slide)[0]!.valueAfterEnd).toBe('unstated');
  });

  it('keeps an effect that states no fill at all, as read-only', async () => {
    const spid = await firstShapeId();
    const { slide } = await withTiming(timingRoot(mainSeq(heldEffect(3, spid, { effect: null }))));
    expect(getSlideAnimations(slide)[0]!.playable).toBe(false);
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
      // `id="20abc"` is not an ST_TLTimeNodeID; that is the case under test.
      { malformed: true },
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
    const { slide } = await withTiming(timingRoot(mainSeq(composite), bldLst), { source: bytes });

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
      `fill="hold" evtFilter="cancelBubble" nodeType="interactiveSeq">` +
      `<p:childTnLst>${presetEffect(31, spid)}</p:childTnLst></p:cTn></p:seq>`;
    const { slide } = await withTiming(timingRoot(mainSeq(presetEffect(3, spid)) + interactive));

    const steps = getSlideAnimations(slide);
    expect(steps).toHaveLength(2);
    expect(steps.map((s) => s.sequence)).toEqual(['mainSeq', 'interactiveSeq']);
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
      { malformed: true },
    );
    const steps = getSlideAnimations(slide);
    expect(steps[0]!.sequence).toBe('mainSeq');
    expect(steps[0]!.id).toBeNull();
    expect(steps[0]!.playable).toBe(true);
    expect(steps[0]!.editable).toBe(false);
  });
});

describe('fn API: setShapeAnimation — start conditions and delay', () => {
  const deck = (): { pres: PresentationData; slide: SlideData } => {
    const pres = createPresentation();
    return { pres, slide: addBlankSlide(pres) };
  };
  const slideXml = (pres: PresentationData): string =>
    decoder.decode(_internalPackageOf(pres).getPart(partName(SLIDE1))!.data);

  it('defaults to a click effect with no delay, unchanged from before', () => {
    const { pres, slide } = deck();
    const shape = addSlideShape(slide, { preset: 'rect', ...box });
    setShapeAnimation(shape, { effect: 'fadeIn' });
    const explicit = deck();
    const other = addSlideShape(explicit.slide, { preset: 'rect', ...box });
    setShapeAnimation(other, { effect: 'fadeIn', start: 'click', delayMs: 0 });
    expect(slideXml(explicit.pres)).toBe(slideXml(pres));
  });

  it('reads back each start condition it wrote', () => {
    const { slide } = deck();
    const a = addSlideShape(slide, { preset: 'rect', ...box });
    const b = addSlideShape(slide, { preset: 'ellipse', ...box });
    const c = addSlideShape(slide, { preset: 'triangle', ...box });
    setShapeAnimation(a, { effect: 'fadeIn' });
    setShapeAnimation(b, { effect: 'fadeIn', start: 'withPrevious' });
    setShapeAnimation(c, { effect: 'fadeIn', start: 'afterPrevious', delayMs: 250 });

    const steps = getSlideAnimations(slide);
    expect(steps.map((s) => s.start)).toEqual(['click', 'withPrevious', 'afterPrevious']);
    expect(steps.map((s) => s.delayMs)).toEqual([0, 0, 250]);
    expect(steps.every((s) => s.playable)).toBe(true);
    expect(steps.map((s) => s.target.shapeId)).toEqual([a, b, c].map(getShapeId));
    expect(new Set(steps.map((s) => s.id)).size).toBe(3);
  });

  // The three effects belong to one click: a viewer clicks once and all of
  // them run, the second alongside the first and the third after it.
  it('folds with/after effects into the click stop already there', () => {
    const { pres, slide } = deck();
    const a = addSlideShape(slide, { preset: 'rect', ...box });
    const b = addSlideShape(slide, { preset: 'ellipse', ...box });
    const c = addSlideShape(slide, { preset: 'triangle', ...box });
    setShapeAnimation(a, { effect: 'fadeIn' });
    setShapeAnimation(b, { effect: 'fadeIn', start: 'withPrevious' });
    setShapeAnimation(c, { effect: 'fadeIn', start: 'afterPrevious' });
    const xml = slideXml(pres);
    expect(xml.match(/delay="indefinite"/g)).toHaveLength(1);
    expect(getSlideAnimations(slide)).toHaveLength(3);
  });

  // With nothing to follow, "with previous" and "after previous" mean "as the
  // slide appears" — the stop must not wait for a click that would otherwise
  // leave the effect stranded.
  it('starts a leading with/after effect as the slide appears', () => {
    for (const start of ['withPrevious', 'afterPrevious'] as const) {
      const { pres, slide } = deck();
      const shape = addSlideShape(slide, { preset: 'rect', ...box });
      setShapeAnimation(shape, { effect: 'fadeIn', start });
      expect(slideXml(pres)).not.toContain('delay="indefinite"');
      expect(getSlideAnimations(slide)[0]!.start).toBe(start);
    }
  });

  it('gives each click effect its own stop', () => {
    const { pres, slide } = deck();
    const a = addSlideShape(slide, { preset: 'rect', ...box });
    const b = addSlideShape(slide, { preset: 'ellipse', ...box });
    setShapeAnimation(a, { effect: 'fadeIn' });
    setShapeAnimation(b, { effect: 'fadeOut' });
    expect(slideXml(pres).match(/delay="indefinite"/g)).toHaveLength(2);
  });

  skipIfNoXmllint('emits a schema-valid tree for every start condition', () => {
    const { pres, slide } = deck();
    const a = addSlideShape(slide, { preset: 'rect', ...box });
    const b = addSlideShape(slide, { preset: 'ellipse', ...box });
    const c = addSlideShape(slide, { preset: 'triangle', ...box });
    setShapeAnimation(a, { effect: 'fadeIn' });
    setShapeAnimation(b, { effect: 'appear', start: 'withPrevious' });
    setShapeAnimation(c, { effect: 'fadeOut', start: 'afterPrevious', delayMs: 750 });
    expectSchemaValid(slideXml(pres), 'pml');
  });

  it('keeps start, delay and order through save and reload', async () => {
    const { pres, slide } = deck();
    const a = addSlideShape(slide, { preset: 'rect', ...box });
    const b = addSlideShape(slide, { preset: 'ellipse', ...box });
    setShapeAnimation(a, { effect: 'fadeIn', delayMs: 100 });
    setShapeAnimation(b, { effect: 'fadeOut', start: 'afterPrevious', delayMs: 300 });
    const before = getSlideAnimations(slide);

    const reloaded = await loadPresentation(await savePresentation(pres));
    expect(getSlideAnimations(getSlides(reloaded)[0]!)).toEqual(before);
  });

  it('rejects a start condition and a delay it cannot write', () => {
    const { slide } = deck();
    const shape = addSlideShape(slide, { preset: 'rect', ...box });
    // @ts-expect-error — the guard is for callers without type checking.
    expect(() => setShapeAnimation(shape, { effect: 'fadeIn', start: 'onHover' })).toThrow(
      /setShapeAnimation: start/,
    );
    expect(() => setShapeAnimation(shape, { effect: 'fadeIn', delayMs: -1 })).toThrow(
      /setShapeAnimation: delayMs/,
    );
    expect(getSlideAnimations(slide)).toEqual([]);
  });

  // Every bounded integer in this library rounds rather than refusing; the
  // delay is not the place to break with that.
  it('rounds a fractional delay to whole milliseconds', () => {
    const { slide } = deck();
    const shape = addSlideShape(slide, { preset: 'rect', ...box });
    setShapeAnimation(shape, { effect: 'fadeIn', delayMs: 1.5 });
    expect(getSlideAnimations(slide)[0]!.delayMs).toBe(2);
  });
});

// A `<p:cBhvr>` on `spid`, with `attrs` going on its `<p:cTn>`.
const behaviour = (id: number, spid: number | string, attrs: string): string =>
  `<p:anim calcmode="lin" valueType="num"><p:cBhvr additive="base">` +
  `<p:cTn id="${id}"${attrs}/><p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl>` +
  `<p:attrNameLst><p:attrName>style.opacity</p:attrName></p:attrNameLst></p:cBhvr>` +
  `<p:tavLst/></p:anim>`;

/** A click stop holding one preset effect, with `body` as its behaviours. */
const stopWith = (body: string, effectStCondLst = '<p:cond delay="0"/>'): string =>
  `<p:par><p:cTn id="3" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst>` +
  `<p:childTnLst><p:par><p:cTn id="4" fill="hold">` +
  `<p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>` +
  `<p:par><p:cTn id="5" presetID="10" presetClass="entr" presetSubtype="0" fill="hold" ` +
  `grpId="0" nodeType="clickEffect"><p:stCondLst>${effectStCondLst}</p:stCondLst>` +
  `<p:childTnLst>${body}</p:childTnLst></p:cTn></p:par>` +
  `</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>`;

/**
 * Start offsets of the wrapper time nodes — the click stops and the groups
 * inside them — in document order. Effect nodes carry `presetID` and are
 * dropped, so what is left is the timeline the wrappers impose.
 *
 * Reading these back out of the saved XML is the point: `start: 'afterPrevious'`
 * is only honoured if the group really is placed after the one before it, and
 * the `nodeType` attribute alone does not place it there.
 */
const wrapperDelays = (xml: string): string[] =>
  [...xml.matchAll(/<p:par><p:cTn\b([^>]*)><p:stCondLst><p:cond delay="([^"]*)"\/>/g)]
    .filter((m) => !m[1]!.includes('presetID'))
    .map((m) => m[2]!);

describe('fn API: setShapeAnimation — when an "after previous" effect actually starts', () => {
  const deck = (): { pres: PresentationData; slide: SlideData } => {
    const pres = createPresentation();
    return { pres, slide: addBlankSlide(pres) };
  };
  const slideXml = (pres: PresentationData): string =>
    decoder.decode(_internalPackageOf(pres).getPart(partName(SLIDE1))!.data);
  const shape = (slide: SlideData): SlideShapeData =>
    addSlideShape(slide, { preset: 'rect', ...box });

  // Sibling `<p:par>` nodes are parallel: they all begin when their stop does.
  // A group left at delay 0 would run *with* the effect it is meant to follow,
  // whatever its `nodeType` says, so the offset is the whole dependency.
  it('offsets the group to where the longest effect before it finishes', () => {
    const { pres, slide } = deck();
    setShapeAnimation(shape(slide), { effect: 'fadeIn', durationMs: 800 });
    setShapeAnimation(shape(slide), { effect: 'fadeIn', durationMs: 400, start: 'withPrevious' });
    setShapeAnimation(shape(slide), { effect: 'fadeIn', durationMs: 300, start: 'afterPrevious' });
    // The stop, its first group, and the group waiting on the 800ms effect.
    expect(wrapperDelays(slideXml(pres))).toEqual(['indefinite', '0', '800']);
    expect(getSlideAnimations(slide).map((s) => s.start)).toEqual([
      'click',
      'withPrevious',
      'afterPrevious',
    ]);
  });

  // A delayed "with previous" pushes the end of its whole group out, so what
  // follows the group has to wait that much longer.
  it('counts a delayed with-previous effect when measuring the group', () => {
    const { pres, slide } = deck();
    setShapeAnimation(shape(slide), { effect: 'fadeIn', durationMs: 500 });
    setShapeAnimation(shape(slide), {
      effect: 'fadeIn',
      durationMs: 200,
      start: 'withPrevious',
      delayMs: 1000,
    });
    setShapeAnimation(shape(slide), { effect: 'fadeIn', start: 'afterPrevious' });
    expect(wrapperDelays(slideXml(pres))).toEqual(['indefinite', '0', '1200']);
  });

  // Two "with previous" effects run together, each measured from the click
  // that started their group — not from one another.
  it('keeps two with-previous effects on the same trigger', () => {
    const { pres, slide } = deck();
    setShapeAnimation(shape(slide), { effect: 'fadeIn', durationMs: 500 });
    setShapeAnimation(shape(slide), {
      effect: 'fadeIn',
      durationMs: 500,
      start: 'withPrevious',
      delayMs: 200,
    });
    setShapeAnimation(shape(slide), {
      effect: 'fadeIn',
      durationMs: 500,
      start: 'withPrevious',
      delayMs: 400,
    });
    // One stop, one group: no wrapper waits on anything.
    expect(wrapperDelays(slideXml(pres))).toEqual(['indefinite', '0']);
    expect(getSlideAnimations(slide).map((s) => s.delayMs)).toEqual([0, 200, 400]);
  });

  // `appear` writes only the 1ms visibility kick, so that is its whole length.
  it('measures an instant effect by its visibility kick', () => {
    const { pres, slide } = deck();
    setShapeAnimation(shape(slide), { effect: 'appear' });
    setShapeAnimation(shape(slide), { effect: 'fadeIn', start: 'afterPrevious' });
    expect(wrapperDelays(slideXml(pres))).toEqual(['indefinite', '0', '1']);
  });

  it('adds the caller delay on top of the wait, without double-counting it', () => {
    const { pres, slide } = deck();
    setShapeAnimation(shape(slide), { effect: 'fadeIn', durationMs: 600 });
    setShapeAnimation(shape(slide), { effect: 'fadeIn', start: 'afterPrevious', delayMs: 150 });
    expect(wrapperDelays(slideXml(pres))).toEqual(['indefinite', '0', '600']);
    // The caller's 150ms stays on the effect, counted from the group's start.
    expect(getSlideAnimations(slide).map((s) => s.delayMs)).toEqual([0, 150]);
  });

  it('chains a second after-previous effect past the first', () => {
    const { pres, slide } = deck();
    setShapeAnimation(shape(slide), { effect: 'fadeIn', durationMs: 400 });
    setShapeAnimation(shape(slide), { effect: 'fadeIn', durationMs: 300, start: 'afterPrevious' });
    setShapeAnimation(shape(slide), { effect: 'fadeIn', durationMs: 100, start: 'afterPrevious' });
    expect(wrapperDelays(slideXml(pres))).toEqual(['indefinite', '0', '400', '700']);
  });

  skipIfNoXmllint('emits a schema-valid tree for a chained sequence', () => {
    const { pres, slide } = deck();
    setShapeAnimation(shape(slide), { effect: 'fadeIn', durationMs: 800 });
    setShapeAnimation(shape(slide), { effect: 'appear', start: 'withPrevious', delayMs: 100 });
    setShapeAnimation(shape(slide), { effect: 'fadeOut', start: 'afterPrevious', delayMs: 50 });
    expectSchemaValid(slideXml(pres), 'pml');
  });

  it('keeps the computed offsets through save and reload', async () => {
    const { pres, slide } = deck();
    setShapeAnimation(shape(slide), { effect: 'fadeIn', durationMs: 800 });
    setShapeAnimation(shape(slide), { effect: 'fadeIn', durationMs: 300, start: 'afterPrevious' });
    const before = slideXml(pres);
    const steps = getSlideAnimations(slide);

    const reloaded = await loadPresentation(await savePresentation(pres));
    expect(slideXml(reloaded)).toBe(before);
    expect(getSlideAnimations(getSlides(reloaded)[0]!)).toEqual(steps);
  });
});

describe('fn API: setShapeAnimation — building text one paragraph at a time', () => {
  const deck = (
    lines: readonly string[],
  ): { pres: PresentationData; slide: SlideData; shape: SlideShapeData } => {
    const pres = createPresentation();
    const slide = addBlankSlide(pres);
    const shape = addSlideTextBox(slide, { ...box, text: lines.join('\n') });
    return { pres, slide, shape };
  };
  const slideXml = (pres: PresentationData): string =>
    decoder.decode(_internalPackageOf(pres).getPart(partName(SLIDE1))!.data);

  it('writes one effect per paragraph, each over its own paragraph range', () => {
    const { slide, shape } = deck(['One', 'Two', 'Three']);
    setShapeAnimation(shape, { effect: 'fadeIn', byParagraph: true });

    const steps = getSlideAnimations(slide);
    expect(steps).toHaveLength(3);
    expect(steps.map((s) => s.target)).toEqual([
      { kind: 'paragraphs', shapeId: getShapeId(shape), firstParagraph: 0, lastParagraph: 0 },
      { kind: 'paragraphs', shapeId: getShapeId(shape), firstParagraph: 1, lastParagraph: 1 },
      { kind: 'paragraphs', shapeId: getShapeId(shape), firstParagraph: 2, lastParagraph: 2 },
    ]);
    expect(steps.every((s) => s.buildByParagraph)).toBe(true);
    expect(steps.every((s) => s.playable)).toBe(true);
    expect(new Set(steps.map((s) => s.id)).size).toBe(3);
  });

  // PowerPoint ties an effect to its build through grpId, and the build is
  // what says "paragraph by paragraph" — one entry for the whole body, not one
  // per paragraph, or the later ones would each re-declare the same build.
  it('gives the whole build a single `<p:bldP build="p">`', () => {
    const { pres, shape } = deck(['One', 'Two', 'Three']);
    setShapeAnimation(shape, { effect: 'fadeIn', byParagraph: true });
    const xml = slideXml(pres);
    expect(xml.match(/<p:bldP\b/g)).toHaveLength(1);
    expect(xml).toContain(`<p:bldP spid="${getShapeId(shape)}" grpId="0" build="p"/>`);
  });

  it('advances a paragraph per click by default', () => {
    const { pres, slide, shape } = deck(['One', 'Two', 'Three']);
    setShapeAnimation(shape, { effect: 'fadeIn', byParagraph: true });
    expect(wrapperDelays(slideXml(pres))).toEqual([
      'indefinite',
      '0',
      'indefinite',
      '0',
      'indefinite',
      '0',
    ]);
    expect(getSlideAnimations(slide).map((s) => s.start)).toEqual(['click', 'click', 'click']);
  });

  it('cascades the paragraphs off one click when they follow each other', () => {
    const { pres, slide, shape } = deck(['One', 'Two', 'Three']);
    setShapeAnimation(shape, {
      effect: 'fadeIn',
      durationMs: 400,
      byParagraph: true,
      start: 'afterPrevious',
    });
    // One stop; each paragraph's group waits out the 400ms before it.
    expect(wrapperDelays(slideXml(pres))).toEqual(['0', '0', '400', '800']);
    expect(getSlideAnimations(slide).map((s) => s.start)).toEqual([
      'afterPrevious',
      'afterPrevious',
      'afterPrevious',
    ]);
  });

  it('runs every paragraph together when they share the previous start', () => {
    const { pres, shape } = deck(['One', 'Two', 'Three']);
    setShapeAnimation(shape, { effect: 'fadeIn', byParagraph: true, start: 'withPrevious' });
    expect(wrapperDelays(slideXml(pres))).toEqual(['0', '0']);
  });

  it('builds a Japanese bulleted body the same way', () => {
    const { pres, slide, shape } = deck(['一つ目', '二つ目', '三つ目']);
    setShapeBullets(shape, 'bullet');
    setShapeAnimation(shape, { effect: 'fadeIn', byParagraph: true });

    const steps = getSlideAnimations(slide);
    expect(steps.map((s) => s.target.kind)).toEqual(['paragraphs', 'paragraphs', 'paragraphs']);
    expect(slideXml(pres).match(/<p:bldP\b/g)).toHaveLength(1);
  });

  skipIfNoXmllint('emits a schema-valid paragraph build', () => {
    const { pres, shape } = deck(['One', 'Two', 'Three']);
    setShapeAnimation(shape, { effect: 'fadeIn', byParagraph: true });
    expectSchemaValid(slideXml(pres), 'pml');
  });

  it('keeps the build through save and reload', async () => {
    const { pres, slide, shape } = deck(['One', 'Two']);
    setShapeAnimation(shape, { effect: 'fadeIn', byParagraph: true });
    const before = slideXml(pres);
    const steps = getSlideAnimations(slide);

    const reloaded = await loadPresentation(await savePresentation(pres));
    expect(slideXml(reloaded)).toBe(before);
    expect(getSlideAnimations(getSlides(reloaded)[0]!)).toEqual(steps);
  });

  // The build is one group; a shape animated as a whole is another. Sharing a
  // grpId would make PowerPoint reveal the text with the other shape's effect.
  it('keeps a build and a whole-shape effect in separate groups', () => {
    const { pres, slide, shape } = deck(['One', 'Two']);
    const other = addSlideShape(slide, { preset: 'rect', ...box });
    setShapeAnimation(shape, { effect: 'fadeIn', byParagraph: true });
    setShapeAnimation(other, { effect: 'fadeIn' });

    const xml = slideXml(pres);
    expect(xml.match(/<p:bldP\b/g)).toHaveLength(2);
    expect(xml).toContain(`<p:bldP spid="${getShapeId(shape)}" grpId="0" build="p"/>`);
    expect(xml).toContain(`<p:bldP spid="${getShapeId(other)}" grpId="1"/>`);
    expect(getShapeAnimation(other)).toBe('fadeIn');
  });

  it('refuses to build a shape that has no text', () => {
    const pres = createPresentation();
    const slide = addBlankSlide(pres);
    const shape = addSlideShape(slide, { preset: 'rect', ...box });
    expect(() => setShapeAnimation(shape, { effect: 'fadeIn', byParagraph: true })).toThrow(
      /byParagraph needs a shape with text/,
    );
    expect(getSlideAnimations(slide)).toEqual([]);
    expect(slideHasAnimations(slide)).toBe(false);
  });

  it('appends a build behind the effects a slide already has', () => {
    const { pres, slide, shape } = deck(['One', 'Two', 'Three']);
    const first = addSlideShape(slide, { preset: 'rect', ...box });
    setShapeAnimation(first, { effect: 'fadeIn' });
    setShapeAnimation(shape, { effect: 'fadeIn', byParagraph: true });

    const steps = getSlideAnimations(slide);
    expect(steps.map((s) => s.target.shapeId)).toEqual([
      getShapeId(first),
      getShapeId(shape),
      getShapeId(shape),
      getShapeId(shape),
    ]);
    expect(steps.map((s) => s.buildByParagraph)).toEqual([false, true, true, true]);
    expect(new Set(steps.map((s) => s.id)).size).toBe(4);
    expect(slideXml(pres).match(/<p:bldP\b/g)).toHaveLength(2);
  });

  skipIfNoXmllint('appends a build to timing this library did not author', async () => {
    const spid = await firstShapeId();
    const { pres, slide } = await withTiming(
      timingRoot(
        mainSeq(presetEffect(3, spid)),
        `<p:bldLst><p:bldP spid="${spid}" grpId="0"/></p:bldLst>`,
      ),
    );
    const target = getSlideShapes(slide)[0]!;
    setShapeAnimation(target, { effect: 'fadeIn', byParagraph: true });

    const steps = getSlideAnimations(slide);
    expect(steps).toHaveLength(1 + getShapeParagraphCount(target));
    expect(steps[0]!.target).toEqual({ kind: 'shape', shapeId: spid });
    expect(steps.slice(1).every((s) => s.target.kind === 'paragraphs')).toBe(true);
    expectSchemaValid(slideXml(pres), 'pml');
  });

  // A build is several effects, so a refusal part-way would be the one case
  // that could strand a slide half-animated. The whole call is written into a
  // copy, so what the slide keeps is either all of it or none.
  it('leaves nothing behind when a build is refused, and still takes the next call', async () => {
    const { pres, slide, shape } = deck(['One', 'Two', 'Three']);
    const before = slideXml(pres);

    expect(() =>
      setShapeAnimation(shape, { effect: 'fadeIn', byParagraph: true, durationMs: -1 }),
    ).toThrow(/durationMs/);
    expect(slideXml(pres)).toBe(before);
    expect(getSlideAnimations(slide)).toEqual([]);
    expect(slideHasAnimations(slide)).toBe(false);

    setShapeAnimation(shape, { effect: 'fadeIn', byParagraph: true });
    const reloaded = await loadPresentation(await savePresentation(pres));
    const steps = getSlideAnimations(getSlides(reloaded)[0]!);
    expect(steps).toHaveLength(3);
    expect(steps.map((s) => s.target.kind)).toEqual(['paragraphs', 'paragraphs', 'paragraphs']);
    expect(
      decoder
        .decode(_internalPackageOf(reloaded).getPart(partName(SLIDE1))!.data)
        .match(/<p:bldP\b/g),
    ).toHaveLength(1);
  });

  it('leaves an authored tree untouched when a build cannot follow it', async () => {
    const spid = await firstShapeId();
    const { pres, slide } = await withTiming(
      timingRoot(
        mainSeq(stopWith(behaviour(6, spid, ' dur="indefinite" fill="hold"'))),
        `<p:bldLst><p:bldP spid="${spid}" grpId="0"/></p:bldLst>`,
      ),
    );
    const before = slideXml(pres);
    const target = getSlideShapes(slide)[0]!;

    expect(() =>
      setShapeAnimation(target, { effect: 'fadeIn', byParagraph: true, start: 'afterPrevious' }),
    ).toThrow(/cannot start an effect after one whose length this library cannot measure/);
    expect(slideXml(pres)).toBe(before);
    expect(getSlideAnimations(slide)).toHaveLength(1);

    // The refusal is about the start condition, not the build: a click build
    // still lands, and lands whole.
    setShapeAnimation(target, { effect: 'fadeIn', byParagraph: true });
    expect(getSlideAnimations(slide)).toHaveLength(1 + getShapeParagraphCount(target));
  });
});

/**
 * An authored slide can time an effect in ways this library does not model.
 * Guessing an end for one of those would place the next effect at a moment
 * PowerPoint never plays it, so the call is refused and the tree left alone.
 */
describe('fn API: setShapeAnimation — timing an after-previous effect cannot measure', () => {
  const slideXml = (pres: PresentationData): string =>
    decoder.decode(_internalPackageOf(pres).getPart(partName(SLIDE1))!.data);

  const refuses = async (body: string, effectStCondLst?: string): Promise<void> => {
    const spid = await firstShapeId();
    const { pres, slide } = await withTiming(
      timingRoot(
        mainSeq(stopWith(body.replaceAll('{spid}', String(spid)), effectStCondLst)),
        `<p:bldLst><p:bldP spid="${spid}" grpId="0"/></p:bldLst>`,
      ),
    );
    const before = slideXml(pres);
    expect(() =>
      setShapeAnimation(getSlideShapes(slide)[0]!, { effect: 'fadeIn', start: 'afterPrevious' }),
    ).toThrow(/cannot start an effect after one whose length this library cannot measure/);
    // Refusing has to leave the authored timing exactly as it was.
    expect(slideXml(pres)).toBe(before);
    expect(getSlideAnimations(slide)).toHaveLength(1);
  };

  // The case that matters most: one readable behaviour beside an unreadable
  // one. Taking the readable maximum would report 500ms for an effect that
  // never ends.
  it('refuses when one behaviour is timed and another runs indefinitely', async () => {
    await refuses(
      behaviour(6, '{spid}', ' dur="500" fill="hold"') +
        behaviour(7, '{spid}', ' dur="indefinite" fill="hold"'),
    );
  });

  it('refuses when one behaviour states no duration at all', async () => {
    await refuses(
      behaviour(6, '{spid}', ' dur="500" fill="hold"') + behaviour(7, '{spid}', ' fill="hold"'),
    );
  });

  it('refuses when a behaviour repeats', async () => {
    await refuses(behaviour(6, '{spid}', ' dur="500" repeatCount="3000" fill="hold"'));
  });

  it('refuses when a behaviour is rescaled by speed or auto-reverse', async () => {
    await refuses(behaviour(6, '{spid}', ' dur="500" spd="50%" fill="hold"'));
    await refuses(behaviour(6, '{spid}', ' dur="500" autoRev="1" fill="hold"'));
  });

  // `accel` and `decel` are shares of `dur`, so they leave the total alone.
  it('still follows an effect that only eases in and out', async () => {
    const spid = await firstShapeId();
    const { pres, slide } = await withTiming(
      timingRoot(
        mainSeq(stopWith(behaviour(6, spid, ' dur="500" accel="20%" decel="20%" fill="hold"'))),
        `<p:bldLst><p:bldP spid="${spid}" grpId="0"/></p:bldLst>`,
      ),
    );
    setShapeAnimation(getSlideShapes(slide)[0]!, { effect: 'fadeIn', start: 'afterPrevious' });
    expect(wrapperDelays(slideXml(pres))).toEqual(['indefinite', '0', '500']);
  });

  // `evt` ties the start to another node's lifetime. The `delay` beside it is
  // counted from that event, not from the group, so it is not an offset.
  it('refuses when the effect starts from another node rather than an offset', async () => {
    await refuses(
      behaviour(6, '{spid}', ' dur="500" fill="hold"'),
      '<p:cond evt="onEnd" delay="0"><p:tn val="2"/></p:cond>',
    );
  });

  it('refuses when the effect has more than one start condition', async () => {
    await refuses(
      behaviour(6, '{spid}', ' dur="500" fill="hold"'),
      '<p:cond delay="0"/><p:cond delay="900"/>',
    );
  });

  // A nested `<p:par>` runs on its own clock: its behaviours start when it
  // does, not when the effect does.
  it('refuses when the effect nests a timeline of its own', async () => {
    await refuses(
      `<p:par><p:cTn id="6" fill="hold"><p:stCondLst><p:cond delay="200"/></p:stCondLst>` +
        `<p:childTnLst>${behaviour(7, '{spid}', ' dur="500" fill="hold"')}</p:childTnLst>` +
        `</p:cTn></p:par>`,
    );
  });

  // Refusal is about measuring, not about who authored the effect: two plain
  // behaviours are measurable whoever wrote them.
  it('follows an authored effect whose behaviours are all timed', async () => {
    const spid = await firstShapeId();
    const { pres, slide } = await withTiming(
      timingRoot(
        mainSeq(
          stopWith(
            behaviour(6, spid, ' dur="500" fill="hold"') +
              behaviour(7, spid, ' dur="900" fill="hold"'),
          ),
        ),
        `<p:bldLst><p:bldP spid="${spid}" grpId="0"/></p:bldLst>`,
      ),
    );
    setShapeAnimation(getSlideShapes(slide)[0]!, { effect: 'fadeIn', start: 'afterPrevious' });
    expect(wrapperDelays(slideXml(pres))).toEqual(['indefinite', '0', '900']);
  });
});

/** A click stop holding one effect, so a hand-written effect has somewhere to sit. */
const clickStop = (effect: string): string =>
  `<p:par><p:cTn id="20" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst>` +
  `<p:childTnLst><p:par><p:cTn id="21" fill="hold">` +
  `<p:stCondLst><p:cond delay="0"/></p:stCondLst>` +
  `<p:childTnLst>${effect}</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>`;

describe('fn API: an emphasis preset is read from its own behaviour', () => {
  // `presetClass="emph" presetID="8"` is PowerPoint's Spin whatever angle it
  // turns through: the amount is on `<p:animRot>`, in sixtieth-thousandths of a
  // degree, and negative for a counter-clockwise turn. Only a single full
  // clockwise turn is an effect this library names, so a tree that says
  // anything else has to come back as one it reads rather than plays.
  const spinEffect = (id: number, spid: number, rotation: string): string =>
    `<p:par><p:cTn id="${id}" presetID="8" presetClass="emph" presetSubtype="0" fill="hold" ` +
    `grpId="0" nodeType="clickEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst>` +
    `<p:childTnLst><p:animRot ${rotation}><p:cBhvr>` +
    `<p:cTn id="${id + 1}" dur="2000" fill="hold"/>` +
    `<p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl>` +
    `<p:attrNameLst><p:attrName>r</p:attrName></p:attrNameLst>` +
    `</p:cBhvr></p:animRot></p:childTnLst></p:cTn></p:par>`;

  it('names a full clockwise turn', async () => {
    const spid = await firstShapeId();
    const { slide } = await withTiming(
      timingRoot(mainSeq(clickStop(spinEffect(3, spid, 'by="21600000"')))),
    );
    const steps = getSlideAnimations(slide);
    expect(steps).toHaveLength(1);
    expect(steps[0]!.effect).toBe('spin');
    expect(steps[0]!.playable).toBe(true);
  });

  it('refuses to name a half turn, a counter-clockwise one, or one with fixed ends', async () => {
    const spid = await firstShapeId();
    for (const rotation of ['by="10800000"', 'by="-21600000"', 'from="0" to="21600000"']) {
      const { slide } = await withTiming(
        timingRoot(mainSeq(clickStop(spinEffect(3, spid, rotation)))),
      );
      const steps = getSlideAnimations(slide);
      expect(steps, rotation).toHaveLength(1);
      // Still listed, with its place in the click order, but not named and not
      // ours to rewrite: turning it a full circle would be a different slide.
      expect(steps[0]!.effect, rotation).toBeNull();
      expect(steps[0]!.presetId, rotation).toBe(8);
      expect(steps[0]!.playable, rotation).toBe(false);
      expect(steps[0]!.editable, rotation).toBe(false);
    }
  });

  it('refuses to rewrite a rotation it did not name, and leaves it as it was', async () => {
    const spid = await firstShapeId();
    const { pres, slide } = await withTiming(
      timingRoot(
        mainSeq(clickStop(spinEffect(3, spid, 'by="10800000"'))),
        `<p:bldLst><p:bldP spid="${spid}" grpId="0"/></p:bldLst>`,
      ),
    );
    const before = decoder.decode(_internalPackageOf(pres).getPart(partName(SLIDE1))!.data);
    // `<p:animRot>` is a behaviour this library writes, so the layout check
    // alone would let this node be replaced. What stops it is that a half turn
    // is not an effect the read model names.
    expect(() => updateSlideAnimation(slide, 3, { effect: 'fadeIn' })).toThrow(/cannot edit/);
    expect(() => updateSlideAnimation(slide, 3, { durationMs: 900 })).toThrow(/cannot edit/);
    expect(decoder.decode(_internalPackageOf(pres).getPart(partName(SLIDE1))!.data)).toBe(before);
  });

  it('refuses a preset that turns the shape and does something else too', async () => {
    const spid = await firstShapeId();
    const composite = spinEffect(3, spid, 'by="21600000"').replace(
      '</p:childTnLst></p:cTn></p:par>',
      `<p:set><p:cBhvr><p:cTn id="9" dur="1" fill="hold"/>` +
        `<p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl>` +
        `<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr>` +
        `<p:to><p:strVal val="hidden"/></p:to></p:set></p:childTnLst></p:cTn></p:par>`,
    );
    const { slide } = await withTiming(timingRoot(mainSeq(clickStop(composite))));
    expect(getSlideAnimations(slide)[0]!.effect).toBeNull();
  });
});

describe('fn API: retiming an exit that hides at the end', () => {
  // This library writes the `<p:set>` that ends an exit one millisecond before
  // the end. A file that puts it somewhere else is saying something we did not
  // write and cannot restate, so a duration change moves the motion and leaves
  // that kick where it is.
  const exitWithHideAt = (id: number, spid: number, hideDelay: number): string =>
    `<p:par><p:cTn id="${id}" presetID="10" presetClass="exit" presetSubtype="0" fill="hold" ` +
    `grpId="0" nodeType="clickEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst>` +
    `<p:childTnLst><p:anim calcmode="lin" valueType="num"><p:cBhvr additive="base">` +
    `<p:cTn id="${id + 1}" dur="500" fill="hold"/>` +
    `<p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl>` +
    `<p:attrNameLst><p:attrName>style.opacity</p:attrName></p:attrNameLst></p:cBhvr>` +
    `<p:tavLst/></p:anim>` +
    `<p:set><p:cBhvr><p:cTn id="${id + 2}" dur="1" fill="hold">` +
    `<p:stCondLst><p:cond delay="${hideDelay}"/></p:stCondLst></p:cTn>` +
    `<p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl>` +
    `<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr>` +
    `<p:to><p:strVal val="hidden"/></p:to></p:set>` +
    `</p:childTnLst></p:cTn></p:par>`;

  const slideXml = (pres: PresentationData): string =>
    decoder.decode(_internalPackageOf(pres).getPart(partName(SLIDE1))!.data);

  const withExit = async (hideDelay: number) => {
    const spid = await firstShapeId();
    return withTiming(
      timingRoot(
        mainSeq(clickStop(exitWithHideAt(3, spid, hideDelay))),
        `<p:bldLst><p:bldP spid="${spid}" grpId="0"/></p:bldLst>`,
      ),
    );
  };

  it('moves a hide this library would have written', async () => {
    const { pres, slide } = await withExit(499);
    updateSlideAnimation(slide, 3, { durationMs: 1200 });
    const xml = slideXml(pres);
    expect(xml).toContain('dur="1200"');
    expect(xml).toContain('<p:cond delay="1199"/>');
  });

  it('leaves a hide the file put somewhere else alone', async () => {
    const { pres, slide } = await withExit(200);
    updateSlideAnimation(slide, 3, { durationMs: 1200 });
    const xml = slideXml(pres);
    expect(xml).toContain('dur="1200"');
    expect(xml).toContain('<p:cond delay="200"/>');
    expect(xml).not.toContain('<p:cond delay="1199"/>');
  });
});
