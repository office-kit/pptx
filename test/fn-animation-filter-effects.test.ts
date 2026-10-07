// The entrance and exit effects that run a transition filter on one object
// (`<p:animEffect filter>`), the peeks that slide while they wipe, and the
// diagonal flies.
//
// The expected rows are PowerPoint's own: preset id, subtype, filter and
// default duration as PowerPoint 16 writes them for each gallery entry and
// Effect Options choice. They are spelled out here rather than derived from the
// writer's table, so a slip in that table cannot agree with itself.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { expectSchemaValid, isSchemaValidationAvailable } from './lib/expect-schema-valid.ts';
import {
  type AnimationOptions,
  type SlideAnimationStep,
  getSlideAnimations,
  getSlideShapes,
  getSlideXmlString,
  getSlides,
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

const reopen = async (pres: Awaited<ReturnType<typeof openDeck>>['pres']) => {
  const slide = getSlides(await loadPresentation(await savePresentation(pres)))[0]!;
  return { xml: getSlideXmlString(slide), steps: getSlideAnimations(slide) };
};

interface Row {
  readonly opts: AnimationOptions;
  readonly presetId: number;
  readonly subtype: number;
  readonly filter: string;
  readonly durationMs: number;
}

const both = (
  name: string,
  opts: Omit<AnimationOptions, 'effect'>,
  presetId: number,
  subtype: number,
  filter: string,
  durationMs = 500,
): Row[] => [
  {
    opts: { effect: `${name}In` as AnimationOptions['effect'], ...opts },
    presetId,
    subtype,
    filter,
    durationMs,
  },
  {
    opts: { effect: `${name}Out` as AnimationOptions['effect'], ...opts },
    presetId,
    subtype,
    filter,
    durationMs,
  },
];

const ROWS: readonly Row[] = [
  ...both('wipe', {}, 22, 4, 'wipe(down)'),
  ...both('wipe', { direction: 'top' }, 22, 1, 'wipe(up)'),
  ...both('wipe', { direction: 'right' }, 22, 2, 'wipe(right)'),
  ...both('wipe', { direction: 'left' }, 22, 8, 'wipe(left)'),
  ...both('blinds', {}, 3, 10, 'blinds(horizontal)'),
  ...both('blinds', { orientation: 'vertical' }, 3, 5, 'blinds(vertical)'),
  ...both('checkerboard', {}, 5, 10, 'checkerboard(across)'),
  ...both('checkerboard', { orientation: 'vertical' }, 5, 5, 'checkerboard(down)'),
  ...both('randomBars', {}, 14, 10, 'randombar(horizontal)'),
  ...both('randomBars', { orientation: 'vertical' }, 14, 5, 'randombar(vertical)'),
  ...both('split', {}, 16, 21, 'barn(inVertical)'),
  ...both('split', { orientation: 'horizontal' }, 16, 26, 'barn(inHorizontal)'),
  ...both('split', { inOut: 'out' }, 16, 37, 'barn(outVertical)'),
  ...both('split', { orientation: 'horizontal', inOut: 'out' }, 16, 42, 'barn(outHorizontal)'),
  ...both('strips', {}, 18, 12, 'strips(downLeft)'),
  ...both('strips', { direction: 'topLeft' }, 18, 9, 'strips(upLeft)'),
  ...both('strips', { direction: 'topRight' }, 18, 3, 'strips(upRight)'),
  ...both('strips', { direction: 'bottomRight' }, 18, 6, 'strips(downRight)'),
  ...both('shape', {}, 6, 16, 'circle(in)', 2000),
  ...both('shape', { inOut: 'out' }, 6, 32, 'circle(out)', 2000),
  ...both('shape', { shape: 'box' }, 4, 16, 'box(in)', 2000),
  ...both('shape', { shape: 'diamond', inOut: 'out' }, 8, 32, 'diamond(out)', 2000),
  ...both('shape', { shape: 'plus' }, 13, 16, 'plus(in)', 2000),
  ...both('dissolve', {}, 9, 0, 'dissolve'),
  ...both('wedge', {}, 20, 0, 'wedge', 2000),
  ...both('wheel', {}, 21, 1, 'wheel(1)', 2000),
  ...both('wheel', { spokes: 3 }, 21, 3, 'wheel(3)', 2000),
  ...both('wheel', { spokes: 8 }, 21, 8, 'wheel(8)', 2000),
];

const OPTION_FIELDS = ['direction', 'orientation', 'inOut', 'shape', 'spokes'] as const;

describe('fn API: filter entrance and exit effects', () => {
  it.each(ROWS.map((row) => [JSON.stringify(row.opts), row] as const))(
    '%s writes PowerPoint’s preset, filter and default duration, and reads back',
    async (_, row) => {
      const { pres, shape } = await openDeck();
      setShapeAnimation(shape, row.opts);
      const { xml, steps } = await reopen(pres);
      const entering = row.opts.effect.endsWith('In');

      expect(xml).toContain(`presetID="${row.presetId}"`);
      expect(xml).toContain(`presetClass="${entering ? 'entr' : 'exit'}"`);
      expect(xml).toContain(`presetSubtype="${row.subtype}"`);
      const effects = xml.match(/<p:animEffect\b[^>]*>/g) ?? [];
      expect(effects).toEqual([
        `<p:animEffect transition="${entering ? 'in' : 'out'}" filter="${row.filter}">`,
      ]);
      expect(xml).toMatch(new RegExp(`<p:animEffect[\\s\\S]*?dur="${row.durationMs}"`));
      // An entrance shows the shape before uncovering it; an exit covers it up
      // and hides it a millisecond before the end.
      const visibility = [...xml.matchAll(/<p:strVal val="(visible|hidden)"/g)].map((m) => m[1]);
      expect(visibility).toEqual([entering ? 'visible' : 'hidden']);
      if (!entering) expect(xml).toContain(`delay="${row.durationMs - 1}"`);

      const [step] = steps;
      expect(step!.effect).toBe(row.opts.effect);
      expect(step!.durationMs).toBe(row.durationMs);
      expect(step!.editable).toBe(true);
      for (const field of OPTION_FIELDS) {
        if (row.opts[field] !== undefined) expect(step![field]).toBe(row.opts[field]);
      }
    },
  );

  it.skipIf(!isSchemaValidationAvailable())(
    'every filter effect is schema-valid in one slide',
    async () => {
      const { pres, slide } = await openDeck();
      const shape = getSlideShapes(slide)[0]!;
      for (const row of ROWS) setShapeAnimation(shape, row.opts);
      const { xml, steps } = await reopen(pres);
      expectSchemaValid(xml, 'pml');
      expect(steps.map((s) => s.effect)).toEqual(ROWS.map((r) => r.opts.effect));
    },
  );

  it('reports the options each effect resolved to, defaults included', async () => {
    const { pres, shape } = await openDeck();
    setShapeAnimation(shape, { effect: 'splitIn' });
    setShapeAnimation(shape, { effect: 'shapeOut' });
    setShapeAnimation(shape, { effect: 'wedgeIn' });
    const { steps } = await reopen(pres);
    const options = (s: SlideAnimationStep) =>
      Object.fromEntries(OPTION_FIELDS.map((f) => [f, s[f]]));
    expect(steps.map(options)).toEqual([
      { direction: null, orientation: 'vertical', inOut: 'in', shape: null, spokes: null },
      { direction: null, orientation: null, inOut: 'in', shape: 'circle', spokes: null },
      { direction: null, orientation: null, inOut: null, shape: null, spokes: null },
    ]);
  });

  it('refuses an option the effect does not take, or a value it does not offer', async () => {
    const { shape } = await openDeck();
    expect(() => setShapeAnimation(shape, { effect: 'wipeIn', direction: 'topLeft' })).toThrow(
      /direction "topLeft" is not one of/,
    );
    expect(() => setShapeAnimation(shape, { effect: 'stripsIn', direction: 'top' })).toThrow(
      /direction "top" is not one of/,
    );
    expect(() => setShapeAnimation(shape, { effect: 'wheelIn', spokes: 5 })).toThrow(
      /spokes 5 is not one of/,
    );
    expect(() =>
      setShapeAnimation(shape, { effect: 'dissolveIn', orientation: 'vertical' }),
    ).toThrow(/orientation only applies to/);
    expect(() => setShapeAnimation(shape, { effect: 'fadeIn', shape: 'box' })).toThrow(
      /shape only applies to shapeIn, shapeOut/,
    );
  });

  it('updateSlideAnimation changes an option and keeps the ones it shares', async () => {
    const { pres, slide, shape } = await openDeck();
    setShapeAnimation(shape, { effect: 'splitIn', orientation: 'horizontal', inOut: 'out' });
    const id = getSlideAnimations(slide)[0]!.id!;
    updateSlideAnimation(slide, id, { inOut: 'in' });
    expect(getSlideAnimations(slide)[0]).toMatchObject({ orientation: 'horizontal', inOut: 'in' });
    // Blinds take an orientation too, so the horizontal survives the switch;
    // the in / out does not, since blinds have none.
    updateSlideAnimation(slide, id, { effect: 'blindsOut' });
    expect(getSlideAnimations(slide)[0]).toMatchObject({
      effect: 'blindsOut',
      orientation: 'horizontal',
      inOut: null,
    });
    // Wipe takes only the four edges, so a fly from a corner turns into a
    // wipe from the default edge rather than an invalid one.
    updateSlideAnimation(slide, id, { effect: 'flyIn', direction: 'topRight' });
    updateSlideAnimation(slide, id, { effect: 'wipeIn' });
    expect(getSlideAnimations(slide)[0]).toMatchObject({ effect: 'wipeIn', direction: 'bottom' });
    expect(() => updateSlideAnimation(slide, id, { spokes: 2 })).toThrow(/spokes only applies/);
    const { steps } = await reopen(pres);
    expect(steps[0]).toMatchObject({ id, effect: 'wipeIn', direction: 'bottom', editable: true });
  });
});

describe('fn API: peek', () => {
  // PowerPoint slides the shape 1.125 of its size on the edge's axis only, and
  // wipes it from the far side as it comes in.
  it.each([
    ['peekIn', 'bottom', 4, 'ppt_y', ['#ppt_y+#ppt_h*1.125000', '#ppt_y'], 'wipe(up)'],
    ['peekIn', 'top', 1, 'ppt_y', ['#ppt_y-#ppt_h*1.125000', '#ppt_y'], 'wipe(down)'],
    ['peekIn', 'right', 2, 'ppt_x', ['#ppt_x+#ppt_w*1.125000', '#ppt_x'], 'wipe(left)'],
    ['peekIn', 'left', 8, 'ppt_x', ['#ppt_x-#ppt_w*1.125000', '#ppt_x'], 'wipe(right)'],
    ['peekOut', 'bottom', 4, 'ppt_y', ['#ppt_y', '#ppt_y+#ppt_h*1.125000'], 'wipe(down)'],
    ['peekOut', 'top', 1, 'ppt_y', ['#ppt_y', '#ppt_y-#ppt_h*1.125000'], 'wipe(up)'],
    ['peekOut', 'right', 2, 'ppt_x', ['#ppt_x', '#ppt_x+#ppt_w*1.125000'], 'wipe(right)'],
    ['peekOut', 'left', 8, 'ppt_x', ['#ppt_x', '#ppt_x-#ppt_w*1.125000'], 'wipe(left)'],
  ] as const)('%s from %s', async (effect, direction, subtype, axis, values, filter) => {
    const { pres, shape } = await openDeck();
    setShapeAnimation(shape, { effect, direction });
    const { xml, steps } = await reopen(pres);
    expect(xml).toContain('presetID="12"');
    expect(xml).toContain(`presetSubtype="${subtype}"`);
    const anims = xml.match(/<p:anim\b[\s\S]*?<\/p:anim>/g) ?? [];
    expect(anims).toHaveLength(1);
    expect(anims[0]).toContain(`<p:attrName>${axis}</p:attrName>`);
    expect([...anims[0]!.matchAll(/<p:strVal val="([^"]*)"/g)].map((m) => m[1])).toEqual(values);
    expect(xml).toContain(`filter="${filter}"`);
    // The slide comes first, then the wipe — PowerPoint's order.
    expect(xml.indexOf('<p:anim ')).toBeLessThan(xml.indexOf('<p:animEffect'));
    expect(steps[0]).toMatchObject({ effect, direction, editable: true });
  });
});

describe('fn API: diagonal fly', () => {
  // A corner moves both axes: the centre starts (or ends) one half-size past
  // both edges. Entrances refer to the shape's own size with `#`, exits do not.
  it.each([
    ['flyIn', 'topLeft', 9, ['0-#ppt_w/2', '#ppt_x'], ['0-#ppt_h/2', '#ppt_y']],
    ['flyIn', 'topRight', 3, ['1+#ppt_w/2', '#ppt_x'], ['0-#ppt_h/2', '#ppt_y']],
    ['flyIn', 'bottomLeft', 12, ['0-#ppt_w/2', '#ppt_x'], ['1+#ppt_h/2', '#ppt_y']],
    ['flyIn', 'bottomRight', 6, ['1+#ppt_w/2', '#ppt_x'], ['1+#ppt_h/2', '#ppt_y']],
    ['flyOut', 'bottomRight', 6, ['ppt_x', '1+ppt_w/2'], ['ppt_y', '1+ppt_h/2']],
    ['flyOut', 'topLeft', 9, ['ppt_x', '0-ppt_w/2'], ['ppt_y', '0-ppt_h/2']],
  ] as const)('%s from %s', async (effect, direction, subtype, x, y) => {
    const { pres, shape } = await openDeck();
    setShapeAnimation(shape, { effect, direction });
    const { xml, steps } = await reopen(pres);
    expect(xml).toContain(`presetSubtype="${subtype}"`);
    const values = (attr: string) => {
      const anim = (xml.match(/<p:anim\b[\s\S]*?<\/p:anim>/g) ?? []).find((a) =>
        a.includes(`<p:attrName>${attr}</p:attrName>`),
      )!;
      return [...anim.matchAll(/<p:strVal val="([^"]*)"/g)].map((m) => m[1]);
    };
    expect(values('ppt_x')).toEqual(x);
    expect(values('ppt_y')).toEqual(y);
    expect(steps[0]).toMatchObject({ effect, direction, editable: true });
  });
});
