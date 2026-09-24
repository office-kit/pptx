// `groupShapes` / `ungroupShapes` — compose a selection of shapes into a
// single `<p:grpSp>` and reverse it.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addBlankSlide,
  addSlide,
  addSlideShape,
  addSlideTextBox,
  createPresentation,
  emu,
  findSlideLayout,
  findSlidePlaceholder,
  getGroupChildren,
  getGroupTransform,
  getShapeBounds,
  getShapeId,
  getShapeKind,
  getShapePosition,
  getShapeRotation,
  getShapeFlip,
  getShapeText,
  getSlideShapes,
  getSlides,
  getSlideXmlString,
  groupShapes,
  getRegroupShapes,
  regroupShapes,
  inches,
  loadPresentation,
  savePresentation,
  removeShape,
  setShapePosition,
  setShapeRotation,
  setShapeFlip,
  setShapeSize,
  ungroupShapes,
} from '../src/api/index.ts';
import { expectSchemaValid, isSchemaValidationAvailable } from './lib/expect-schema-valid.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const skipIfNoXmllint = isSchemaValidationAvailable() ? it : it.skip;

// A from-scratch deck's blank slide carries zero shapes to start — unlike a
// loaded template's "Blank" layout, which may still contribute inherited
// date/footer/slide-number placeholder shells.
const blankSlide = () => {
  const pres = createPresentation();
  const slide = addBlankSlide(pres);
  return { pres, slide };
};

describe('fn API: groupShapes / ungroupShapes', () => {
  it('regroups former siblings from one selected member, retaining current geometry', async () => {
    const { pres, slide } = blankSlide();
    const add = (x: number) =>
      addSlideShape(slide, {
        preset: 'rect',
        x: inches(x),
        y: inches(1),
        w: inches(1),
        h: inches(1),
      });
    const group = groupShapes([add(1), add(3)], { name: 'Old name' });
    const members = ungroupShapes(group);
    setShapePosition(members[0]!, inches(2), inches(2));
    const before = members.map(getShapeBounds);
    expect(getRegroupShapes(members[1]!).map(getShapeId)).toEqual(members.map(getShapeId));
    const loaded = await loadPresentation(await savePresentation(pres));
    expect(getSlideShapes(getSlides(loaded)[0]!).every((s) => !getRegroupShapes(s).length)).toBe(
      true,
    );
    const regrouped = regroupShapes([members[1]!]);
    expect(getGroupChildren(regrouped).map(getShapeBounds)).toEqual(before);
    expect(getGroupChildren(regrouped).every((s) => !getRegroupShapes(s).length)).toBe(true);
    expect(getShapeKind(regrouped)).toBe('group');
  });

  it('restores only the first eligible former group and rejects detached membership', () => {
    const { slide } = blankSlide();
    const add = (x: number) =>
      addSlideShape(slide, {
        preset: 'rect',
        x: inches(x),
        y: inches(1),
        w: inches(1),
        h: inches(1),
      });
    const first = ungroupShapes(groupShapes([add(1), add(3)]));
    const second = ungroupShapes(groupShapes([add(5), add(7)]));
    const result = regroupShapes([second[0]!, first[0]!]);
    expect(getGroupChildren(result).map(getShapeId)).toEqual(second.map(getShapeId));
    expect(getRegroupShapes(first[0]!)).toHaveLength(2);
    removeShape(first[1]!);
    add(9);
    expect(getRegroupShapes(first[0]!)).toEqual([]);
    expect(() => regroupShapes([first[0]!])).toThrow('no previously ungrouped group');
  });

  it('wraps two shapes in a <p:grpSp> whose bounds are the union of its members', () => {
    const { slide } = blankSlide();
    const box = addSlideShape(slide, {
      preset: 'rect',
      x: inches(1),
      y: inches(1),
      w: inches(2),
      h: inches(1),
    });
    const label = addSlideTextBox(slide, {
      x: inches(1.25),
      y: inches(1.25),
      w: inches(1.5),
      h: inches(0.5),
      text: 'KPI card',
    });

    const group = groupShapes([box, label], { name: 'KPI Card' });

    expect(getShapeKind(group)).toBe('group');
    expect(getShapeBounds(group)).toEqual({
      x: inches(1),
      y: inches(1),
      w: inches(2),
      h: inches(1),
    });

    // The originals are gone from the top level — only the group remains
    // there — but reappear one level down via getGroupChildren, with their
    // own bounds intact.
    expect(getSlideShapes(slide).map((s) => getShapeKind(s))).toEqual(['group', 'shape', 'shape']);
    const children = getGroupChildren(group);
    expect(children).toHaveLength(2);
    expect(children.map((c) => getShapeText(c))).toContain('KPI card');
    expect(getShapePosition(children[1]!)).toEqual({ x: inches(1.25), y: inches(1.25) });
  });

  it('rejects grouping fewer than two shapes', () => {
    const { slide } = blankSlide();
    const only = addSlideShape(slide, {
      preset: 'rect',
      x: emu(0),
      y: emu(0),
      w: inches(1),
      h: inches(1),
    });
    expect(() => groupShapes([only])).toThrow(/at least 2 shapes/);
  });

  it('rejects a shape that has no explicit position/size', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const layout = findSlideLayout(pres, 'Title Slide');
    const slide = addSlide(pres, { layout: layout! });
    // Placeholders inherit geometry from the layout; they carry no <a:xfrm>
    // until explicitly resized.
    const title = findSlidePlaceholder(slide, 'ctrTitle');
    const box = addSlideShape(slide, {
      preset: 'rect',
      x: emu(0),
      y: emu(0),
      w: inches(1),
      h: inches(1),
    });
    expect(title).not.toBeNull();
    expect(() => groupShapes([title!, box])).toThrow(/no explicit position\/size/);
  });

  it('rejects the same shape passed twice (would duplicate its id)', () => {
    const { slide } = blankSlide();
    const a = addSlideShape(slide, {
      preset: 'rect',
      x: emu(0),
      y: emu(0),
      w: inches(1),
      h: inches(1),
    });
    const b = addSlideShape(slide, {
      preset: 'rect',
      x: inches(1),
      y: emu(0),
      w: inches(1),
      h: inches(1),
    });
    expect(() => groupShapes([a, b, a])).toThrow(/was passed twice/);
  });

  it('rejects grouping shapes with different parents', () => {
    const { slide } = blankSlide();
    const a = addSlideShape(slide, {
      preset: 'rect',
      x: emu(0),
      y: emu(0),
      w: inches(1),
      h: inches(1),
    });
    const b = addSlideShape(slide, {
      preset: 'rect',
      x: inches(1),
      y: emu(0),
      w: inches(1),
      h: inches(1),
    });
    const c = addSlideShape(slide, {
      preset: 'rect',
      x: inches(2),
      y: emu(0),
      w: inches(1),
      h: inches(1),
    });
    groupShapes([a, b]);
    // `a` is now nested one level down; grouping it directly (rather than
    // its enclosing group) must fail instead of silently doing nothing.
    expect(() => groupShapes([a, c])).toThrow(/not a direct child/);
  });

  it('groups and ungroups siblings inside a transformed parent', async () => {
    const { pres, slide } = blankSlide();
    const a = addSlideShape(slide, {
      preset: 'rect',
      x: inches(1),
      y: inches(2),
      w: inches(2),
      h: inches(1),
    });
    const b = addSlideShape(slide, {
      preset: 'ellipse',
      x: inches(4),
      y: inches(3),
      w: inches(2),
      h: inches(1),
    });
    const parent = groupShapes([a, b]);
    setShapeRotation(parent, 30);
    setShapeSize(parent, inches(9), inches(4));
    const before = getSlideXmlString(slide);
    const bounds = [getShapeBounds(a), getShapeBounds(b)];
    const inner = groupShapes([a, b]);
    expect(getGroupChildren(parent).map(getShapeId)).toEqual([getShapeId(inner)]);
    expect(getGroupChildren(inner).map(getShapeId)).toEqual([getShapeId(a), getShapeId(b)]);
    expect(getGroupChildren(inner).map(getShapeBounds)).toEqual(bounds);
    const restored = ungroupShapes(inner);
    expect(restored.map(getShapeId)).toEqual([getShapeId(a), getShapeId(b)]);
    expect(restored.map(getShapeBounds)).toEqual(bounds);
    expect(getSlideXmlString(slide)).toBe(before);
    const loaded = await loadPresentation(await savePresentation(pres));
    expect(getGroupChildren(getSlideShapes(getSlides(loaded)[0]!)[0]!).length).toBe(2);
  });

  it('allows grouping an existing group with another top-level shape (nested groups)', () => {
    const { slide } = blankSlide();
    const a = addSlideShape(slide, {
      preset: 'rect',
      x: emu(0),
      y: emu(0),
      w: inches(1),
      h: inches(1),
    });
    const b = addSlideShape(slide, {
      preset: 'rect',
      x: inches(1),
      y: emu(0),
      w: inches(1),
      h: inches(1),
    });
    const c = addSlideShape(slide, {
      preset: 'rect',
      x: inches(2),
      y: emu(0),
      w: inches(1),
      h: inches(1),
    });
    const inner = groupShapes([a, b]);
    const outer = groupShapes([inner, c]);
    expect(getShapeKind(outer)).toBe('group');
    expect(getGroupChildren(outer).map((s) => getShapeKind(s))).toEqual(['group', 'shape']);
  });

  it('round-trips through ungroupShapes when the group is never moved', () => {
    const { slide } = blankSlide();
    const a = addSlideShape(slide, {
      preset: 'ellipse',
      x: inches(1),
      y: inches(1),
      w: inches(1),
      h: inches(1),
    });
    const b = addSlideShape(slide, {
      preset: 'ellipse',
      x: inches(2.5),
      y: inches(1.5),
      w: inches(0.75),
      h: inches(0.75),
    });
    const boundsBefore = [getShapeBounds(a), getShapeBounds(b)];

    const group = groupShapes([a, b]);
    const restored = ungroupShapes(group);

    expect(restored).toHaveLength(2);
    expect(restored.map((s) => getShapeBounds(s))).toEqual(boundsBefore);
    expect(getSlideShapes(slide)).toHaveLength(2);
  });

  it('rescales children on ungroup after the group itself was moved and resized', () => {
    const { slide } = blankSlide();
    const a = addSlideShape(slide, {
      preset: 'rect',
      x: emu(0),
      y: emu(0),
      w: inches(1),
      h: inches(1),
    });
    const b = addSlideShape(slide, {
      preset: 'rect',
      x: inches(1),
      y: emu(0),
      w: inches(1),
      h: inches(1),
    });
    const group = groupShapes([a, b]);

    // Original group bounds: (0,0) 2in x 1in. Move to (1in, 1in) and double
    // the size — every child should end up scaled 2x and offset by (1in,1in).
    setShapePosition(group, inches(1), inches(1));
    setShapeSize(group, inches(4), inches(2));

    const [first, second] = ungroupShapes(group);
    expect(getShapeBounds(first!)).toEqual({
      x: inches(1),
      y: inches(1),
      w: inches(2),
      h: inches(2),
    });
    expect(getShapeBounds(second!)).toEqual({
      x: inches(3),
      y: inches(1),
      w: inches(2),
      h: inches(2),
    });
  });

  it.each([
    [90, false, false],
    [37, true, false],
    [215, false, true],
    [15, true, true],
  ] as const)(
    'preserves child corners when ungrouping rotation %s and flips %s/%s',
    async (angle, horizontal, vertical) => {
      const { pres, slide } = blankSlide();
      const a = addSlideShape(slide, {
        preset: 'rect',
        x: inches(1),
        y: inches(2),
        w: inches(2),
        h: inches(1),
      });
      const b = addSlideShape(slide, {
        preset: 'ellipse',
        x: inches(5),
        y: inches(4),
        w: inches(1),
        h: inches(2),
      });
      setShapeRotation(a, 23);
      setShapeFlip(a, { horizontal: true });
      const group = groupShapes([a, b]);
      setShapeRotation(group, angle);
      setShapeFlip(group, { horizontal, vertical });
      setShapePosition(group, inches(3), inches(1));
      const initial = getShapeBounds(group)!;
      setShapeSize(group, emu(initial.w * 2), emu(initial.h * 2));
      const project = (shape: typeof a, x: number, y: number) => {
        const box = getShapeBounds(shape)!;
        const flip = getShapeFlip(shape);
        const rad = (getShapeRotation(shape) * Math.PI) / 180;
        const dx = (x - box.w / 2) * (flip?.horizontal ? -1 : 1);
        const dy = (y - box.h / 2) * (flip?.vertical ? -1 : 1);
        return {
          x: box.x + box.w / 2 + dx * Math.cos(rad) - dy * Math.sin(rad),
          y: box.y + box.h / 2 + dx * Math.sin(rad) + dy * Math.cos(rad),
        };
      };
      const frame = getGroupTransform(group)!;
      const before = [a, b].map((child) => {
        const box = getShapeBounds(child)!;
        return [
          [0, 0],
          [1, 0],
          [1, 1],
          [0, 1],
        ].map(([x, y]) => {
          const point = project(child, x! * box.w, y! * box.h);
          return project(group, (point.x - frame.inner.x) * 2, (point.y - frame.inner.y) * 2);
        });
      });
      const restored = ungroupShapes(group);
      restored.forEach((child, index) => {
        const box = getShapeBounds(child)!;
        [
          [0, 0],
          [1, 0],
          [1, 1],
          [0, 1],
        ].forEach(([x, y], corner) => {
          const actual = project(child, x! * box.w, y! * box.h);
          expect(Math.abs(actual.x - before[index]![corner]!.x)).toBeLessThan(2);
          expect(Math.abs(actual.y - before[index]![corner]!.y)).toBeLessThan(2);
        });
      });
      const loaded = await loadPresentation(await savePresentation(pres));
      const saved = getSlideShapes(getSlides(loaded)[0]!);
      expect(saved.map(getShapeRotation)).toEqual(restored.map(getShapeRotation));
      expect(saved.map(getShapeFlip)).toEqual(restored.map(getShapeFlip));
    },
  );

  it('preserves each member rotation and z-order position across group/ungroup', () => {
    const { slide } = blankSlide();
    const back = addSlideShape(slide, {
      preset: 'rect',
      x: emu(0),
      y: emu(0),
      w: inches(1),
      h: inches(1),
    });
    const a = addSlideShape(slide, {
      preset: 'rect',
      x: inches(2),
      y: emu(0),
      w: inches(1),
      h: inches(1),
    });
    const b = addSlideShape(slide, {
      preset: 'rect',
      x: inches(3),
      y: emu(0),
      w: inches(1),
      h: inches(1),
    });
    const front = addSlideShape(slide, {
      preset: 'rect',
      x: inches(4),
      y: emu(0),
      w: inches(1),
      h: inches(1),
    });
    setShapeRotation(a, 45);

    const group = groupShapes([a, b]);
    // Group replaces [a, b] at their original position, between back and
    // front — group children flatten right after the group itself.
    expect(getSlideShapes(slide).map((s) => getShapeId(s))).toEqual([
      getShapeId(back),
      getShapeId(group),
      getShapeId(a),
      getShapeId(b),
      getShapeId(front),
    ]);

    const restored = ungroupShapes(group);
    expect(getShapeRotation(restored[0]!)).toBe(45);
    expect(getShapeRotation(restored[1]!)).toBe(0);
  });

  it('exposes getGroupTransform with chOff/chExt equal to off/ext right after creation', () => {
    const { slide } = blankSlide();
    const a = addSlideShape(slide, {
      preset: 'rect',
      x: inches(1),
      y: inches(1),
      w: inches(1),
      h: inches(1),
    });
    const b = addSlideShape(slide, {
      preset: 'rect',
      x: inches(2),
      y: inches(1),
      w: inches(1),
      h: inches(1),
    });
    const group = groupShapes([a, b]);
    const transform = getGroupTransform(group);
    expect(transform).toEqual({ outer: transform!.inner, inner: transform!.inner });
  });

  skipIfNoXmllint('produces schema-valid XML after group + move + ungroup', async () => {
    const { pres, slide } = blankSlide();
    const a = addSlideShape(slide, {
      preset: 'roundRect',
      x: inches(1),
      y: inches(1),
      w: inches(1.5),
      h: inches(1),
      text: 'Phase 1',
    });
    const b = addSlideShape(slide, {
      preset: 'roundRect',
      x: inches(3),
      y: inches(1),
      w: inches(1.5),
      h: inches(1),
      text: 'Phase 2',
    });
    const group = groupShapes([a, b], { name: 'Process' });
    setShapePosition(group, inches(0.5), inches(2));
    setShapeSize(group, inches(6), inches(2));
    expectSchemaValid(getSlideXmlString(slide), 'pml');

    ungroupShapes(group);
    expectSchemaValid(getSlideXmlString(slide), 'pml');

    await savePresentation(pres);
  });
});
