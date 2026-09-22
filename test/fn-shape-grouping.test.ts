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
  getSlides,
  getShapeText,
  getSlideShapes,
  getSlideXmlString,
  groupShapes,
  inches,
  loadPresentation,
  savePresentation,
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
  it('regroups siblings inside a transformed ancestor in stacking order and round-trips', async () => {
    const { pres, slide } = blankSlide();
    const members = [0, 1, 2].map((i) =>
      addSlideTextBox(slide, {
        x: inches(i + 1),
        y: inches(1),
        w: inches(1),
        h: inches(1),
        text: String(i),
      }),
    );
    const outer = groupShapes(members);
    setShapeRotation(outer, 35);
    setShapeFlip(outer, { horizontal: true, vertical: false });
    setShapeSize(outer, inches(6), inches(3));
    const before = getGroupTransform(outer);
    const bounds = members.map(getShapeBounds);
    const inner = groupShapes([members[1]!, members[0]!]);
    expect(getGroupChildren(inner).map(getShapeId)).toEqual(members.slice(0, 2).map(getShapeId));
    expect(getGroupChildren(outer).map(getShapeId)).toEqual([
      getShapeId(inner),
      getShapeId(members[2]!),
    ]);
    expect(getGroupTransform(outer)).toEqual(before);
    const restored = ungroupShapes(inner);
    expect(restored.map(getShapeBounds)).toEqual(bounds.slice(0, 2));
    expect(getGroupChildren(outer).map(getShapeId)).toEqual(members.map(getShapeId));
    const loaded = getSlides(await loadPresentation(await savePresentation(pres)))[0]!;
    const loadedOuter = getSlideShapes(loaded)[0]!;
    expect(getGroupTransform(loadedOuter)).toEqual(before);
    expect(getShapeRotation(loadedOuter)).toBe(35);
    expect(getShapeFlip(loadedOuter)).toEqual({ horizontal: true, vertical: false });
    expect(getGroupChildren(loadedOuter).map(getShapeBounds)).toEqual(bounds);
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

  it('rejects grouping shapes from different immediate parents', () => {
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
    // `a` is nested while `c` is still at slide level; they cannot share a new group.
    expect(() => groupShapes([a, c])).toThrow(/not a direct child/);
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

  for (const rotation of [45, 90, 270]) {
    for (const horizontal of [false, true]) {
      for (const vertical of [false, true]) {
        it(`preserves rotated group geometry after ungroup and reload (${rotation}, ${horizontal}, ${vertical})`, async () => {
          const { pres, slide } = blankSlide();
          const a = addSlideShape(slide, {
            preset: 'triangle',
            x: inches(1),
            y: inches(1),
            w: inches(2),
            h: inches(1),
          });
          const b = addSlideShape(slide, {
            preset: 'rect',
            x: inches(5),
            y: inches(2),
            w: inches(1),
            h: inches(2),
          });
          setShapeRotation(a, 30);
          setShapeFlip(a, { horizontal: true });
          const group = groupShapes([a, b]);
          setShapePosition(group, inches(2), inches(3));
          setShapeSize(group, inches(10), inches(6));
          setShapeRotation(group, rotation);
          setShapeFlip(group, { horizontal, vertical });
          const expected = [
            { cx: 4, cy: 4, w: 4, h: 2, rotation: 30, horizontal: true },
            { cx: 11, cy: 7, w: 2, h: 4, rotation: 0, horizontal: false },
          ].map((child) => {
            const dx = (child.cx - 7) * (horizontal ? -1 : 1);
            const dy = (child.cy - 6) * (vertical ? -1 : 1);
            const angle = (rotation * Math.PI) / 180;
            return {
              x: Math.round(
                (7 + dx * Math.cos(angle) - dy * Math.sin(angle) - child.w / 2) * 914400,
              ),
              y: Math.round(
                (6 + dx * Math.sin(angle) + dy * Math.cos(angle) - child.h / 2) * 914400,
              ),
              w: inches(child.w),
              h: inches(child.h),
              rotation:
                (rotation + (horizontal !== vertical ? -child.rotation : child.rotation) + 360) %
                360,
              flip: { horizontal: horizontal !== child.horizontal, vertical },
            };
          });
          const check = (shapes: readonly (typeof a)[]) => {
            shapes.forEach((shape, i) => {
              const { rotation: angle, flip, ...bounds } = expected[i]!;
              const actual = getShapeBounds(shape)!;
              expect(Math.abs(actual.x - bounds.x)).toBeLessThanOrEqual(1);
              expect(Math.abs(actual.y - bounds.y)).toBeLessThanOrEqual(1);
              expect(actual.w).toBe(bounds.w);
              expect(actual.h).toBe(bounds.h);
              expect(getShapeRotation(shape)).toBe(angle);
              expect(getShapeFlip(shape)).toEqual(flip);
            });
          };
          check(ungroupShapes(group));
          const loaded = await loadPresentation(await savePresentation(pres));
          check(getSlideShapes(getSlides(loaded)[0]!));
        });
      }
    }
  }

  it('composes nested group rotations without changing its child coordinate space', async () => {
    const { pres, slide } = blankSlide();
    const boxes = [1, 3, 5].map((x) =>
      addSlideShape(slide, {
        preset: 'rect',
        x: inches(x),
        y: inches(1),
        w: inches(1),
        h: inches(1),
      }),
    );
    const inner = groupShapes(boxes.slice(0, 2));
    setShapeRotation(inner, 30);
    const before = getGroupChildren(inner).map(getShapeBounds);
    const outer = groupShapes([inner, boxes[2]!]);
    setShapeRotation(outer, 90);
    const [restored] = ungroupShapes(outer);
    expect(getShapeKind(restored!)).toBe('group');
    expect(getShapeRotation(restored!)).toBe(120);
    expect(getShapeBounds(restored!)).toEqual({
      x: inches(2),
      y: emu(0),
      w: inches(3),
      h: inches(1),
    });
    expect(getGroupChildren(restored!).map(getShapeBounds)).toEqual(before);
    const loaded = await loadPresentation(await savePresentation(pres));
    const children = ungroupShapes(getSlideShapes(getSlides(loaded)[0]!)[0]!);
    children.forEach((child, index) => {
      const bounds = getShapeBounds(child)!;
      expect(Math.abs(bounds.x - inches(index ? 2.5 : 3.5))).toBeLessThanOrEqual(1);
      expect(
        Math.abs(bounds.y - Math.round((((index ? 1 : -1) * Math.sqrt(3)) / 2) * 914400)),
      ).toBeLessThanOrEqual(1);
      expect(getShapeRotation(child)).toBe(120);
    });
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
