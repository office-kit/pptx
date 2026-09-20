// Resizing a group moves and resizes its children but leaves their text at the
// authored point size — PowerPoint and LibreOffice both render a group child's
// glyphs undistorted however far `<a:ext>` drifts from `<a:chExt>`.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  addSlideTextBox,
  findSlideLayout,
  groupShapes,
  inches,
  loadPresentation,
  setShapeSize,
  setShapeText,
  setShapeTextFormat,
} from '@office-kit/pptx';
import { renderSlideToSvg } from '../packages/preview/src/index.ts';

const fixturePath = fileURLToPath(new URL('./fixtures/minimal/blank.pptx', import.meta.url));

const TEXT_PT = 20;

const fontSizesIn = (svg: string): string[] => [
  ...new Set([...svg.matchAll(/font-size:([\d.]+)px/g)].map((m) => m[1]!)),
];

// Two stacked boxes, optionally grouped and squashed to half the group's
// natural height (the shape Google Slides writes for a hand-resized group).
const deckWith = async (opts: { group: boolean }) => {
  const pres = await loadPresentation(await readFile(fixturePath));
  const slide = addSlide(pres, { layout: findSlideLayout(pres, 'Blank')! });
  const top = addSlideTextBox(slide, {
    x: inches(1),
    y: inches(1),
    w: inches(3),
    h: inches(1),
    text: 'Alpha',
  });
  setShapeText(top, 'Alpha');
  setShapeTextFormat(top, { size: TEXT_PT });
  const bottom = addSlideTextBox(slide, {
    x: inches(1),
    y: inches(2.5),
    w: inches(3),
    h: inches(1),
    text: 'Beta',
  });
  setShapeText(bottom, 'Beta');
  setShapeTextFormat(bottom, { size: TEXT_PT });
  if (opts.group) {
    const group = groupShapes([top, bottom]);
    setShapeSize(group, inches(3), inches(1.25));
  }
  return { pres, slide };
};

describe('renderSlideToSvg: group scale and child text', () => {
  it('keeps a squashed group’s text at its authored size', async () => {
    const plain = await deckWith({ group: false });
    const grouped = await deckWith({ group: true });

    const plainSvg = renderSlideToSvg(plain.pres, plain.slide);
    const groupedSvg = renderSlideToSvg(grouped.pres, grouped.slide);

    expect(fontSizesIn(groupedSvg)).toEqual(fontSizesIn(plainSvg));
  });

  it('cancels the group scale over the text so glyphs stay undistorted', async () => {
    const { pres, slide } = await deckWith({ group: true });
    const svg = renderSlideToSvg(pres, slide);

    // The group halves its children vertically; the text carries the inverse.
    expect(svg).toContain('scale(1.000000 0.500000)');
    expect(svg).toContain('<g transform="scale(1.000000 2.000000)">');
  });

  it('lays the text out in the group-scaled rect', async () => {
    const { pres, slide } = await deckWith({ group: true });
    const svg = renderSlideToSvg(pres, slide);

    // Boxes are 1in tall inside a group squashed to half height, so each frame
    // covers half an inch (48px at 96 DPI). The default 0.05in top/bottom
    // insets are text-body properties, so they no more scale than the font
    // does: 48 - 2 * 4.8 = 38.4px of text area.
    const heights = [...svg.matchAll(/<foreignObject[^>]*height="([\d.]+)"/g)].map((m) =>
      Number(m[1]),
    );
    expect(heights.length).toBeGreaterThan(0);
    for (const h of heights) expect(h).toBeCloseTo(38.4, 1);
  });
});
