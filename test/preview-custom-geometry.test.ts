// Renderer coverage for custom geometry (`<a:custGeom>`): a custGeom shape
// renders as a real SVG `<path>` (not the labelled rect fallback), inherits
// the shape's fill, and carries no fallback marker — while a malformed
// custGeom still falls back with `data-pptx-fallback="custGeom"`.
//
// There is no public API to author custGeom, so the geometry is injected at
// the OPC zip layer (the same internal hook the chart-fallback test uses).

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
// Import the authoring API from `@office-kit/pptx` (not `../src/api`) so the deck's
// `PresentationData` shares type identity with the `renderSlideToSvg`
// signature, which also imports from `@office-kit/pptx`. The vitest alias resolves
// the package to the same source module at runtime.
import {
  addSlide,
  addSlideShape,
  findSlideLayout,
  getSlides,
  groupShapes,
  inches,
  loadPresentation,
  savePresentation,
  setShapeFill,
  setShapeSize,
} from '@office-kit/pptx';
import { readZip, writeZip } from '../src/internal/opc/index.ts';
import { renderSlideToSvg } from '../packages/preview/src/index.ts';

const fixturePath = fileURLToPath(new URL('./fixtures/minimal/blank.pptx', import.meta.url));

/** Builds a blank-layout deck with one red rect, then rewrites that rect's
 * prstGeom to `custGeom` at the zip layer and returns the reloaded deck. */
const renderWithCustGeom = async (custGeom: string): Promise<string> => {
  const pres = await loadPresentation(await readFile(fixturePath));
  const layout = findSlideLayout(pres, 'Blank');
  if (!layout) throw new Error('Blank layout missing');
  const slide = addSlide(pres, { layout });
  const shape = addSlideShape(slide, {
    preset: 'rect',
    x: inches(1),
    y: inches(1),
    w: inches(2),
    h: inches(2),
  });
  setShapeFill(shape, '#FF0000');
  const bytes = await savePresentation(pres);
  const { entries } = readZip(bytes);
  const dec = new TextDecoder();
  const enc = new TextEncoder();
  const modified = entries.map((e) => {
    if (!(e.name.includes('slides/slide') && e.name.endsWith('.xml'))) return e;
    const xml = dec
      .decode(e.data)
      .replace('<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>', custGeom);
    return { name: e.name, data: enc.encode(xml) };
  });
  const reloaded = await loadPresentation(writeZip(modified));
  return renderSlideToSvg(reloaded, getSlides(reloaded)[0]!, { textLayout: 'svg' });
};

// A right triangle in a 100×100 path space. Mapped onto a 2-inch (1828800
// EMU) square at offset (914400, 914400) EMU → 96 px box at (96, 96) px.
const TRIANGLE =
  '<a:custGeom><a:avLst/><a:gdLst/><a:pathLst><a:path w="100" h="100">' +
  '<a:moveTo><a:pt x="0" y="0"/></a:moveTo>' +
  '<a:lnTo><a:pt x="100" y="0"/></a:lnTo>' +
  '<a:lnTo><a:pt x="50" y="100"/></a:lnTo>' +
  '<a:close/></a:path></a:pathLst></a:custGeom>';

// `<a:rect>` is optional and states where the shape's text goes, in the same
// guide space as the path commands. Top-right quadrant of the box, written
// with built-in guides so the test also pins that they resolve.
const TEXT_RECT = '<a:rect l="hc" t="t" r="r" b="vc"/>';

const withTextRect = (custGeom: string, rect: string): string =>
  custGeom.replace('<a:pathLst>', `${rect}<a:pathLst>`);

/**
 * The text rectangle the HTML text path resolves for a custGeom shape: the
 * `<foreignObject>` is that rectangle, in px. `grouped` puts the shape in a
 * group squashed to half its natural height, because a group's scale reaches
 * the text bounds but not the shape's own `<a:ext>` — the space `<a:rect>` is
 * written in.
 */
const textRectWithCustGeom = async (
  custGeom: string,
  opts: { grouped?: boolean } = {},
): Promise<{ x: number; y: number; w: number; h: number }> => {
  const pres = await loadPresentation(await readFile(fixturePath));
  const layout = findSlideLayout(pres, 'Blank');
  if (!layout) throw new Error('Blank layout missing');
  const slide = addSlide(pres, { layout });
  const box = (w: number, h: number, text?: string) =>
    addSlideShape(slide, {
      preset: 'rect',
      x: inches(1),
      y: inches(1),
      w: inches(w),
      h: inches(h),
      ...(text === undefined ? {} : { text }),
    });
  const shape = box(2, 2, 'Alpha');
  if (opts.grouped) {
    // The sibling sits inside the custGeom shape, so the group's box is that
    // shape's own box and only the squash below separates ext from chExt.
    const group = groupShapes([shape, box(0.5, 0.5)]);
    setShapeSize(group, inches(2), inches(1));
  }
  const bytes = await savePresentation(pres);
  const { entries } = readZip(bytes);
  const dec = new TextDecoder();
  const enc = new TextEncoder();
  const modified = entries.map((e) => {
    if (!(e.name.includes('slides/slide') && e.name.endsWith('.xml'))) return e;
    const xml = dec
      .decode(e.data)
      .replace('<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>', custGeom);
    return { name: e.name, data: enc.encode(xml) };
  });
  const reloaded = await loadPresentation(writeZip(modified));
  const svg = renderSlideToSvg(reloaded, getSlides(reloaded)[0]!, {
    textLayout: 'foreignObject',
  });
  const attrs =
    /<foreignObject x="([\d.-]+)" y="([\d.-]+)" width="([\d.-]+)" height="([\d.-]+)"/.exec(svg);
  if (!attrs) throw new Error('no text body was rendered');
  return { x: +attrs[1]!, y: +attrs[2]!, w: +attrs[3]!, h: +attrs[4]! };
};

const pathDsIn = (svg: string): string[] =>
  [...svg.matchAll(/<path\b[^>]*\bd="([^"]*)"/g)].map((m) => m[1]!);

describe('renderSlideToSvg: custom geometry', () => {
  it('renders a triangle custGeom as a <path> with the expected command letters', async () => {
    const svg = await renderWithCustGeom(TRIANGLE);
    expect(svg).not.toContain('data-pptx-fallback="custGeom"');
    const ds = pathDsIn(svg);
    const triangleD = ds.find((d) => d.startsWith('M96'));
    expect(triangleD).toBeDefined();
    // M (moveTo) → L L (lnTo) → Z (close). The 2-inch (192 px) square sits at
    // (96, 96) px; path coords 0/50/100 map to 96/192/288 px across that box.
    expect(triangleD).toMatch(/^M96\.00,96\.00 L288\.00,96\.00 L192\.00,288\.00 Z$/);
  });

  it("carries the shape's solid fill on the geometry path", async () => {
    const svg = await renderWithCustGeom(TRIANGLE);
    const triPath = [...svg.matchAll(/<path\b[^>]*>/g)]
      .map((m) => m[0])
      .find((p) => p.includes('M96.00,96.00'));
    expect(triPath).toBeDefined();
    expect(triPath!.toLowerCase()).toMatch(/fill="#ff0000"/);
  });

  // The 2-inch square sits at (96, 96) px and spans 192 px. PowerPoint's
  // default insets are 0.1 in horizontally (9.6 px) and 0.05 in vertically
  // (4.8 px), and they come off the text rect, not the shape box.
  describe('its own text rectangle', () => {
    it('lays the text inside the stated rect', async () => {
      expect(await textRectWithCustGeom(withTextRect(TRIANGLE, TEXT_RECT))).toEqual({
        x: 201.6,
        y: 100.8,
        w: 76.8,
        h: 86.4,
      });
    });

    it('keeps the whole box when the geometry states no rect', async () => {
      // Not the triangle region the preset table would approximate: a custom
      // shape that says nothing about its text is saying the whole box.
      expect(await textRectWithCustGeom(TRIANGLE)).toEqual({
        x: 105.6,
        y: 100.8,
        w: 172.8,
        h: 182.4,
      });
    });

    it('keeps its proportions when a group squashes the shape', async () => {
      // Half height, so the rect covers the same quarter of the child: the
      // insets stay absolute, which is why the height is not simply halved.
      expect(
        await textRectWithCustGeom(withTextRect(TRIANGLE, TEXT_RECT), { grouped: true }),
      ).toEqual({ x: 201.6, y: 52.8, w: 76.8, h: 38.4 });
    });
  });

  it('falls back to the labelled rect for a malformed custGeom', async () => {
    const broken =
      '<a:custGeom><a:pathLst><a:path w="100" h="100">' +
      '<a:moveTo><a:pt x="nope" y="0"/></a:moveTo></a:path></a:pathLst></a:custGeom>';
    const svg = await renderWithCustGeom(broken);
    expect(svg).toContain('data-pptx-fallback="custGeom"');
  });
});
