// Text-body bevels (`<a:bodyPr><a:scene3d>…<a:sp3d><a:bevelT/>`), as Mac
// the reference desktop app's Soft Bevel and Sharp Bevel text-art presets write them — see
// test/fixtures/native/text-art-capture.md.

import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import {
  addBlankSlide,
  addSlideTable,
  addSlideTextBox,
  createPresentation,
  getTableCells,
  getTableCellText3D,
  inches,
  setShapeText3D,
  setShapeTextFormat,
  type ReadText3D,
  type Text3D,
} from '../src/api/index.ts';
import { CELL_ELEMENT } from '../src/api/_internal-symbols.ts';
import { applyText3D, readText3D } from '../src/internal/drawingml/index.ts';
import { type XmlElement, parseXml } from '../src/internal/xml/index.ts';
import { renderSlideToSvg, renderTextEffectsSvg } from '../packages/preview/src/index.ts';
import { textBevelOf } from '../packages/preview/src/text-bevel.ts';

const NATIVE = new URL('./fixtures/native/', import.meta.url);

const find = (root: XmlElement, local: string): XmlElement => {
  const stack = [root];
  while (stack.length) {
    const node = stack.pop()!;
    if (node.name.localName === local) return node;
    for (const child of node.children) if (child.kind === 'element') stack.push(child);
  }
  throw new Error(`no a:${local}`);
};

// The fixture's body 3-D, minus the contour (not rendered, and its read-back
// color token is not a writable `Color`).
const nativeText3D = (name: string): Text3D => {
  const read = readText3D(
    find(parseXml(readFileSync(new URL(name, NATIVE), 'utf8')).root, 'bodyPr'),
  )!;
  const {
    contourColor: _color,
    contourColorTransforms: _transforms,
    extrusionColor: _extrusion,
    ...rest
  } = read;
  return rest;
};
const SOFT = 'text-art-accent4-soft-bevel-shape.xml';
const SHARP = 'text-art-accent3-sharp-bevel-shape.xml';

const deck = (text3d: Text3D | null) => {
  const pres = createPresentation();
  const slide = addBlankSlide(pres);
  const shape = addSlideTextBox(slide, {
    x: inches(1),
    y: inches(1),
    w: inches(6),
    h: inches(2),
    text: 'Outline title',
  });
  setShapeTextFormat(shape, { size: 40, bold: true, color: 'accent4' });
  if (text3d) setShapeText3D(shape, text3d);
  return { pres, slide, shape };
};

// The bevel filter each beveled glyph layer of `text` references.
const bevelFilters = (svg: string, text: string): string[] =>
  [
    ...svg.matchAll(
      /<text [^>]*filter="url\(#([^)]+)\)"[^>]*data-pptx-bevel="text">((?:(?!<\/text>).)*)<\/text>/g,
    ),
  ]
    .filter(([, , body]) => body!.includes(`>${text}</tspan>`))
    .map(([, id]) => new RegExp(`<filter id="${id}"[^>]*>.*?</filter>`).exec(svg)![0]);

const attr = (xml: string, element: string, name: string): string | undefined =>
  new RegExp(`<${element}\\b[^>]*\\b${name}="([^"]*)"`).exec(xml)?.[1];

describe('renderSlideToSvg: text bevels', () => {
  it('shades the Soft Bevel glyphs from its width, height and light rig', () => {
    const { pres, slide } = deck(nativeText3D(SOFT));
    const svg = renderSlideToSvg(pres, slide, { textLayout: 'svg' });
    const [filter, ...rest] = bevelFilters(svg, 'Outline title');
    expect(rest).toEqual([]);
    // w = 25400 EMU = 2.67 px; the alpha blur spans it with σ = w / 2.
    expect(attr(filter!, 'feGaussianBlur', 'stdDeviation')).toBe('1.33');
    // h = 38100 EMU = 4 px, times the relief factor.
    const relief = textBevelOf(nativeText3D(SOFT))!.relief;
    expect(attr(filter!, 'feDiffuseLighting', 'surfaceScale')).toBe(String(4 * relief));
    // Soft rig from the top (270°), turned by its 260° revolution.
    expect(attr(filter!, 'feDistantLight', 'azimuth')).toBe('170');
    expect(attr(filter!, 'feDistantLight', 'elevation')).toBe('45');
    // softEdge material: a faint specular highlight.
    expect(filter).toContain('<feSpecularLighting');
    expect(filter).toContain('<feComposite in="bevelMerged" in2="SourceAlpha" operator="in"/>');
    // The bevel is painted over the solid glyphs, not instead of them.
    expect(svg.indexOf('data-pptx-bevel="text"')).toBeGreaterThan(
      svg.search(/<text [^>]*xml:space="preserve"><tspan[^>]*>Outline title/),
    );
  });

  it('shades the Sharp Bevel with the angle profile under a harsh top light', () => {
    const { pres, slide } = deck(nativeText3D(SHARP));
    const [filter] = bevelFilters(
      renderSlideToSvg(pres, slide, { textLayout: 'svg' }),
      'Outline title',
    );
    // w = 63500 EMU = 6.67 px.
    expect(attr(filter!, 'feGaussianBlur', 'stdDeviation')).toBe('3.33');
    expect(attr(filter!, 'feDistantLight', 'azimuth')).toBe('270');
    expect(attr(filter!, 'feDistantLight', 'elevation')).toBe('30');
    // matte: no specular. `angle` is a straight chamfer.
    expect(filter).not.toContain('feSpecularLighting');
    expect(attr(filter!, 'feFuncA', 'tableValues')).toBe(
      '0 0 0 0 0 0 0 0 0 0.13 0.25 0.38 0.5 0.63 0.75 0.88 1',
    );
  });

  it('draws the bevel over the browser-layout HTML and in the editing overlay', () => {
    const { pres, slide, shape } = deck(nativeText3D(SOFT));
    const svg = renderSlideToSvg(pres, slide);
    const after = svg.slice(svg.indexOf('</foreignObject>'));
    expect(bevelFilters(after, 'Outline title')).toHaveLength(1);
    const overlay = renderTextEffectsSvg(pres, slide, shape, { w: inches(6), h: inches(2) });
    expect(bevelFilters(overlay, 'Outline title')).toHaveLength(1);
  });

  it('bevels table cells whose text body carries one', () => {
    const pres = createPresentation();
    const slide = addBlankSlide(pres);
    const table = addSlideTable(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(6),
      h: inches(1),
      rows: [['Cell', 'Plain']],
    });
    const [beveled, plain] = getTableCells(table)[0]!;
    applyText3D(find(beveled![CELL_ELEMENT], 'bodyPr'), nativeText3D(SHARP), 'test');
    expect(getTableCellText3D(beveled!)?.bevelTop).toEqual({
      widthEmu: 63500,
      heightEmu: 12700,
      preset: 'angle',
    });
    expect(getTableCellText3D(plain!)).toBeNull();
    for (const textLayout of ['svg', 'foreignObject'] as const) {
      const svg = renderSlideToSvg(pres, slide, { textLayout });
      expect(bevelFilters(svg, 'Cell')).toHaveLength(1);
      expect(bevelFilters(svg, 'Plain')).toEqual([]);
    }
    const overlay = renderTextEffectsSvg(
      pres,
      slide,
      table,
      { w: inches(3), h: inches(1) },
      { cell: { row: 0, col: 0 } },
    );
    expect(bevelFilters(overlay, 'Cell')).toHaveLength(1);
  });

  it('leaves text without a bevel unchanged', () => {
    const flat = deck(null);
    for (const textLayout of ['svg', 'foreignObject'] as const) {
      const svg = renderSlideToSvg(flat.pres, flat.slide, { textLayout });
      expect(svg).not.toContain('data-pptx-bevel');
      expect(svg).not.toContain('feDiffuseLighting');
    }
    // A scene without a top bevel, and a flat rig, draw nothing either.
    const sceneOnly = deck({
      scene: { camera: 'orthographicFront', lightRig: { type: 'threePt', direction: 't' } },
    });
    expect(renderSlideToSvg(sceneOnly.pres, sceneOnly.slide)).not.toContain('data-pptx-bevel');
    expect(
      textBevelOf({
        scene: { camera: 'orthographicFront', lightRig: { type: 'flat', direction: 't' } },
        bevelTop: {},
      }),
    ).toBeNull();
  });
});

describe('textBevelOf', () => {
  const PRESETS = [
    'relaxedInset',
    'circle',
    'slope',
    'cross',
    'angle',
    'softRound',
    'convex',
    'coolSlant',
    'divot',
    'riblet',
    'hardEdge',
    'artDeco',
  ] as const;

  it('applies the CT_Bevel defaults: circle, 76200 × 76200 EMU, threePt from the top', () => {
    const bevel = textBevelOf({ bevelTop: {} })!;
    expect(bevel.widthEmu).toBe(76200);
    expect(bevel.heightEmu).toBe(76200);
    expect(bevel.profile).toEqual(textBevelOf({ bevelTop: { preset: 'circle' } })!.profile);
    expect(bevel.azimuthDeg).toBe(270);
  });

  it('gives every preset a distinct height profile from the edge up', () => {
    const profiles = PRESETS.map((preset) => textBevelOf({ bevelTop: { preset } })!.profile);
    for (const profile of profiles) {
      expect(profile[0]).toBe(0);
      for (const height of profile) {
        expect(height).toBeGreaterThanOrEqual(0);
        expect(height).toBeLessThanOrEqual(1);
      }
    }
    // `slope` and `angle` are both straight chamfers; every other pair differs.
    const keys = new Set(profiles.map((profile) => profile.join(',')));
    expect(keys.size).toBe(PRESETS.length - 1);
  });

  it('maps every light direction onto a screen azimuth', () => {
    const azimuth = (direction: NonNullable<ReadText3D['scene']>['lightRig']['direction']) =>
      textBevelOf({
        scene: { camera: 'orthographicFront', lightRig: { type: 'threePt', direction } },
        bevelTop: {},
      })!.azimuthDeg;
    expect(
      (['r', 'br', 'b', 'bl', 'l', 'tl', 't', 'tr'] as const).map((direction) =>
        azimuth(direction),
      ),
    ).toEqual([0, 45, 90, 135, 180, 225, 270, 315]);
  });

  it('only shiny materials add a specular highlight', () => {
    const specular = (material: NonNullable<ReadText3D['material']>) =>
      textBevelOf({ bevelTop: {}, material })!.specular;
    expect(specular('matte')).toBeNull();
    expect(specular('warmMatte')).toBeNull();
    expect(specular('metal')).not.toBeNull();
    expect(specular('plastic')).not.toBeNull();
    expect(textBevelOf({ bevelTop: {}, material: 'flat' })).toBeNull();
  });
});
