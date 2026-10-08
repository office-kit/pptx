// Gradient and pattern glyph fills (`<a:gradFill>` / `<a:pattFill>` on a run),
// as the reference desktop app's text-art presets write them — see
// test/fixtures/native/text-art-capture.md.

import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import {
  addBlankSlide,
  addSlideTable,
  addSlideTextBox,
  getTableCells,
  createPresentation,
  inches,
  setShapeTextFormat,
  setTableCellTextFormat,
  toWritableTextFormat,
  type TextFormat,
} from '../src/api/index.ts';
import { parseRPrLikeElement } from '../src/api/fn/shape-color.ts';
import { type XmlElement, parseXml } from '../src/internal/xml/index.ts';
import { renderSlideToSvg, renderTextEffectsSvg } from '../packages/preview/src/index.ts';

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

const nativeFormat = (name: string): TextFormat =>
  toWritableTextFormat(
    parseRPrLikeElement(find(parseXml(readFileSync(new URL(name, NATIVE), 'utf8')).root, 'rPr')),
  );

const deck = (format: TextFormat, text = 'Outline title') => {
  const pres = createPresentation();
  const slide = addBlankSlide(pres);
  const shape = addSlideTextBox(slide, {
    x: inches(1),
    y: inches(1),
    w: inches(6),
    h: inches(2),
    text,
  });
  setShapeTextFormat(shape, { size: 40, ...format });
  return { pres, slide, shape };
};

// A tspan whose glyphs paint with the given paint server.
const painted = (id: string, text: string): RegExp =>
  new RegExp(`<tspan [^>]*fill="url\\(#${id}\\)"[^>]*>${text}</tspan>`);

const paintId = (svg: string, element: string): string => {
  const id = new RegExp(`<${element} id="([^"]+)"`).exec(svg)?.[1];
  if (!id) throw new Error(`no <${element}>`);
  return id;
};

describe('renderSlideToSvg: gradient and pattern text fills', () => {
  it('paints the gray text-art gradient across the text block', () => {
    const { pres, slide } = deck(nativeFormat('text-art-gray-gradient-shape.xml'));
    const svg = renderSlideToSvg(pres, slide, { textLayout: 'svg' });
    const id = paintId(svg, 'linearGradient');
    const gradient = new RegExp(`<linearGradient id="${id}"[^>]*>.*?</linearGradient>`).exec(
      svg,
    )![0];
    // The block's user-space box, not each line's own bounding box.
    expect(gradient).toContain('gradientUnits="userSpaceOnUse"');
    expect(gradient).toMatch(/gradientTransform="matrix\([\d.]+ 0\.0000 0\.0000 [\d.]+ /);
    // `ang="5400000"` = 90°: top to bottom.
    expect(gradient).toContain('x1="0.5000" y1="0.0000" x2="0.5000" y2="1.0000"');
    expect(gradient).toContain('<stop offset="0.2100" stop-color="#53575C"/>');
    expect(gradient).toContain('<stop offset="0.8800" stop-color="#C5C7CA"/>');
    expect(svg).toMatch(painted(id, 'Outline title'));
  });

  it('spans the gradient over every line of the body, not each line', () => {
    const { pres, slide } = deck(
      nativeFormat('text-art-gray-gradient-shape.xml'),
      'First line\nSecond line',
    );
    const svg = renderSlideToSvg(pres, slide, { textLayout: 'svg' });
    // One paint server per run (here, per paragraph), all mapped to one box.
    const transforms = [...svg.matchAll(/gradientTransform="([^"]+)"/g)].map((m) => m[1]);
    expect(transforms.length).toBeGreaterThan(0);
    expect(new Set(transforms).size).toBe(1);
    const ids = [...svg.matchAll(/<linearGradient id="([^"]+)"/g)].map((m) => m[1]!);
    expect(ids.some((id) => painted(id, 'First line').test(svg))).toBe(true);
    expect(ids.some((id) => painted(id, 'Second line').test(svg))).toBe(true);
    const [, h] = /gradientTransform="matrix\([\d.]+ 0\.0000 0\.0000 ([\d.]+) /.exec(svg)!;
    // Two 40pt lines, not one.
    expect(Number(h)).toBeGreaterThan(40 * (96 / 72) * 1.5);
  });

  it('resolves scheme stops and their transforms through the theme', () => {
    const { pres, slide } = deck(nativeFormat('text-art-accent5-gradient-reflection-shape.xml'));
    const svg = renderSlideToSvg(pres, slide, { textLayout: 'svg' });
    const gradient = /<linearGradient id="[^"]+"[^>]*>.*?<\/linearGradient>/.exec(svg)![0];
    expect(gradient).not.toContain('scheme');
    expect(gradient.match(/stop-color="#[0-9A-F]{6}"/gi)?.length).toBeGreaterThan(1);
    // The reflection copies the same gradient paint.
    const id = paintId(svg, 'linearGradient');
    expect(svg).toMatch(new RegExp(`data-pptx-reflection="text"[^>]*>.*?fill="url\\(#${id}\\)"`));
  });

  it('paints a pattern preset with its resolved colors', () => {
    const { pres, slide } = deck(nativeFormat('text-art-white-pattern-shadow-shape.xml'));
    const svg = renderSlideToSvg(pres, slide, { textLayout: 'svg' });
    const id = paintId(svg, 'pattern');
    const pattern = new RegExp(`<pattern id="${id}"[^>]*>.*?</pattern>`).exec(svg)![0];
    // bg1 lumMod 50% = #808080 strokes; tx1 lumMod 75% lumOff 25% = #404040 ground.
    expect(pattern).toContain('fill="#404040"');
    expect(pattern).toContain('stroke="#808080" stroke-width="2"');
    expect(svg).toMatch(painted(id, 'Outline title'));
  });

  it.each([
    'text-art-accent1-pattern-hard-shadow-shape.xml',
    'text-art-accent3-pattern-inner-shadow-shape.xml',
    'text-art-accent5-pattern-outline-shape.xml',
    'text-art-dark-blue-pattern-hard-shadow-shape.xml',
  ])('paints %s with a pattern', (name) => {
    const { pres, slide } = deck(nativeFormat(name));
    const svg = renderSlideToSvg(pres, slide, { textLayout: 'svg' });
    const id = paintId(svg, 'pattern');
    expect(svg).toMatch(painted(id, 'Outline title'));
  });

  it('leaves solid runs untouched', () => {
    const { pres, slide } = deck({ color: '#336699' });
    const svg = renderSlideToSvg(pres, slide, { textLayout: 'svg' });
    expect(svg).toMatch(/fill="#336699"[^>]*>Outline title<\/tspan>/);
    expect(svg).not.toContain('<linearGradient');
    expect(svg).not.toContain('<pattern');
    const html = renderSlideToSvg(pres, slide);
    expect(html).toContain('color:#336699');
    expect(html).not.toContain('<text');
  });

  it('draws browser-layout glyphs in an SVG layer over transparent HTML', () => {
    const { pres, slide } = deck(nativeFormat('text-art-gray-gradient-shape.xml'));
    const svg = renderSlideToSvg(pres, slide);
    // The editable HTML keeps the glyphs for layout and selection only.
    expect(svg).toMatch(/<span style="[^"]*color:transparent[^"]*">Outline title<\/span>/);
    const after = svg.slice(svg.indexOf('</foreignObject>'));
    const id = paintId(after, 'linearGradient');
    expect(after).toMatch(painted(id, 'Outline title'));
  });

  it('includes the fill layer in the editor text-effects overlay', () => {
    const { pres, slide, shape } = deck(nativeFormat('text-art-white-pattern-shadow-shape.xml'));
    const overlay = renderTextEffectsSvg(pres, slide, shape, { w: inches(6), h: inches(2) });
    const id = paintId(overlay, 'pattern');
    expect(overlay).toMatch(painted(id, 'Outline title'));
    // The run's shadow travels with the painted glyphs.
    expect(overlay).toContain('<filter id="text-outer-shadow-');
  });

  it('paints table cell runs too', () => {
    const pres = createPresentation();
    const slide = addBlankSlide(pres);
    const table = addSlideTable(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(6),
      h: inches(1),
      rows: [['Cell', 'Plain']],
    });
    setTableCellTextFormat(
      getTableCells(table)[0]![0]!,
      nativeFormat('text-art-gray-gradient-shape.xml'),
    );
    for (const textLayout of ['svg', 'foreignObject'] as const) {
      const svg = renderSlideToSvg(pres, slide, { textLayout });
      expect(svg).toMatch(painted(paintId(svg, 'linearGradient'), 'Cell'));
      expect(svg).not.toMatch(/fill="url\([^)]*\)"[^>]*>Plain</);
    }
  });
});
