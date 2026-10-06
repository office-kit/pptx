// The identifiers an animation player needs in the rendered SVG:
// `data-pptx-shape-id` on every object the slide's timing can target, and
// `data-pptx-paragraph` on each paragraph of a text body.
//
// Both text paths are covered — the HTML `<foreignObject>` one the browser
// uses and the pure-SVG one the server rasterises through — because a build
// has to be reachable in whichever the caller asked for. The assertions also
// pin down what the markers must NOT do: change where anything is drawn, and
// name a layout or master shape whose id means nothing to this slide.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  _internalPackageOf,
  addSlide,
  addSlideImage,
  addSlideShape,
  addSlideTable,
  addSlideTextBox,
  findSlideLayout,
  getGroupChildren,
  getShapeId,
  getShapeKind,
  getShapeName,
  getSlideLayout,
  getSlideLayoutShapes,
  getSlideShapes,
  getSlides,
  groupShapes,
  inches,
  loadPresentation,
  savePresentation,
  setShapeText,
} from '../src/api/index.ts';
import { partName } from '../src/internal/opc/index.ts';
import { type XmlElement, parseXml } from '../src/internal/xml/index.ts';
import { renderSlideToSvg } from '../packages/preview/src/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const blankSlide = async () => {
  const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
  const layout = findSlideLayout(pres, 'Blank');
  if (!layout) throw new Error('Blank layout not found');
  return { pres, slide: addSlide(pres, { layout }) };
};

const box = { x: inches(1), y: inches(1), w: inches(3), h: inches(2) };

const PNG = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
  0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
  0x89, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x63, 0x00, 0x01, 0x00, 0x00,
  0x05, 0x00, 0x01, 0x0d, 0x0a, 0x2d, 0xb4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae,
  0x42, 0x60, 0x82,
]);

/** Every shape of the slide, groups and their children alike, in document order. */
const everyShapeId = (slide: Parameters<typeof getSlideShapes>[0]): string[] => {
  const out: string[] = [];
  const walk = (shapes: readonly Parameters<typeof getShapeId>[0][]): void => {
    for (const shape of shapes) {
      out.push(String(getShapeId(shape)));
      if (getShapeKind(shape) === 'group') walk(getGroupChildren(shape));
    }
  };
  walk(getSlideShapes(slide));
  return out;
};

const shapeIds = (svg: string): string[] =>
  [...svg.matchAll(/data-pptx-shape-id="(\d+)"/g)].map((m) => m[1]!);

const paragraphIndexes = (svg: string): string[] =>
  [...svg.matchAll(/data-pptx-paragraph="(\d+)"/g)].map((m) => m[1]!);

/** Every `<g>` tag that carries a marker, exactly as it was written. */
const markerTags = (svg: string): string[] =>
  svg.match(/<g\b[^>]*data-pptx-(?:shape-id|paragraph)[^>]*>/g) ?? [];

const attr = (el: XmlElement, name: string): string | null =>
  el.attrs.find((a) => a.name.localName === name)?.value ?? null;

/**
 * The markers enclosing the first element that satisfies `found`, outermost
 * first, read off the parsed document rather than from where the strings
 * happen to fall — nesting is the thing being asserted, and string order does
 * not tell containment from mere sequence.
 */
const markersAround = (
  svg: string,
  found: (el: XmlElement) => boolean,
): readonly string[] | null => {
  const walk = (el: XmlElement, acc: readonly string[]): readonly string[] | null => {
    const shapeId = attr(el, 'data-pptx-shape-id');
    const paragraph = attr(el, 'data-pptx-paragraph');
    const here =
      shapeId !== null ? `shape:${shapeId}` : paragraph !== null ? `para:${paragraph}` : null;
    const chain = here === null ? acc : [...acc, here];
    if (found(el)) return chain;
    for (const child of el.children) {
      if (child.kind !== 'element') continue;
      const hit = walk(child, chain);
      if (hit !== null) return hit;
    }
    return null;
  };
  return walk(parseXml(svg).root, []);
};

const withShapeId = (id: number) => (el: XmlElement) =>
  attr(el, 'data-pptx-shape-id') === String(id);

const withName = (name: string) => (el: XmlElement) => attr(el, 'data-pptx-shape-name') === name;

describe('preview: data-pptx-shape-id', () => {
  it('names every kind of object on the slide, once each', async () => {
    const { pres, slide } = await blankSlide();
    const rect = addSlideShape(slide, { preset: 'rect', ...box });
    const picture = addSlideImage(slide, PNG, box);
    const table = addSlideTable(slide, {
      ...box,
      rows: [
        ['A', 'B'],
        ['C', 'D'],
      ],
    });
    const text = addSlideTextBox(slide, { ...box, text: 'One' });

    const ids = shapeIds(renderSlideToSvg(pres, slide));
    for (const shape of [rect, picture, table, text]) {
      expect(ids).toContain(String(getShapeId(shape)));
    }
    // Every shape the slide has, named once each and in document order — the
    // layout's placeholders this slide inherited included, since the timing
    // can target those too.
    expect(ids).toEqual(everyShapeId(slide));
  });

  it('names a group and the shapes inside it, nesting included', async () => {
    const { pres, slide } = await blankSlide();
    const inner1 = addSlideShape(slide, { preset: 'rect', ...box });
    const inner2 = addSlideShape(slide, { preset: 'ellipse', ...box });
    const inner = groupShapes([inner1, inner2]);
    const outerMate = addSlideShape(slide, { preset: 'triangle', ...box });
    const outer = groupShapes([inner, outerMate]);

    const svg = renderSlideToSvg(pres, slide);
    const ids = shapeIds(svg);
    // Every shape of the nest is reachable, each named once.
    expect(new Set(ids).size).toBe(ids.length);
    for (const shape of [outer, inner, inner1, inner2, outerMate]) {
      expect(ids).toContain(String(getShapeId(shape)));
    }
    // A group's marker encloses its children's, so hiding the group hides
    // them — read off the parsed tree, not from where the strings fall.
    expect(markersAround(svg, withShapeId(getShapeId(inner1)))).toEqual([
      `shape:${getShapeId(outer)}`,
      `shape:${getShapeId(inner)}`,
      `shape:${getShapeId(inner1)}`,
    ]);
    expect(markersAround(svg, withShapeId(getShapeId(outerMate)))).toEqual([
      `shape:${getShapeId(outer)}`,
      `shape:${getShapeId(outerMate)}`,
    ]);
  });

  it('leaves the layout and master decoration unnamed', async () => {
    // This template's layout carries a logo and a bar the renderer draws into
    // the slide's background — real shapes, not placeholders, so they are not
    // dropped on the way.
    const pres = await loadPresentation(await readFile(fixture('layout-decoration.pptx')));
    const slide = getSlides(pres)[0]!;
    const layout = getSlideLayout(slide);
    if (!layout) throw new Error('no layout');
    const decoration = getSlideLayoutShapes(pres, layout).filter((s) =>
      ['Logo', 'Template Bar'].includes(getShapeName(s) ?? ''),
    );
    expect(decoration).toHaveLength(2);

    const svg = renderSlideToSvg(pres, slide);
    expect(svg).toContain('data-pptx-shape-name="Template Bar"');
    // Drawn, but not named: a layout shape's id is its own, and a
    // `<p:spTgt spid>` on this slide means a shape of this slide. One marker
    // per shape of the slide and not one more, although the background put
    // several more shapes in the drawing.
    expect(shapeIds(svg)).toEqual(everyShapeId(slide));
    // The bar is in the drawing with no marker above it at all.
    expect(markersAround(svg, withName('Template Bar'))).toEqual([]);
  });

  it('marks only the slide shape when the layout carries the same id', async () => {
    // The layout's bar is cNvPr id 101. Giving a shape of the slide that same
    // number is what the renderer must not be confused by: both are drawn,
    // and only the slide's answers to the id its timing would name.
    const pres = await loadPresentation(await readFile(fixture('layout-decoration.pptx')));
    const slide = getSlides(pres)[0]!;
    const layout = getSlideLayout(slide)!;
    const bar = getSlideLayoutShapes(pres, layout).find((s) => getShapeName(s) === 'Template Bar')!;
    const barId = getShapeId(bar);
    const title = getSlideShapes(slide)[0]!;

    const part = _internalPackageOf(pres).getPart(partName('/ppt/slides/slide1.xml'))!;
    part.data = new TextEncoder().encode(
      new TextDecoder()
        .decode(part.data)
        .replace(`<p:cNvPr id="${getShapeId(title)}"`, `<p:cNvPr id="${barId}"`),
    );
    const reloaded = await loadPresentation(await savePresentation(pres));
    const only = getSlides(reloaded)[0]!;

    const svg = renderSlideToSvg(reloaded, only);
    expect(svg).toContain('data-pptx-shape-name="Template Bar"');
    expect(shapeIds(svg).filter((id) => id === String(barId))).toHaveLength(1);
    // And the one marker is around the slide's shape, not the layout's bar.
    expect(markersAround(svg, withName('Template Bar'))).toEqual([]);
    expect(markersAround(svg, withShapeId(barId))).toEqual([`shape:${barId}`]);
  });

  it('repeats an id a deck repeats, rather than inventing a unique one', async () => {
    const { pres, slide } = await blankSlide();
    const a = addSlideShape(slide, { preset: 'rect', ...box });
    const b = addSlideShape(slide, { preset: 'ellipse', ...box });
    // Nothing in the schema makes `<p:cNvPr id>` unique. The renderer marks
    // what the deck says rather than inventing a number no timing could name;
    // what a player should do when one id answers for two shapes is not
    // settled here, and this library does not claim to support it. The id is
    // rewritten in the saved XML because the authoring API keeps them apart.
    const slideIndex = getSlides(pres).length;
    const part = _internalPackageOf(pres).getPart(partName(`/ppt/slides/slide${slideIndex}.xml`))!;
    const xml = new TextDecoder()
      .decode(part.data)
      .replace(`<p:cNvPr id="${getShapeId(b)}"`, `<p:cNvPr id="${getShapeId(a)}"`);
    part.data = new TextEncoder().encode(xml);
    const reloaded = await loadPresentation(await savePresentation(pres));

    const only = getSlides(reloaded).at(-1)!;
    const ids = shapeIds(renderSlideToSvg(reloaded, only));
    expect(ids).toEqual(everyShapeId(only));
    // Both shapes carry the marker, because both carry the id.
    expect(ids.filter((id) => id === String(getShapeId(a)))).toHaveLength(2);
  });

  it('adds nothing to the drawing but the markers themselves', async () => {
    const { pres, slide } = await blankSlide();
    addSlideShape(slide, { preset: 'rect', ...box });
    addSlideTextBox(slide, { ...box, text: 'One\nTwo' });
    addSlideTable(slide, {
      ...box,
      rows: [
        ['A', 'B'],
        ['C', 'D'],
      ],
    });

    for (const textLayout of ['foreignObject', 'svg'] as const) {
      const svg = renderSlideToSvg(pres, slide, { textLayout });
      // A marker wrapper carries the marker and nothing else — no transform,
      // no style, no opacity — so it cannot move or repaint what it wraps.
      expect(markerTags(svg)).toEqual(
        svg.match(/<g data-pptx-(?:shape-id|paragraph)="\d+">/g) ?? [],
      );
      expect(markerTags(svg).length).toBeGreaterThan(0);
      // And the document still parses, so every wrapper is closed where it
      // opened. `<foreignObject>` holds XHTML, which parses as XML too.
      expect(() => parseXml(svg)).not.toThrow();
    }
  });
});

describe('preview: data-pptx-paragraph', () => {
  it('numbers the paragraphs of a text body in the HTML text path', async () => {
    const { pres, slide } = await blankSlide();
    addSlideTextBox(slide, { ...box, text: 'One\nTwo\nThree' });
    const svg = renderSlideToSvg(pres, slide, { textLayout: 'foreignObject' });

    expect(paragraphIndexes(svg)).toEqual(['0', '1', '2']);
    // The marker sits on the paragraph element itself, next to its text.
    expect(svg).toMatch(/<p data-pptx-paragraph="1"[^>]*>(?:(?!<\/p>).)*Two/);
  });

  it('numbers the paragraphs of a text body in the SVG text path', async () => {
    const { pres, slide } = await blankSlide();
    addSlideTextBox(slide, { ...box, text: 'One\nTwo\nThree' });
    const svg = renderSlideToSvg(pres, slide, { textLayout: 'svg' });

    expect(paragraphIndexes(svg)).toEqual(['0', '1', '2']);
    expect(svg).toMatch(/<g data-pptx-paragraph="1">(?:(?!<\/g>).)*Two/);
  });

  it('keeps a wrapped paragraph’s lines under one marker', async () => {
    const { pres, slide } = await blankSlide();
    addSlideTextBox(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(1),
      h: inches(2),
      text: 'wrap wrap wrap wrap wrap\nsecond',
    });
    const svg = renderSlideToSvg(pres, slide, { textLayout: 'svg' });

    expect(paragraphIndexes(svg)).toEqual(['0', '1']);
    const first = svg.slice(
      svg.indexOf('<g data-pptx-paragraph="0">'),
      svg.indexOf('<g data-pptx-paragraph="1">'),
    );
    // The paragraph wrapped, so one marker covers several lines.
    expect((first.match(/<text\b/g) ?? []).length).toBeGreaterThan(1);
  });

  it('numbers the paragraphs of a shape inside a group', async () => {
    const { pres, slide } = await blankSlide();
    const text = addSlideTextBox(slide, { ...box, text: 'One\nTwo' });
    const other = addSlideShape(slide, { preset: 'rect', ...box });
    const group = groupShapes([text, other]);

    for (const textLayout of ['foreignObject', 'svg'] as const) {
      const svg = renderSlideToSvg(pres, slide, { textLayout });
      expect(paragraphIndexes(svg)).toEqual(['0', '1']);
      // The text's own marker is inside the group's, and its paragraphs
      // inside that one.
      expect(markersAround(svg, (el) => attr(el, 'data-pptx-paragraph') === '1')).toEqual([
        `shape:${getShapeId(group)}`,
        `shape:${getShapeId(text)}`,
        'para:1',
      ]);
    }
  });

  it('numbers a table cell’s paragraphs without naming the cell as a shape', async () => {
    const { pres, slide } = await blankSlide();
    const table = addSlideTable(slide, { ...box, rows: [['A', 'B']] });
    const svg = renderSlideToSvg(pres, slide, { textLayout: 'svg' });

    // The table is one object to the timing; its cells are not shapes.
    expect(shapeIds(svg)).toEqual(everyShapeId(slide));
    expect(shapeIds(svg)).toContain(String(getShapeId(table)));
  });

  it('leaves a shape without text unmarked', async () => {
    const { pres, slide } = await blankSlide();
    addSlideShape(slide, { preset: 'rect', ...box });
    for (const textLayout of ['foreignObject', 'svg'] as const) {
      expect(paragraphIndexes(renderSlideToSvg(pres, slide, { textLayout }))).toEqual([]);
    }
  });

  it('survives save and reload with the same numbering', async () => {
    const { pres, slide } = await blankSlide();
    const text = addSlideTextBox(slide, { ...box, text: 'One' });
    setShapeText(text, 'One\nTwo\nThree');
    const before = renderSlideToSvg(pres, slide, { textLayout: 'svg' });

    const reloaded = await loadPresentation(await savePresentation(pres));
    const after = renderSlideToSvg(reloaded, getSlides(reloaded).at(-1)!, { textLayout: 'svg' });
    expect(paragraphIndexes(after)).toEqual(paragraphIndexes(before));
    expect(shapeIds(after)).toEqual(shapeIds(before));
  });
});
