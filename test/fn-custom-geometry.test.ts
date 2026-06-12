// Custom geometry (`<a:custGeom>`, ECMA-376 §20.1.9): the guide-formula
// evaluator, the public `getShapeCustomGeometry` reader, and round-trip
// safety of a custGeom shape through load → save.
//
// There is no public API to author custGeom, so the reader/round-trip
// tests inject the geometry at the OPC zip layer (the same internal hook
// the chart-fallback renderer test uses): build a rect shape, then swap
// its `<a:prstGeom>` for a `<a:custGeom>` in the saved slide XML.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  addSlideShape,
  findSlideLayout,
  getShapeCustomGeometry,
  getShapeXmlString,
  getSlideShapes,
  getSlides,
  inches,
  loadPresentation,
  savePresentation,
  type SlideShapeData,
} from '../src/api/index.ts';
import { readZip, writeZip } from '../src/internal/opc/index.ts';
import { NS, parseXml } from '../src/internal/xml/index.ts';
import { parseCustomGeometry } from '../src/internal/drawingml/index.ts';

// ---------------------------------------------------------------------------
// Guide-formula evaluator (black-box: a guide's value is read back through
// the x coordinate of a single moveTo point).

const A_NS = ` xmlns:a="${NS.dml}"`;

/** Builds a `<a:custGeom>` element from gdLst + a single moveTo whose x is `xRef`. */
const custGeomXml = (gdLst: string, xRef: string, opts?: { avLst?: string }): string =>
  `<a:custGeom${A_NS}>` +
  `${opts?.avLst ?? ''}` +
  `<a:gdLst>${gdLst}</a:gdLst>` +
  `<a:pathLst><a:path w="100" h="100">` +
  `<a:moveTo><a:pt x="${xRef}" y="0"/></a:moveTo>` +
  `</a:path></a:pathLst></a:custGeom>`;

/** Parses a custGeom fragment and returns the resolved x of its first moveTo. */
const evalGuide = (gdLst: string, xRef: string, w = 100, h = 200): number => {
  const root = parseXml(custGeomXml(gdLst, xRef)).root;
  const geom = parseCustomGeometry(root, w, h);
  if (geom === null) throw new Error('geometry failed to evaluate');
  const cmd = geom.paths[0]!.commands[0]!;
  if (cmd.kind !== 'moveTo') throw new Error('expected moveTo');
  return cmd.pt.x;
};

const gd = (name: string, fmla: string): string => `<a:gd name="${name}" fmla="${fmla}"/>`;

describe('custGeom guide evaluator: fmla operators (ECMA-376 §20.1.9.11)', () => {
  it('val — literal', () => {
    expect(evalGuide(gd('g', 'val 1234'), 'g')).toBe(1234);
  });

  it('*/ — (a*b)/c', () => {
    expect(evalGuide(gd('g', '*/ 10 20 5'), 'g')).toBe(40);
  });

  it('+- — a+b-c', () => {
    expect(evalGuide(gd('g', '+- 10 20 5'), 'g')).toBe(25);
  });

  it('+/ — (a+b)/c', () => {
    expect(evalGuide(gd('g', '+/ 10 20 6'), 'g')).toBe(5);
  });

  it('?: — positive picks the second arg, non-positive the third', () => {
    expect(evalGuide(gd('g', '?: 1 7 9'), 'g')).toBe(7);
    expect(evalGuide(gd('g', '?: 0 7 9'), 'g')).toBe(9);
    expect(evalGuide(gd('g', '?: -3 7 9'), 'g')).toBe(9);
  });

  it('abs / sqrt / max / min', () => {
    expect(evalGuide(gd('g', 'abs -42'), 'g')).toBe(42);
    expect(evalGuide(gd('g', 'sqrt 144'), 'g')).toBe(12);
    expect(evalGuide(gd('g', 'max 3 8'), 'g')).toBe(8);
    expect(evalGuide(gd('g', 'min 3 8'), 'g')).toBe(3);
  });

  it('mod — magnitude of the 3-vector', () => {
    expect(evalGuide(gd('g', 'mod 3 4 0'), 'g')).toBe(5);
  });

  it('pin — clamps the middle value into [lo, hi]', () => {
    expect(evalGuide(gd('g', 'pin 10 5 20'), 'g')).toBe(10);
    expect(evalGuide(gd('g', 'pin 10 15 20'), 'g')).toBe(15);
    expect(evalGuide(gd('g', 'pin 10 25 20'), 'g')).toBe(20);
  });

  it('sin / cos / tan — angle arg in 60000ths of a degree', () => {
    // 90° = 90 × 60000 = 5_400_000.
    expect(evalGuide(gd('g', 'sin 1000 5400000'), 'g')).toBeCloseTo(1000, 6);
    // 180° = 10_800_000 → cos = -1.
    expect(evalGuide(gd('g', 'cos 1000 10800000'), 'g')).toBeCloseTo(-1000, 6);
    // 45° = 2_700_000 → tan = 1.
    expect(evalGuide(gd('g', 'tan 1000 2700000'), 'g')).toBeCloseTo(1000, 6);
  });

  it('at2 — arctan(y/x) returned in 60000ths of a degree', () => {
    // atan2(1, 1) = 45° → 2_700_000.
    expect(evalGuide(gd('g', 'at2 1 1'), 'g')).toBeCloseTo(2_700_000, 3);
  });

  it('cat2 / sat2 — a · cos|sin(arctan(c/b))', () => {
    // atan2(0, 1) = 0 → cos = 1, sin = 0.
    expect(evalGuide(gd('g', 'cat2 100 1 0'), 'g')).toBeCloseTo(100, 6);
    expect(evalGuide(gd('g', 'sat2 100 1 0'), 'g')).toBeCloseTo(0, 6);
  });

  it('division by zero yields 0 (spec-defined result)', () => {
    expect(evalGuide(gd('g', '*/ 10 20 0'), 'g')).toBe(0);
    expect(evalGuide(gd('g', '+/ 10 20 0'), 'g')).toBe(0);
  });

  it('forward references — a guide may read an earlier guide', () => {
    const gds = gd('a', 'val 100') + gd('b', '*/ a 2 1');
    expect(evalGuide(gds, 'b')).toBe(200);
  });

  it('built-in guides derive from the shape extents (w=100, h=200)', () => {
    expect(evalGuide('', 'w')).toBe(100);
    expect(evalGuide('', 'h')).toBe(200);
    expect(evalGuide('', 'ss')).toBe(100); // min(w, h)
    expect(evalGuide('', 'ls')).toBe(200); // max(w, h)
    expect(evalGuide('', 'hc')).toBe(50);
    expect(evalGuide('', 'vc')).toBe(100);
    expect(evalGuide('', 'cd4')).toBe(5_400_000); // 90° constant
  });

  it('avLst values can shadow a built-in guide', () => {
    const root = parseXml(
      custGeomXml(gd('g', '*/ w 1 1'), 'g', { avLst: `<a:avLst>${gd('w', 'val 7')}</a:avLst>` }),
    ).root;
    const geom = parseCustomGeometry(root, 100, 200)!;
    const cmd = geom.paths[0]!.commands[0]!;
    expect(cmd.kind === 'moveTo' && cmd.pt.x).toBe(7);
  });

  it('an unresolved guide reference makes the whole geometry null', () => {
    const root = parseXml(custGeomXml('', 'doesNotExist')).root;
    expect(parseCustomGeometry(root, 100, 200)).toBeNull();
  });

  it('an unknown fmla operator makes the whole geometry null', () => {
    const root = parseXml(custGeomXml(gd('g', 'bogus 1 2'), 'g')).root;
    expect(parseCustomGeometry(root, 100, 200)).toBeNull();
  });
});

// ---------------------------------------------------------------------------
// Path parsing — command kinds, fill/stroke attributes.

describe('custGeom path parsing', () => {
  const parse = (inner: string, w = 100, h = 100) =>
    parseCustomGeometry(parseXml(`<a:custGeom${A_NS}>${inner}</a:custGeom>`).root, w, h);

  it('triangle — moveTo / lnTo / close', () => {
    const geom = parse(
      `<a:pathLst><a:path w="100" h="100">` +
        `<a:moveTo><a:pt x="0" y="0"/></a:moveTo>` +
        `<a:lnTo><a:pt x="100" y="0"/></a:lnTo>` +
        `<a:lnTo><a:pt x="50" y="100"/></a:lnTo>` +
        `<a:close/></a:path></a:pathLst>`,
    )!;
    expect(geom.paths).toHaveLength(1);
    expect(geom.paths[0]!.commands.map((c) => c.kind)).toEqual(['moveTo', 'lnTo', 'lnTo', 'close']);
  });

  it('cubicBezTo — three control points, fully evaluated', () => {
    const geom = parse(
      `<a:pathLst><a:path>` +
        `<a:moveTo><a:pt x="0" y="0"/></a:moveTo>` +
        `<a:cubicBezTo>` +
        `<a:pt x="10" y="20"/><a:pt x="30" y="40"/><a:pt x="50" y="60"/>` +
        `</a:cubicBezTo></a:path></a:pathLst>`,
    )!;
    const c = geom.paths[0]!.commands[1]!;
    expect(c.kind).toBe('cubicBezTo');
    if (c.kind === 'cubicBezTo') {
      expect(c.pts).toEqual([
        { x: 10, y: 20 },
        { x: 30, y: 40 },
        { x: 50, y: 60 },
      ]);
    }
  });

  it('arcTo — radii and angles captured (no point list)', () => {
    const geom = parse(
      `<a:pathLst><a:path>` +
        `<a:moveTo><a:pt x="0" y="0"/></a:moveTo>` +
        `<a:arcTo wR="50" hR="25" stAng="0" swAng="5400000"/>` +
        `</a:path></a:pathLst>`,
    )!;
    const c = geom.paths[0]!.commands[1]!;
    expect(c).toEqual({ kind: 'arcTo', wR: 50, hR: 25, stAng: 0, swAng: 5_400_000 });
  });

  it('point coordinates may reference guides', () => {
    const geom = parse(
      `<a:gdLst><a:gd name="cx" fmla="*/ w 1 2"/></a:gdLst>` +
        `<a:pathLst><a:path><a:moveTo><a:pt x="cx" y="0"/></a:moveTo></a:path></a:pathLst>`,
      80,
      80,
    )!;
    const c = geom.paths[0]!.commands[0]!;
    expect(c.kind === 'moveTo' && c.pt.x).toBe(40);
  });

  it('multiple paths with fill="none" / stroke="0" carry the right flags', () => {
    const geom = parse(
      `<a:pathLst>` +
        `<a:path fill="none"><a:moveTo><a:pt x="0" y="0"/></a:moveTo></a:path>` +
        `<a:path stroke="0"><a:moveTo><a:pt x="0" y="0"/></a:moveTo></a:path>` +
        `<a:path fill="darken"><a:moveTo><a:pt x="0" y="0"/></a:moveTo></a:path>` +
        `</a:pathLst>`,
    )!;
    expect(geom.paths).toHaveLength(3);
    expect(geom.paths[0]!.fill).toBe('none');
    expect(geom.paths[0]!.stroke).toBe(true);
    expect(geom.paths[1]!.fill).toBe('norm');
    expect(geom.paths[1]!.stroke).toBe(false);
    expect(geom.paths[2]!.fill).toBe('darken');
  });
});

// ---------------------------------------------------------------------------
// Public reader + round-trip through a real deck.

const fixturePath = fileURLToPath(new URL('./fixtures/minimal/blank.pptx', import.meta.url));

const TRIANGLE_CUST_GEOM =
  '<a:custGeom><a:avLst/><a:gdLst/><a:pathLst><a:path w="100" h="100">' +
  '<a:moveTo><a:pt x="0" y="0"/></a:moveTo>' +
  '<a:lnTo><a:pt x="100" y="0"/></a:lnTo>' +
  '<a:lnTo><a:pt x="50" y="100"/></a:lnTo>' +
  '<a:close/></a:path></a:pathLst></a:custGeom>';

/**
 * Loads blank.pptx, adds a rect shape, then rewrites that shape's prstGeom
 * to `custGeom` at the zip layer. Returns the custom-geometry shape (the
 * Blank layout contributes its own placeholder shapes, so we pick the one
 * whose XML actually carries the injected `custGeom`).
 */
const custGeomShapeOf = (pres: Awaited<ReturnType<typeof loadPresentation>>): SlideShapeData => {
  const shape = getSlideShapes(getSlides(pres)[0]!).find((s) =>
    getShapeXmlString(s).includes('custGeom'),
  );
  if (!shape) throw new Error('injected custGeom shape not found');
  return shape;
};

const deckWithCustGeom = async (custGeom: string) => {
  const pres = await loadPresentation(await readFile(fixturePath));
  const layout = findSlideLayout(pres, 'Blank');
  if (!layout) throw new Error('Blank layout missing');
  const slide = addSlide(pres, { layout });
  addSlideShape(slide, { preset: 'rect', x: inches(1), y: inches(1), w: inches(2), h: inches(2) });
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
  return { pres: reloaded, shape: custGeomShapeOf(reloaded) };
};

describe('getShapeCustomGeometry (public reader)', () => {
  it('returns null for a preset-geometry shape', async () => {
    const pres = await loadPresentation(await readFile(fixturePath));
    const layout = findSlideLayout(pres, 'Blank')!;
    const slide = addSlide(pres, { layout });
    const shape = addSlideShape(slide, {
      preset: 'rect',
      x: inches(1),
      y: inches(1),
      w: inches(2),
      h: inches(2),
    });
    expect(getShapeCustomGeometry(shape)).toBeNull();
  });

  it('parses a triangle custGeom', async () => {
    const { shape } = await deckWithCustGeom(TRIANGLE_CUST_GEOM);
    const geom = getShapeCustomGeometry(shape);
    expect(geom).not.toBeNull();
    expect(geom!.paths).toHaveLength(1);
    expect(geom!.paths[0]!.commands.map((c) => c.kind)).toEqual([
      'moveTo',
      'lnTo',
      'lnTo',
      'close',
    ]);
  });

  it('returns null for a malformed custGeom (unresolved guide)', async () => {
    const broken =
      '<a:custGeom><a:pathLst><a:path w="100" h="100">' +
      '<a:moveTo><a:pt x="nope" y="0"/></a:moveTo></a:path></a:pathLst></a:custGeom>';
    const { shape } = await deckWithCustGeom(broken);
    expect(getShapeCustomGeometry(shape)).toBeNull();
  });

  it('round-trips a custGeom shape through load → save (semantically equal)', async () => {
    const { pres, shape } = await deckWithCustGeom(TRIANGLE_CUST_GEOM);
    const before = getShapeCustomGeometry(shape);
    const reloaded = await loadPresentation(await savePresentation(pres));
    const after = getShapeCustomGeometry(custGeomShapeOf(reloaded));
    expect(after).toEqual(before);
    expect(after).not.toBeNull();
  });
});
