// Format pane line and text paint the editor used to show disabled: gradient
// lines (shape outline and text outline), PowerPoint's Sketched style, text
// No fill / Picture fill, and Keep text flat.

import { describe, expect, it } from 'vitest';
import {
  addBlankSlide,
  addSlideShape,
  createPresentation,
  getShapeRunFormat,
  getShapeRunFormatEffective,
  getShapeStroke,
  getShapeStrokeEffective,
  getShapeStrokeGradient,
  getShapeStrokeSketch,
  getShapeTextFlat,
  getShapeText3D,
  getShapeXmlString,
  getSlides,
  getSlideShapes,
  getSlideXmlString,
  getTableCell,
  getTableCellRunFormatEffective,
  addSlideTable,
  inches,
  listPackageParts,
  loadPresentation,
  readPackagePart,
  savePresentation,
  setShapeParagraphs,
  setShapeRunFormat,
  setShapeStroke,
  setShapeStrokeDash,
  setShapeStrokeSketch,
  setShapeText,
  setShapeText3D,
  setShapeTextFlat,
  setShapeTextFormat,
  setTableCellText,
  setTableCellTextFormat,
  toWritableTextFormat,
  validatePresentation,
  type LineFill,
  type PresentationData,
  type SlideShapeData,
} from '../src/api/index.ts';
import { SHAPE_ELEMENT } from '../src/api/_internal-symbols.ts';
import { NS, parseFragment } from '../src/internal/xml/index.ts';
import { expectSchemaValid, isSchemaValidationAvailable } from './lib/expect-schema-valid.ts';

const PNG = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
  0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
  0x89, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x62, 0x00, 0x01, 0x00, 0x00,
  0x05, 0x00, 0x01, 0x0d, 0x0a, 0x2d, 0xb4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae,
  0x42, 0x60, 0x82,
]);

const GRADIENT: LineFill = {
  kind: 'gradient',
  path: 'linear',
  angleDeg: 90,
  scaled: true,
  stops: [
    { offset: 0, color: 'accent1', brightness: 0.95 },
    { offset: 1, color: '#FF0000', opacity: 0.5 },
  ],
};

const rect = () => {
  const pres = createPresentation();
  const slide = addBlankSlide(pres);
  const shape = addSlideShape(slide, {
    preset: 'rect',
    x: inches(1),
    y: inches(1),
    w: inches(2),
    h: inches(1),
  });
  return { pres, slide, shape };
};

const compact = (xml: string) => xml.replace(/>\s+</g, '><').replace(/\s+\/>/g, '/>');
const ln = (shape: SlideShapeData) =>
  compact(getShapeXmlString(shape)).match(/<a:ln\b[\s\S]*?<\/a:ln>|<a:ln\b[^>]*\/>/)![0];
const reload = async (pres: PresentationData) => {
  const reloaded = await loadPresentation(await savePresentation(pres));
  expect(validatePresentation(reloaded).filter((issue) => issue.severity === 'error')).toEqual([]);
  return getSlideShapes(getSlides(reloaded)[0]!)[0]!;
};

// The Office-written form, from a document Word 2021 saved: the sketched path
// replaces the geometry and the original preset stays inside the props.
const OFFICE_SKETCH =
  `<a:ln xmlns:a="${NS.dml}" w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill>` +
  `<a:extLst><a:ext uri="{FFFF0000-0000-0000-0000-000000000000}"><x:other xmlns:x="urn:example"/></a:ext>` +
  `<a:ext uri="{C807C97D-BFC1-408E-A445-0C87EB9F89A2}">` +
  `<ask:lineSketchStyleProps xmlns:ask="http://schemas.microsoft.com/office/drawing/2018/sketchyshapes" sd="1219033472">` +
  `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><ask:type><ask:lineSketchScribble/></ask:type>` +
  `</ask:lineSketchStyleProps></a:ext></a:extLst></a:ln>`;

const officeSketched = () => {
  const made = rect();
  const spPr = made.shape[SHAPE_ELEMENT].children.find(
    (child) => child.kind === 'element' && child.name.localName === 'spPr',
  );
  if (spPr?.kind !== 'element') throw new Error('no spPr');
  const geometry = spPr.children.findIndex(
    (child) => child.kind === 'element' && child.name.localName === 'prstGeom',
  );
  spPr.children.splice(
    geometry,
    1,
    parseFragment(
      `<a:custGeom xmlns:a="${NS.dml}"><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/>` +
        `<a:pathLst><a:path w="10" h="10"><a:moveTo><a:pt x="0" y="1"/></a:moveTo><a:lnTo><a:pt x="10" y="0"/></a:lnTo>` +
        `<a:lnTo><a:pt x="9" y="10"/></a:lnTo><a:close/></a:path></a:pathLst></a:custGeom>`,
    ),
  );
  spPr.children = spPr.children.filter(
    (child) => !(child.kind === 'element' && child.name.localName === 'ln'),
  );
  spPr.children.push(parseFragment(OFFICE_SKETCH));
  return made;
};

describe('gradient line', () => {
  it('writes a:gradFill in the line fill slot and reads it back', async () => {
    const { pres, shape } = rect();
    setShapeStrokeDash(shape, 'dash');
    setShapeStroke(shape, { fill: GRADIENT, widthEmu: 28575 });
    expect(ln(shape)).toBe(
      '<a:ln w="28575"><a:gradFill flip="none" rotWithShape="1"><a:gsLst>' +
        '<a:gs pos="0"><a:schemeClr val="accent1"><a:lumMod val="5000"/><a:lumOff val="95000"/></a:schemeClr></a:gs>' +
        '<a:gs pos="100000"><a:srgbClr val="FF0000"><a:alpha val="50000"/></a:srgbClr></a:gs>' +
        '</a:gsLst><a:lin ang="5400000" scaled="1"/></a:gradFill><a:prstDash val="dash"/></a:ln>',
    );
    expect(getShapeStroke(shape)).toEqual({ kind: 'gradient', widthEmu: 28575 });
    expect(getShapeStrokeEffective(pres, shape)).toEqual({ kind: 'gradient', widthEmu: 28575 });
    const read = getShapeStrokeGradient(shape)!;
    expect(read).toMatchObject({ angleDeg: 90, scaled: true, rotateWithShape: true });
    expect(read.stops.map((stop) => [stop.offset, stop.color, stop.opacity])).toEqual([
      [0, 'scheme:accent1', undefined],
      [1, '#FF0000', 0.5],
    ]);
    expect(getShapeStrokeGradient(shape, pres)!.stops[1]!.resolvedColor).toBe('#FF0000');
    expect(getShapeStrokeGradient(await reload(pres))).toEqual(read);
  });

  it('switches between a gradient and a solid line through the one fill choice', () => {
    const { shape } = rect();
    setShapeStroke(shape, { fill: GRADIENT });
    setShapeStroke(shape, { color: 'accent2' });
    expect(ln(shape)).toBe('<a:ln><a:solidFill><a:schemeClr val="accent2"/></a:solidFill></a:ln>');
    expect(getShapeStrokeGradient(shape)).toBeNull();
    setShapeStroke(shape, { fill: GRADIENT });
    expect(ln(shape)).not.toContain('solidFill');
  });

  it('rejects a second fill choice and invalid stops before changing the line', () => {
    const { shape } = rect();
    setShapeStroke(shape, { color: '#000000' });
    const before = ln(shape);
    expect(() => setShapeStroke(shape, { fill: GRADIENT, color: '#FF0000' })).toThrow(/exclusive/);
    expect(() => setShapeStroke(shape, { fill: GRADIENT, opacity: 0.5 })).toThrow(/exclusive/);
    expect(() =>
      setShapeStroke(shape, { fill: { ...GRADIENT, stops: [GRADIENT.stops[0]!] } }),
    ).toThrow(/two stops/);
    expect(ln(shape)).toBe(before);
  });

  it('writes, reads and round-trips a gradient text outline', async () => {
    const { pres, shape } = rect();
    setShapeText(shape, 'Outline');
    setShapeTextFormat(shape, { outline: { fill: GRADIENT, widthEmu: 12700 } });
    expect(compact(getShapeXmlString(shape))).toContain(
      '<a:rPr><a:ln w="12700"><a:gradFill flip="none" rotWithShape="1">',
    );
    const outline = getShapeRunFormat(shape, 0, 0)!.outline!;
    expect(outline.widthEmu).toBe(12700);
    expect(outline.fill).toMatchObject({ kind: 'gradient', angleDeg: 90, scaled: true });
    expect(outline.fill!.stops.map((stop) => stop.color)).toEqual(['accent1', '#FF0000']);
    const effective = getShapeRunFormatEffective(pres, shape, 0, 0).outline!;
    expect(effective.fill!.stops[0]!.resolvedColor).toMatch(/^#[0-9A-F]{6}$/);
    // A format read back writes the same outline again.
    const copy = rect();
    setShapeText(copy.shape, 'Copy');
    setShapeTextFormat(copy.shape, toWritableTextFormat(getShapeRunFormat(shape, 0, 0)!));
    expect(getShapeRunFormat(copy.shape, 0, 0)!.outline).toEqual(outline);
    expect(getShapeRunFormat(await reload(pres), 0, 0)!.outline).toEqual(outline);
    expect(() => setShapeTextFormat(shape, { outline: { fill: GRADIENT, color: 'tx1' } })).toThrow(
      /exclusive/,
    );
  });
});

describe('sketched line style', () => {
  it('writes the Office extension with a stable seed and reads the preset', async () => {
    const { pres, shape } = rect();
    setShapeStroke(shape, { color: '#000000' });
    setShapeStrokeSketch(shape, 'freehand');
    const written = ln(shape);
    expect(written).toMatch(
      /^<a:ln><a:solidFill><a:srgbClr val="000000"\/><\/a:solidFill><a:extLst><a:ext uri="\{C807C97D-BFC1-408E-A445-0C87EB9F89A2\}"><ask:lineSketchStyleProps xmlns:ask="http:\/\/schemas\.microsoft\.com\/office\/drawing\/2018\/sketchyshapes" sd="\d+"><ask:type><ask:lineSketchFreehand\/><\/ask:type><\/ask:lineSketchStyleProps><\/a:ext><\/a:extLst><\/a:ln>$/,
    );
    expect(getShapeStrokeSketch(shape)).toBe('freehand');
    // Changing the preset keeps the seed; the same deck writes the same seed.
    const seed = written.match(/sd="(\d+)"/)![1];
    setShapeStrokeSketch(shape, 'scribble');
    expect(ln(shape)).toContain(`sd="${seed}"><ask:type><ask:lineSketchScribble/>`);
    expect(ln(withSketch(rect().shape))).toContain(`sd="${seed}"`);
    const reloaded = await reload(pres);
    expect(getShapeStrokeSketch(reloaded)).toBe('scribble');
    setShapeStrokeSketch(shape, null);
    expect(ln(shape)).toBe('<a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>');
    expect(getShapeStrokeSketch(shape)).toBeNull();
    expect(() => setShapeStrokeSketch(shape, 'wobbly' as 'curved')).toThrow(/sketch/);
  });

  it('keeps other extensions and the dash order when the sketch is added after them', () => {
    const { shape } = rect();
    setShapeStrokeSketch(shape, 'curved');
    setShapeStrokeDash(shape, 'dash');
    expect(ln(shape)).toMatch(/^<a:ln><a:prstDash val="dash"\/><a:extLst>/);
  });

  it('restores the geometry Office kept and preserves unknown extensions', async () => {
    const { pres, shape } = officeSketched();
    expect(getShapeStrokeSketch(shape)).toBe('scribble');
    setShapeStrokeSketch(shape, 'curved');
    const xml = compact(getShapeXmlString(shape));
    expect(xml).toContain('<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:ln');
    expect(xml).not.toContain('custGeom');
    expect(xml).toContain(
      '<a:ext uri="{FFFF0000-0000-0000-0000-000000000000}"><x:other xmlns:x="urn:example"/></a:ext>',
    );
    expect(xml).toContain('sd="1219033472"><ask:type><ask:lineSketchCurved/></ask:type>');
    setShapeStrokeSketch(shape, null);
    expect(ln(shape).replace(` xmlns:a="${NS.dml}"`, '')).toBe(
      '<a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:extLst>' +
        '<a:ext uri="{FFFF0000-0000-0000-0000-000000000000}"><x:other xmlns:x="urn:example"/></a:ext></a:extLst></a:ln>',
    );
    expect(getShapeStrokeSketch(await reload(pres))).toBeNull();
  });

  it('removing a sketch from a shape without a line writes nothing', () => {
    const { shape } = rect();
    const before = getShapeXmlString(shape);
    setShapeStrokeSketch(shape, null);
    expect(getShapeXmlString(shape)).toBe(before);
  });
});

const withSketch = (shape: SlideShapeData) => {
  setShapeStrokeSketch(shape, 'curved');
  return shape;
};

describe('text No fill and Picture fill', () => {
  it('writes a:noFill and reads it as the none fill', async () => {
    const { pres, shape } = rect();
    setShapeText(shape, 'Hollow');
    setShapeTextFormat(shape, { textFill: { kind: 'none' }, outline: { color: 'tx1' } });
    expect(compact(getShapeXmlString(shape))).toContain('</a:ln><a:noFill/></a:rPr>');
    expect(getShapeRunFormat(shape, 0, 0)!.textFill).toEqual({ kind: 'none' });
    expect(getShapeRunFormatEffective(pres, shape, 0, 0)).toMatchObject({
      textFill: { kind: 'none' },
    });
    expect(getShapeRunFormatEffective(pres, shape, 0, 0).color).toBeUndefined();
    setShapeTextFormat(shape, { color: 'accent1' });
    expect(compact(getShapeXmlString(shape))).not.toContain('noFill');
    setShapeTextFormat(shape, { textFill: { kind: 'none' } });
    expect(getShapeRunFormat(await reload(pres), 0, 0)!.textFill).toEqual({ kind: 'none' });
  });

  it('embeds the picture once for every run it fills and reads its bytes back', async () => {
    const { pres, slide, shape } = rect();
    setShapeText(shape, '');
    setShapeParagraphs(shape, [
      { runs: [{ text: 'One' }, { text: 'Two', format: { bold: true } }] },
      { runs: [{ text: 'Three' }] },
    ]);
    const mediaBefore = listPackageParts(pres).filter((part) =>
      part.name.startsWith('/ppt/media/'),
    );
    setShapeTextFormat(shape, { textFill: { kind: 'image', bytes: PNG } });
    const media = listPackageParts(pres).filter((part) => part.name.startsWith('/ppt/media/'));
    expect(media.length).toBe(mediaBefore.length + 1);
    const xml = compact(getShapeXmlString(shape));
    const embeds = [
      ...xml.matchAll(
        /<a:blipFill rotWithShape="1"><a:blip r:embed="(rId\d+)"\/><a:stretch><a:fillRect\/><\/a:stretch><\/a:blipFill>/g,
      ),
    ];
    // Three runs and two paragraph end marks share one relationship.
    expect(embeds.length).toBe(5);
    expect(new Set(embeds.map((match) => match[1])).size).toBe(1);
    const read = getShapeRunFormatEffective(pres, shape, 0, 1).textFill;
    expect(read).toEqual({ kind: 'image', bytes: PNG });
    expect(getShapeRunFormat(shape, 1, 0)!.textFill).toEqual({ kind: 'image', bytes: PNG });
    expect(toWritableTextFormat(getShapeRunFormat(shape, 0, 0)!).textFill).toEqual({
      kind: 'image',
      bytes: PNG,
    });
    const reloaded = await reload(pres);
    expect(getShapeRunFormat(reloaded, 0, 0)!.textFill).toEqual({ kind: 'image', bytes: PNG });
    expect(readPackagePart(pres, media.at(-1)!.name)).toEqual(PNG);
    void slide;
  });

  it('fills one run and a table cell through their own editors', () => {
    const { pres, slide, shape } = rect();
    setShapeText(shape, 'Run');
    setShapeRunFormat(shape, 0, 0, { textFill: { kind: 'image', bytes: PNG } });
    expect(getShapeRunFormat(shape, 0, 0)!.textFill).toEqual({ kind: 'image', bytes: PNG });
    const table = addSlideTable(slide, {
      x: inches(1),
      y: inches(3),
      w: inches(4),
      h: inches(1),
      rows: [['']],
    });
    const cell = getTableCell(table, 0, 0);
    setTableCellText(cell, 'Cell');
    setTableCellTextFormat(cell, { textFill: { kind: 'image', bytes: PNG } });
    expect(getTableCellRunFormatEffective(pres, cell, 0, 0).textFill).toEqual({
      kind: 'image',
      bytes: PNG,
    });
  });

  it('rejects bytes that are not a raster picture before changing the text', () => {
    const { pres, shape } = rect();
    setShapeText(shape, 'Text');
    const before = getShapeXmlString(shape);
    const parts = listPackageParts(pres).length;
    const svg = new TextEncoder().encode('<svg xmlns="http://www.w3.org/2000/svg"/>');
    expect(() => setShapeTextFormat(shape, { textFill: { kind: 'image', bytes: svg } })).toThrow(
      /PNG, JPEG/,
    );
    expect(() =>
      setShapeParagraphs(shape, [
        { runs: [{ text: 'a', format: { textFill: { kind: 'image', bytes: PNG } } }] },
        { runs: [{ text: 'b', format: { underline: 'nope' } }] },
      ]),
    ).toThrow(/underline/);
    expect(getShapeXmlString(shape)).toBe(before);
    expect(listPackageParts(pres).length).toBe(parts);
  });
});

describe('keep text flat', () => {
  it('writes a:flatTx, replacing the text body 3-D choice', async () => {
    const { pres, shape } = rect();
    setShapeText(shape, 'Flat');
    setShapeText3D(shape, { bevelTop: { preset: 'circle' } });
    setShapeTextFlat(shape, true);
    expect(compact(getShapeXmlString(shape))).toMatch(
      /<a:bodyPr[^>]*>(?:(?!<\/a:bodyPr>).)*<a:flatTx\/><\/a:bodyPr>/,
    );
    expect(getShapeText3D(shape)).toBeNull();
    expect(getShapeTextFlat(shape)).toBe(true);
    expect(getShapeTextFlat(await reload(pres))).toBe(true);
    setShapeTextFlat(shape, false);
    expect(getShapeTextFlat(shape)).toBe(false);
    expect(getShapeXmlString(shape)).not.toContain('flatTx');
  });
});

describe.skipIf(!isSchemaValidationAvailable())('schema validity', () => {
  it('validates gradient lines, sketches, text fills and flat text against pml.xsd', () => {
    const { slide, shape } = rect();
    setShapeStroke(shape, { fill: GRADIENT, widthEmu: 19050 });
    setShapeStrokeDash(shape, 'sysDot');
    setShapeStrokeSketch(shape, 'scribble');
    setShapeText(shape, '');
    setShapeParagraphs(shape, [
      { runs: [{ text: 'a' }, { text: 'b' }] },
      { runs: [{ text: 'c' }] },
    ]);
    setShapeTextFormat(shape, { outline: { fill: GRADIENT } }, { range: { start: 0, end: 1 } });
    setShapeTextFormat(shape, { textFill: { kind: 'none' } }, { range: { start: 1, end: 2 } });
    setShapeTextFormat(shape, { textFill: { kind: 'image', bytes: PNG } }, { paragraphEnd: 1 });
    setShapeRunFormat(shape, 1, 0, { textFill: { kind: 'image', bytes: PNG } });
    setShapeTextFlat(shape, true);
    const office = officeSketched();
    setShapeStrokeSketch(office.shape, 'freehand');
    expectSchemaValid(getSlideXmlString(slide), 'pml');
    expectSchemaValid(getSlideXmlString(office.slide), 'pml');
  });
});
