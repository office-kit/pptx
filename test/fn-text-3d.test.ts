// Text-body 3-D — `<a:scene3d>` / `<a:sp3d>` inside `<a:bodyPr>`, where
// the reference desktop app writes its text-art bevels.

import { describe, expect, it } from 'vitest';
import {
  addBlankSlide,
  addSlideTextBox,
  createPresentation,
  getShapeText3D,
  getShapeXmlString,
  getSlides,
  getSlideShapes,
  getSlideXmlString,
  inches,
  loadPresentation,
  savePresentation,
  setShapeText3D,
  setShapeTextAutoFit,
  type Text3D,
} from '../src/api/index.ts';
import { SHAPE_ELEMENT } from '../src/api/_internal-symbols.ts';
import { NS, firstChildElement, parseFragment, qname } from '../src/internal/xml/index.ts';
import { expectSchemaValid, isSchemaValidationAvailable } from './lib/expect-schema-valid.ts';

const textBox = () => {
  const pres = createPresentation();
  const slide = addBlankSlide(pres);
  const shape = addSlideTextBox(slide, {
    x: inches(1),
    y: inches(1),
    w: inches(4),
    h: inches(1),
    text: 'Outline title',
  });
  return { pres, slide, shape };
};

const compact = (xml: string) => xml.replace(/>\s+</g, '><').replace(/\s+\/>/g, '/>');
const bodyPr = (xml: string) =>
  compact(xml).match(/<a:bodyPr\b[^>]*?(?:\/>|>[\s\S]*?<\/a:bodyPr>)/)![0];

// The reference desktop app's Sharp Bevel
// (test/fixtures/native/text-art-accent3-sharp-bevel-shape.xml).
const SHARP: Text3D = {
  scene: { camera: 'orthographicFront', lightRig: { type: 'harsh', direction: 't' } },
  bevelTop: { widthEmu: 63500, heightEmu: 12700, preset: 'angle' },
  extrusionHeightEmu: 57150,
  material: 'matte',
  contourColor: 'bg1',
  contourColorTransforms: [{ kind: 'lumMod', value: 0.65 }],
};

describe('setShapeText3D / getShapeText3D', () => {
  it('writes scene3d and sp3d and reads them back', async () => {
    const { pres, shape } = textBox();
    setShapeText3D(shape, SHARP);
    expect(bodyPr(getShapeXmlString(shape))).toContain(
      '<a:scene3d><a:camera prst="orthographicFront"/><a:lightRig rig="harsh" dir="t"/></a:scene3d>' +
        '<a:sp3d extrusionH="57150" prstMaterial="matte"><a:bevelT w="63500" h="12700" prst="angle"/>' +
        '<a:contourClr><a:schemeClr val="bg1"><a:lumMod val="65000"/></a:schemeClr></a:contourClr></a:sp3d>',
    );
    expect(getShapeText3D(shape)).toEqual(SHARP);
    const reloaded = await loadPresentation(await savePresentation(pres));
    expect(getShapeText3D(getSlideShapes(getSlides(reloaded)[0]!)[0]!)).toEqual(SHARP);
  });

  it('round trips a light-rig rotation in degrees', () => {
    const { shape } = textBox();
    const value: Text3D = {
      scene: {
        camera: 'orthographicFront',
        lightRig: {
          type: 'soft',
          direction: 't',
          rotation: { latitudeDeg: 0, longitudeDeg: 0, revolutionDeg: 260 },
        },
      },
      bevelTop: { widthEmu: 25400, heightEmu: 38100 },
    };
    setShapeText3D(shape, value);
    expect(bodyPr(getShapeXmlString(shape))).toContain(
      '<a:lightRig rig="soft" dir="t"><a:rot lat="0" lon="0" rev="15600000"/></a:lightRig>',
    );
    expect(getShapeText3D(shape)).toEqual(value);
  });

  it('returns null without 3-D and removes it on null', () => {
    const { shape } = textBox();
    expect(getShapeText3D(shape)).toBeNull();
    setShapeText3D(shape, SHARP);
    setShapeText3D(shape, null);
    expect(getShapeText3D(shape)).toBeNull();
    expect(getShapeXmlString(shape)).not.toMatch(/scene3d|sp3d/);
  });

  it('writes only the parts it is given', () => {
    const { shape } = textBox();
    setShapeText3D(shape, SHARP);
    setShapeText3D(shape, { bevelTop: {} });
    const xml = bodyPr(getShapeXmlString(shape));
    expect(xml).not.toContain('scene3d');
    expect(xml).toContain('<a:sp3d><a:bevelT/></a:sp3d>');
    expect(getShapeText3D(shape)).toEqual({ bevelTop: {} });
  });

  it('orders bodyPr children per CT_TextBodyProperties and keeps the rest', () => {
    const { shape } = textBox();
    const txBody = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'txBody', NS.pml))!;
    const body = firstChildElement(txBody, qname('a', 'bodyPr', NS.dml))!;
    // A warp, flat text (the sp3d alternative) and an extension list.
    body.children = [
      `<a:prstTxWarp xmlns:a="${NS.dml}" prst="textArchUp"><a:avLst/></a:prstTxWarp>`,
      `<a:flatTx xmlns:a="${NS.dml}"/>`,
      `<a:extLst xmlns:a="${NS.dml}"><a:ext uri="{test}"/></a:extLst>`,
    ].map(parseFragment);
    setShapeText3D(shape, SHARP);
    setShapeTextAutoFit(shape, 'normal');
    const order = bodyPr(getShapeXmlString(shape))
      .match(/<a:(\w+)/g)!
      .map((tag) => tag.slice(3))
      .filter((name) =>
        ['prstTxWarp', 'normAutofit', 'scene3d', 'sp3d', 'flatTx', 'extLst'].includes(name),
      );
    // flatTx and sp3d are one choice, so the bevel replaces the flat text.
    expect(order).toEqual(['prstTxWarp', 'normAutofit', 'scene3d', 'sp3d', 'extLst']);

    setShapeText3D(shape, null);
    expect(bodyPr(getShapeXmlString(shape))).toContain('<a:ext uri="{test}"/>');
    expect(bodyPr(getShapeXmlString(shape))).toContain('prstTxWarp');
  });

  it('keeps the backdrop it does not model and replaces the fields it does', () => {
    const { shape } = textBox();
    const txBody = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'txBody', NS.pml))!;
    const body = firstChildElement(txBody, qname('a', 'bodyPr', NS.dml))!;
    body.children = [
      `<a:scene3d xmlns:a="${NS.dml}"><a:camera prst="perspectiveFront" fov="2700000"/>` +
        `<a:lightRig rig="threePt" dir="t"/><a:backdrop><a:anchor x="0" y="0" z="0"/>` +
        `<a:norm dx="0" dy="0" dz="1"/><a:up dx="0" dy="1" dz="0"/></a:backdrop></a:scene3d>`,
      `<a:sp3d xmlns:a="${NS.dml}" z="12700" contourW="6350"><a:bevelB/></a:sp3d>`,
    ].map(parseFragment);
    expect(getShapeText3D(shape)).toMatchObject({
      scene: { camera: 'perspectiveFront', fieldOfViewDeg: 45 },
      bevelBottom: {},
      contourWidthEmu: 6350,
      distanceFromGroundEmu: 12700,
    });
    setShapeText3D(shape, SHARP);
    const xml = bodyPr(getShapeXmlString(shape));
    expect(xml).toContain('<a:camera prst="orthographicFront"/>');
    expect(xml).toContain('<a:backdrop>');
    expect(xml).not.toContain('contourW');
    expect(xml).not.toContain('bevelB');
    expect(getShapeText3D(shape)).toEqual(SHARP);
  });

  it('rejects tokens outside the schema enums and leaves the body untouched', () => {
    const { shape } = textBox();
    setShapeText3D(shape, SHARP);
    const before = getShapeXmlString(shape);
    const bad = (value: unknown) => () => setShapeText3D(shape, value as Text3D);
    expect(bad({ ...SHARP, scene: { ...SHARP.scene, camera: 'front' } })).toThrow(/scene.camera/);
    expect(
      bad({
        ...SHARP,
        scene: { camera: 'orthographicFront', lightRig: { type: 'x', direction: 't' } },
      }),
    ).toThrow(/lightRig.type/);
    expect(bad({ ...SHARP, bevelTop: { preset: 'round' } })).toThrow(/bevelTop.preset/);
    expect(bad({ ...SHARP, material: 'wood' })).toThrow(/material/);
    expect(bad({ ...SHARP, extrusionHeightEmu: -1 })).toThrow(/extrusionHeightEmu/);
    expect(bad({ contourColorTransforms: [] })).toThrow(/requires contourColor/);
    expect(getShapeXmlString(shape)).toBe(before);
  });

  it.skipIf(!isSchemaValidationAvailable())('validates against pml.xsd', () => {
    const { slide, shape } = textBox();
    setShapeTextAutoFit(shape, 'shape');
    setShapeText3D(shape, {
      ...SHARP,
      scene: {
        camera: 'orthographicFront',
        lightRig: {
          type: 'soft',
          direction: 't',
          rotation: { latitudeDeg: 0, longitudeDeg: 0, revolutionDeg: 260 },
        },
      },
    });
    expectSchemaValid(getSlideXmlString(slide), 'pml');
  });
});
