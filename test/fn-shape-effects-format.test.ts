// The shape effects PowerPoint's Format Shape pane edits one at a time: outer
// and inner shadow (with scale and skew), reflection, soft edge, glow removal
// and the shape's own 3-D in `<p:spPr>`.

import { describe, expect, it } from 'vitest';
import {
  addBlankSlide,
  addSlideShape,
  createPresentation,
  getShape3D,
  getShapeEffects,
  getShapeXmlString,
  getSlides,
  getSlideShapes,
  getSlideXmlString,
  inches,
  loadPresentation,
  savePresentation,
  setShape3D,
  setShapeFill,
  setShapeGlow,
  setShapeInnerShadow,
  setShapeReflection,
  setShapeShadow,
  setShapeSoftEdge,
  setShapeStroke,
  type Shape3D,
} from '../src/api/index.ts';
import { SHAPE_ELEMENT } from '../src/api/_internal-symbols.ts';
import { NS, parseFragment } from '../src/internal/xml/index.ts';
import { expectSchemaValid, isSchemaValidationAvailable } from './lib/expect-schema-valid.ts';

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
const spPr = (xml: string) => compact(xml).match(/<p:spPr\b[\s\S]*?<\/p:spPr>/)![0];

// PowerPoint's Shadow ▸ Perspective: Upper Left preset carries scale and skew.
const PERSPECTIVE = {
  color: '#000000',
  opacity: 0.2,
  blurEmu: 76200,
  offsetEmu: 0,
  angleDeg: 315,
  scaleY: 0.23,
  skewX: -20,
  alignment: 'bl',
} as const;

// A bevelled, extruded, rotated shape with every modelled 3-D field.
const BEVELLED: Shape3D = {
  scene: {
    camera: 'perspectiveFront',
    cameraRotation: { latitudeDeg: 20, longitudeDeg: 30, revolutionDeg: 0 },
    fieldOfViewDeg: 45,
    lightRig: {
      type: 'threePt',
      direction: 't',
      rotation: { latitudeDeg: 0, longitudeDeg: 0, revolutionDeg: 40 },
    },
  },
  bevelTop: { widthEmu: 76200, heightEmu: 76200, preset: 'circle' },
  bevelBottom: { widthEmu: 50800, heightEmu: 25400, preset: 'angle' },
  extrusionHeightEmu: 127000,
  extrusionColor: 'accent2',
  contourWidthEmu: 12700,
  contourColor: '#FF0000',
  distanceFromGroundEmu: -25400,
  material: 'metal',
};

describe('shape effects edited one at a time', () => {
  it('writes and reads shadow scale and skew', () => {
    const { pres, shape } = rect();
    setShapeShadow(shape, PERSPECTIVE);
    expect(spPr(getShapeXmlString(shape))).toContain(
      '<a:outerShdw blurRad="76200" dist="0" dir="18900000" sy="23000" kx="-1200000" algn="bl" rotWithShape="0">',
    );
    expect(getShapeEffects(pres, shape)[0]).toMatchObject({
      kind: 'outerShdw',
      scaleY: 0.23,
      skewX: -20,
      alignment: 'bl',
    });
    expect(() => setShapeShadow(shape, { skewX: 90 })).toThrow(/skewX/);
  });

  it('removes one effect with null and keeps the others in schema order', () => {
    const { pres, shape } = rect();
    setShapeSoftEdge(shape, 63500);
    setShapeReflection(shape, {
      blurEmu: 6350,
      startOpacity: 0.5,
      opacity: 0.003,
      endPosition: 0.35,
      angleDeg: 90,
      scaleY: -1,
      alignment: 'bl',
      rotateWithShape: false,
    });
    setShapeInnerShadow(shape, { blurEmu: 63500, offsetEmu: 50800, angleDeg: 225, opacity: 0.5 });
    setShapeGlow(shape, { color: 'accent1', radiusEmu: 63500, opacity: 0.4 });
    setShapeShadow(shape);
    expect(getShapeEffects(pres, shape).map((effect) => effect.kind)).toEqual([
      'glow',
      'innerShdw',
      'outerShdw',
      'reflection',
      'softEdge',
    ]);
    setShapeGlow(shape, null);
    setShapeShadow(shape, null);
    setShapeInnerShadow(shape, null);
    expect(getShapeEffects(pres, shape).map((effect) => effect.kind)).toEqual([
      'reflection',
      'softEdge',
    ]);
    setShapeReflection(shape, null);
    setShapeSoftEdge(shape, null);
    expect(spPr(getShapeXmlString(shape))).not.toContain('effectLst');
    expect(() => setShapeSoftEdge(shape, -1)).toThrow(/radiusEmu/);
  });

  it('keeps an empty effect list when the shape style references a theme effect', () => {
    const { pres, shape } = rect();
    const style = parseFragment(
      `<p:style xmlns:p="${NS.pml}" xmlns:a="${NS.dml}"><a:lnRef idx="2"><a:schemeClr val="accent1"/></a:lnRef>` +
        `<a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="2"><a:schemeClr val="accent1"/></a:effectRef>` +
        `<a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style>`,
    );
    const element = shape[SHAPE_ELEMENT];
    const at = element.children.findIndex(
      (child) => child.kind === 'element' && child.name.localName === 'txBody',
    );
    element.children.splice(at === -1 ? element.children.length : at, 0, style);
    setShapeShadow(shape);
    setShapeShadow(shape, null);
    expect(spPr(getShapeXmlString(shape))).toContain('<a:effectLst/>');
    expect(getShapeEffects(pres, shape)).toEqual([]);
  });

  it('writes the shape 3-D after the effects and reads it back after a save', async () => {
    const { pres, shape } = rect();
    setShapeFill(shape, { color: 'accent1' });
    setShapeStroke(shape, { color: '#000000' });
    setShape3D(shape, BEVELLED);
    setShapeShadow(shape);
    const xml = spPr(getShapeXmlString(shape));
    expect(xml).toMatch(/<a:ln\b[\s\S]*<\/a:ln><a:effectLst>[\s\S]*<\/a:effectLst><a:scene3d>/);
    expect(xml).toContain(
      '<a:scene3d><a:camera prst="perspectiveFront" fov="2700000"><a:rot lat="1200000" lon="1800000" rev="0"/></a:camera>' +
        '<a:lightRig rig="threePt" dir="t"><a:rot lat="0" lon="0" rev="2400000"/></a:lightRig></a:scene3d>' +
        '<a:sp3d z="-25400" extrusionH="127000" contourW="12700" prstMaterial="metal">' +
        '<a:bevelT w="76200" h="76200" prst="circle"/><a:bevelB w="50800" h="25400" prst="angle"/>' +
        '<a:extrusionClr><a:schemeClr val="accent2"/></a:extrusionClr>' +
        '<a:contourClr><a:srgbClr val="FF0000"/></a:contourClr></a:sp3d>',
    );
    expect(getShape3D(shape)).toEqual(BEVELLED);
    const reloaded = await loadPresentation(await savePresentation(pres));
    expect(getShape3D(getSlideShapes(getSlides(reloaded)[0]!)[0]!)).toEqual(BEVELLED);
    setShape3D(shape, { scene: BEVELLED.scene! });
    expect(spPr(getShapeXmlString(shape))).not.toContain('sp3d');
    setShape3D(shape, null);
    expect(getShape3D(shape)).toBeNull();
    expect(() => setShape3D(shape, { distanceFromGroundEmu: 0.5 })).toThrow(/distanceFromGround/);
    expect(() => setShape3D(shape, { scene: { ...BEVELLED.scene!, fieldOfViewDeg: 181 } })).toThrow(
      /fieldOfViewDeg/,
    );
  });

  it.skipIf(!isSchemaValidationAvailable())('validates against pml.xsd', () => {
    const { slide, shape } = rect();
    setShapeStroke(shape, { color: '#000000' });
    setShape3D(shape, BEVELLED);
    setShapeShadow(shape, PERSPECTIVE);
    setShapeInnerShadow(shape, {});
    setShapeGlow(shape, { color: 'accent1' });
    setShapeReflection(shape, { scaleY: -1, alignment: 'bl', endPosition: 0.5 });
    setShapeSoftEdge(shape, 12700);
    expectSchemaValid(getSlideXmlString(slide), 'pml');
  });
});
