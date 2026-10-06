// Text-body 3-D: `<a:scene3d>` and `<a:sp3d>` inside `<a:bodyPr>`.
//
// PowerPoint's WordArt bevels (Soft Bevel, Sharp Bevel) live here, on the text
// body, not on the shape's `<p:spPr>`: a shape-level bevel would bevel the box.
// The schema types are shared with shape 3-D (CT_Scene3D, CT_Shape3D,
// ECMA-376 §20.1.4.1.26 / §20.1.5.12), so the vocabulary below is DrawingML's.

import type { Color } from './color.ts';
import { buildColorElement } from './color.ts';
import {
  type ColorTransform,
  buildColorTransforms,
  readColorTransforms,
} from './color-transforms.ts';
import { emuExtent, oneOf } from '../bounds.ts';
import {
  BEVEL_PRESETS,
  CAMERA_PRESETS,
  LIGHT_RIG_DIRECTIONS,
  LIGHT_RIG_TYPES,
  PRESET_MATERIALS,
} from '../enum-values.ts';
import {
  NS,
  type XmlElement,
  attr,
  elem,
  firstChildElement,
  getAttrValue,
  insertChildByRank,
  qname,
} from '../xml/index.ts';

/** ECMA-376 `ST_PresetCameraType`; PowerPoint's flat WordArt uses `'orthographicFront'`. */
export type CameraPreset = (typeof CAMERA_PRESETS)[number];
/** ECMA-376 `ST_LightRigType` (`'threePt'`, `'soft'`, `'harsh'`, ...). */
export type LightRigType = (typeof LIGHT_RIG_TYPES)[number];
/** ECMA-376 `ST_LightRigDirection`: where the light rig sits relative to the scene. */
export type LightRigDirection = (typeof LIGHT_RIG_DIRECTIONS)[number];
/** ECMA-376 `ST_BevelPresetType` (`'circle'`, `'angle'`, `'softRound'`, ...). */
export type BevelPreset = (typeof BEVEL_PRESETS)[number];
/** ECMA-376 `ST_PresetMaterialType` (`'matte'`, `'softEdge'`, `'metal'`, ...). */
export type PresetMaterial = (typeof PRESET_MATERIALS)[number];

/** `CT_SphereCoords`: a rotation in degrees, each in `[0, 360)`. */
export interface Rotation3D {
  readonly latitudeDeg: number;
  readonly longitudeDeg: number;
  readonly revolutionDeg: number;
}

/** `<a:lightRig>`. */
export interface LightRig {
  readonly type: LightRigType;
  readonly direction: LightRigDirection;
  /** `<a:rot>`; omitted, the rig keeps its preset orientation. */
  readonly rotation?: Rotation3D;
}

/** `<a:scene3d>`: the camera and light the 3-D text is rendered with. */
export interface Scene3D {
  readonly camera: CameraPreset;
  readonly lightRig: LightRig;
}

/** `CT_Bevel`. Omitted fields take the schema defaults: `circle`, 76200 × 76200 EMU. */
export interface Bevel {
  readonly preset?: BevelPreset;
  readonly widthEmu?: number;
  readonly heightEmu?: number;
}

/**
 * 3-D on a text body — what PowerPoint writes for WordArt bevels. `scene`
 * becomes `<a:scene3d>`; the other fields become `<a:sp3d>`, written when any
 * of them is set.
 */
export interface Text3D {
  readonly scene?: Scene3D;
  /** Top bevel (`<a:bevelT>`). */
  readonly bevelTop?: Bevel;
  /** Extrusion depth (`extrusionH`) in EMU. */
  readonly extrusionHeightEmu?: number;
  /** Surface material (`prstMaterial`). */
  readonly material?: PresetMaterial;
  /** Contour color (`<a:contourClr>`). Same accepted forms as `TextFormat.color`. */
  readonly contourColor?: Color;
  /** Ordered adjustments to `contourColor`. Requires `contourColor`. */
  readonly contourColorTransforms?: readonly ColorTransform[];
}

/** A text body's 3-D read back from a deck; `contourColor` is the raw `#RRGGBB` or scheme token. */
export type ReadText3D = Omit<Text3D, 'contourColor'> & { readonly contourColor?: string };

const dml = (local: string) => qname('a', local, NS.dml);
const NAME_SCENE_3D = dml('scene3d');
const NAME_SP_3D = dml('sp3d');
const NAME_CAMERA = dml('camera');
const NAME_LIGHT_RIG = dml('lightRig');
const NAME_ROT = dml('rot');
const NAME_BEVEL_T = dml('bevelT');
const NAME_CONTOUR_CLR = dml('contourClr');
const plain = (local: string) => qname('', local, '');

const ANGLE_UNITS_PER_DEGREE = 60000;
const FULL_TURN_DEGREES = 360;
const FULL_TURN_UNITS = FULL_TURN_DEGREES * ANGLE_UNITS_PER_DEGREE;

// CT_TextBodyProperties is a sequence: prstTxWarp, the autofit choice,
// scene3d, the sp3d|flatTx choice, extLst. Unknown children keep their place
// behind everything the schema names.
const BODY_PR_CHILD_RANK: Record<string, number> = {
  prstTxWarp: 0,
  noAutofit: 1,
  normAutofit: 1,
  spAutoFit: 1,
  scene3d: 2,
  sp3d: 3,
  flatTx: 3,
  extLst: 4,
};
/** @internal */
export const bodyPrChildRank = (el: XmlElement): number =>
  el.name.namespaceURI === NS.dml ? (BODY_PR_CHILD_RANK[el.name.localName] ?? 99) : 99;

// CT_Shape3D: bevelT, bevelB, extrusionClr, contourClr, extLst.
const SP_3D_CHILD_RANK: Record<string, number> = {
  bevelT: 0,
  bevelB: 1,
  extrusionClr: 2,
  contourClr: 3,
  extLst: 4,
};
const sp3dChildRank = (el: XmlElement): number =>
  el.name.namespaceURI === NS.dml ? (SP_3D_CHILD_RANK[el.name.localName] ?? 99) : 99;

const isDml = (el: XmlElement, ...locals: string[]): boolean =>
  el.name.namespaceURI === NS.dml && locals.includes(el.name.localName);

const withoutChildren = (host: XmlElement, ...locals: string[]): void => {
  host.children = host.children.filter((c) => c.kind !== 'element' || !isDml(c, ...locals));
};

const setAttr = (host: XmlElement, local: string, value: string | null): void => {
  host.attrs = host.attrs.filter((a) => a.name.namespaceURI !== '' || a.name.localName !== local);
  if (value !== null) host.attrs.push(attr(plain(local), value));
};

// ST_PositiveFixedAngle is [0, 21600000); a turn is normalised into it the
// way the effect directions are.
const positiveFixedAngle = (degrees: number, field: string): string => {
  if (!Number.isFinite(degrees)) throw new RangeError(`${field} must be finite`);
  const turn = ((degrees % FULL_TURN_DEGREES) + FULL_TURN_DEGREES) % FULL_TURN_DEGREES;
  return String(Math.round(turn * ANGLE_UNITS_PER_DEGREE) % FULL_TURN_UNITS);
};

const hasShape3D = (value: Text3D): boolean =>
  value.bevelTop !== undefined ||
  value.extrusionHeightEmu !== undefined ||
  value.material !== undefined ||
  value.contourColor !== undefined;

const buildScene = (scene: Scene3D, previous: XmlElement | null, caller: string): XmlElement => {
  const camera = oneOf(scene.camera, CAMERA_PRESETS, `${caller}: scene.camera`);
  const rig = oneOf(scene.lightRig.type, LIGHT_RIG_TYPES, `${caller}: scene.lightRig.type`);
  const dir = oneOf(
    scene.lightRig.direction,
    LIGHT_RIG_DIRECTIONS,
    `${caller}: scene.lightRig.direction`,
  );
  const rotation = scene.lightRig.rotation;
  const lightRig = elem(NAME_LIGHT_RIG, {
    attrs: [attr(plain('rig'), rig), attr(plain('dir'), dir)],
    children:
      rotation === undefined
        ? []
        : [
            elem(NAME_ROT, {
              attrs: [
                attr(
                  plain('lat'),
                  positiveFixedAngle(rotation.latitudeDeg, `${caller}: latitudeDeg`),
                ),
                attr(
                  plain('lon'),
                  positiveFixedAngle(rotation.longitudeDeg, `${caller}: longitudeDeg`),
                ),
                attr(
                  plain('rev'),
                  positiveFixedAngle(rotation.revolutionDeg, `${caller}: revolutionDeg`),
                ),
              ],
            }),
          ],
  });
  // The camera's field of view, zoom and rotation, and the scene's backdrop,
  // are not modelled; editing the scene keeps whatever the deck had.
  const result = previous
    ? { ...previous, attrs: [...previous.attrs], children: [...previous.children] }
    : elem(NAME_SCENE_3D);
  const previousCamera = firstChildElement(result, NAME_CAMERA);
  const cameraElement = previousCamera
    ? { ...previousCamera, attrs: [...previousCamera.attrs] }
    : elem(NAME_CAMERA);
  setAttr(cameraElement, 'prst', camera);
  withoutChildren(result, 'camera', 'lightRig');
  // camera and lightRig open the CT_Scene3D sequence.
  result.children.unshift(cameraElement, lightRig);
  return result;
};

const buildShape3D = (value: Text3D, previous: XmlElement | null, caller: string): XmlElement => {
  const extrusion =
    value.extrusionHeightEmu === undefined
      ? null
      : String(emuExtent(value.extrusionHeightEmu, `${caller}: extrusionHeightEmu`));
  const material =
    value.material === undefined
      ? null
      : oneOf(value.material, PRESET_MATERIALS, `${caller}: material`);
  const bevel = value.bevelTop;
  const bevelElement =
    bevel === undefined
      ? null
      : elem(NAME_BEVEL_T, {
          attrs: [
            ...(bevel.widthEmu === undefined
              ? []
              : [
                  attr(
                    plain('w'),
                    String(emuExtent(bevel.widthEmu, `${caller}: bevelTop.widthEmu`)),
                  ),
                ]),
            ...(bevel.heightEmu === undefined
              ? []
              : [
                  attr(
                    plain('h'),
                    String(emuExtent(bevel.heightEmu, `${caller}: bevelTop.heightEmu`)),
                  ),
                ]),
            ...(bevel.preset === undefined
              ? []
              : [
                  attr(
                    plain('prst'),
                    oneOf(bevel.preset, BEVEL_PRESETS, `${caller}: bevelTop.preset`),
                  ),
                ]),
          ],
        });
  const contourElement = (() => {
    if (value.contourColor === undefined) return null;
    const color = buildColorElement(value.contourColor);
    color.children = buildColorTransforms(value.contourColorTransforms ?? []);
    return elem(NAME_CONTOUR_CLR, { children: [color] });
  })();
  // `z`, `contourW`, the bottom bevel and the extrusion color are not
  // modelled; editing keeps whatever the deck had.
  const result = previous
    ? { ...previous, attrs: [...previous.attrs], children: [...previous.children] }
    : elem(NAME_SP_3D);
  setAttr(result, 'extrusionH', extrusion);
  setAttr(result, 'prstMaterial', material);
  withoutChildren(result, 'bevelT', 'contourClr');
  if (bevelElement) insertChildByRank(result, bevelElement, sp3dChildRank);
  if (contourElement) insertChildByRank(result, contourElement, sp3dChildRank);
  return result;
};

/**
 * Replaces the 3-D on `bodyPr`. Everything is validated before `bodyPr`
 * changes, so a rejected value leaves it as it was.
 */
export const applyText3D = (bodyPr: XmlElement, value: Text3D | null, caller: string): void => {
  if (value?.contourColorTransforms !== undefined && value.contourColor === undefined)
    throw new Error(`${caller}: contourColorTransforms requires contourColor`);
  const previousScene = firstChildElement(bodyPr, NAME_SCENE_3D);
  const previousShape = firstChildElement(bodyPr, NAME_SP_3D);
  const scene = value?.scene === undefined ? null : buildScene(value.scene, previousScene, caller);
  const shape =
    value === null || !hasShape3D(value) ? null : buildShape3D(value, previousShape, caller);
  // sp3d and flatTx are one choice (EG_Text3D); a bevel replaces flat text.
  withoutChildren(bodyPr, 'scene3d', 'sp3d', ...(shape ? ['flatTx'] : []));
  if (scene) insertChildByRank(bodyPr, scene, bodyPrChildRank);
  if (shape) insertChildByRank(bodyPr, shape, bodyPrChildRank);
};

const readEnum = <T extends string>(
  element: XmlElement,
  local: string,
  allowed: readonly T[],
): T | undefined => {
  const raw = getAttrValue(element, plain(local));
  return raw !== null && (allowed as readonly string[]).includes(raw) ? (raw as T) : undefined;
};

const readInteger = (element: XmlElement, local: string): number | undefined => {
  const raw = getAttrValue(element, plain(local));
  const value = raw === null ? Number.NaN : Number.parseInt(raw, 10);
  return Number.isFinite(value) ? value : undefined;
};

const readDegrees = (element: XmlElement, local: string): number | undefined => {
  const units = readInteger(element, local);
  return units === undefined ? undefined : units / ANGLE_UNITS_PER_DEGREE;
};

const readScene = (scene: XmlElement): Scene3D | undefined => {
  const cameraElement = firstChildElement(scene, NAME_CAMERA);
  const rigElement = firstChildElement(scene, NAME_LIGHT_RIG);
  if (cameraElement === null || rigElement === null) return undefined;
  const camera = readEnum(cameraElement, 'prst', CAMERA_PRESETS);
  const type = readEnum(rigElement, 'rig', LIGHT_RIG_TYPES);
  const direction = readEnum(rigElement, 'dir', LIGHT_RIG_DIRECTIONS);
  if (camera === undefined || type === undefined || direction === undefined) return undefined;
  const rot = firstChildElement(rigElement, NAME_ROT);
  const latitudeDeg = rot === null ? undefined : readDegrees(rot, 'lat');
  const longitudeDeg = rot === null ? undefined : readDegrees(rot, 'lon');
  const revolutionDeg = rot === null ? undefined : readDegrees(rot, 'rev');
  const rotation =
    latitudeDeg === undefined || longitudeDeg === undefined || revolutionDeg === undefined
      ? undefined
      : { latitudeDeg, longitudeDeg, revolutionDeg };
  return { camera, lightRig: { type, direction, ...(rotation ? { rotation } : {}) } };
};

const readBevel = (bevel: XmlElement): Bevel => {
  const preset = readEnum(bevel, 'prst', BEVEL_PRESETS);
  const widthEmu = readInteger(bevel, 'w');
  const heightEmu = readInteger(bevel, 'h');
  return {
    ...(preset === undefined ? {} : { preset }),
    ...(widthEmu === undefined ? {} : { widthEmu }),
    ...(heightEmu === undefined ? {} : { heightEmu }),
  };
};

// The literal color, as the run readers report one without a theme.
const readContour = (
  sp3d: XmlElement,
): { contourColor?: string; contourColorTransforms?: ColorTransform[] } => {
  const holder = firstChildElement(sp3d, NAME_CONTOUR_CLR);
  const color = holder?.children.find(
    (c): c is XmlElement => c.kind === 'element' && isDml(c, 'srgbClr', 'schemeClr'),
  );
  const val = color && getAttrValue(color, plain('val'));
  if (!color || val == null) return {};
  const transforms = readColorTransforms(color);
  return {
    contourColor: color.name.localName === 'srgbClr' ? `#${val.toUpperCase()}` : val,
    ...(transforms.length ? { contourColorTransforms: transforms } : {}),
  };
};

/** Reads `bodyPr`'s 3-D, or `null` when it has neither `<a:scene3d>` nor `<a:sp3d>`. */
export const readText3D = (bodyPr: XmlElement): ReadText3D | null => {
  const sceneElement = firstChildElement(bodyPr, NAME_SCENE_3D);
  const shape = firstChildElement(bodyPr, NAME_SP_3D);
  if (sceneElement === null && shape === null) return null;
  const scene = sceneElement === null ? undefined : readScene(sceneElement);
  if (shape === null) return scene ? { scene } : {};
  const bevelElement = firstChildElement(shape, NAME_BEVEL_T);
  const bevelTop = bevelElement === null ? undefined : readBevel(bevelElement);
  const extrusionHeightEmu = readInteger(shape, 'extrusionH');
  const material = readEnum(shape, 'prstMaterial', PRESET_MATERIALS);
  return {
    ...(scene ? { scene } : {}),
    ...(bevelTop ? { bevelTop } : {}),
    ...(extrusionHeightEmu === undefined ? {} : { extrusionHeightEmu }),
    ...(material === undefined ? {} : { material }),
    ...readContour(shape),
  };
};
