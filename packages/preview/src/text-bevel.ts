// Text-body bevel (`<a:bodyPr><a:scene3d>…<a:sp3d><a:bevelT/>`) → the lighting
// parameters the text engine shades glyphs with.
//
// PowerPoint renders a beveled text body as real 3-D geometry. The preview
// approximates it in 2-D: the glyph alpha becomes a height field whose edge
// profile follows the bevel preset, lit by one distant light derived from the
// light rig. Everything below is a coarse, hand-tuned approximation, not a
// port of PowerPoint's lighting model:
//
// - Each light rig is one distant light (PowerPoint's rigs have two to four)
//   whose elevation and contrast stand in for the rig's character.
// - The camera is always treated as orthographicFront. Perspective and
//   isometric cameras would tilt the text; that is not drawn.
// - Extrusion (`extrusionH`) is invisible from the front, so it is ignored.
//   So are the contour (`contourW`/`contourClr`) and the bottom bevel.
// - Materials only vary the specular highlight and the overall contrast.

import type { ReadText3D } from '@office-kit/pptx';
import type { TextBevelInput } from './text-layout.ts';

type BevelPreset = NonNullable<NonNullable<ReadText3D['bevelTop']>['preset']>;
type LightRig = NonNullable<ReadText3D['scene']>['lightRig'];
type Material = NonNullable<ReadText3D['material']>;

// CT_Bevel defaults (ECMA-376 dml-main.xsd): circle, 76200 × 76200 EMU.
const DEFAULT_BEVEL_EMU = 76200;
const DEFAULT_PRESET: BevelPreset = 'circle';
// A text body with sp3d but no scene3d keeps PowerPoint's default rig.
const DEFAULT_RIG: LightRig = { type: 'threePt', direction: 't' };
const PROFILE_SAMPLES = 9;

// Height profiles from the glyph edge (t = 0) to the bevel's inner edge
// (t = 1), shaped after the Bevel gallery thumbnails. The blur that feeds the
// table already rounds every corner, so only the broad shape survives:
// convex presets brighten the lit edge sharply, linear ones evenly, and the
// grooved ones (divot, riblet, artDeco, cross) add a second highlight band.
const PROFILES: Record<BevelPreset, (t: number) => number> = {
  circle: (t) => Math.sqrt(1 - (1 - t) ** 2),
  relaxedInset: (t) => Math.sin((t * Math.PI) / 2) ** 2,
  slope: (t) => t,
  angle: (t) => t,
  softRound: (t) => t * t * (3 - 2 * t),
  convex: (t) => Math.sqrt(Math.sqrt(1 - (1 - t) ** 2)),
  coolSlant: (t) => Math.min(1, t * 1.6),
  hardEdge: (t) => Math.min(1, t * 3),
  cross: (t) => (t < 0.5 ? t * 1.6 : 0.8 + (t - 0.5) * 0.4),
  divot: (t) => (t < 0.35 ? t / 0.35 : 1 - 0.55 * Math.sin(((t - 0.35) / 0.65) * Math.PI)),
  riblet: (t) => t * 0.7 + 0.3 * Math.abs(Math.sin(t * 2.5 * Math.PI)),
  artDeco: (t) => Math.min(1, Math.floor(t * 3 + 0.5) / 3 + (t * 3 - Math.floor(t * 3)) * 0.12),
};

// ST_LightRigDirection → where the light comes from, as an SVG azimuth
// (degrees clockwise from +x; +y points down the slide).
const DIRECTION_AZIMUTH: Record<LightRig['direction'], number> = {
  r: 0,
  br: 45,
  b: 90,
  bl: 135,
  l: 180,
  tl: 225,
  t: 270,
  tr: 315,
};

interface RigLook {
  readonly elevationDeg: number;
  readonly contrast: number;
}
// Low lights rake across the bevel (high contrast); high, diffuse rigs
// barely shade it. `flat` and `legacyFlat*` light the surface head-on.
const RIG_LOOK = (type: LightRig['type']): RigLook | null => {
  if (type === 'flat' || type.startsWith('legacyFlat')) return null;
  if (type === 'harsh' || type.startsWith('legacyHarsh') || type === 'contrasting') {
    return { elevationDeg: 30, contrast: 0.65 };
  }
  if (type === 'soft' || type === 'flood' || type === 'brightRoom' || type === 'glow') {
    return { elevationDeg: 45, contrast: 0.55 };
  }
  return { elevationDeg: 38, contrast: 0.65 };
};

interface MaterialLook {
  readonly contrastScale: number;
  readonly specular: TextBevelInput['specular'];
}
const MATTE: MaterialLook = { contrastScale: 1, specular: null };
const MATERIAL_LOOK: Partial<Record<Material, MaterialLook>> = {
  plastic: { contrastScale: 1, specular: { constant: 0.6, exponent: 20 } },
  legacyPlastic: { contrastScale: 1, specular: { constant: 0.6, exponent: 20 } },
  metal: { contrastScale: 1.2, specular: { constant: 0.9, exponent: 12 } },
  legacyMetal: { contrastScale: 1.2, specular: { constant: 0.9, exponent: 12 } },
  softmetal: { contrastScale: 1.1, specular: { constant: 0.6, exponent: 8 } },
  clear: { contrastScale: 0.9, specular: { constant: 0.8, exponent: 30 } },
  softEdge: { contrastScale: 0.9, specular: { constant: 0.25, exponent: 10 } },
  powder: { contrastScale: 0.8, specular: null },
  translucentPowder: { contrastScale: 0.7, specular: null },
  // `flat` drops all shading, as PowerPoint's Flat material does.
  flat: { contrastScale: 0, specular: null },
};

// The authored h / w slope reads too shallow once squeezed into a few pixels
// of blurred alpha; PowerPoint's renders show a clearly embossed edge.
const RELIEF = 2.5;

/**
 * The bevel to shade a text body with, or `null` when it has no top bevel or
 * its rig/material renders it unshaded.
 */
export const textBevelOf = (text3d: ReadText3D | null): TextBevelInput | null => {
  const bevel = text3d?.bevelTop;
  if (!bevel) return null;
  const widthEmu = bevel.widthEmu ?? DEFAULT_BEVEL_EMU;
  const heightEmu = bevel.heightEmu ?? DEFAULT_BEVEL_EMU;
  if (widthEmu <= 0 || heightEmu <= 0) return null;
  const rig = text3d.scene?.lightRig ?? DEFAULT_RIG;
  const look = RIG_LOOK(rig.type);
  const material = (text3d.material && MATERIAL_LOOK[text3d.material]) ?? MATTE;
  if (!look || material.contrastScale === 0) return null;
  const profile = PROFILES[bevel.preset ?? DEFAULT_PRESET];
  // The rig's revolution turns it about the view axis; DrawingML angles run
  // clockwise on screen, like the azimuth.
  const azimuthDeg = (DIRECTION_AZIMUTH[rig.direction] + (rig.rotation?.revolutionDeg ?? 0)) % 360;
  return {
    widthEmu,
    heightEmu,
    profile: Array.from({ length: PROFILE_SAMPLES }, (_, i) =>
      Math.max(0, Math.min(1, profile(i / (PROFILE_SAMPLES - 1)))),
    ),
    azimuthDeg,
    elevationDeg: look.elevationDeg,
    contrast: Math.min(1, look.contrast * material.contrastScale),
    relief: RELIEF,
    specular: material.specular,
  };
};
