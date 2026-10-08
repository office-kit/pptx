// Flat approximations of a shape's own 3-D (`<a:scene3d>` / `<a:sp3d>`) for
// SVG, which has no 3-D: the camera rotation becomes the affine image of the
// rotated plane under an orthographic projection, and a top bevel becomes a
// lighting filter over the shape's edge. Perspective foreshortening,
// extrusion, contours and materials are not drawn.
import type { ReadText3D } from '@office-kit/pptx';

const EMU_PER_PX = 9525;
// ECMA-376 CT_Bevel: `w` and `h` default to 76200 EMU (6 pt).
const DEFAULT_BEVEL_EMU = 76200;
// Where the reference desktop app's default rigs light a bevel from: the upper left.
const LIGHT_AZIMUTH_DEG = 225;
const LIGHT_ELEVATION_DEG = 45;

const radians = (degrees: number) => (degrees * Math.PI) / 180;

/**
 * An SVG `matrix(...)` for the camera's explicit rotation about the point
 * (`cx`, `cy`) in px, or `null` when the camera keeps the shape face-on.
 * The camera turns, so the shape appears turned the opposite way: a camera
 * revolution of +6° (Picture Style "Rotated, White") tilts the picture 6°
 * counter-clockwise.
 */
export const cameraTransform = (
  shape3d: ReadText3D | null,
  cx: number,
  cy: number,
): string | null => {
  const rotation = shape3d?.scene?.cameraRotation;
  if (!rotation) return null;
  const { latitudeDeg, longitudeDeg, revolutionDeg } = rotation;
  if (latitudeDeg === 0 && longitudeDeg === 0 && revolutionDeg === 0) return null;
  // The plane's x and y axes after turning about y (longitude) and then x
  // (latitude), projected onto the screen; then the in-screen revolution.
  const lon = radians(-longitudeDeg);
  const lat = radians(-latitudeDeg);
  const rev = radians(-revolutionDeg);
  const ex = [Math.cos(lon), Math.sin(lon) * Math.sin(lat)] as const;
  const ey = [0, Math.cos(lat)] as const;
  const cos = Math.cos(rev);
  const sin = Math.sin(rev);
  const turn = (v: readonly [number, number]) => [cos * v[0] - sin * v[1], sin * v[0] + cos * v[1]];
  const [a, b] = turn(ex);
  const [c, d] = turn(ey);
  const e = cx - a! * cx - c! * cy;
  const f = cy - b! * cx - d! * cy;
  return `matrix(${[a, b, c, d, e, f].map((n) => n!.toFixed(6)).join(' ')})`;
};

/**
 * The `<filter>` body that shades a top bevel onto whatever the shape painted
 * (picture, fill and border alike): a lit height map built from the blurred
 * outline, multiplied in so flat areas keep their color and slopes facing
 * away from the light darken, plus a specular highlight on the lit slopes.
 * `null` when the shape has no top bevel.
 */
export const bevelFilterPrimitives = (shape3d: ReadText3D | null): string | null => {
  const bevel = shape3d?.bevelTop;
  if (!bevel) return null;
  const widthPx = (bevel.widthEmu ?? DEFAULT_BEVEL_EMU) / EMU_PER_PX;
  const heightPx = (bevel.heightEmu ?? DEFAULT_BEVEL_EMU) / EMU_PER_PX;
  if (widthPx <= 0 || heightPx <= 0) return null;
  const blur = widthPx / 2;
  // A steeper bevel (taller for its width) catches more light and shade.
  const surfaceScale = Math.min(10, Math.max(1, (heightPx / widthPx) * 6));
  // Divide out the flat surface's own diffuse term so it multiplies by 1.
  const diffuse = 1 / Math.sin(radians(LIGHT_ELEVATION_DEG));
  const light = `<feDistantLight azimuth="${LIGHT_AZIMUTH_DEG}" elevation="${LIGHT_ELEVATION_DEG}"/>`;
  return (
    `<feGaussianBlur in="SourceAlpha" stdDeviation="${blur.toFixed(2)}" result="bevelHeight"/>` +
    `<feDiffuseLighting in="bevelHeight" surfaceScale="${surfaceScale.toFixed(2)}" diffuseConstant="${diffuse.toFixed(4)}" lighting-color="#fff" result="bevelShade">${light}</feDiffuseLighting>` +
    `<feComposite in="SourceGraphic" in2="bevelShade" operator="arithmetic" k1="1" result="bevelLit"/>` +
    `<feSpecularLighting in="bevelHeight" surfaceScale="${surfaceScale.toFixed(2)}" specularConstant="0.6" specularExponent="20" lighting-color="#fff" result="bevelSpec">${light}</feSpecularLighting>` +
    `<feComposite in="bevelSpec" in2="SourceAlpha" operator="in" result="bevelGloss"/>` +
    `<feComposite in="bevelLit" in2="bevelGloss" operator="arithmetic" k2="1" k3="1"/>`
  );
};
