// Picture Styles and Compress Pictures' per-picture compression state.
import {
  applyPictureStyle,
  type BuiltinPictureStyleName,
  detectPictureStyle,
  isBuiltinPictureStyleName,
} from '../../internal/drawingml/index.ts';
import {
  NS,
  type XmlElement,
  attr,
  elem,
  firstChildElement,
  getAttrValue,
  qname,
} from '../../internal/xml/index.ts';
import { SHAPE_ELEMENT, SHAPE_SNAPSHOT, type SlideShapeData } from '../_internal-symbols.ts';
import { commitAndRefresh, requireSpPr } from './_helpers.ts';

const requirePicture = (shape: SlideShapeData, fnName: string): void => {
  if (shape[SHAPE_SNAPSHOT].kind !== 'picture')
    throw new Error(
      `${fnName} only works on picture shapes; ${shape[SHAPE_SNAPSHOT].kind} is not one`,
    );
};

/**
 * Applies one of the reference desktop app's built-in Picture Styles (see
 * `BUILTIN_PICTURE_STYLES`) to a picture. Like the reference desktop app, it replaces the
 * picture's geometry, fill, border, effects and 3-D in `<p:spPr>` with the
 * style's own markup — which names no theme color, so the result does not
 * follow the theme — and keeps the picture, its crop and its position. Throws
 * when the shape is not a picture or `style` is not a built-in name.
 */
export const setShapePictureStyle = (
  shape: SlideShapeData,
  style: BuiltinPictureStyleName,
): void => {
  requirePicture(shape, 'setShapePictureStyle');
  if (!isBuiltinPictureStyleName(style))
    throw new Error(`setShapePictureStyle: "${String(style)}" is not a built-in picture style`);
  applyPictureStyle(requireSpPr(shape), style);
  commitAndRefresh(shape);
};

/**
 * The built-in picture style whose markup the picture's `<p:spPr>` carries
 * exactly (position, size and extensions aside), or `null`. A deck does not
 * record which style was applied, so a picture styled and then edited — a
 * different border, an extra effect — reads as `null`.
 */
export const getShapePictureStyle = (shape: SlideShapeData): BuiltinPictureStyleName | null => {
  if (shape[SHAPE_SNAPSHOT].kind !== 'picture') return null;
  const spPr = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'spPr', NS.pml));
  return spPr ? detectPictureStyle(spPr) : null;
};

/**
 * `ST_BlipCompression`: the target a picture was compressed for. The reference desktop app's
 * Compress Pictures writes `print` (220 ppi), `screen` (150 ppi) and `email`
 * (96 ppi); `hqprint` and `none` are the schema's other values.
 */
export type ImageCompressionState = 'email' | 'screen' | 'print' | 'hqprint' | 'none';

const COMPRESSION_STATES: ReadonlySet<string> = new Set<ImageCompressionState>([
  'email',
  'screen',
  'print',
  'hqprint',
  'none',
]);
const isCompressionState = (value: string | null): value is ImageCompressionState =>
  value !== null && COMPRESSION_STATES.has(value);
const ATTR_CSTATE = qname('', 'cstate', '');
// [MS-ODRAWXML] 2.3.1.13 `useLocalDpi` and the ext URI the reference desktop app files it under.
const USE_LOCAL_DPI_URI = '{28A0092B-C50C-407E-A947-70E740481C1C}';
const NAME_EXT_LST = qname('a', 'extLst', NS.dml);
const NAME_EXT = qname('a', 'ext', NS.dml);

const pictureBlip = (shape: SlideShapeData): XmlElement | null => {
  const fill = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'blipFill', NS.pml));
  return fill ? firstChildElement(fill, qname('a', 'blip', NS.dml)) : null;
};

/** The picture's `a:blip/@cstate`, or `null` when it has none. */
export const getShapeImageCompressionState = (
  shape: SlideShapeData,
): ImageCompressionState | null => {
  if (shape[SHAPE_SNAPSHOT].kind !== 'picture') return null;
  const blip = pictureBlip(shape);
  const value = blip ? getAttrValue(blip, ATTR_CSTATE) : null;
  return isCompressionState(value) ? value : null;
};

/**
 * Records what a picture was compressed for, as Compress Pictures does: it
 * sets `a:blip/@cstate` and adds the `a14:useLocalDpi` extension (without a
 * `val`, meaning true) that the reference desktop app writes alongside it. `null` removes
 * both. This only labels the picture; resample the pixels with
 * `setShapeImage`.
 */
export const setShapeImageCompressionState = (
  shape: SlideShapeData,
  state: ImageCompressionState | null,
): void => {
  requirePicture(shape, 'setShapeImageCompressionState');
  if (state !== null && !isCompressionState(state))
    throw new Error(
      `setShapeImageCompressionState: "${String(state)}" is not an ST_BlipCompression value`,
    );
  const blip = pictureBlip(shape);
  if (!blip) throw new Error('setShapeImageCompressionState: the picture has no <a:blip>');
  blip.attrs = blip.attrs.filter(
    (a) => !(a.name.namespaceURI === '' && a.name.localName === 'cstate'),
  );
  let extLst = firstChildElement(blip, NAME_EXT_LST);
  if (extLst) {
    extLst.children = extLst.children.filter(
      (child) =>
        !(
          child.kind === 'element' &&
          child.name.namespaceURI === NS.dml &&
          child.name.localName === 'ext' &&
          getAttrValue(child, qname('', 'uri', ''))?.toUpperCase() === USE_LOCAL_DPI_URI
        ),
    );
    if (!extLst.children.some((child) => child.kind === 'element')) {
      blip.children = blip.children.filter((child) => child !== extLst);
      extLst = null;
    }
  }
  if (state !== null) {
    blip.attrs.push(attr(ATTR_CSTATE, state));
    const ext = elem(NAME_EXT, {
      attrs: [attr(qname('', 'uri', ''), USE_LOCAL_DPI_URI)],
      children: [
        elem(qname('a14', 'useLocalDpi', NS.a14), { prefixDecls: new Map([['a14', NS.a14]]) }),
      ],
    });
    // extLst is CT_Blip's last child, after every effect.
    if (extLst) extLst.children.push(ext);
    else blip.children.push(elem(NAME_EXT_LST, { children: [ext] }));
  }
  commitAndRefresh(shape);
};
