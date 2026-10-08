// Writes `<a:custGeom>` (ECMA-376 §20.1.9.8), the counterpart of
// `getShapeCustomGeometry`: freehand ink, traced outlines and any outline no
// preset describes.

import type {
  GeomCommand,
  GeomPoint,
  PathFillMode,
} from '../../internal/drawingml/custom-geometry.ts';
import { NS, attr, elem, qname, type XmlElement } from '../../internal/xml/index.ts';
import { SHAPE_SNAPSHOT, type SlideShapeData } from '../_internal-symbols.ts';
import { commitAndRefresh, requireSpPr } from './_helpers.ts';

/** One path to write; `fill` defaults to `'norm'` and `stroke` to `true`, as in the schema. */
export interface CustomGeometryPathInput {
  /** Path coordinate-space width and height (`ST_PositiveCoordinate`). */
  readonly w: number;
  readonly h: number;
  readonly fill?: PathFillMode;
  readonly stroke?: boolean;
  readonly commands: readonly GeomCommand[];
}
export interface CustomGeometryInput {
  readonly paths: readonly CustomGeometryPathInput[];
}

const a = (name: string) => qname('a', name, NS.dml);
// ST_PositiveCoordinate and ST_AdjCoordinate share the ST_Coordinate range.
const MAX_COORDINATE = 27273042316900;

function coordinate(value: number, label: string, positive = false): string {
  if (!Number.isFinite(value)) throw new Error(`setShapeCustomGeometry: ${label} must be finite`);
  const rounded = Math.round(value);
  if (Math.abs(rounded) > MAX_COORDINATE || (positive && rounded <= 0))
    throw new Error(`setShapeCustomGeometry: ${label} is out of range`);
  return String(rounded);
}
const point = (pt: GeomPoint, label: string): XmlElement =>
  elem(a('pt'), {
    attrs: [
      attr(qname('', 'x', ''), coordinate(pt.x, `${label}.x`)),
      attr(qname('', 'y', ''), coordinate(pt.y, `${label}.y`)),
    ],
  });

function command(item: GeomCommand, label: string): XmlElement {
  switch (item.kind) {
    case 'moveTo':
    case 'lnTo':
      return elem(a(item.kind), { children: [point(item.pt, label)] });
    case 'quadBezTo':
    case 'cubicBezTo':
      return elem(a(item.kind), {
        children: item.pts.map((pt, i) => point(pt, `${label}.pts[${i}]`)),
      });
    case 'arcTo':
      return elem(a('arcTo'), {
        attrs: [
          attr(qname('', 'wR', ''), coordinate(item.wR, `${label}.wR`)),
          attr(qname('', 'hR', ''), coordinate(item.hR, `${label}.hR`)),
          attr(qname('', 'stAng', ''), coordinate(item.stAng, `${label}.stAng`)),
          attr(qname('', 'swAng', ''), coordinate(item.swAng, `${label}.swAng`)),
        ],
      });
    case 'close':
      return elem(a('close'));
    default: {
      const never: never = item;
      throw new Error(`setShapeCustomGeometry: unknown command ${JSON.stringify(never)}`);
    }
  }
}

function path(input: CustomGeometryPathInput, index: number): XmlElement {
  const label = `paths[${index}]`;
  if (input.commands.length === 0 || input.commands[0]!.kind !== 'moveTo')
    throw new Error(`setShapeCustomGeometry: ${label} must start with moveTo`);
  const attrs = [
    attr(qname('', 'w', ''), coordinate(input.w, `${label}.w`, true)),
    attr(qname('', 'h', ''), coordinate(input.h, `${label}.h`, true)),
  ];
  if (input.fill !== undefined && input.fill !== 'norm')
    attrs.push(attr(qname('', 'fill', ''), input.fill));
  if (input.stroke === false) attrs.push(attr(qname('', 'stroke', ''), '0'));
  return elem(a('path'), {
    attrs,
    children: input.commands.map((item, i) => command(item, `${label}.commands[${i}]`)),
  });
}

/**
 * Replaces the shape's geometry (preset or custom) with the given paths. The
 * text rectangle is the whole shape, as the reference desktop app writes for freeforms.
 * Coordinates are rounded to integers, which the schema requires.
 */
export const setShapeCustomGeometry = (
  shape: SlideShapeData,
  geometry: CustomGeometryInput,
): void => {
  if (shape[SHAPE_SNAPSHOT].kind !== 'shape')
    throw new Error('setShapeCustomGeometry: only shapes (p:sp) carry custom geometry');
  if (geometry.paths.length === 0)
    throw new Error('setShapeCustomGeometry: at least one path is required');
  const custGeom = elem(a('custGeom'), {
    children: [
      elem(a('avLst')),
      elem(a('gdLst')),
      elem(a('ahLst')),
      elem(a('cxnLst')),
      elem(a('rect'), {
        attrs: ['l', 't', 'r', 'b'].map((name) => attr(qname('', name, ''), name)),
      }),
      elem(a('pathLst'), { children: geometry.paths.map(path) }),
    ],
  });
  const spPr = requireSpPr(shape);
  spPr.children = spPr.children.filter(
    (child) =>
      !(
        child.kind === 'element' &&
        child.name.namespaceURI === NS.dml &&
        ['prstGeom', 'custGeom'].includes(child.name.localName)
      ),
  );
  // CT_ShapeProperties: xfrm?, then the geometry choice, then fill and line.
  const transform = spPr.children.findIndex(
    (child) =>
      child.kind === 'element' &&
      child.name.namespaceURI === NS.dml &&
      child.name.localName === 'xfrm',
  );
  spPr.children.splice(transform + 1, 0, custGeom);
  commitAndRefresh(shape);
};
