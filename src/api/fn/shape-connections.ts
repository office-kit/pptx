import {
  NS,
  attr,
  elem,
  firstChildElement,
  getAttrValue,
  parseXml,
  qname,
  type XmlElement,
} from '../../internal/xml/index.ts';
import {
  parseConnectionSites,
  type ConnectionSite,
} from '../../internal/drawingml/custom-geometry.ts';
import { presetConnectionSites } from '../../internal/drawingml/preset-connection-sites.ts';
import {
  SHAPE_ELEMENT,
  SHAPE_SLIDE,
  SHAPE_SNAPSHOT,
  SLIDE_SHAPES,
  type SlideShapeData,
} from '../_internal-symbols.ts';
import { getShapeId, getShapePreset, getShapeSize } from './shape-read-base.ts';
import { commitAndRefresh } from './_helpers.ts';

export type { ConnectionSite } from '../../internal/drawingml/custom-geometry.ts';
export interface ShapeConnection {
  shapeId: number;
  siteIndex: number;
}
const dml = (name: string) => qname('a', name, NS.dml);
const pml = (name: string) => qname('p', name, NS.pml);
const attribute = (name: string) => qname('', name, '');
const definitions = new Map<string, XmlElement>();

/** Connection sites in unrotated, unflipped local EMU coordinates, in OOXML index order.
 * Supply resolved extents for placeholders inheriting their size. */
export function getShapeConnectionSites(
  shape: SlideShapeData,
  size = getShapeSize(shape),
): readonly ConnectionSite[] {
  if (!size || shape[SHAPE_SNAPSHOT].kind === 'connector') return [];
  const spPr = firstChildElement(shape[SHAPE_ELEMENT], pml('spPr'));
  if (!spPr) return [];
  const custom = firstChildElement(spPr, dml('custGeom'));
  if (custom) return parseConnectionSites(custom, size.w, size.h) ?? [];
  const preset = getShapePreset(shape);
  if (!preset || !Object.hasOwn(presetConnectionSites, preset)) return [];
  let definition = definitions.get(preset);
  if (!definition) {
    definition = parseXml(presetConnectionSites[preset]!).root;
    definitions.set(preset, definition);
  }
  const geometry = firstChildElement(spPr, dml('prstGeom'));
  return (
    parseConnectionSites(
      definition,
      size.w,
      size.h,
      geometry && firstChildElement(geometry, dml('avLst')),
    ) ?? []
  );
}

function properties(shape: SlideShapeData) {
  if (shape[SHAPE_SNAPSHOT].kind !== 'connector') return null;
  const nv = firstChildElement(shape[SHAPE_ELEMENT], pml('nvCxnSpPr'));
  return nv && firstChildElement(nv, pml('cNvCxnSpPr'));
}

/** Reads a connector's native start/end attachment, or null when disconnected. */
export function getShapeConnection(
  shape: SlideShapeData,
  end: 'start' | 'end',
): ShapeConnection | null {
  const parent = properties(shape);
  const connection = parent && firstChildElement(parent, dml(end === 'start' ? 'stCxn' : 'endCxn'));
  if (!connection) return null;
  const id = getAttrValue(connection, attribute('id'));
  const index = getAttrValue(connection, attribute('idx'));
  if (id === null || index === null) return null;
  const shapeId = Number(id),
    siteIndex = Number(index);
  return Number.isInteger(shapeId) && shapeId > 0 && Number.isInteger(siteIndex) && siteIndex >= 0
    ? { shapeId, siteIndex }
    : null;
}

/** Writes a native connector attachment. Passing null detaches that endpoint.
 * This stores topology; callers recompute geometry when the connected shape moves. */
export function setShapeConnection(
  shape: SlideShapeData,
  end: 'start' | 'end',
  connection: ShapeConnection | null,
): void {
  const parent = properties(shape);
  if (!parent) throw new Error('setShapeConnection: expected a connector');
  if (end !== 'start' && end !== 'end') throw new Error('Invalid connector endpoint.');
  if (connection !== null) {
    const { shapeId, siteIndex } = connection;
    if (
      !Number.isInteger(shapeId) ||
      shapeId <= 0 ||
      shapeId > 0xffffffff ||
      !Number.isInteger(siteIndex) ||
      siteIndex < 0 ||
      siteIndex > 0xffffffff
    )
      throw new Error('Invalid connector attachment.');
    const target = shape[SHAPE_SLIDE][SLIDE_SHAPES].find((s) => getShapeId(s) === shapeId);
    if (!target || target === shape || target[SHAPE_SNAPSHOT].kind === 'connector')
      throw new Error('Connection target must be a non-connector on the same slide.');
  }
  const name = end === 'start' ? 'stCxn' : 'endCxn';
  parent.children = parent.children.filter(
    (child) =>
      child.kind !== 'element' ||
      child.name.namespaceURI !== NS.dml ||
      child.name.localName !== name,
  );
  if (connection !== null) {
    const node = elem(dml(name), {
      attrs: [
        attr(attribute('id'), String(connection.shapeId)),
        attr(attribute('idx'), String(connection.siteIndex)),
      ],
    });
    // CT_NonVisualConnectorProperties: locks, start, end, extension list.
    const before = parent.children.findIndex(
      (child) =>
        child.kind === 'element' &&
        (child.name.localName === 'extLst' ||
          (end === 'start' && child.name.localName === 'endCxn')),
    );
    parent.children.splice(before < 0 ? parent.children.length : before, 0, node);
  }
  commitAndRefresh(shape);
}
