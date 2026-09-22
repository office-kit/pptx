// Builds the XML payload for a new slide from scratch.
//
// PowerPoint's "new slide from layout" emits a minimal `<p:sld>` whose
// shape tree contains only the placeholder stubs the layout exposes
// (title, body, footer, etc.). Each stub carries a `<p:ph>` element with
// the layout's same idx/type so the slide inherits geometry, fill, and
// default text style from the layout. The stub's `<p:spPr/>` is empty
// (no `<a:xfrm>`), and its `<p:txBody>` carries one empty `<a:p/>` so
// the placeholder is mutable via `setText` immediately.
//
// We do NOT copy the layout shape itself — that would freeze the
// inheritance and silently desync the slide from layout changes.

import {
  type XmlDocument,
  type XmlElement,
  type XmlAttr,
  NS,
  attr,
  elem,
  firstChildElement,
  getAttrValue,
  qname,
  text as textNode,
} from '../xml/index.ts';

const NAME_SLD = qname('p', 'sld', NS.pml);
const NAME_CSLD = qname('p', 'cSld', NS.pml);
const NAME_SP_TREE = qname('p', 'spTree', NS.pml);
const NAME_NV_GRP_SP_PR = qname('p', 'nvGrpSpPr', NS.pml);
const NAME_C_NV_PR = qname('p', 'cNvPr', NS.pml);
const NAME_C_NV_GRP_SP_PR = qname('p', 'cNvGrpSpPr', NS.pml);
const NAME_NV_PR = qname('p', 'nvPr', NS.pml);
const NAME_GRP_SP_PR = qname('p', 'grpSpPr', NS.pml);
const NAME_SP = qname('p', 'sp', NS.pml);
const NAME_NV_SP_PR = qname('p', 'nvSpPr', NS.pml);
const NAME_C_NV_SP_PR = qname('p', 'cNvSpPr', NS.pml);
const NAME_SP_PR = qname('p', 'spPr', NS.pml);
const NAME_TX_BODY = qname('p', 'txBody', NS.pml);
const NAME_PH = qname('p', 'ph', NS.pml);
const NAME_BODY_PR = qname('a', 'bodyPr', NS.dml);
const NAME_LST_STYLE = qname('a', 'lstStyle', NS.dml);
const NAME_P = qname('a', 'p', NS.dml);
const NAME_SP_LOCKS = qname('a', 'spLocks', NS.dml);
const ATTR_ID = qname('', 'id', '');
const ATTR_NAME = qname('', 'name', '');
const ATTR_TYPE = qname('', 'type', '');
const PLACEHOLDER_BINDING_ATTRIBUTES = new Set(['type', 'idx', 'orient', 'sz']);
const ATTR_NO_GRP = qname('', 'noGrp', '');

/**
 * Builds a `<p:sld>` document containing the canonical slide-root group
 * (id=1, name="") and one placeholder stub per `<p:ph>` found on
 * `layoutSpTree`.
 *
 * The stubs retain the layout's placeholder type, index, orientation and size, and
 * the names follow PowerPoint's `"Title 1"`, `"Content Placeholder 2"`
 * convention based on the placeholder type and an incrementing index.
 */
export const buildSlideFromLayout = (layoutSpTree: XmlElement): XmlDocument => {
  // Walk the layout's shape tree to find placeholders.
  const placeholders: XmlElement[] = [];
  for (const child of layoutSpTree.children) {
    if (child.kind !== 'element') continue;
    if (child.name.namespaceURI !== NS.pml || child.name.localName !== 'sp') continue;
    const nvSpPr = firstChildElement(child, NAME_NV_SP_PR);
    if (nvSpPr === null) continue;
    const nvPr = firstChildElement(nvSpPr, NAME_NV_PR);
    if (nvPr === null) continue;
    const ph = firstChildElement(nvPr, NAME_PH);
    if (ph === null) continue;
    placeholders.push(ph);
  }

  // Shape-id allocator: id=1 is the slide-root group, id=2+ are the
  // placeholder stubs. PowerPoint's emission order is `nvGrpSpPr` (root
  // group metadata) → `grpSpPr` (root group properties) → placeholder
  // shapes as siblings, so we mirror that.
  let nextShapeId = 2;
  const stubs: XmlElement[] = [];
  for (const ph of placeholders) {
    stubs.push(buildPlaceholderStub(nextShapeId, ph));
    nextShapeId++;
  }

  const spTree = elem(NAME_SP_TREE, {
    children: [buildNvGrpSpPr(1), buildGrpSpPr(), ...stubs],
  });
  const cSld = elem(NAME_CSLD, { children: [spTree] });
  const sld = elem(NAME_SLD, {
    prefixDecls: new Map([
      ['a', NS.dml],
      ['r', NS.officeDocRels],
      ['p', NS.pml],
    ]),
    children: [cSld],
  });
  return {
    kind: 'document',
    decl: { version: '1.0', encoding: 'UTF-8', standalone: 'yes' },
    prolog: [],
    root: sld,
    epilog: [],
  };
};

// `nvGrpSpPr` for the slide root group: id=1, name="" (PowerPoint requires
// the empty name; other writers' non-empty names get treated as quirky).
const buildNvGrpSpPr = (id: number): XmlElement =>
  elem(NAME_NV_GRP_SP_PR, {
    children: [
      elem(NAME_C_NV_PR, {
        attrs: [attr(ATTR_ID, String(id)), attr(ATTR_NAME, '')],
      }),
      elem(NAME_C_NV_GRP_SP_PR),
      elem(NAME_NV_PR),
    ],
  });

const buildGrpSpPr = (): XmlElement => elem(NAME_GRP_SP_PR);

// Returns a placeholder-shape stub: <p:sp> with cNvPr, nvSpPr, ph, empty
// spPr, and an empty txBody. Geometry, fill, and default style all flow
// from the corresponding layout placeholder via inheritance.
const buildPlaceholderStub = (id: number, layoutPlaceholder: XmlElement): XmlElement => {
  const phType = getAttrValue(layoutPlaceholder, ATTR_TYPE);
  // Prompt flags describe layout-only prompt content, not the empty slide body.
  const phAttrs: XmlAttr[] = layoutPlaceholder.attrs
    .filter(
      (a) => a.name.namespaceURI === '' && PLACEHOLDER_BINDING_ATTRIBUTES.has(a.name.localName),
    )
    .map((a) => attr(a.name, a.value));

  const ph = elem(NAME_PH, { attrs: phAttrs });
  const nvPr = elem(NAME_NV_PR, { children: [ph] });
  const cNvSpPr = elem(NAME_C_NV_SP_PR, {
    children: [elem(NAME_SP_LOCKS, { attrs: [attr(ATTR_NO_GRP, '1')] })],
  });
  const name = inferPlaceholderName(id, phType);
  const cNvPr = elem(NAME_C_NV_PR, {
    attrs: [attr(ATTR_ID, String(id)), attr(ATTR_NAME, name)],
  });
  const nvSpPr = elem(NAME_NV_SP_PR, { children: [cNvPr, cNvSpPr, nvPr] });
  const spPr = elem(NAME_SP_PR);
  const txBody = elem(NAME_TX_BODY, {
    children: [
      elem(NAME_BODY_PR),
      elem(NAME_LST_STYLE),
      elem(NAME_P, { children: [textNode('')] }),
    ],
  });
  return elem(NAME_SP, { children: [nvSpPr, spPr, txBody] });
};

/**
 * Maps a placeholder type to a human-readable name PowerPoint emits. The
 * `id` becomes part of the name to keep them unique within the slide.
 */
const inferPlaceholderName = (id: number, phType: string | null): string => {
  switch (phType) {
    case 'title':
      return `Title ${id - 1}`;
    case 'ctrTitle':
      return `Centered Title ${id - 1}`;
    case 'subTitle':
      return `Subtitle ${id - 1}`;
    case 'body':
      return `Content Placeholder ${id - 1}`;
    case 'pic':
      return `Picture Placeholder ${id - 1}`;
    case 'chart':
      return `Chart Placeholder ${id - 1}`;
    case 'tbl':
      return `Table Placeholder ${id - 1}`;
    case 'dt':
      return `Date Placeholder ${id - 1}`;
    case 'sldNum':
      return `Slide Number Placeholder ${id - 1}`;
    case 'ftr':
      return `Footer Placeholder ${id - 1}`;
    case 'hdr':
      return `Header Placeholder ${id - 1}`;
    default:
      return `Placeholder ${id - 1}`;
  }
};
