// Read-only view over a slide-layout part (`/ppt/slideLayouts/slideLayoutN.xml`).
//
// ECMA-376 Part 1 §19.3.1.39 — `<p:sldLayout>` is the same shape as `<p:sld>`
// (it wraps a `<p:cSld>` with a `<p:spTree>`) plus a layout-specific `type`
// attribute and an optional `<p:cSld name="...">` that user-visible code
// looks at (PowerPoint shows it in the layout picker).
//
// Layouts inherit non-overridden placement and formatting from their slide
// master; in turn, each slide inherits from its layout. We expose the raw
// per-layout shape tree here; effective inheritance resolution belongs to
// a higher layer that has the full graph.

import { NS, firstChildElement, getAttrValue, qname } from '../xml/index.ts';
import type { XmlElement } from '../xml/index.ts';
import { readShapeTreeFromCsldRoot, type SlideShape } from './slide-part.ts';

const NAME_CSLD = qname('p', 'cSld', NS.pml);
const ATTR_NAME = qname('', 'name', '');
const ATTR_TYPE = qname('', 'type', '');

/**
 * Layout type token per ECMA-376 §19.7.15 `ST_SlideLayoutType`. Listed for
 * autocomplete; layouts in real templates sometimes use vendor-specific
 * tokens not in this enum, so the field stays a plain string at the API.
 */
export type SlideLayoutType =
  | 'title'
  | 'tx'
  | 'twoColTx'
  | 'tbl'
  | 'txAndChart'
  | 'chartAndTx'
  | 'dgm'
  | 'chart'
  | 'txAndClipArt'
  | 'clipArtAndTx'
  | 'titleOnly'
  | 'blank'
  | 'txAndObj'
  | 'objAndTx'
  | 'objOnly'
  | 'obj'
  | 'txAndMedia'
  | 'mediaAndTx'
  | 'objOverTx'
  | 'txOverObj'
  | 'txAndTwoObj'
  | 'twoObjAndTx'
  | 'twoObjOverTx'
  | 'fourObj'
  | 'vertTx'
  | 'clipArtAndVertTx'
  | 'vertTitleAndTx'
  | 'vertTitleAndTxOverChart'
  | 'twoObj'
  | 'objAndTwoObj'
  | 'twoObjAndObj'
  | 'cust'
  | 'secHead'
  | 'twoTxTwoObj'
  | 'objTx'
  | 'picTx';

export interface SlideLayoutPart {
  /** Human-visible layout name from `<p:cSld name="...">`. */
  readonly name: string;
  /**
   * Layout type token. `null` when the attribute is absent — the spec
   * default in that case is `cust` (custom).
   */
  readonly layoutType: string | null;
  /** Shapes on the layout in document order, groups flattened. */
  readonly shapes: ReadonlyArray<SlideShape>;
  readonly root: XmlElement;
}

/** Parses a `<p:sldLayout>` root element. */
export const readSlideLayoutPart = (
  root: XmlElement,
  options: { recurseIntoGroups?: boolean } = {},
): SlideLayoutPart => {
  const { shapes } = readShapeTreeFromCsldRoot(root, 'sldLayout', options);
  const cSld = firstChildElement(root, NAME_CSLD);
  const name = cSld !== null ? getAttrValue(cSld, ATTR_NAME) : null;
  const layoutType = getAttrValue(root, ATTR_TYPE);
  return {
    name: name ?? '',
    layoutType,
    shapes,
    root,
  };
};
