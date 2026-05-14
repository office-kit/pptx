// Relationships parts (`*.rels`) per OPC §9.3.
//
// Structure:
//
//   <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
//     <Relationship Id="rId1" Type="http://...." Target="slides/slide1.xml"/>
//     <Relationship Id="rId2" Type="http://...." Target="https://x" TargetMode="External"/>
//   </Relationships>
//
// `Id` is unique within the part; PPTX uses `rId<N>` by convention but the
// schema accepts any xsd:ID-valid string. `Type` is a URI identifying the
// relationship semantics. `Target` is a URI: relative to the part's location
// when `TargetMode` is `Internal` (default), or absolute when `External`.

import { NS, type XmlDocument, elem, parseXml, qname, serializeXml } from '../xml/index.ts';

export type TargetMode = 'Internal' | 'External';

export interface Relationship {
  id: string;
  type: string;
  target: string;
  /** Defaults to `'Internal'` when omitted from the XML. */
  targetMode: TargetMode;
}

export interface Relationships {
  items: Relationship[];
}

const REL_NS = NS.relationships;
const NAME_RELATIONSHIPS = qname('', 'Relationships', REL_NS);
const NAME_RELATIONSHIP = qname('', 'Relationship', REL_NS);
const ATTR_ID = qname('', 'Id', '');
const ATTR_TYPE = qname('', 'Type', '');
const ATTR_TARGET = qname('', 'Target', '');
const ATTR_TARGET_MODE = qname('', 'TargetMode', '');

const findAttr = (
  e: { attrs: ReadonlyArray<{ name: { localName: string }; value: string }> },
  local: string,
): string | null => {
  for (const a of e.attrs) {
    if (a.name.localName === local) return a.value;
  }
  return null;
};

export const parseRels = (xml: string): Relationships => {
  const doc = parseXml(xml);
  if (doc.root.name.namespaceURI !== REL_NS || doc.root.name.localName !== 'Relationships') {
    throw new Error(
      `expected root <Relationships xmlns="${REL_NS}">, got <${doc.root.name.localName}>`,
    );
  }
  const items: Relationship[] = [];
  for (const child of doc.root.children) {
    if (child.kind !== 'element') continue;
    if (child.name.namespaceURI !== REL_NS || child.name.localName !== 'Relationship') {
      continue;
    }
    const id = findAttr(child, 'Id');
    const type = findAttr(child, 'Type');
    const target = findAttr(child, 'Target');
    if (id === null || type === null || target === null) {
      throw new Error('Relationship missing Id, Type, or Target');
    }
    const modeRaw = findAttr(child, 'TargetMode');
    let targetMode: TargetMode = 'Internal';
    if (modeRaw !== null) {
      if (modeRaw !== 'Internal' && modeRaw !== 'External') {
        throw new Error(`invalid TargetMode "${modeRaw}"`);
      }
      targetMode = modeRaw;
    }
    items.push({ id, type, target, targetMode });
  }
  return { items };
};

export const serializeRels = (rels: Relationships): string => {
  const children = rels.items.map((r) => {
    const attrs = [
      { name: ATTR_ID, value: r.id },
      { name: ATTR_TYPE, value: r.type },
      { name: ATTR_TARGET, value: r.target },
    ];
    // OOXML convention: only emit TargetMode when External. Tools that
    // do otherwise sometimes get flagged by Office's diff comparison.
    if (r.targetMode === 'External') {
      attrs.push({ name: ATTR_TARGET_MODE, value: 'External' });
    }
    return elem(NAME_RELATIONSHIP, { attrs });
  });
  const root = elem(NAME_RELATIONSHIPS, {
    prefixDecls: new Map([['', REL_NS]]),
    children,
  });
  const doc: XmlDocument = {
    kind: 'document',
    decl: { version: '1.0', encoding: 'UTF-8', standalone: 'yes' },
    prolog: [],
    root,
    epilog: [],
  };
  return serializeXml(doc);
};

/**
 * Convenience: an empty Relationships value with no items.
 */
export const emptyRels = (): Relationships => ({ items: [] });

/**
 * Returns the next free `rId<N>` id given a list of existing ids. The OPC
 * spec doesn't require this naming pattern, but PowerPoint emits it and our
 * authoring code follows suit so output diffs cleanly against PowerPoint.
 */
export const nextRelId = (existing: ReadonlyArray<string>): string => {
  let max = 0;
  for (const id of existing) {
    const m = id.match(/^rId(\d+)$/);
    if (m?.[1] !== undefined) {
      const n = Number.parseInt(m[1], 10);
      if (Number.isFinite(n) && n > max) max = n;
    }
  }
  return `rId${max + 1}`;
};
