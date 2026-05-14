// Read-only view over `/ppt/presentation.xml`.
//
// ECMA-376 Part 1 §19.2.1.26 defines the root `p:presentation` element:
//
//   <p:presentation xmlns:p="...">
//     <p:sldMasterIdLst><p:sldMasterId id="..." r:id="rId1"/>...</p:sldMasterIdLst>
//     <p:sldIdLst><p:sldId id="256" r:id="rId2"/>...</p:sldIdLst>
//     <p:notesMasterIdLst><p:notesMasterId r:id="..."/></p:notesMasterIdLst>
//     <p:sldSz cx="9144000" cy="6858000" type="screen4x3"/>
//     <p:notesSz cx="..." cy="..."/>
//     <p:defaultTextStyle>...</p:defaultTextStyle>
//   </p:presentation>
//
// We only model the structural list elements at this phase. Default text
// style cascade is the subject of P3+ — too much surface area to bake in
// before the drawingml layer is ready.

import {
  type XmlElement,
  NS,
  allChildElements,
  firstChildElement,
  getAttrValue,
  qname,
} from '../xml/index.ts';

export interface SlideMasterId {
  /** Allocated id, ≥2147483648 per ECMA-376 §19.2.1.36. */
  readonly id: number;
  /** Relationship id pointing at the slideMaster part. */
  readonly rId: string;
}

export interface SlideId {
  /** Allocated id, in [256, 2³¹-1024] per the spec + PowerPoint quirk. */
  readonly id: number;
  /** Relationship id pointing at the slide part. */
  readonly rId: string;
}

export interface NotesMasterId {
  readonly rId: string;
}

export interface SlideSize {
  /** Width in EMU. */
  readonly cx: number;
  /** Height in EMU. */
  readonly cy: number;
  /**
   * Aspect-ratio hint. PowerPoint emits one of the standard tokens —
   * `screen4x3`, `screen16x9`, `letter`, `A4`, `custom`, etc. — but the
   * value is informational; the actual size is `cx`/`cy`.
   */
  readonly type?: string;
}

export interface PresentationPart {
  readonly slideMasters: ReadonlyArray<SlideMasterId>;
  readonly slides: ReadonlyArray<SlideId>;
  readonly notesMaster: NotesMasterId | null;
  readonly slideSize: SlideSize | null;
  readonly notesSize: SlideSize | null;
}

const NAME_PRESENTATION = qname('p', 'presentation', NS.pml);
const NAME_SLD_MASTER_ID_LST = qname('p', 'sldMasterIdLst', NS.pml);
const NAME_SLD_MASTER_ID = qname('p', 'sldMasterId', NS.pml);
const NAME_SLD_ID_LST = qname('p', 'sldIdLst', NS.pml);
const NAME_SLD_ID = qname('p', 'sldId', NS.pml);
const NAME_NOTES_MASTER_ID_LST = qname('p', 'notesMasterIdLst', NS.pml);
const NAME_NOTES_MASTER_ID = qname('p', 'notesMasterId', NS.pml);
const NAME_SLD_SZ = qname('p', 'sldSz', NS.pml);
const NAME_NOTES_SZ = qname('p', 'notesSz', NS.pml);
const ATTR_ID = qname('', 'id', '');
const ATTR_R_ID = qname('r', 'id', NS.officeDocRels);
const ATTR_CX = qname('', 'cx', '');
const ATTR_CY = qname('', 'cy', '');
const ATTR_TYPE = qname('', 'type', '');

const requireRId = (element: XmlElement, context: string): string => {
  const value = getAttrValue(element, ATTR_R_ID);
  if (value === null) {
    throw new Error(`${context}: missing r:id attribute`);
  }
  return value;
};

const readSize = (parent: XmlElement, name: typeof NAME_SLD_SZ): SlideSize | null => {
  const el = firstChildElement(parent, name);
  if (el === null) return null;
  const cxRaw = getAttrValue(el, ATTR_CX);
  const cyRaw = getAttrValue(el, ATTR_CY);
  if (cxRaw === null || cyRaw === null) {
    throw new Error(`<${name.localName}>: missing cx/cy attribute`);
  }
  const type = getAttrValue(el, ATTR_TYPE);
  return {
    cx: Number.parseInt(cxRaw, 10),
    cy: Number.parseInt(cyRaw, 10),
    ...(type !== null ? { type } : {}),
  };
};

/**
 * Parses a `p:presentation` root element into the typed view above. Throws
 * if the root is not `p:presentation` in the PresentationML namespace.
 */
export const readPresentationPart = (root: XmlElement): PresentationPart => {
  if (
    root.name.namespaceURI !== NAME_PRESENTATION.namespaceURI ||
    root.name.localName !== 'presentation'
  ) {
    throw new Error(`expected <p:presentation>, got <${root.name.prefix}:${root.name.localName}>`);
  }

  const sldMasters: SlideMasterId[] = [];
  const sldMasterLst = firstChildElement(root, NAME_SLD_MASTER_ID_LST);
  if (sldMasterLst !== null) {
    for (const item of allChildElements(sldMasterLst, NAME_SLD_MASTER_ID)) {
      const idRaw = getAttrValue(item, ATTR_ID);
      if (idRaw === null) {
        throw new Error('<p:sldMasterId>: missing id attribute');
      }
      sldMasters.push({
        id: Number.parseInt(idRaw, 10),
        rId: requireRId(item, '<p:sldMasterId>'),
      });
    }
  }

  const slides: SlideId[] = [];
  const sldLst = firstChildElement(root, NAME_SLD_ID_LST);
  if (sldLst !== null) {
    for (const item of allChildElements(sldLst, NAME_SLD_ID)) {
      const idRaw = getAttrValue(item, ATTR_ID);
      if (idRaw === null) {
        throw new Error('<p:sldId>: missing id attribute');
      }
      slides.push({
        id: Number.parseInt(idRaw, 10),
        rId: requireRId(item, '<p:sldId>'),
      });
    }
  }

  let notesMaster: NotesMasterId | null = null;
  const notesMasterLst = firstChildElement(root, NAME_NOTES_MASTER_ID_LST);
  if (notesMasterLst !== null) {
    const item = firstChildElement(notesMasterLst, NAME_NOTES_MASTER_ID);
    if (item !== null) {
      notesMaster = { rId: requireRId(item, '<p:notesMasterId>') };
    }
  }

  return {
    slideMasters: sldMasters,
    slides,
    notesMaster,
    slideSize: readSize(root, NAME_SLD_SZ),
    notesSize: readSize(root, NAME_NOTES_SZ),
  };
};
