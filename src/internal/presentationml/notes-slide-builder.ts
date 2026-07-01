// Builds the canonical empty notes-slide part PowerPoint emits.
//
// The part's root element is `<p:notes>` (ECMA-376 §19.3.1.26 — the global
// element for CT_NotesSlide is named `notes`, not `notesSlide`; the
// `notesSlide` token only appears in the part name and content type). Emitting
// `<p:notesSlide>` here makes the part fail pml.xsd validation.
//
// Notes slides carry two placeholders: a `sldImg` placeholder that
// PowerPoint renders as the slide thumbnail, and a `body` placeholder
// (idx="1") that holds the speaker-notes text. We emit both with no
// geometry of their own so they inherit position from the notes master.

import {
  type XmlDocument,
  type XmlElement,
  NS,
  attr,
  elem,
  qname,
  text as textNode,
} from '../xml/index.ts';

const NAME_NOTES_SLIDE = qname('p', 'notes', NS.pml);
const NAME_CSLD = qname('p', 'cSld', NS.pml);
const NAME_SP_TREE = qname('p', 'spTree', NS.pml);
const NAME_NV_GRP_SP_PR = qname('p', 'nvGrpSpPr', NS.pml);
const NAME_C_NV_PR = qname('p', 'cNvPr', NS.pml);
const NAME_C_NV_GRP_SP_PR = qname('p', 'cNvGrpSpPr', NS.pml);
const NAME_C_NV_SP_PR = qname('p', 'cNvSpPr', NS.pml);
const NAME_NV_PR = qname('p', 'nvPr', NS.pml);
const NAME_GRP_SP_PR = qname('p', 'grpSpPr', NS.pml);
const NAME_SP = qname('p', 'sp', NS.pml);
const NAME_NV_SP_PR = qname('p', 'nvSpPr', NS.pml);
const NAME_SP_PR = qname('p', 'spPr', NS.pml);
const NAME_PH = qname('p', 'ph', NS.pml);
const NAME_TX_BODY = qname('p', 'txBody', NS.pml);
const NAME_CLR_MAP_OVR = qname('p', 'clrMapOvr', NS.pml);
const NAME_MASTER_CLR_MAPPING = qname('a', 'masterClrMapping', NS.dml);
const NAME_BODY_PR = qname('a', 'bodyPr', NS.dml);
const NAME_LST_STYLE = qname('a', 'lstStyle', NS.dml);
const NAME_P = qname('a', 'p', NS.dml);
const NAME_R = qname('a', 'r', NS.dml);
const NAME_RPR = qname('a', 'rPr', NS.dml);
const NAME_T = qname('a', 't', NS.dml);
const NAME_SP_LOCKS = qname('a', 'spLocks', NS.dml);
const ATTR_ID = qname('', 'id', '');
const ATTR_NAME = qname('', 'name', '');
const ATTR_NO_GRP = qname('', 'noGrp', '');
const ATTR_NO_ROT = qname('', 'noRot', '');
const ATTR_NO_CHANGE_ASPECT = qname('', 'noChangeAspect', '');
const ATTR_TYPE = qname('', 'type', '');
const ATTR_IDX = qname('', 'idx', '');
const ATTR_LANG = qname('', 'lang', '');

const buildRootGroup = (): XmlElement => {
  const cNvPr = elem(NAME_C_NV_PR, { attrs: [attr(ATTR_ID, '1'), attr(ATTR_NAME, '')] });
  return elem(NAME_NV_GRP_SP_PR, {
    children: [cNvPr, elem(NAME_C_NV_GRP_SP_PR), elem(NAME_NV_PR)],
  });
};

const buildSldImgPlaceholder = (id: number): XmlElement => {
  const cNvPr = elem(NAME_C_NV_PR, {
    attrs: [attr(ATTR_ID, String(id)), attr(ATTR_NAME, `Slide Image Placeholder ${id - 1}`)],
  });
  const spLocks = elem(NAME_SP_LOCKS, {
    attrs: [attr(ATTR_NO_GRP, '1'), attr(ATTR_NO_ROT, '1'), attr(ATTR_NO_CHANGE_ASPECT, '1')],
  });
  const cNvSpPr = elem(NAME_C_NV_SP_PR, { children: [spLocks] });
  const ph = elem(NAME_PH, { attrs: [attr(ATTR_TYPE, 'sldImg')] });
  const nvSpPr = elem(NAME_NV_SP_PR, {
    children: [cNvPr, cNvSpPr, elem(NAME_NV_PR, { children: [ph] })],
  });
  return elem(NAME_SP, { children: [nvSpPr, elem(NAME_SP_PR)] });
};

const buildNotesParagraphs = (notes: string): XmlElement[] => {
  const lines = notes.split('\n');
  return lines.map((line) => {
    const t = elem(NAME_T, { children: line.length > 0 ? [textNode(line)] : [] });
    const r = elem(NAME_R, {
      children: [elem(NAME_RPR, { attrs: [attr(ATTR_LANG, 'en-US')] }), t],
    });
    return elem(NAME_P, { children: [r] });
  });
};

const buildNotesBodyPlaceholder = (id: number, notes: string): XmlElement => {
  const cNvPr = elem(NAME_C_NV_PR, {
    attrs: [attr(ATTR_ID, String(id)), attr(ATTR_NAME, `Notes Placeholder ${id - 1}`)],
  });
  const cNvSpPr = elem(NAME_C_NV_SP_PR, {
    children: [elem(NAME_SP_LOCKS, { attrs: [attr(ATTR_NO_GRP, '1')] })],
  });
  const ph = elem(NAME_PH, {
    attrs: [attr(ATTR_TYPE, 'body'), attr(ATTR_IDX, '1')],
  });
  const nvSpPr = elem(NAME_NV_SP_PR, {
    children: [cNvPr, cNvSpPr, elem(NAME_NV_PR, { children: [ph] })],
  });
  const txBody = elem(NAME_TX_BODY, {
    children: [elem(NAME_BODY_PR), elem(NAME_LST_STYLE), ...buildNotesParagraphs(notes)],
  });
  return elem(NAME_SP, { children: [nvSpPr, elem(NAME_SP_PR), txBody] });
};

/**
 * Returns a fresh `<p:notes>` document with the given notes text in
 * the body placeholder. Designed for callers that need to create the
 * notesSlide part from scratch (no existing notes file yet for the
 * slide).
 */
export const buildEmptyNotesSlide = (notes: string): XmlDocument => {
  const spTree = elem(NAME_SP_TREE, {
    children: [
      buildRootGroup(),
      elem(NAME_GRP_SP_PR),
      buildSldImgPlaceholder(2),
      buildNotesBodyPlaceholder(3, notes),
    ],
  });
  const cSld = elem(NAME_CSLD, { children: [spTree] });
  const clrMapOvr = elem(NAME_CLR_MAP_OVR, { children: [elem(NAME_MASTER_CLR_MAPPING)] });
  const root = elem(NAME_NOTES_SLIDE, {
    prefixDecls: new Map([
      ['a', NS.dml],
      ['r', NS.officeDocRels],
      ['p', NS.pml],
    ]),
    children: [cSld, clrMapOvr],
  });
  return {
    kind: 'document',
    decl: { version: '1.0', encoding: 'UTF-8', standalone: 'yes' },
    prolog: [],
    root,
    epilog: [],
  };
};
