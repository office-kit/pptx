// The notes master and the handout master (ECMA-376 Part 1 §19.3.1.27
// `notesMaster`, §19.3.1.24 `handoutMaster`): the printed-page counterparts of
// the slide master. The reference desktop app's Notes Master and Handout Master views edit
// their placeholders, the page orientation and the handout's slides per page.
//
// A presentation has at most one of each (`<p:notesMasterIdLst>` and
// `<p:handoutMasterIdLst>` hold one entry). Neither is required, so the
// setters create the master the reference desktop app would — its default placeholders and
// its own copy of the slide theme — the first time it is edited.

import {
  type PartName,
  emptyRels,
  nextRelId,
  partName,
  resolveTarget,
} from '../../internal/opc/index.ts';
import type { OpcPackage } from '../../internal/parts/index.ts';
import {
  HANDOUT_MASTER_PLACEHOLDER_TYPES,
  type HandoutMasterPlaceholderType,
  NOTES_MASTER_PLACEHOLDER_TYPES,
  type NotesMasterPlaceholderType,
  REL_TYPES,
  defaultHandoutMasterXml,
  defaultNotesMasterXml,
  handoutMasterPlaceholderXml,
  notesMasterPlaceholderXml,
  readShapeTreeFromCsldRoot,
} from '../../internal/presentationml/index.ts';
import {
  NS,
  type XmlElement,
  attr,
  elem,
  firstChildElement,
  getAttrValue,
  qname,
} from '../../internal/xml/index.ts';
import type { Emu } from '../units.ts';
import { INTERNAL_PACKAGE, type PresentationData } from '../_internal-symbols.ts';
import {
  ATTR_R_ID,
  NAME_SLD_MASTER_ID_LST,
  PRES_PART_NAME,
  allocatePartNames,
  encode,
  nextShapeIdInTree,
  relativeTarget,
} from './_helpers.ts';
import {
  type SlideLayoutPlaceholder,
  commitPartDoc,
  hasPlaceholderOfType,
  insertPlaceholderInOrder,
  placeholderElement,
  placeholderTypeOf,
  placeholderViews,
  readPartDoc,
  removePlaceholdersOfType,
  requirePartSpTree,
} from './_part-placeholders.ts';
import { readPresentationProperties, writePresentationProperties } from './show-properties.ts';
import { getSlideSize } from './slide-size.ts';

export type { HandoutMasterPlaceholderType, NotesMasterPlaceholderType };

const NOTES_MASTER_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml';
const HANDOUT_MASTER_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.presentationml.handoutMaster+xml';
const NOTES_SLIDE_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml';
const THEME_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.theme+xml';

const NAME_NOTES_SZ = qname('p', 'notesSz', NS.pml);
const NAME_PRN_PR = qname('p', 'prnPr', NS.pml);
const ATTR_CX = qname('', 'cx', '');
const ATTR_CY = qname('', 'cy', '');
const ATTR_PRN_WHAT = qname('', 'prnWhat', '');

// `<p:notesSz>` is required, but a hand-built deck may still omit it; the reference desktop app
// then lays notes and handouts out on its 7.5 × 10 in portrait page.
const DEFAULT_PAGE = { w: 6858000, h: 9144000 };
const DEFAULT_SLIDE = { w: 12192000, h: 6858000 };

type PageMasterKind = 'notesMaster' | 'handoutMaster';

const KIND = {
  notesMaster: {
    rel: REL_TYPES.notesMaster,
    contentType: NOTES_MASTER_CONTENT_TYPE,
    stem: '/ppt/notesMasters/notesMaster',
    list: qname('p', 'notesMasterIdLst', NS.pml),
    entry: qname('p', 'notesMasterId', NS.pml),
    // CT_Presentation: sldMasterIdLst, notesMasterIdLst, handoutMasterIdLst, …
    after: [NAME_SLD_MASTER_ID_LST],
  },
  handoutMaster: {
    rel: REL_TYPES.handoutMaster,
    contentType: HANDOUT_MASTER_CONTENT_TYPE,
    stem: '/ppt/handoutMasters/handoutMaster',
    list: qname('p', 'handoutMasterIdLst', NS.pml),
    entry: qname('p', 'handoutMasterId', NS.pml),
    after: [NAME_SLD_MASTER_ID_LST, qname('p', 'notesMasterIdLst', NS.pml)],
  },
} as const;

const relTarget = (from: PartName, target: string): PartName =>
  target.startsWith('/') ? partName(target) : resolveTarget(from, target);

/** The master's part, through the presentation's relationship, or `null`. */
const pageMasterName = (pkg: OpcPackage, kind: PageMasterKind): PartName | null => {
  const rel = pkg.getRels(PRES_PART_NAME)?.items.find((r) => r.type === KIND[kind].rel);
  if (!rel) return null;
  const name = relTarget(PRES_PART_NAME, rel.target);
  return pkg.getPart(name) === null ? null : name;
};

const pageSize = (pkg: OpcPackage): { w: number; h: number } => {
  const notesSz = firstChildElement(readPartDoc(pkg, PRES_PART_NAME).doc.root, NAME_NOTES_SZ);
  const w = Number(notesSz && getAttrValue(notesSz, ATTR_CX));
  const h = Number(notesSz && getAttrValue(notesSz, ATTR_CY));
  return w > 0 && h > 0 ? { w, h } : DEFAULT_PAGE;
};

const slideCanvas = (pres: PresentationData): { w: number; h: number } => {
  const size = getSlideSize(pres);
  return size ? { w: size.width, h: size.height } : DEFAULT_SLIDE;
};

/** The slide theme, which a new page master copies as the reference desktop app does. */
const slideTheme = (pkg: OpcPackage, presentation: XmlElement): Uint8Array | null => {
  const rels = pkg.getRels(PRES_PART_NAME)?.items ?? [];
  const firstMaster = firstChildElement(presentation, NAME_SLD_MASTER_ID_LST)?.children.find(
    (c): c is XmlElement => c.kind === 'element',
  );
  const masterRel =
    rels.find((r) => firstMaster && r.id === getAttrValue(firstMaster, ATTR_R_ID)) ??
    rels.find((r) => r.type === REL_TYPES.slideMaster);
  const owner = masterRel ? relTarget(PRES_PART_NAME, masterRel.target) : PRES_PART_NAME;
  const themeRel = pkg.getRels(owner)?.items.find((r) => r.type === REL_TYPES.theme);
  const theme = themeRel && pkg.getPart(relTarget(owner, themeRel.target));
  return theme ? theme.data : null;
};

/** The master's part name, creating the reference desktop app's default master first when absent. */
const ensurePageMaster = (pres: PresentationData, kind: PageMasterKind): PartName => {
  const pkg = pres[INTERNAL_PACKAGE];
  const existing = pageMasterName(pkg, kind);
  if (existing !== null) return existing;

  const spec = KIND[kind];
  const presentation = readPartDoc(pkg, PRES_PART_NAME);
  const page = pageSize(pkg);
  const [name] = allocatePartNames(pkg, spec.stem, 1) as [PartName];
  const xml =
    kind === 'notesMaster'
      ? defaultNotesMasterXml(page, slideCanvas(pres))
      : defaultHandoutMasterXml(page);
  pkg.addPart(name, spec.contentType, encode(xml));

  // Every master needs a theme (§19.3.1.27 / §19.3.1.24 parts list one).
  const themeBytes = slideTheme(pkg, presentation.doc.root);
  if (themeBytes === null) throw new Error(`${kind}: the presentation has no theme to copy`);
  const [themeName] = allocatePartNames(pkg, '/ppt/theme/theme', 1) as [PartName];
  pkg.addPart(themeName, THEME_CONTENT_TYPE, themeBytes.slice());
  pkg.setRels(name, {
    items: [
      {
        id: 'rId1',
        type: REL_TYPES.theme,
        target: relativeTarget(name, themeName),
        targetMode: 'Internal',
      },
    ],
  });

  const presRels = pkg.getRels(PRES_PART_NAME) ?? emptyRels();
  const rId = nextRelId(presRels.items.map((r) => r.id));
  presRels.items.push({
    id: rId,
    type: spec.rel,
    target: relativeTarget(PRES_PART_NAME, name),
    targetMode: 'Internal',
  });
  pkg.setRels(PRES_PART_NAME, presRels);
  const root = presentation.doc.root;
  let list = firstChildElement(root, spec.list);
  if (list === null) {
    list = elem(spec.list);
    let at = 0;
    for (const before of spec.after) {
      const found = firstChildElement(root, before);
      if (found) at = root.children.indexOf(found) + 1;
    }
    root.children.splice(at, 0, list);
  }
  list.children = [elem(spec.entry, { attrs: [attr(ATTR_R_ID, rId)] })];
  commitPartDoc(presentation);

  // Notes slides made before the deck had a notes master point at none; the
  // notes-slide part is meant to relate to the notes master (Part 1 §13.3.5).
  if (kind === 'notesMaster') {
    for (const part of pkg.parts) {
      if (part.contentType !== NOTES_SLIDE_CONTENT_TYPE) continue;
      const rels = pkg.getRels(part.name) ?? emptyRels();
      if (rels.items.some((r) => r.type === REL_TYPES.notesMaster)) continue;
      rels.items.push({
        id: nextRelId(rels.items.map((r) => r.id)),
        type: REL_TYPES.notesMaster,
        target: relativeTarget(part.name, name),
        targetMode: 'Internal',
      });
      pkg.setRels(part.name, rels);
    }
  }
  return name;
};

const readPlaceholders = (
  pres: PresentationData,
  kind: PageMasterKind,
): ReadonlyArray<SlideLayoutPlaceholder> | null => {
  const pkg = pres[INTERNAL_PACKAGE];
  const name = pageMasterName(pkg, kind);
  if (name === null) return null;
  return placeholderViews(readShapeTreeFromCsldRoot(readPartDoc(pkg, name).doc.root, kind).shapes);
};

/**
 * The notes master's placeholders (header, date, slide image, notes body,
 * footer, slide number) and where they sit on the page, or `null` when the
 * presentation has no notes master.
 */
export const getNotesMasterPlaceholders = (
  pres: PresentationData,
): ReadonlyArray<SlideLayoutPlaceholder> | null => readPlaceholders(pres, 'notesMaster');

/**
 * The handout master's placeholders (header, date, footer, page number), or
 * `null` when the presentation has no handout master.
 */
export const getHandoutMasterPlaceholders = (
  pres: PresentationData,
): ReadonlyArray<SlideLayoutPlaceholder> | null => readPlaceholders(pres, 'handoutMaster');

const setIncluded = (
  pres: PresentationData,
  kind: PageMasterKind,
  type: string,
  order: ReadonlyArray<string>,
  build: (id: number) => string,
  included: boolean,
): void => {
  if (!order.includes(type)) throw new Error(`${kind}: unknown placeholder type ${type}`);
  const target = readPartDoc(pres[INTERNAL_PACKAGE], ensurePageMaster(pres, kind));
  const spTree = requirePartSpTree(target.doc.root);
  const tokens = new Set([type]);
  if (hasPlaceholderOfType(spTree, tokens) !== included) {
    if (included)
      insertPlaceholderInOrder(
        spTree,
        placeholderElement(build(nextShapeIdInTree(spTree))),
        type,
        order,
      );
    else removePlaceholdersOfType(spTree, tokens);
  }
  commitPartDoc(target);
};

/**
 * The Notes Master tab's placeholder checkboxes: removes a placeholder from
 * the notes master, or puts it back where the reference desktop app's default puts it.
 * Creates the default notes master when the presentation has none.
 */
export const setNotesMasterPlaceholderIncluded = (
  pres: PresentationData,
  type: NotesMasterPlaceholderType,
  included: boolean,
): void => {
  const page = pageSize(pres[INTERNAL_PACKAGE]);
  setIncluded(
    pres,
    'notesMaster',
    type,
    NOTES_MASTER_PLACEHOLDER_TYPES,
    (id) => notesMasterPlaceholderXml(type, id, page, slideCanvas(pres)),
    included,
  );
};

/** The Handout Master tab's placeholder checkboxes; see `setNotesMasterPlaceholderIncluded`. */
export const setHandoutMasterPlaceholderIncluded = (
  pres: PresentationData,
  type: HandoutMasterPlaceholderType,
  included: boolean,
): void => {
  const page = pageSize(pres[INTERNAL_PACKAGE]);
  setIncluded(
    pres,
    'handoutMaster',
    type,
    HANDOUT_MASTER_PLACEHOLDER_TYPES,
    (id) => handoutMasterPlaceholderXml(type, id, page),
    included,
  );
};

// ---------------------------------------------------------------------------
// Page size and orientation.

/**
 * The notes and handout page (`<p:notesSz>`), in EMU. Notes pages, handouts
 * and the outline print on it; it is portrait unless it is wider than tall.
 */
export const getNotesPageSize = (pres: PresentationData): { width: Emu; height: Emu } => {
  const page = pageSize(pres[INTERNAL_PACKAGE]);
  return { width: page.w as Emu, height: page.h as Emu };
};

const NAME_XFRM = qname('a', 'xfrm', NS.dml);
const NAME_OFF = qname('a', 'off', NS.dml);
const NAME_EXT = qname('a', 'ext', NS.dml);
const NAME_SP_PR = qname('p', 'spPr', NS.pml);

/**
 * Moves every top-level box of a page part to the turned page: boxes scale
 * per axis, except the slide image, which keeps the slide's proportions and
 * stays centred where it was.
 */
const turnPage = (root: XmlElement, sx: number, sy: number): void => {
  for (const shape of requirePartSpTree(root).children) {
    if (shape.kind !== 'element') continue;
    const spPr = firstChildElement(shape, NAME_SP_PR);
    const xfrm = spPr && firstChildElement(spPr, NAME_XFRM);
    const off = xfrm && firstChildElement(xfrm, NAME_OFF);
    const ext = xfrm && firstChildElement(xfrm, NAME_EXT);
    if (!off || !ext) continue;
    const x = Number(getAttrValue(off, qname('', 'x', '')));
    const y = Number(getAttrValue(off, qname('', 'y', '')));
    const w = Number(getAttrValue(ext, ATTR_CX));
    const h = Number(getAttrValue(ext, ATTR_CY));
    let box: [number, number, number, number];
    if (placeholderTypeOf(shape) === 'sldImg') {
      const s = Math.min(sx, sy);
      const cx = (x + w / 2) * sx;
      const cy = (y + h / 2) * sy;
      box = [cx - (w * s) / 2, cy - (h * s) / 2, w * s, h * s];
    } else box = [x * sx, y * sy, w * sx, h * sy];
    const [nx, ny, nw, nh] = box.map(Math.round) as [number, number, number, number];
    off.attrs = [attr(qname('', 'x', ''), String(nx)), attr(qname('', 'y', ''), String(ny))];
    ext.attrs = [attr(ATTR_CX, String(nw)), attr(ATTR_CY, String(nh))];
  }
};

/**
 * Handout Orientation / Notes Page Orientation. Both turn the one notes page
 * (`<p:notesSz>`) that notes pages and handouts share, and the reference desktop app moves
 * the notes master, the handout master and every notes page with it.
 */
export const setNotesPageOrientation = (
  pres: PresentationData,
  orientation: 'portrait' | 'landscape',
): void => {
  if (orientation !== 'portrait' && orientation !== 'landscape')
    throw new Error(`setNotesPageOrientation: unknown orientation ${orientation}`);
  const pkg = pres[INTERNAL_PACKAGE];
  const page = pageSize(pkg);
  if (page.w > page.h === (orientation === 'landscape')) return;
  const turned = { w: page.h, h: page.w };
  const sx = turned.w / page.w;
  const sy = turned.h / page.h;
  for (const part of pkg.parts) {
    const kind =
      part.contentType === NOTES_MASTER_CONTENT_TYPE ||
      part.contentType === HANDOUT_MASTER_CONTENT_TYPE ||
      part.contentType === NOTES_SLIDE_CONTENT_TYPE;
    if (!kind) continue;
    const target = readPartDoc(pkg, part.name);
    turnPage(target.doc.root, sx, sy);
    commitPartDoc(target);
  }
  const presentation = readPartDoc(pkg, PRES_PART_NAME);
  let notesSz = firstChildElement(presentation.doc.root, NAME_NOTES_SZ);
  if (notesSz === null) {
    notesSz = elem(NAME_NOTES_SZ);
    const sldSz = firstChildElement(presentation.doc.root, qname('p', 'sldSz', NS.pml));
    const at = sldSz
      ? presentation.doc.root.children.indexOf(sldSz) + 1
      : presentation.doc.root.children.length;
    presentation.doc.root.children.splice(at, 0, notesSz);
  }
  notesSz.attrs = [attr(ATTR_CX, String(turned.w)), attr(ATTR_CY, String(turned.h))];
  commitPartDoc(presentation);
};

// ---------------------------------------------------------------------------
// Slides per page.

/** The Handout Master tab's Slides Per Page choices (`'outline'` is Slide Outline). */
export type HandoutSlidesPerPage = 1 | 2 | 3 | 4 | 6 | 9 | 'outline';

const PER_PAGE_TOKENS: ReadonlyArray<readonly [HandoutSlidesPerPage, string]> = [
  [1, 'handouts1'],
  [2, 'handouts2'],
  [3, 'handouts3'],
  [4, 'handouts4'],
  [6, 'handouts6'],
  [9, 'handouts9'],
  ['outline', 'outline'],
];
// The handout master shows six slides a page until another layout is chosen.
const DEFAULT_SLIDES_PER_PAGE = 6;

/**
 * How many slides the handout master lays out per page. The handout master
 * part has nowhere to keep it; the schema's only home for a handout layout is
 * the print settings' `<p:prnPr prnWhat>` in the presentation properties, so
 * that is where it is read from (six when it names no handout layout).
 */
export const getHandoutSlidesPerPage = (pres: PresentationData): HandoutSlidesPerPage => {
  const root = readPresentationProperties(pres);
  const prnPr = root && firstChildElement(root, NAME_PRN_PR);
  const token = prnPr && getAttrValue(prnPr, ATTR_PRN_WHAT);
  return PER_PAGE_TOKENS.find(([, t]) => t === token)?.[0] ?? DEFAULT_SLIDES_PER_PAGE;
};

/** Sets the handout's slides per page; see `getHandoutSlidesPerPage`. */
export const setHandoutSlidesPerPage = (
  pres: PresentationData,
  slidesPerPage: HandoutSlidesPerPage,
): void => {
  const token = PER_PAGE_TOKENS.find(([value]) => value === slidesPerPage)?.[1];
  if (token === undefined)
    throw new Error(`setHandoutSlidesPerPage: unsupported value ${String(slidesPerPage)}`);
  writePresentationProperties(pres, (root) => {
    let prnPr = firstChildElement(root, NAME_PRN_PR);
    if (prnPr === null) {
      prnPr = elem(NAME_PRN_PR);
      // CT_PresentationProperties: htmlPubPr, webPr, prnPr, showPr, clrMru, extLst.
      const following = root.children.findIndex(
        (child) =>
          child.kind === 'element' &&
          child.name.namespaceURI === NS.pml &&
          ['showPr', 'clrMru', 'extLst'].includes(child.name.localName),
      );
      root.children.splice(following < 0 ? root.children.length : following, 0, prnPr);
    }
    prnPr.attrs = [
      ...prnPr.attrs.filter((a) => !(a.name.namespaceURI === '' && a.name.localName === 'prnWhat')),
      attr(ATTR_PRN_WHAT, token),
    ];
  });
};
