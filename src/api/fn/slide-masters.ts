// Slide masters and the structure of their layouts: what the reference desktop app's Slide
// Master tab edits (Insert Slide Master, Insert Layout, Delete, Rename,
// Preserve, Master Layout, Insert Placeholder, the Title and Footers
// checkboxes and Hide Background Graphics).
//
// A master is named by its package part name (`getSlideMasterPartNames`), the
// same handle `getSlideMasterPartName(layout)` returns. Layouts are
// `SlideLayoutData` handles, as everywhere else.
//
// ECMA-376 Part 1 §19.3.1.42 `sldMaster` lists its layouts in
// `<p:sldLayoutIdLst>`; §19.2.1.34 / §19.3.1.40 give masters and layouts ids
// from one space of at least 2³¹, which the reference desktop app requires to be unique across
// every master and layout of the presentation.

import { boundedInt } from '../../internal/bounds.ts';
import {
  type PartName,
  emptyRels,
  nextRelId,
  partName,
  partNamesEqual,
  relsPartNameFor,
  resolveTarget,
} from '../../internal/opc/index.ts';
import { DEFAULT_THEME_XML, type OpcPackage } from '../../internal/parts/index.ts';
import {
  type LayoutPlaceholderKind,
  MASTER_PLACEHOLDER_TYPES,
  type MasterPlaceholderType,
  REL_TYPES,
  customSlideLayoutXml,
  defaultSlideLayoutsXml,
  defaultSlideMasterXml,
  insertedPlaceholderXml,
  layoutFooterPlaceholdersXml,
  layoutTitlePlaceholderXml,
  masterPlaceholderXml,
  readShapeTreeFromCsldRoot,
  readSlideLayoutPart,
} from '../../internal/presentationml/index.ts';
import {
  NS,
  type XmlElement,
  allChildElements,
  attr,
  elem,
  firstChildElement,
  getAttrValue,
  parseXml,
  qname,
  serializeXml,
} from '../../internal/xml/index.ts';
import {
  INTERNAL_PACKAGE,
  LAYOUT_DOCUMENT,
  LAYOUT_PART,
  LAYOUT_PART_NAME,
  type PresentationData,
  type SlideLayoutData,
} from '../_internal-symbols.ts';
import {
  ATTR_ID,
  ATTR_R_ID,
  NAME_SLD_MASTER_ID_LST,
  PRES_PART_NAME,
  SLIDE_CONTENT_TYPE,
  SLIDE_LAYOUT_CONTENT_TYPE,
  allocatePartNames,
  commitLayoutData,
  decode,
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
  placeholderIndices,
  placeholderShapes,
  placeholderViews,
  readPartDoc,
  removePlaceholdersOfType,
  requirePartSpTree,
} from './_part-placeholders.ts';
import type { ShapeBounds } from './shape-read-base.ts';
import { getSlideMasterPartNames } from './shape-read-base.ts';
import { getSlideSize } from './slide-size.ts';

export type { LayoutPlaceholderKind, MasterPlaceholderType };

const SLIDE_MASTER_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml';
const THEME_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.theme+xml';

const NAME_SLD_MASTER_ID = qname('p', 'sldMasterId', NS.pml);
const NAME_SLD_LAYOUT_ID_LST = qname('p', 'sldLayoutIdLst', NS.pml);
const NAME_SLD_LAYOUT_ID = qname('p', 'sldLayoutId', NS.pml);
const ATTR_NAME = qname('', 'name', '');
const ATTR_PRESERVE = qname('', 'preserve', '');
const ATTR_SHOW_MASTER_SP = qname('', 'showMasterSp', '');

// ST_SlideMasterId / ST_SlideLayoutId: xsd:unsignedInt, at least 2³¹.
const MASTER_LAYOUT_ID_MIN = 2147483648;
const MASTER_LAYOUT_ID_MAX = 4294967295;

// What the reference desktop app names a master it inserts and a layout it inserts. Repeats
// are numbered in front: "1_Custom Design", "2_Custom Design", …
const INSERTED_MASTER_NAME = 'Custom Design';
const INSERTED_LAYOUT_NAME = 'Custom Layout';

const DEFAULT_SLIDE = { w: 12192000, h: 6858000 };

const relTarget = (from: PartName, target: string): PartName =>
  target.startsWith('/') ? partName(target) : resolveTarget(from, target);

const slideCanvas = (pres: PresentationData): { w: number; h: number } => {
  const size = getSlideSize(pres);
  return size ? { w: size.width, h: size.height } : DEFAULT_SLIDE;
};

const requireMaster = (pres: PresentationData, master: string): PartName => {
  const name = getSlideMasterPartNames(pres).find((n) =>
    partNamesEqual(partName(n), partName(master)),
  );
  if (name === undefined) throw new Error(`${master} is not a slide master of this presentation`);
  return partName(name);
};

/** Layout part names of a master, in its `<p:sldLayoutIdLst>` order. */
const masterLayoutNames = (pkg: OpcPackage, master: PartName, root: XmlElement): PartName[] => {
  const rels = pkg.getRels(master)?.items ?? [];
  const list = firstChildElement(root, NAME_SLD_LAYOUT_ID_LST);
  if (list === null) return [];
  const out: PartName[] = [];
  for (const entry of allChildElements(list, NAME_SLD_LAYOUT_ID)) {
    const rel = rels.find((r) => r.id === getAttrValue(entry, ATTR_R_ID));
    if (rel) out.push(relTarget(master, rel.target));
  }
  return out;
};

const layoutData = (pkg: OpcPackage, name: PartName): SlideLayoutData => {
  const part = pkg.getPart(name);
  if (part === null) throw new Error(`slide layout part missing: ${name}`);
  const doc = parseXml(decode(part.data));
  return {
    [INTERNAL_PACKAGE]: pkg,
    [LAYOUT_PART_NAME]: name,
    [LAYOUT_DOCUMENT]: doc,
    [LAYOUT_PART]: readSlideLayoutPart(doc.root),
  };
};

/** Slide-layout part name → number of slides on it, for every slide in the package. */
const layoutUse = (pkg: OpcPackage): Map<string, number> => {
  const use = new Map<string, number>();
  for (const part of pkg.parts) {
    if (part.contentType !== SLIDE_CONTENT_TYPE) continue;
    const rel = pkg.getRels(part.name)?.items.find((r) => r.type === REL_TYPES.slideLayout);
    if (!rel) continue;
    const key = relTarget(part.name, rel.target).toLowerCase();
    use.set(key, (use.get(key) ?? 0) + 1);
  }
  return use;
};

const themePartOf = (pkg: OpcPackage, from: PartName): PartName | null => {
  const rel = pkg.getRels(from)?.items.find((r) => r.type === REL_TYPES.theme);
  return rel ? relTarget(from, rel.target) : null;
};

/**
 * `count` fresh ids for `<p:sldMasterId>` / `<p:sldLayoutId>`, above every id
 * the presentation and its masters already use.
 */
const allocateMasterLayoutIds = (
  pkg: OpcPackage,
  presentation: XmlElement,
  count: number,
): number[] => {
  let max = MASTER_LAYOUT_ID_MIN - 1;
  const consider = (entries: ReadonlyArray<XmlElement>) => {
    for (const entry of entries) {
      const id = Number(getAttrValue(entry, ATTR_ID));
      if (Number.isFinite(id) && id > max) max = id;
    }
  };
  const masterList = firstChildElement(presentation, NAME_SLD_MASTER_ID_LST);
  if (masterList) consider(allChildElements(masterList, NAME_SLD_MASTER_ID));
  for (const part of pkg.parts) {
    if (part.contentType !== SLIDE_MASTER_CONTENT_TYPE) continue;
    const list = firstChildElement(parseXml(decode(part.data)).root, NAME_SLD_LAYOUT_ID_LST);
    if (list) consider(allChildElements(list, NAME_SLD_LAYOUT_ID));
  }
  if (max + count > MASTER_LAYOUT_ID_MAX)
    throw new Error('the presentation has run out of slide master and layout ids');
  return Array.from({ length: count }, (_, i) => max + 1 + i);
};

/** `base`, or the reference desktop app's `1_base`, `2_base`, … when `base` is taken. */
const uniqueName = (base: string, taken: ReadonlySet<string>): string => {
  if (!taken.has(base)) return base;
  for (let n = 1; ; n++) if (!taken.has(`${n}_${base}`)) return `${n}_${base}`;
};

// ---------------------------------------------------------------------------
// Reading.

/** The master's layouts, in the order the reference desktop app's Slide Master pane lists them. */
export const getSlideMasterLayouts = (
  pres: PresentationData,
  master: string,
): ReadonlyArray<SlideLayoutData> => {
  const pkg = pres[INTERNAL_PACKAGE];
  const name = requireMaster(pres, master);
  return masterLayoutNames(pkg, name, readPartDoc(pkg, name).doc.root).map((layout) =>
    layoutData(pkg, layout),
  );
};

/**
 * The master's own placeholders (title, body, date, footer, slide number) and
 * where it puts them. A layout placeholder without a transform of its own sits
 * where the master's placeholder of the same type does.
 */
export const getSlideMasterPlaceholders = (
  pres: PresentationData,
  master: string,
): ReadonlyArray<SlideLayoutPlaceholder> => {
  const pkg = pres[INTERNAL_PACKAGE];
  const { doc } = readPartDoc(pkg, requireMaster(pres, master));
  return placeholderViews(readShapeTreeFromCsldRoot(doc.root, 'sldMaster').shapes);
};

/**
 * The master's name, as the reference desktop app shows it in Slide Master view ("<theme
 * name> Slide Master"). The reference desktop app keeps it as the name of the master's theme
 * (`<a:theme name>`), not in the master part, so renaming a master renames its
 * theme.
 */
export const getSlideMasterName = (pres: PresentationData, master: string): string => {
  const pkg = pres[INTERNAL_PACKAGE];
  const theme = themePartOf(pkg, requireMaster(pres, master));
  if (theme === null) return '';
  return getAttrValue(readPartDoc(pkg, theme).doc.root, ATTR_NAME) ?? '';
};

/**
 * Whether the master is preserved (`<p:sldMaster preserve="1">`): the reference desktop app
 * deletes a master that is not preserved once no slide uses it.
 */
export const isSlideMasterPreserved = (pres: PresentationData, master: string): boolean => {
  const { doc } = readPartDoc(pres[INTERNAL_PACKAGE], requireMaster(pres, master));
  const value = getAttrValue(doc.root, ATTR_PRESERVE);
  return value === '1' || value === 'true';
};

// ---------------------------------------------------------------------------
// Masters.

/** Renames the master (its theme; see `getSlideMasterName`). */
export const setSlideMasterName = (pres: PresentationData, master: string, name: string): void => {
  const pkg = pres[INTERNAL_PACKAGE];
  const themeName = themePartOf(pkg, requireMaster(pres, master));
  if (themeName === null) throw new Error(`setSlideMasterName: ${master} has no theme`);
  const theme = readPartDoc(pkg, themeName);
  setAttribute(theme.doc.root, ATTR_NAME, name);
  commitPartDoc(theme);
};

/** Sets or clears Preserve Master. */
export const setSlideMasterPreserved = (
  pres: PresentationData,
  master: string,
  preserved: boolean,
): void => {
  const target = readPartDoc(pres[INTERNAL_PACKAGE], requireMaster(pres, master));
  setAttribute(target.doc.root, ATTR_PRESERVE, preserved ? '1' : null);
  commitPartDoc(target);
};

const setAttribute = (
  element: XmlElement,
  name: ReturnType<typeof qname>,
  value: string | null,
): void => {
  element.attrs = element.attrs.filter(
    (a) => !(a.name.namespaceURI === name.namespaceURI && a.name.localName === name.localName),
  );
  if (value !== null) element.attrs.push(attr(name, value));
};

/**
 * Insert Slide Master: adds the reference desktop app's default master, with its own
 * copy of the default theme named "Custom Design" and the eleven default
 * layouts, after the existing masters. The master is preserved, as the reference desktop app
 * marks a master it inserts. Returns the new master's part name.
 */
export const addSlideMaster = (pres: PresentationData): string => {
  const pkg = pres[INTERNAL_PACKAGE];
  const presentation = readPartDoc(pkg, PRES_PART_NAME);
  const size = slideCanvas(pres);
  const layouts = defaultSlideLayoutsXml(size);
  const [masterName] = allocatePartNames(pkg, '/ppt/slideMasters/slideMaster', 1) as [PartName];
  const layoutNames = allocatePartNames(pkg, '/ppt/slideLayouts/slideLayout', layouts.length);
  const [themeName] = allocatePartNames(pkg, '/ppt/theme/theme', 1) as [PartName];
  const [masterId, ...layoutIds] = allocateMasterLayoutIds(
    pkg,
    presentation.doc.root,
    1 + layouts.length,
  ) as [number, ...number[]];

  const masterNames = new Set(
    getSlideMasterPartNames(pres).map((name) => getSlideMasterName(pres, name)),
  );
  const theme = parseXml(DEFAULT_THEME_XML);
  setAttribute(theme.root, ATTR_NAME, uniqueName(INSERTED_MASTER_NAME, masterNames));
  pkg.addPart(themeName, THEME_CONTENT_TYPE, encode(serializeXml(theme)));

  const masterRels = emptyRels();
  const entries = layoutNames.map((layoutName, i) => {
    const rId = `rId${i + 1}`;
    masterRels.items.push({
      id: rId,
      type: REL_TYPES.slideLayout,
      target: relativeTarget(masterName, layoutName),
      targetMode: 'Internal',
    });
    pkg.addPart(layoutName, SLIDE_LAYOUT_CONTENT_TYPE, encode(layouts[i]!));
    pkg.setRels(layoutName, layoutBackRels(layoutName, masterName));
    return { id: layoutIds[i]!, rId };
  });
  masterRels.items.push({
    id: `rId${layoutNames.length + 1}`,
    type: REL_TYPES.theme,
    target: relativeTarget(masterName, themeName),
    targetMode: 'Internal',
  });
  pkg.addPart(masterName, SLIDE_MASTER_CONTENT_TYPE, encode(defaultSlideMasterXml(size, entries)));
  pkg.setRels(masterName, masterRels);

  const presRels = pkg.getRels(PRES_PART_NAME) ?? emptyRels();
  const rId = nextRelId(presRels.items.map((r) => r.id));
  presRels.items.push({
    id: rId,
    type: REL_TYPES.slideMaster,
    target: relativeTarget(PRES_PART_NAME, masterName),
    targetMode: 'Internal',
  });
  pkg.setRels(PRES_PART_NAME, presRels);
  let list = firstChildElement(presentation.doc.root, NAME_SLD_MASTER_ID_LST);
  if (list === null) {
    list = elem(NAME_SLD_MASTER_ID_LST);
    presentation.doc.root.children.unshift(list);
  }
  list.children.push(
    elem(NAME_SLD_MASTER_ID, { attrs: [attr(ATTR_ID, String(masterId)), attr(ATTR_R_ID, rId)] }),
  );
  commitPartDoc(presentation);
  return masterName;
};

const layoutBackRels = (layout: PartName, master: PartName) => ({
  items: [
    {
      id: 'rId1',
      type: REL_TYPES.slideMaster,
      target: relativeTarget(layout, master),
      targetMode: 'Internal' as const,
    },
  ],
});

/** Every part some relationship in the package points at. */
const referencedParts = (pkg: OpcPackage, except: ReadonlySet<string>): Set<string> => {
  const out = new Set<string>();
  for (const part of pkg.parts) {
    if (!part.name.endsWith('.rels')) continue;
    const owner = ownerOfRels(part.name);
    if (owner === null || except.has(owner.toLowerCase())) continue;
    for (const rel of pkg.getRels(owner)?.items ?? []) {
      if (rel.targetMode === 'External') continue;
      out.add(relTarget(owner, rel.target).toLowerCase());
    }
  }
  return out;
};

// `/a/b/_rels/c.xml.rels` belongs to `/a/b/c.xml`; the package's `/_rels/.rels` to no part.
const ownerOfRels = (rels: PartName): PartName | null => {
  const match = /^(.*)\/_rels\/([^/]+)\.rels$/.exec(rels);
  return match?.[2] ? partName(`${match[1]}/${match[2]}`) : null;
};

const removePartWithRels = (pkg: OpcPackage, name: PartName): void => {
  pkg.removePart(name);
  pkg.removePart(relsPartNameFor(name));
};

/**
 * Deletes a master with its layouts and, when nothing else uses it, its theme.
 * Like the reference desktop app's Delete, this refuses (throws) while any slide uses one of
 * the master's layouts, and for the presentation's only master.
 */
export const removeSlideMaster = (pres: PresentationData, master: string): void => {
  const pkg = pres[INTERNAL_PACKAGE];
  const name = requireMaster(pres, master);
  if (getSlideMasterPartNames(pres).length < 2)
    throw new Error('removeSlideMaster: a presentation needs at least one slide master');
  const target = readPartDoc(pkg, name);
  const layouts = masterLayoutNames(pkg, name, target.doc.root);
  const use = layoutUse(pkg);
  if (layouts.some((layout) => (use.get(layout.toLowerCase()) ?? 0) > 0))
    throw new Error('removeSlideMaster: slides still use this master');

  const theme = themePartOf(pkg, name);
  const removed = new Set([name, ...layouts].map((n) => n.toLowerCase()));
  for (const layout of layouts) removePartWithRels(pkg, layout);
  removePartWithRels(pkg, name);
  if (theme !== null && !referencedParts(pkg, removed).has(theme.toLowerCase()))
    removePartWithRels(pkg, theme);

  const presRels = pkg.getRels(PRES_PART_NAME) ?? emptyRels();
  const rel = presRels.items.find(
    (r) =>
      r.type === REL_TYPES.slideMaster && partNamesEqual(relTarget(PRES_PART_NAME, r.target), name),
  );
  pkg.setRels(PRES_PART_NAME, { items: presRels.items.filter((r) => r !== rel) });
  const presentation = readPartDoc(pkg, PRES_PART_NAME);
  const list = firstChildElement(presentation.doc.root, NAME_SLD_MASTER_ID_LST);
  if (list && rel)
    list.children = list.children.filter(
      (child) => child.kind !== 'element' || getAttrValue(child, ATTR_R_ID) !== rel.id,
    );
  commitPartDoc(presentation);
};

// ---------------------------------------------------------------------------
// Layouts.

/**
 * Insert Layout: adds the reference desktop app's "Custom Layout" — a title and the date,
 * footer and slide-number placeholders — to `master`, at `index` in its layout
 * list (the end by default). Repeats are named "1_Custom Layout", and so on.
 */
export const addSlideLayout = (
  pres: PresentationData,
  master: string,
  options: { index?: number } = {},
): SlideLayoutData => {
  const pkg = pres[INTERNAL_PACKAGE];
  const masterName = requireMaster(pres, master);
  const target = readPartDoc(pkg, masterName);
  const existing = masterLayoutNames(pkg, masterName, target.doc.root);
  const names = new Set(existing.map((layout) => layoutData(pkg, layout)[LAYOUT_PART].name));
  const presentation = readPartDoc(pkg, PRES_PART_NAME);
  const [id] = allocateMasterLayoutIds(pkg, presentation.doc.root, 1) as [number];
  const [name] = allocatePartNames(pkg, '/ppt/slideLayouts/slideLayout', 1) as [PartName];

  pkg.addPart(
    name,
    SLIDE_LAYOUT_CONTENT_TYPE,
    encode(customSlideLayoutXml(uniqueName(INSERTED_LAYOUT_NAME, names), slideCanvas(pres))),
  );
  pkg.setRels(name, layoutBackRels(name, masterName));
  const masterRels = pkg.getRels(masterName) ?? emptyRels();
  const rId = nextRelId(masterRels.items.map((r) => r.id));
  masterRels.items.push({
    id: rId,
    type: REL_TYPES.slideLayout,
    target: relativeTarget(masterName, name),
    targetMode: 'Internal',
  });
  pkg.setRels(masterName, masterRels);

  let list = firstChildElement(target.doc.root, NAME_SLD_LAYOUT_ID_LST);
  if (list === null) {
    list = elem(NAME_SLD_LAYOUT_ID_LST);
    // CT_SlideMaster: cSld, clrMap, sldLayoutIdLst, …
    const clrMap = firstChildElement(target.doc.root, qname('p', 'clrMap', NS.pml));
    target.doc.root.children.splice(target.doc.root.children.indexOf(clrMap!) + 1, 0, list);
  }
  const entries = allChildElements(list, NAME_SLD_LAYOUT_ID);
  const at = Math.max(0, Math.min(Math.trunc(options.index ?? entries.length), entries.length));
  const entry = elem(NAME_SLD_LAYOUT_ID, {
    attrs: [attr(ATTR_ID, String(id)), attr(ATTR_R_ID, rId)],
  });
  const before = entries[at];
  list.children.splice(before ? list.children.indexOf(before) : list.children.length, 0, entry);
  commitPartDoc(target);
  return layoutData(pkg, name);
};

const masterOfLayout = (layout: SlideLayoutData): PartName => {
  const name = layout[LAYOUT_PART_NAME];
  const rel = layout[INTERNAL_PACKAGE]
    .getRels(name)
    ?.items.find((r) => r.type === REL_TYPES.slideMaster);
  if (!rel) throw new Error(`${name} has no slide master`);
  return relTarget(name, rel.target);
};

/**
 * Deletes a layout. Like the reference desktop app's Delete, this refuses (throws) while any
 * slide uses the layout, and for a master's only layout.
 */
export const removeSlideLayout = (layout: SlideLayoutData): void => {
  const pkg = layout[INTERNAL_PACKAGE];
  const name = layout[LAYOUT_PART_NAME];
  if ((layoutUse(pkg).get(name.toLowerCase()) ?? 0) > 0)
    throw new Error('removeSlideLayout: slides still use this layout');
  const masterName = masterOfLayout(layout);
  const master = readPartDoc(pkg, masterName);
  if (masterLayoutNames(pkg, masterName, master.doc.root).length < 2)
    throw new Error('removeSlideLayout: a slide master needs at least one layout');
  const masterRels = pkg.getRels(masterName) ?? emptyRels();
  const rel = masterRels.items.find(
    (r) =>
      r.type === REL_TYPES.slideLayout && partNamesEqual(relTarget(masterName, r.target), name),
  );
  pkg.setRels(masterName, { items: masterRels.items.filter((r) => r !== rel) });
  const list = firstChildElement(master.doc.root, NAME_SLD_LAYOUT_ID_LST);
  if (list && rel)
    list.children = list.children.filter(
      (child) => child.kind !== 'element' || getAttrValue(child, ATTR_R_ID) !== rel.id,
    );
  commitPartDoc(master);
  removePartWithRels(pkg, name);
};

// ---------------------------------------------------------------------------
// Placeholders.

// Which `<p:ph type>` tokens each Master Layout checkbox stands for.
const MASTER_TYPE_TOKENS: Record<MasterPlaceholderType, ReadonlySet<string>> = {
  title: new Set(['title', 'ctrTitle']),
  body: new Set(['body', 'obj']),
  dt: new Set(['dt']),
  ftr: new Set(['ftr']),
  sldNum: new Set(['sldNum']),
};

/**
 * Master Layout: removes one of the master's five placeholders, or puts it
 * back where the reference desktop app's default master has it.
 */
export const setSlideMasterPlaceholderIncluded = (
  pres: PresentationData,
  master: string,
  type: MasterPlaceholderType,
  included: boolean,
): void => {
  const tokens = MASTER_TYPE_TOKENS[type];
  if (tokens === undefined)
    throw new Error(`setSlideMasterPlaceholderIncluded: unknown type ${type}`);
  const target = readPartDoc(pres[INTERNAL_PACKAGE], requireMaster(pres, master));
  const spTree = requirePartSpTree(target.doc.root);
  if (hasPlaceholderOfType(spTree, tokens) === included) return;
  if (included) {
    const element = placeholderElement(
      masterPlaceholderXml(type, nextShapeIdInTree(spTree), slideCanvas(pres)),
    );
    insertPlaceholderInOrder(spTree, element, type, MASTER_PLACEHOLDER_TYPES);
  } else {
    removePlaceholdersOfType(spTree, tokens);
  }
  commitPartDoc(target);
};

const TITLE_TOKENS: ReadonlySet<string> = new Set(['title', 'ctrTitle']);
const FOOTER_TOKENS = ['dt', 'ftr', 'sldNum'] as const;
// The reference desktop app's layouts give their footers these indices; a layout that already
// uses one for something else gets the next free index instead.
const FOOTER_IDX = [10, 11, 12] as const;

const layoutSpTree = (layout: SlideLayoutData): XmlElement =>
  requirePartSpTree(layout[LAYOUT_DOCUMENT].root);

/** The layout's Title checkbox: removes its title, or adds one that follows the master's. */
export const setSlideLayoutTitleIncluded = (layout: SlideLayoutData, included: boolean): void => {
  const spTree = layoutSpTree(layout);
  if (hasPlaceholderOfType(spTree, TITLE_TOKENS) === included) return;
  if (included) {
    const element = placeholderElement(layoutTitlePlaceholderXml(nextShapeIdInTree(spTree)));
    // The title comes first, straight after the group properties.
    const first = spTree.children.findIndex(
      (child) =>
        child.kind === 'element' &&
        child.name.localName !== 'nvGrpSpPr' &&
        child.name.localName !== 'grpSpPr',
    );
    spTree.children.splice(first < 0 ? spTree.children.length : first, 0, element);
  } else {
    removePlaceholdersOfType(spTree, TITLE_TOKENS);
  }
  commitLayoutData(layout);
};

/**
 * The layout's Footers checkbox: removes its date, footer and slide-number
 * placeholders, or adds whichever are missing (they follow the master's).
 */
export const setSlideLayoutFootersIncluded = (layout: SlideLayoutData, included: boolean): void => {
  const spTree = layoutSpTree(layout);
  const tokens = new Set<string>(FOOTER_TOKENS);
  if (!included) {
    if (!hasPlaceholderOfType(spTree, tokens)) return;
    removePlaceholdersOfType(spTree, tokens);
    commitLayoutData(layout);
    return;
  }
  const missing = FOOTER_TOKENS.filter((token) => !hasPlaceholderOfType(spTree, new Set([token])));
  if (missing.length === 0) return;
  const used = new Set(placeholderIndices(spTree));
  let spare = Math.max(12, ...used);
  const idx = FOOTER_IDX.map((preferred) => (used.has(preferred) ? ++spare : preferred)) as [
    number,
    number,
    number,
  ];
  const xml = layoutFooterPlaceholdersXml(nextShapeIdInTree(spTree), idx);
  FOOTER_TOKENS.forEach((token, i) => {
    if (missing.includes(token)) spTree.children.push(placeholderElement(xml[i]!));
  });
  commitLayoutData(layout);
};

/**
 * Hide Background Graphics on a layout (`showMasterSp="0"`): slides on it stop
 * showing the master's decorative shapes; the layout's own shapes stay.
 */
export const setSlideLayoutBackgroundGraphicsHidden = (
  layout: SlideLayoutData,
  hidden: boolean,
): void => {
  setAttribute(layout[LAYOUT_DOCUMENT].root, ATTR_SHOW_MASTER_SP, hidden ? '0' : null);
  commitLayoutData(layout);
};

/**
 * Insert Placeholder: draws a placeholder of `kind` on the layout at `bounds`.
 * Slides added on the layout get the placeholder; content and text
 * placeholders carry the master's five prompt levels, the others their
 * kind's prompt ("Picture", "Chart", …). Returns the placeholder's index in
 * `getSlideLayoutPlaceholders`.
 */
export const addSlideLayoutPlaceholder = (
  layout: SlideLayoutData,
  kind: LayoutPlaceholderKind,
  bounds: ShapeBounds,
): number => {
  const box = {
    x: boundedInt(bounds.x, 'coordinate', 'addSlideLayoutPlaceholder: x'),
    y: boundedInt(bounds.y, 'coordinate', 'addSlideLayoutPlaceholder: y'),
    w: boundedInt(bounds.w, 'positiveCoordinate', 'addSlideLayoutPlaceholder: w'),
    h: boundedInt(bounds.h, 'positiveCoordinate', 'addSlideLayoutPlaceholder: h'),
  };
  const spTree = layoutSpTree(layout);
  // Custom layout placeholders start after the footers' 10–12, as the reference desktop app's do.
  const idx = Math.max(12, ...placeholderIndices(spTree)) + 1;
  const element = placeholderElement(
    insertedPlaceholderXml(kind, nextShapeIdInTree(spTree), idx, box),
  );
  spTree.children.push(element);
  commitLayoutData(layout);
  return placeholderShapes(layout[LAYOUT_PART].shapes).length - 1;
};
