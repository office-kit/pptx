import { inkContentPart } from '../../internal/drawingml/ink-content.ts';
// Shape click action.
import { getCustomShows } from './custom-shows.ts';
import { getSlides } from './slide-query.ts';

import {
  basename,
  emptyRels,
  nextRelId,
  partName,
  resolveTarget,
} from '../../internal/opc/index.ts';
import type { OpcPackage } from '../../internal/parts/index.ts';
import { REL_TYPES } from '../../internal/presentationml/index.ts';
import {
  NS,
  type XmlElement,
  attr,
  elem,
  firstChildElement,
  getAttrValue,
  insertChildByRank,
  qname,
} from '../../internal/xml/index.ts';
import {
  INTERNAL_PACKAGE,
  type PresentationData,
  SHAPE_ELEMENT,
  SHAPE_SLIDE,
  SHAPE_SNAPSHOT,
  SLIDE_PART_NAME,
  type SlideData,
  type SlideShapeData,
} from '../_internal-symbols.ts';
import { commitAndRefresh } from './_helpers.ts';
// ---------------------------------------------------------------------------
// Shape click action — `<a:hlinkClick>` on the shape's cNvPr.
//
// Supports external URLs, internal slide destinations and preset navigation.

/** What clicking the shape should do. */
export type ShapeClickAction =
  | { readonly kind: 'url'; readonly url: string }
  | { readonly kind: 'slide'; readonly slide: SlideData }
  | { readonly kind: 'customShow'; readonly id: number; readonly showAndReturn: boolean }
  | { readonly kind: 'nextSlide' }
  | { readonly kind: 'prevSlide' }
  | { readonly kind: 'firstSlide' }
  | { readonly kind: 'lastSlide' }
  | { readonly kind: 'lastSlideViewed' }
  | { readonly kind: 'endShow' };

/** Parses the custom-show action defined by MS-OI29500 section 2.1.1395. */
export const parseCustomShowAction = (value: string | null): ShapeClickAction | null => {
  if (!value?.startsWith('ppaction://customshow?')) return null;
  const params = new URLSearchParams(value.slice('ppaction://customshow?'.length));
  const raw = params.get('id');
  if (raw === null || !/^\d+$/.test(raw)) return null;
  const id = Number(raw);
  if (!Number.isInteger(id) || id > 0xffffffff) return null;
  return { kind: 'customShow', id, showAndReturn: params.get('return') === 'true' };
};

export const NAME_HLINK_CLICK_FN = qname('a', 'hlinkClick', NS.dml);
const NAME_HLINK_HOVER = qname('a', 'hlinkHover', NS.dml);
type ActionTrigger = 'hlinkClick' | 'hlinkHover';

// cNvPr lives at different paths depending on shape kind. Returns null
// for kinds we don't know how to navigate yet.
export const findCNvPr = (shape: SlideShapeData): XmlElement | null => {
  const root = shape[SHAPE_ELEMENT];
  const kind = shape[SHAPE_SNAPSHOT].kind;
  if (kind === 'ink') {
    const content = inkContentPart(root);
    const nv = content && firstChildElement(content, qname('p14', 'nvContentPartPr', NS.p14));
    return nv && firstChildElement(nv, qname('p14', 'cNvPr', NS.p14));
  }
  const wrapperName =
    kind === 'shape'
      ? 'nvSpPr'
      : kind === 'picture'
        ? 'nvPicPr'
        : kind === 'connector'
          ? 'nvCxnSpPr'
          : kind === 'graphicFrame'
            ? 'nvGraphicFramePr'
            : kind === 'group'
              ? 'nvGrpSpPr'
              : null;
  if (wrapperName === null) return null;
  const wrapper = firstChildElement(root, qname('p', wrapperName, NS.pml));
  if (!wrapper) return null;
  return firstChildElement(wrapper, qname('p', 'cNvPr', NS.pml));
};

const removeExistingAction = (cNvPr: XmlElement, trigger: ActionTrigger): void => {
  cNvPr.children = cNvPr.children.filter(
    (c) =>
      !(c.kind === 'element' && c.name.namespaceURI === NS.dml && c.name.localName === trigger),
  );
};

const findExistingHyperlinkRel = (
  rels: ReturnType<OpcPackage['getRels']>,
  url: string,
): string | null => {
  if (rels === null) return null;
  const existing = rels.items.find(
    (rl) => rl.type === REL_TYPES.hyperlink && rl.target === url && rl.targetMode === 'External',
  );
  return existing?.id ?? null;
};

/**
 * Reads the click action attached to the shape's cNvPr, or `null` if
 * none. Mirrors `setShapeClickAction`:
 *
 *   - `{ kind: 'url', url }`     — `hyperlink` rel + targetMode=External
 *   - `{ kind: 'slide', slide }` — `slide` rel + `ppaction://hlinksldjump`
 *   - `{ kind: 'nextSlide' | 'prevSlide' | 'firstSlide' | 'lastSlide' }`
 *     — preset show-navigation `ppaction`.
 *
 * For `kind: 'slide'`, the matching slide is resolved by part name.
 * Returns `null` for unknown `ppaction` strings.
 */
export const getShapeClickAction = (shape: SlideShapeData): ShapeClickAction | null =>
  readShapeAction(shape, 'hlinkClick');

/** Reads the action triggered when the pointer enters the shape during a slideshow. */
export const getShapeHoverAction = (shape: SlideShapeData): ShapeClickAction | null =>
  readShapeAction(shape, 'hlinkHover');

const readShapeAction = (
  shape: SlideShapeData,
  trigger: ActionTrigger,
): ShapeClickAction | null => {
  const cNvPr = findCNvPr(shape);
  if (!cNvPr) return null;
  const hlink = firstChildElement(
    cNvPr,
    trigger === 'hlinkClick' ? NAME_HLINK_CLICK_FN : NAME_HLINK_HOVER,
  );
  if (!hlink) return null;
  const action = getAttrValue(hlink, qname('', 'action', ''));
  const rId = getAttrValue(hlink, qname('r', 'id', NS.officeDocRels));

  const customShow = parseCustomShowAction(action);
  if (customShow) return customShow;

  if (action === 'ppaction://hlinkshowjump?jump=nextslide') return { kind: 'nextSlide' };
  if (action === 'ppaction://hlinkshowjump?jump=previousslide') return { kind: 'prevSlide' };
  if (action === 'ppaction://hlinkshowjump?jump=firstslide') return { kind: 'firstSlide' };
  if (action === 'ppaction://hlinkshowjump?jump=lastslide') return { kind: 'lastSlide' };
  if (action === 'ppaction://hlinkshowjump?jump=lastslideviewed')
    return { kind: 'lastSlideViewed' };
  if (action === 'ppaction://hlinkshowjump?jump=endshow') return { kind: 'endShow' };

  if (rId !== null && rId !== '') {
    const slide = shape[SHAPE_SLIDE];
    const pkg = slide[INTERNAL_PACKAGE];
    const rels = pkg.getRels(slide[SLIDE_PART_NAME]);
    if (!rels) return null;
    const rel = rels.items.find((r) => r.id === rId);
    if (!rel) return null;

    if (action === 'ppaction://hlinksldjump' && rel.type === REL_TYPES.slide) {
      // Resolve to the SlideData of the target slide.
      const targetPartName = rel.target.startsWith('/')
        ? partName(rel.target)
        : resolveTarget(slide[SLIDE_PART_NAME], rel.target);
      const pres: PresentationData = { [INTERNAL_PACKAGE]: pkg, _slidesCache: null };
      for (const candidate of getSlides(pres)) {
        if (candidate[SLIDE_PART_NAME] === targetPartName) {
          return { kind: 'slide', slide: candidate };
        }
      }
      return null;
    }
    if (rel.type === REL_TYPES.hyperlink && rel.targetMode === 'External') {
      return { kind: 'url', url: rel.target };
    }
  }
  return null;
};

/** Reads only the object click ScreenTip, independently of text-run links. */
export const getShapeClickActionTooltip = (shape: SlideShapeData): string | null => {
  const metadata = findCNvPr(shape);
  const link = metadata && firstChildElement(metadata, NAME_HLINK_CLICK_FN);
  return link ? getAttrValue(link, qname('', 'tooltip', '')) : null;
};

/**
 * Sets (or clears) the click action on the shape. Side effects:
 *
 *   - For `kind: 'url'`, a `hyperlink` rel is added (or reused) on the
 *     slide's rels with `targetMode="External"`. `<a:hlinkClick r:id=…/>`
 *     points at it.
 *   - For `kind: 'slide'`, a `slide` rel is added pointing at the
 *     target slide's part. The `<a:hlinkClick>` carries
 *     `action="ppaction://hlinksldjump"`.
 *   - For the preset navigations (`nextSlide`, `prevSlide`, ...), no rel
 *     is allocated; just the `action` attribute carries the preset.
 *   - `null` removes any existing `<a:hlinkClick>`.
 *
 * Supports shapes, pictures, connectors, graphic frames and groups.
 * The optional tooltip is the ScreenTip stored on the object link.
 */
export const setShapeClickAction = (
  shape: SlideShapeData,
  action: ShapeClickAction | null,
  tooltip?: string,
): void => {
  setShapeAction(shape, action, tooltip, 'hlinkClick');
};

/** Sets or clears the mouse-over action without changing the click action. */
export const setShapeHoverAction = (
  shape: SlideShapeData,
  action: ShapeClickAction | null,
  tooltip?: string,
): void => setShapeAction(shape, action, tooltip, 'hlinkHover');

/** Reads the ScreenTip attached to the mouse-over action. */
export const getShapeHoverActionTooltip = (shape: SlideShapeData): string | null => {
  const metadata = findCNvPr(shape);
  const link = metadata && firstChildElement(metadata, NAME_HLINK_HOVER);
  return link ? getAttrValue(link, qname('', 'tooltip', '')) : null;
};

const setShapeAction = (
  shape: SlideShapeData,
  action: ShapeClickAction | null,
  tooltip: string | undefined,
  trigger: ActionTrigger,
): void => {
  const cNvPr = findCNvPr(shape);
  if (!cNvPr) {
    throw new Error(
      `setShapeClickAction: ${shape[SHAPE_SNAPSHOT].kind} shape has no cNvPr to attach to`,
    );
  }

  const link = createShapeClickLink(shape, action, tooltip);
  const existing = firstChildElement(
    cNvPr,
    trigger === 'hlinkClick' ? NAME_HLINK_CLICK_FN : NAME_HLINK_HOVER,
  );
  // Updating the destination must retain independent settings such as sound,
  // endSnd and extension data. A null action explicitly removes the whole link.
  if (link && existing) {
    link.attrs.push(
      ...existing.attrs.filter(
        ({ name }) =>
          !(name.namespaceURI === NS.officeDocRels && name.localName === 'id') &&
          !(
            name.namespaceURI === '' && ['action', 'tooltip', 'invalidUrl'].includes(name.localName)
          ),
      ),
    );
    link.children = existing.children;
    link.prefixDecls = new Map(existing.prefixDecls);
  }
  removeExistingAction(cNvPr, trigger);
  if (link) {
    link.name = trigger === 'hlinkClick' ? NAME_HLINK_CLICK_FN : NAME_HLINK_HOVER;
    insertChildByRank(cNvPr, link, (child) => {
      if (child.name.namespaceURI !== NS.dml) return 2;
      if (child.name.localName === 'hlinkClick') return 0;
      if (child.name.localName === 'hlinkHover') return 1;
      return 2;
    });
  }
  commitAndRefresh(shape);
};

/** Builds a click link and allocates its slide relationship without changing object metadata. */
export const createShapeClickLink = (
  shape: SlideShapeData,
  action: ShapeClickAction | null,
  tooltip?: string,
): XmlElement | null => {
  // Validate a destination before touching an existing link or its relationships.
  const slide = shape[SHAPE_SLIDE];
  const pkg = slide[INTERNAL_PACKAGE];
  if (
    action?.kind === 'slide' &&
    (action.slide[INTERNAL_PACKAGE] !== pkg ||
      !getSlides({ [INTERNAL_PACKAGE]: pkg, _slidesCache: null }).some(
        (candidate) => candidate[SLIDE_PART_NAME] === action.slide[SLIDE_PART_NAME],
      ))
  )
    throw new Error('setShapeClickAction: target slide must belong to this presentation.');

  if (
    action?.kind === 'customShow' &&
    (!Number.isInteger(action.id) ||
      action.id < 0 ||
      action.id > 0xffffffff ||
      typeof action.showAndReturn !== 'boolean' ||
      !getCustomShows({ [INTERNAL_PACKAGE]: pkg, _slidesCache: null }).some(
        (show) => show.id === action.id,
      ))
  )
    throw new Error('setShapeClickAction: custom show must belong to this presentation.');
  if (action === null) return null;
  let rId: string | null = null;
  let actionAttr: string | null = null;

  switch (action.kind) {
    case 'customShow':
      actionAttr = `ppaction://customshow?id=${action.id}${action.showAndReturn ? '&return=true' : ''}`;
      break;
    case 'url': {
      const rels = pkg.getRels(slide[SLIDE_PART_NAME]) ?? emptyRels();
      const reused = findExistingHyperlinkRel(rels, action.url);
      if (reused !== null) {
        rId = reused;
      } else {
        const newId = nextRelId(rels.items.map((r) => r.id));
        rels.items.push({
          id: newId,
          type: REL_TYPES.hyperlink,
          target: action.url,
          targetMode: 'External',
        });
        pkg.setRels(slide[SLIDE_PART_NAME], rels);
        rId = newId;
      }
      break;
    }
    case 'slide': {
      const target = action.slide[SLIDE_PART_NAME];
      const targetBase = basename(target);
      const rels = pkg.getRels(slide[SLIDE_PART_NAME]) ?? emptyRels();
      const existing = rels.items.find(
        (rl) =>
          rl.type === REL_TYPES.slide &&
          rl.target === `../slides/${targetBase}` &&
          rl.targetMode === 'Internal',
      );
      if (existing) {
        rId = existing.id;
      } else {
        const newId = nextRelId(rels.items.map((r) => r.id));
        rels.items.push({
          id: newId,
          type: REL_TYPES.slide,
          target: `../slides/${targetBase}`,
          targetMode: 'Internal',
        });
        pkg.setRels(slide[SLIDE_PART_NAME], rels);
        rId = newId;
      }
      actionAttr = 'ppaction://hlinksldjump';
      break;
    }
    case 'nextSlide':
      actionAttr = 'ppaction://hlinkshowjump?jump=nextslide';
      break;
    case 'prevSlide':
      actionAttr = 'ppaction://hlinkshowjump?jump=previousslide';
      break;
    case 'firstSlide':
      actionAttr = 'ppaction://hlinkshowjump?jump=firstslide';
      break;
    case 'lastSlideViewed':
      actionAttr = 'ppaction://hlinkshowjump?jump=lastslideviewed';
      break;
    case 'endShow':
      actionAttr = 'ppaction://hlinkshowjump?jump=endshow';
      break;
    case 'lastSlide':
      actionAttr = 'ppaction://hlinkshowjump?jump=lastslide';
      break;
  }

  const attrs = [] as Array<ReturnType<typeof attr>>;
  if (rId !== null) attrs.push(attr(qname('r', 'id', NS.officeDocRels), rId));
  else attrs.push(attr(qname('r', 'id', NS.officeDocRels), ''));
  if (actionAttr !== null) attrs.push(attr(qname('', 'action', ''), actionAttr));
  if (tooltip !== undefined) attrs.push(attr(qname('', 'tooltip', ''), tooltip));

  return elem(NAME_HLINK_CLICK_FN, { attrs });
};
