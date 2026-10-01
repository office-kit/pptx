import { nextRelId, partName, resolveTarget } from '../../internal/opc/index.ts';
import { REL_TYPES } from '../../internal/presentationml/relationship-types.ts';
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
  SLIDE_PART_NAME,
  type SlideData,
  type PresentationData,
} from '../_internal-symbols.ts';
import { PRES_PART_NAME, decode, encode } from './_helpers.ts';
import { getSlides } from './slide-query.ts';

/** A presentation-level drawing guide. Position is in EMU from the left/top edge. */
export interface DrawingGuide {
  id: number;
  axis: 'x' | 'y';
  position: number;
  color: string;
}
const namespace = 'http://schemas.microsoft.com/office/powerpoint/2012/main';
const uri = '{EFAFB233-063F-42B5-8137-9DF3F51BA10A}';
// MS-PPTX CT_ExtendedGuide uses master units (1/576 inch), not EMU.
const unit = 914400 / 576;
const p = (name: string) => qname('p', name, NS.pml);
const p15 = (name: string) => qname('p15', name, namespace);
const a = (name: string) => qname('a', name, NS.dml);
const attribute = (name: string) => qname('', name, '');
const value = (node: XmlElement, name: string) => getAttrValue(node, attribute(name));
function list(root: XmlElement) {
  const extensions = firstChildElement(root, p('extLst'));
  const extension =
    extensions && allChildElements(extensions, p('ext')).find((node) => value(node, 'uri') === uri);
  return extension ? firstChildElement(extension, p15('sldGuideLst')) : null;
}
function viewProperties(presentation: PresentationData | SlideData) {
  const pkg = presentation[INTERNAL_PACKAGE];
  const rel = pkg
    .getRels(PRES_PART_NAME)
    ?.items.find((item) => item.type === REL_TYPES.viewProps && item.targetMode === 'Internal');
  const part = rel && pkg.getPart(resolveTarget(PRES_PART_NAME, rel.target));
  if (!part) return null;
  const doc = parseXml(decode(part.data));
  const slideView = firstChildElement(doc.root, p('slideViewPr'));
  const common = slideView && firstChildElement(slideView, p('cSldViewPr'));
  return { part, doc, common };
}
function legacyList(presentation: PresentationData) {
  const view = viewProperties(presentation);
  const guides = view?.common && firstChildElement(view.common, p('guideLst'));
  return view && guides ? { ...view, guides } : null;
}
/** Returns null when the presentation has no explicit guide list; [] means cleared. */
export function getDrawingGuides(presentation: PresentationData): DrawingGuide[] | null {
  const part = presentation[INTERNAL_PACKAGE].getPart(PRES_PART_NAME);
  if (!part) return null;
  const extended = list(parseXml(decode(part.data)).root);
  const guides = extended ?? legacyList(presentation)?.guides;
  if (!guides) return null;
  const result: DrawingGuide[] = [];
  const ids = new Set<number>();
  for (const [index, node] of allChildElements(
    guides,
    extended ? p15('guide') : p('guide'),
  ).entries()) {
    const id = extended ? Number(value(node, 'id')) : index + 1;
    const position = Number(value(node, 'pos') ?? 0);
    if (
      (extended && value(node, 'id') === null) ||
      !Number.isInteger(id) ||
      id < 0 ||
      id > 0xffffffff ||
      !Number.isInteger(position) ||
      position < -2147483648 ||
      position > 2147483647 ||
      ids.has(id)
    )
      continue;
    ids.add(id);
    const clr = firstChildElement(node, p15('clr'));
    const rgb = clr && firstChildElement(clr, a('srgbClr'));
    const hex = rgb ? value(rgb, 'val') : null;
    result.push({
      id,
      axis: value(node, 'orient') === 'horz' ? 'y' : 'x',
      position: position * unit,
      color: '#' + (hex && /^[0-9a-f]{6}$/i.test(hex) ? hex.toLowerCase() : '888888'),
    });
  }
  return result;
}
/** Replaces presentation guides, retaining metadata/extensions on surviving guide IDs. */
export const setDrawingGuides = (
  presentation: PresentationData,
  guides: readonly DrawingGuide[],
): void => {
  if (
    !Array.isArray(guides) ||
    new Set(guides.map((guide) => guide?.id)).size !== guides.length ||
    guides.some(
      (guide) =>
        !guide ||
        !Number.isInteger(guide.id) ||
        guide.id < 0 ||
        guide.id > 0xffffffff ||
        !['x', 'y'].includes(guide.axis) ||
        !Number.isFinite(guide.position) ||
        Math.round(guide.position / unit) < -2147483648 ||
        Math.round(guide.position / unit) > 2147483647 ||
        !/^#[0-9a-f]{6}$/i.test(guide.color),
    )
  )
    throw new Error('Invalid drawing guides.');
  const part = presentation[INTERNAL_PACKAGE].getPart(PRES_PART_NAME);
  if (!part) throw new Error('presentation.xml is missing');
  const doc = parseXml(decode(part.data));
  const legacy = legacyList(presentation);
  let target = list(doc.root);
  if (!target) {
    let extensions = firstChildElement(doc.root, p('extLst'));
    if (!extensions) {
      extensions = elem(p('extLst'));
      doc.root.children.push(extensions);
    }
    target = elem(p15('sldGuideLst'), {
      prefixDecls: new Map([
        ['p15', namespace],
        ['a', NS.dml],
      ]),
    });
    const existing = allChildElements(extensions, p('ext')).find(
      (node) => value(node, 'uri') === uri,
    );
    if (existing) existing.children.push(target);
    else
      extensions.children.push(
        elem(p('ext'), { attrs: [attr(attribute('uri'), uri)], children: [target] }),
      );
  }
  target.prefixDecls.set('p15', namespace);
  const old = new Map<number, XmlElement>();
  for (const node of allChildElements(target, p15('guide'))) {
    const id = Number(value(node, 'id'));
    if (!old.has(id)) old.set(id, node);
  }
  const children = guides.map((guide) => {
    const node =
      old.get(guide.id) ?? elem(p15('guide'), { attrs: [attr(attribute('userDrawn'), '1')] });
    for (const [key, val] of Object.entries({
      id: guide.id,
      orient: guide.axis === 'x' ? 'vert' : 'horz',
      pos: Math.round(guide.position / unit),
    })) {
      node.attrs = node.attrs.filter(
        (item) => item.name.namespaceURI !== '' || item.name.localName !== key,
      );
      node.attrs.push(attr(attribute(key), String(val)));
    }
    let clr = firstChildElement(node, p15('clr'));
    const rgb = clr && firstChildElement(clr, a('srgbClr'));
    const previous = rgb ? '#' + value(rgb, 'val')?.toLowerCase() : '#888888';
    if (!clr || previous !== guide.color.toLowerCase()) {
      const replacement = elem(p15('clr'), {
        prefixDecls: new Map([['a', NS.dml]]),
        children: [
          elem(a('srgbClr'), {
            attrs: [attr(attribute('val'), guide.color.slice(1).toUpperCase())],
          }),
        ],
      });
      if (clr) node.children[node.children.indexOf(clr)] = replacement;
      else node.children.unshift(replacement);
    }
    return node;
  });
  target.children = [
    ...children,
    ...target.children.filter(
      (node) =>
        node.kind !== 'element' ||
        node.name.namespaceURI !== namespace ||
        node.name.localName !== 'guide',
    ),
  ];
  if (legacy) {
    legacy.guides.children = guides.map((guide) =>
      elem(p('guide'), {
        prefixDecls: new Map([['p', NS.pml]]),
        attrs: [
          attr(attribute('orient'), guide.axis === 'x' ? 'vert' : 'horz'),
          attr(attribute('pos'), String(Math.round(guide.position / unit))),
        ],
      }),
    );
  }
  const data = encode(serializeXml(doc));
  const legacyData = legacy && encode(serializeXml(legacy.doc));
  part.data = data;
  if (legacy && legacyData) legacy.part.data = legacyData;
};

/** Whether drawing guides are visible in the presentation's slide editing view. */
export function getDrawingGuidesVisible(presentation: PresentationData): boolean {
  const common = viewProperties(presentation)?.common;
  const visible = common && value(common, 'showGuides');
  return visible === '1' || visible === 'true';
}
/** Persists slide-view guide visibility without changing other view preferences. */
export const setDrawingGuidesVisible = (presentation: PresentationData, visible: boolean): void => {
  if (typeof visible !== 'boolean') throw new Error('Invalid guide visibility.');
  updateViewProperties(presentation, (_root, common) => {
    common.attrs = common.attrs.filter(
      (item) => item.name.namespaceURI !== '' || item.name.localName !== 'showGuides',
    );
    common.attrs.push(attr(attribute('showGuides'), visible ? '1' : '0'));
  });
};

/** Document grid snapping; null means no slide-view preferences are stored. */
export function getSnapToGrid(presentation: PresentationData): boolean | null {
  const common = viewProperties(presentation)?.common;
  if (!common) return null;
  const snap = value(common, 'snapToGrid');
  // The OOXML default is true, including native files that omit this attribute.
  return snap !== '0' && snap !== 'false';
}
export const setSnapToGrid = (presentation: PresentationData, enabled: boolean): void => {
  if (typeof enabled !== 'boolean') throw new Error('Invalid grid snapping.');
  updateViewProperties(presentation, (_root, common) => {
    common.attrs = common.attrs.filter(
      (item) => item.name.namespaceURI !== '' || item.name.localName !== 'snapToGrid',
    );
    common.attrs.push(attr(attribute('snapToGrid'), enabled ? '1' : '0'));
  });
};

/** Grid spacing in EMU, or null when no document spacing is specified. */
export function getGridSpacing(presentation: PresentationData): { x: number; y: number } | null {
  const view = viewProperties(presentation);
  const node = view && firstChildElement(view.doc.root, p('gridSpacing'));
  if (!node) return null;
  const x = Number(value(node, 'cx')),
    y = Number(value(node, 'cy'));
  return Number.isSafeInteger(x) && x > 0 && Number.isSafeInteger(y) && y > 0 ? { x, y } : null;
}

export const setGridSpacing = (
  presentation: PresentationData,
  spacing: { x: number; y: number },
): void => {
  if (
    ![spacing?.x, spacing?.y].every((n) => Number.isSafeInteger(n) && n > 0 && n <= 27273042316900)
  )
    throw new Error('Invalid grid spacing.');
  updateViewProperties(presentation, (root) => {
    let node = firstChildElement(root, p('gridSpacing'));
    if (!node) {
      node = elem(p('gridSpacing'));
      const index = root.children.findIndex(
        (child) =>
          child.kind === 'element' &&
          child.name.namespaceURI === NS.pml &&
          child.name.localName === 'extLst',
      );
      root.children.splice(index < 0 ? root.children.length : index, 0, node);
    }
    node.attrs = node.attrs.filter(
      (item) => item.name.namespaceURI !== '' || !['cx', 'cy'].includes(item.name.localName),
    );
    node.attrs.push(
      attr(attribute('cx'), String(spacing.x)),
      attr(attribute('cy'), String(spacing.y)),
    );
  });
};

function updateViewProperties(
  presentation: PresentationData | SlideData,
  update: (root: XmlElement, common: XmlElement) => void,
): void {
  const pkg = presentation[INTERNAL_PACKAGE];
  const rels = pkg.getRels(PRES_PART_NAME) ?? { items: [] };
  const rel = rels.items.find((item) => item.type === REL_TYPES.viewProps);
  if (rel?.targetMode === 'External') throw new Error('View properties must be internal.');
  const existing = viewProperties(presentation);
  const doc = existing?.doc ?? parseXml(`<p:viewPr xmlns:p="${NS.pml}" xmlns:a="${NS.dml}"/>`);
  doc.root.prefixDecls.set('p', NS.pml);
  let slideView = firstChildElement(doc.root, p('slideViewPr'));
  if (!slideView) {
    slideView = elem(p('slideViewPr'));
    const after = doc.root.children.findIndex(
      (node) =>
        node.kind === 'element' &&
        node.name.namespaceURI === NS.pml &&
        node.name.localName !== 'normalViewPr',
    );
    doc.root.children.splice(after < 0 ? doc.root.children.length : after, 0, slideView);
  }
  let common = firstChildElement(slideView, p('cSldViewPr'));
  if (!common) {
    common = elem(p('cSldViewPr'), {
      attrs: [attr(attribute('snapToGrid'), '0')],
      children: [
        elem(p('cViewPr'), {
          prefixDecls: new Map([['a', NS.dml]]),
          children: [
            elem(p('scale'), {
              children: ['sx', 'sy'].map((name) =>
                elem(a(name), {
                  attrs: [attr(attribute('n'), '100'), attr(attribute('d'), '100')],
                }),
              ),
            }),
            elem(p('origin'), { attrs: [attr(attribute('x'), '0'), attr(attribute('y'), '0')] }),
          ],
        }),
      ],
    });
    slideView.children.unshift(common);
  }
  update(doc.root, common);
  const data = encode(serializeXml(doc));
  if (existing) existing.part.data = data;
  else {
    let name = rel ? resolveTarget(PRES_PART_NAME, rel.target) : partName('/ppt/viewProps.xml');
    if (!rel) {
      for (let index = 1; pkg.getPart(name); index++) name = partName(`/ppt/viewProps${index}.xml`);
    }
    pkg.addPart(
      name,
      'application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml',
      data,
    );
    if (!rel) {
      rels.items.push({
        id: nextRelId(rels.items.map((item) => item.id)),
        type: REL_TYPES.viewProps,
        target: name,
        targetMode: 'Internal',
      });
      pkg.setRels(PRES_PART_NAME, rels);
    }
  }
}

/** Slides whose bodies are collapsed in PowerPoint's outline view, in deck order. */
export function getCollapsedOutlineSlides(presentation: PresentationData): readonly SlideData[] {
  const view = viewProperties(presentation);
  const outline = view && firstChildElement(view.doc.root, p('outlineViewPr'));
  const slides = outline && firstChildElement(outline, p('sldLst'));
  if (!view || !slides) return [];
  const rels = presentation[INTERNAL_PACKAGE].getRels(view.part.name);
  const targets = new Map(
    rels?.items
      .filter((rel) => rel.type === REL_TYPES.slide && rel.targetMode === 'Internal')
      .map((rel) => [rel.id, resolveTarget(view.part.name, rel.target)]),
  );
  const collapsed = new Set(
    allChildElements(slides, p('sld'))
      .filter((node) => ['1', 'true'].includes(value(node, 'collapse') ?? ''))
      .map((node) => targets.get(getAttrValue(node, qname('r', 'id', NS.officeDocRels)) ?? '')),
  );
  return getSlides(presentation).filter((slide) => collapsed.has(slide[SLIDE_PART_NAME]));
}

/** Persists outline collapse and retains other view settings. */
export const setSlideOutlineCollapsed = (
  slide: SlideData | readonly SlideData[],
  collapsed: boolean,
): void => {
  if (typeof collapsed !== 'boolean') throw new Error('Invalid outline collapse state.');
  const selected: readonly SlideData[] = 'length' in slide ? slide : [slide];
  if (!selected.length) return;
  const first = selected[0]!;
  const pkg = first[INTERNAL_PACKAGE];
  if (selected.some((item) => item[INTERNAL_PACKAGE] !== pkg))
    throw new Error('Outline slides must belong to the same presentation.');
  updateViewProperties(first, () => {});
  const view = viewProperties(first)!;
  const rels = pkg.getRels(view.part.name) ?? { items: [] };
  const relationships = new Map(
    rels.items
      .filter((rel) => rel.type === REL_TYPES.slide && rel.targetMode === 'Internal')
      .map((rel) => [resolveTarget(view.part.name, rel.target), rel]),
  );
  let availableId = nextRelId(rels.items.map((rel) => rel.id));
  let outline = firstChildElement(view.doc.root, p('outlineViewPr'));
  if (!outline) {
    outline = elem(p('outlineViewPr'), {
      children: [
        elem(p('cViewPr'), {
          children: [
            elem(p('scale'), {
              children: ['sx', 'sy'].map((name) =>
                elem(a(name), {
                  attrs: [attr(attribute('n'), '100'), attr(attribute('d'), '100')],
                }),
              ),
            }),
            elem(p('origin'), { attrs: [attr(attribute('x'), '0'), attr(attribute('y'), '0')] }),
          ],
        }),
      ],
    });
    const after = view.doc.root.children.findIndex(
      (node) =>
        node.kind === 'element' &&
        node.name.namespaceURI === NS.pml &&
        !['normalViewPr', 'slideViewPr'].includes(node.name.localName),
    );
    view.doc.root.children.splice(after < 0 ? view.doc.root.children.length : after, 0, outline);
  }
  let slides = firstChildElement(outline, p('sldLst'));
  if (!slides) {
    slides = elem(p('sldLst'));
    const extension = outline.children.findIndex(
      (node) =>
        node.kind === 'element' &&
        node.name.localName === 'extLst' &&
        node.name.namespaceURI === NS.pml,
    );
    outline.children.splice(extension < 0 ? outline.children.length : extension, 0, slides);
  }
  const entries = new Map(
    allChildElements(slides, p('sld')).map((node) => [
      getAttrValue(node, qname('r', 'id', NS.officeDocRels)),
      node,
    ]),
  );
  for (const selectedSlide of new Set(selected)) {
    const target = selectedSlide[SLIDE_PART_NAME];
    let rel = relationships.get(target);
    if (!rel) {
      rel = { id: availableId, type: REL_TYPES.slide, target, targetMode: 'Internal' };
      availableId = nextRelId([rel.id]);
      rels.items.push(rel);
      relationships.set(target, rel);
    }
    let node = entries.get(rel.id);
    if (!node) {
      node = elem(p('sld'), { attrs: [attr(qname('r', 'id', NS.officeDocRels), rel.id)] });
      slides.children.push(node);
      entries.set(rel.id, node);
    }
    node.attrs = node.attrs.filter(
      (item) => item.name.namespaceURI !== '' || item.name.localName !== 'collapse',
    );
    node.attrs.push(attr(attribute('collapse'), collapsed ? '1' : '0'));
  }
  view.doc.root.prefixDecls.set('r', NS.officeDocRels);
  view.part.data = encode(serializeXml(view.doc));
  pkg.setRels(view.part.name, rels);
};
