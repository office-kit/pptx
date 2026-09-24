import {
  LAYOUT_PART,
  SHAPE_ELEMENT,
  type PresentationData,
  type SlideData,
  type SlideLayoutData,
} from '../_internal-symbols.ts';
import { buildSlideFromLayout } from '../../internal/presentationml/slide-builder.ts';
import {
  NS,
  attr,
  elem,
  firstChildElement,
  getAttrValue,
  qname,
  type XmlElement,
} from '../../internal/xml/index.ts';
import {
  commitSlideData,
  rebuildShapesFromDocument,
  requireSpTree,
  nextShapeId,
} from './_helpers.ts';
import { getSlideLayout, getSlideShapes, setSlideLayout } from './shape-slide-read.ts';
import { getShapeBoundsResolved, getShapeRotation, getShapeFlip } from './shape-read-base.ts';

const p = (name: string) => qname('p', name, NS.pml);
const a = (name: string) => qname('a', name, NS.dml);
const value = (node: XmlElement, name: string) => getAttrValue(node, qname('', name, ''));
const ph = (node: XmlElement): XmlElement | null => {
  for (const kind of ['nvSpPr', 'nvPicPr', 'nvGraphicFramePr']) {
    const nv = firstChildElement(node, p(kind));
    const props = nv && firstChildElement(nv, p('nvPr'));
    const placeholder = props && firstChildElement(props, p('ph'));
    if (placeholder) return placeholder;
  }
  return null;
};
const category = (node: XmlElement) => {
  const type = value(node, 'type') ?? 'obj';
  return ['title', 'ctrTitle'].includes(type)
    ? 'title'
    : ['body', 'obj', 'subTitle'].includes(type)
      ? 'text'
      : type;
};
const remove = (node: XmlElement, namespace: string, names: string[]) => {
  node.children = node.children.filter(
    (child) =>
      child.kind !== 'element' ||
      child.name.namespaceURI !== namespace ||
      !names.includes(child.name.localName),
  );
};

/** Apply a layout to existing content, reconciling text placeholders by slot and role.
 * Unmatched content and ordinary shapes are retained. Reset additionally clears
 * placeholder formatting overrides while preserving text, fields and links.
 */
export const applySlideLayout = (
  pres: PresentationData,
  slide: SlideData,
  layout: SlideLayoutData,
  options: { reset?: boolean } = {},
): void => {
  const csld = firstChildElement(layout[LAYOUT_PART].root, p('cSld'));
  const layoutTree = csld && firstChildElement(csld, p('spTree'));
  if (!layoutTree) throw new Error('Layout has no shape tree.');
  const stubsDoc = buildSlideFromLayout(layoutTree);
  const stubTree = firstChildElement(firstChildElement(stubsDoc.root, p('cSld'))!, p('spTree'))!;
  const tree = requireSpTree(slide);
  const existing = tree.children.filter(
    (child): child is XmlElement => child.kind === 'element' && ph(child) !== null,
  );
  const unused = new Set(existing);
  let nextId = nextShapeId(slide);
  // Preserve inherited geometry for content that has no counterpart in the new layout.
  const frames = new Map(
    getSlideShapes(slide).map((shape) => [
      shape[SHAPE_ELEMENT],
      {
        bounds: getShapeBoundsResolved(pres, shape),
        rotation: getShapeRotation(shape),
        flip: getShapeFlip(shape),
      },
    ]),
  );
  setSlideLayout(slide, layout);
  for (const stub of stubTree.children) {
    if (stub.kind !== 'element') continue;
    const target = ph(stub);
    if (!target) continue;
    const match =
      [...unused].find(
        (node) =>
          category(ph(node)!) === category(target) &&
          (value(ph(node)!, 'idx') ?? '0') === (value(target, 'idx') ?? '0'),
      ) ?? [...unused].find((node) => category(ph(node)!) === category(target));
    if (!match) {
      if (['dt', 'ftr', 'sldNum', 'hdr'].includes(value(target, 'type') ?? '')) continue;
      const nv = firstChildElement(stub, p('nvSpPr'))!;
      const paragraph = firstChildElement(firstChildElement(stub, p('txBody'))!, a('p'));
      if (paragraph) paragraph.children = [];
      const id = firstChildElement(nv, p('cNvPr'))!;
      id.attrs = id.attrs.map((item) =>
        item.name.localName === 'id' ? attr(item.name, String(nextId++)) : item,
      );
      tree.children.push(stub);
      continue;
    }
    unused.delete(match);
    ph(match)!.attrs = target.attrs.map((item) => ({ ...item }));
    const props = firstChildElement(match, p('spPr'));
    if (props) {
      if (options.reset) props.children = [];
      else remove(props, NS.dml, ['xfrm']);
    }
    if (options.reset) {
      remove(match, NS.pml, ['style']);
      const body = firstChildElement(match, p('txBody'));
      if (body) {
        for (const child of body.children) {
          if (child.kind !== 'element') continue;
          if (['bodyPr', 'lstStyle'].includes(child.name.localName)) {
            child.attrs = [];
            child.children = [];
          }
          if (child.name.localName !== 'p') continue;
          const props = firstChildElement(child, a('pPr'));
          if (props) {
            props.attrs = props.attrs.filter((item) => item.name.localName === 'lvl');
            props.children = [];
          }
          for (const run of child.children) {
            if (run.kind !== 'element') continue;
            const rpr =
              run.name.localName === 'endParaRPr' ? run : firstChildElement(run, a('rPr'));
            if (!rpr) continue;
            rpr.attrs = rpr.attrs.filter((item) =>
              ['lang', 'altLang'].includes(item.name.localName),
            );
            rpr.children = rpr.children.filter(
              (item) =>
                item.kind === 'element' &&
                ['hlinkClick', 'hlinkMouseOver'].includes(item.name.localName),
            );
          }
        }
      }
    }
  }
  for (const node of unused) {
    const frame = frames.get(node);
    const props = firstChildElement(node, p('spPr'));
    if (!frame?.bounds || !props || firstChildElement(props, a('xfrm'))) continue;
    const box = frame.bounds;
    props.children.unshift(
      elem(a('xfrm'), {
        attrs: [
          attr(qname('', 'rot', ''), String(frame.rotation * 60000)),
          ...(frame.flip?.horizontal ? [attr(qname('', 'flipH', ''), '1')] : []),
          ...(frame.flip?.vertical ? [attr(qname('', 'flipV', ''), '1')] : []),
        ],
        children: [
          elem(a('off'), {
            attrs: [
              attr(qname('', 'x', ''), String(box.x)),
              attr(qname('', 'y', ''), String(box.y)),
            ],
          }),
          elem(a('ext'), {
            attrs: [
              attr(qname('', 'cx', ''), String(box.w)),
              attr(qname('', 'cy', ''), String(box.h)),
            ],
          }),
        ],
      }),
    );
  }
  commitSlideData(slide);
  rebuildShapesFromDocument(slide);
};

export const resetSlideLayout = (pres: PresentationData, slide: SlideData): void => {
  const layout = getSlideLayout(slide);
  if (layout) applySlideLayout(pres, slide, layout, { reset: true });
};
