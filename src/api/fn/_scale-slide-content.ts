// Uniform page fitting. Restrict edits to dimensional DrawingML properties;
// image bytes, crop fractions, geometry paths, relationships and timing stay intact.
import { boundedInt } from '../../internal/bounds.ts';
import {
  NS,
  attr,
  elem,
  getAttrValue,
  qname,
  firstChildElement,
  type XmlElement,
} from '../../internal/xml/index.ts';

const SHAPES = new Set(['sp', 'pic', 'cxnSp', 'graphicFrame', 'grpSp']);
const TEXT_PROPS = new Set(['rPr', 'defRPr', 'endParaRPr']);
const PARAGRAPH_PROPS = new Set([
  'pPr',
  'defPPr',
  ...Array.from({ length: 9 }, (_, i) => `lvl${i + 1}pPr`),
]);

// Unsized table text uses the editor's 18pt default. Record it before scaling
// only where no local font-size source exists, so authored inheritance survives.
const materializeTableFontDefaults = (body: XmlElement): void => {
  const child = (node: XmlElement | null, name: string) =>
    node && firstChildElement(node, qname('a', name, NS.dml));
  const size = (node: XmlElement | null) => node && getAttrValue(node, qname('', 'sz', ''));
  const ensureSize = (parent: XmlElement, name: string): void => {
    let props = child(parent, name);
    if (!props) {
      props = elem(qname('a', name, NS.dml));
      if (name === 'rPr') parent.children.unshift(props);
      else parent.children.push(props);
    }
    if (size(props) === null) props.attrs.push(attr(qname('', 'sz', ''), '1800'));
  };
  const list = child(body, 'lstStyle');
  for (const paragraph of body.children) {
    if (
      paragraph.kind !== 'element' ||
      paragraph.name.namespaceURI !== NS.dml ||
      paragraph.name.localName !== 'p'
    )
      continue;
    const pPr = child(paragraph, 'pPr');
    const level = Number(pPr && getAttrValue(pPr, qname('', 'lvl', ''))) || 0;
    const levelProps =
      child(list, `lvl${level + 1}pPr`) ?? (level === 0 ? child(list, 'defPPr') : null);
    if (size(child(pPr, 'defRPr')) !== null || size(child(levelProps, 'defRPr')) !== null) continue;
    const runs = paragraph.children.filter(
      (node): node is XmlElement =>
        node.kind === 'element' &&
        node.name.namespaceURI === NS.dml &&
        ['r', 'fld', 'br'].includes(node.name.localName),
    );
    const endSize = size(child(paragraph, 'endParaRPr'));
    for (const run of runs) {
      if (run === runs.at(-1) && endSize !== null) continue;
      ensureSize(run, 'rPr');
    }
    ensureSize(paragraph, 'endParaRPr');
  }
};

/** Scales coordinates and physical formatting; centers only top-level objects. */
export const scaleSlideContent = (
  root: XmlElement,
  scale: number,
  dx: number,
  dy: number,
): void => {
  const visit = (
    node: XmlElement,
    shapeDepth: number,
    parent: string,
    inheritedBody = false,
  ): void => {
    const { localName: name, namespaceURI: ns } = node.name;
    const depth = shapeDepth + (ns === NS.pml && SHAPES.has(name) ? 1 : 0);
    const drawing = ns === NS.dml;
    if (drawing && name === 'txBody' && parent === 'tc') materializeTableFontDefaults(node);
    if (ns === NS.pml && SHAPES.has(name)) {
      const nonVisual = firstChildElement(node, qname('p', 'nvSpPr', NS.pml));
      const application = nonVisual && firstChildElement(nonVisual, qname('p', 'nvPr', NS.pml));
      inheritedBody =
        root.name.localName !== 'sldMaster' &&
        !!application &&
        firstChildElement(application, qname('p', 'ph', NS.pml)) !== null;
    }
    const defaults = (values: Record<string, number>) => {
      for (const [key, value] of Object.entries(values)) {
        if (!node.attrs.some((a) => a.name.namespaceURI === '' && a.name.localName === key)) {
          node.attrs.push(attr(qname('', key, ''), String(value)));
        }
      }
    };
    // Materialize defaults on standalone bodies and master placeholders only;
    // slide/layout placeholders must keep inheriting their authored margins.
    if (drawing && name === 'bodyPr' && !inheritedBody)
      defaults({ lIns: 91440, rIns: 91440, tIns: 45720, bIns: 45720 });
    if (drawing && name === 'tcPr')
      defaults({ marL: 91440, marR: 91440, marT: 45720, marB: 45720 });
    for (let index = 0; index < node.attrs.length; index++) {
      const a = node.attrs[index]!;
      if (a.name.namespaceURI !== '') continue;
      const key = a.name.localName;
      let range: Parameters<typeof boundedInt>[1] | undefined;
      let offset = 0;
      if (drawing && parent === 'xfrm') {
        if ((name === 'off' || name === 'chOff') && (key === 'x' || key === 'y')) {
          range = 'coordinate';
          if (name === 'off' && depth === 1) offset = key === 'x' ? dx : dy;
        }
        if ((name === 'ext' || name === 'chExt') && (key === 'cx' || key === 'cy'))
          range = 'positiveCoordinate';
      }
      if (drawing && TEXT_PROPS.has(name)) {
        if (key === 'sz') range = 'fontSize';
        if (key === 'spc') range = 'textPoint';
        if (key === 'kern') range = 'textNonNegativePoint';
      }
      if (
        drawing &&
        PARAGRAPH_PROPS.has(name) &&
        ['marL', 'marR', 'indent', 'defTabSz'].includes(key)
      )
        range =
          key === 'indent' ? 'textIndent' : key === 'defTabSz' ? 'coordinate32' : 'textMargin';
      if (drawing && name === 'tab' && key === 'pos') range = 'coordinate32';
      if (drawing && name === 'spcPts' && key === 'val') range = 'textSpacingPoint';
      if (drawing && name === 'buSzPts' && key === 'val') range = 'fontSize';
      if (drawing && name === 'bodyPr' && ['lIns', 'rIns', 'tIns', 'bIns', 'spcCol'].includes(key))
        range = key === 'spcCol' ? 'positiveCoordinate32' : 'coordinate32';
      if (drawing && name === 'tcPr' && ['marL', 'marR', 'marT', 'marB'].includes(key))
        range = 'coordinate32';
      if (drawing && ((name === 'gridCol' && key === 'w') || (name === 'tr' && key === 'h')))
        range = 'positiveCoordinate';
      if (
        drawing &&
        (name === 'ln' ||
          (name.startsWith('ln') &&
            ['lnL', 'lnR', 'lnT', 'lnB', 'lnTlToBr', 'lnBlToTr'].includes(name))) &&
        key === 'w'
      )
        range = 'lineWidth';
      if (
        drawing &&
        ['outerShdw', 'innerShdw', 'reflection'].includes(name) &&
        ['blurRad', 'dist'].includes(key)
      )
        range = 'positiveCoordinate';
      if (
        drawing &&
        ['glow', 'softEdge', 'blur'].includes(name) &&
        ['rad', 'blurRad'].includes(key)
      )
        range = 'positiveCoordinate';
      if (range)
        node.attrs[index] = attr(
          a.name,
          String(
            boundedInt(Number(a.value) * scale + offset, range, `setSlideSize: ${name}@${key}`),
          ),
        );
    }
    for (const child of node.children)
      if (child.kind === 'element') visit(child, depth, name, inheritedBody);
  };
  visit(root, 0, '');
};
