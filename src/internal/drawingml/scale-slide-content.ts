import { NS, attr, qname, type XmlElement } from '../xml/index.ts';

// Absolute DrawingML measurements only. Percentages, angles, crop fractions,
// geometry path coordinates, shape IDs and connector references stay unchanged.
const measurements: Record<string, readonly string[]> = {
  bodyPr: ['lIns', 'rIns', 'tIns', 'bIns', 'spcCol'],
  pPr: ['marL', 'marR', 'indent', 'defTabSz'],
  tab: ['pos'],
  rPr: ['sz', 'spc', 'kern'],
  defRPr: ['sz', 'spc', 'kern'],
  endParaRPr: ['sz', 'spc', 'kern'],
  spcPts: ['val'],
  buSzPts: ['val'],
  ln: ['w'],
  lnL: ['w'],
  lnR: ['w'],
  lnT: ['w'],
  lnB: ['w'],
  lnTlToBr: ['w'],
  lnBlToTr: ['w'],
  gridCol: ['w'],
  tr: ['h'],
  tcPr: ['marL', 'marR', 'marT', 'marB'],
  blur: ['rad'],
  glow: ['rad'],
  softEdge: ['rad'],
  outerShdw: ['blurRad', 'dist'],
  innerShdw: ['blurRad', 'dist'],
  reflection: ['blurRad', 'dist'],
};

/** Scale explicit slide measurements while keeping nested group coordinates coherent. */
export function scaleSlideContent(root: XmlElement, factor: number, dx: number, dy: number) {
  const hasPlaceholder = (node: XmlElement): boolean =>
    (node.name.namespaceURI === NS.pml && node.name.localName === 'ph') ||
    node.children.some((child) => child.kind === 'element' && hasPlaceholder(child));
  function visit(node: XmlElement, ancestors: XmlElement[]) {
    const name = node.name.localName;
    const parent = ancestors.at(-1);
    const drawing = node.name.namespaceURI === NS.dml;
    // Materialize schema inset defaults where there is no placeholder cascade.
    // Master placeholders supply these to layout/slide placeholders, which must
    // retain missing values so explicit inherited insets continue to take effect.
    if (drawing && factor !== 1 && (name === 'bodyPr' || name === 'tcPr')) {
      const shape = ancestors.at(-2);
      if (
        name === 'tcPr' ||
        root.name.localName === 'sldMaster' ||
        (shape && !hasPlaceholder(shape))
      ) {
        const defaults =
          name === 'bodyPr'
            ? { lIns: 91440, rIns: 91440, tIns: 45720, bIns: 45720 }
            : { marL: 91440, marR: 91440, marT: 45720, marB: 45720 };
        for (const [key, value] of Object.entries(defaults)) {
          if (
            !node.attrs.some(
              (attribute) => !attribute.name.namespaceURI && attribute.name.localName === key,
            )
          )
            node.attrs.push(attr(qname('', key, ''), String(value)));
        }
      }
    }
    const transform =
      parent?.name.localName === 'xfrm' &&
      (parent.name.namespaceURI === NS.dml || parent.name.namespaceURI === NS.pml);
    const rootGroupTransform =
      transform &&
      ancestors.at(-2)?.name.localName === 'grpSpPr' &&
      ancestors.at(-3)?.name.localName === 'spTree';
    let keys: readonly string[] = drawing ? (measurements[name] ?? []) : [];
    if (drawing && /^lvl[1-9]pPr$/.test(name)) keys = measurements.pPr!;
    if (drawing && transform && !rootGroupTransform) {
      if (name === 'off' || name === 'chOff') keys = ['x', 'y'];
      if (name === 'ext' || name === 'chExt') keys = ['cx', 'cy'];
    }
    // The top-level shape owns its slide-space offset. All offsets inside
    // a group use that group's child space and must not receive the translation.
    const groups = ancestors.filter(
      (a) => a.name.namespaceURI === NS.pml && a.name.localName === 'grpSp',
    );
    const owner =
      ancestors.at(-2)?.name.localName === 'graphicFrame' ? ancestors.at(-2) : ancestors.at(-3);
    const topOffset =
      drawing &&
      transform &&
      name === 'off' &&
      !rootGroupTransform &&
      groups.length === (owner?.name.localName === 'grpSp' ? 1 : 0);
    node.attrs = node.attrs.map((attribute) => {
      if (attribute.name.namespaceURI || !keys.includes(attribute.name.localName)) return attribute;
      const value = Number(attribute.value);
      if (!Number.isFinite(value)) throw new Error('Invalid slide measurement.');
      let scaled = value * factor;
      if (topOffset) scaled += attribute.name.localName === 'x' ? dx : dy;
      scaled = Math.round(scaled);
      if (!Number.isSafeInteger(scaled))
        throw new RangeError('Scaled slide measurement exceeds the supported range.');
      return { ...attribute, value: String(scaled) };
    });
    for (const child of node.children)
      if (child.kind === 'element') visit(child, [...ancestors, node]);
  }
  visit(root, []);
}
