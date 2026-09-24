import { NS, attr, elem, qname, type XmlElement } from '../xml/index.ts';
import { buildPicture } from './picture-builder.ts';
import type { BuiltInk } from './ink-builder.ts';

/** Office 2010 ink content with a picture fallback for older readers. */
export const buildInkContent = (options: {
  id: number;
  name?: string;
  inkRelId: string;
  pictureRelId: string;
  bounds: BuiltInk['bounds'];
}): XmlElement => {
  if (!Number.isInteger(options.id) || options.id < 1 || options.id > 4294967295)
    throw new RangeError('ink shape id must be a positive unsigned integer');
  const label = options.name ?? `Ink ${options.id}`;
  const attribute = (key: string, value: string | number) =>
    attr(qname('', key, ''), String(value));
  const { x, y, cx, cy } = options.bounds;
  // Build the fallback first so its geometry validation also covers the choice.
  const picture = buildPicture({
    id: options.id,
    name: label,
    rEmbed: options.pictureRelId,
    x,
    y,
    w: cx,
    h: cy,
    lockAspect: false,
  });
  const content = elem(qname('p', 'contentPart', NS.pml), {
    attrs: [
      attr(qname('p14', 'bwMode', NS.p14), 'auto'),
      attr(qname('r', 'id', NS.officeDocRels), options.inkRelId),
    ],
    children: [
      elem(qname('p14', 'nvContentPartPr', NS.p14), {
        children: [
          elem(qname('p14', 'cNvPr', NS.p14), {
            attrs: [attribute('id', options.id), attribute('name', label)],
          }),
          elem(qname('p14', 'cNvContentPartPr', NS.p14)),
          elem(qname('p14', 'nvPr', NS.p14)),
        ],
      }),
      elem(qname('p14', 'xfrm', NS.p14), {
        children: [
          elem(qname('a', 'off', NS.dml), {
            attrs: [attribute('x', Math.round(x)), attribute('y', Math.round(y))],
          }),
          elem(qname('a', 'ext', NS.dml), {
            attrs: [attribute('cx', Math.round(cx)), attribute('cy', Math.round(cy))],
          }),
        ],
      }),
    ],
  });
  return elem(qname('mc', 'AlternateContent', NS.mc), {
    prefixDecls: new Map([
      ['mc', NS.mc],
      ['p', NS.pml],
      ['a', NS.dml],
      ['r', NS.officeDocRels],
    ]),
    children: [
      elem(qname('mc', 'Choice', NS.mc), {
        prefixDecls: new Map([['p14', NS.p14]]),
        attrs: [attribute('Requires', 'p14')],
        children: [content],
      }),
      elem(qname('mc', 'Fallback', NS.mc), { children: [picture] }),
    ],
  });
};
