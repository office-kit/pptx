import {
  INTERNAL_PACKAGE,
  SLIDE_PART_NAME,
  SLIDE_SHAPES,
  type SlideData,
  type SlideShapeData,
} from '../_internal-symbols.ts';
import { buildInk, type InkStroke } from '../../internal/presentationml/ink-builder.ts';
import { buildInkContent } from '../../internal/presentationml/ink-content-builder.ts';
import { REL_TYPES } from '../../internal/presentationml/relationship-types.ts';
import {
  emptyRels,
  nextRelId,
  partName,
  detectImageFormat,
  readImagePixelSize,
} from '../../internal/opc/index.ts';
import { serializeFragment } from '../../internal/xml/index.ts';
import {
  commitSlideData,
  encode,
  requireSpTree,
  nextShapeId,
  rebuildShapesFromDocument,
} from './_helpers.ts';

export type { InkStroke } from '../../internal/presentationml/ink-builder.ts';

/** Physical ink bounds, including the pen radius, in slide EMU coordinates. */
export const getInkBounds = (strokes: ReadonlyArray<InkStroke>) => buildInk(strokes).bounds;

/** Adds editable native ink with a PNG fallback cropped to getInkBounds(strokes). */
export const addSlideInk = (
  slide: SlideData,
  strokes: ReadonlyArray<InkStroke>,
  fallbackPng: Uint8Array,
): SlideShapeData => {
  appendSlideInk(slide, strokes, fallbackPng);
  return slide[SLIDE_SHAPES].at(-1)!;
};

/** Append the native content part and its compatibility picture.
 * The PNG must depict the strokes cropped to buildInk(strokes).bounds.
 */
export const appendSlideInk = (
  slide: SlideData,
  strokes: ReadonlyArray<InkStroke>,
  fallbackPng: Uint8Array,
): number => {
  const ink = buildInk(strokes);
  if (detectImageFormat(fallbackPng) !== 'png' || !readImagePixelSize(fallbackPng))
    throw new Error('ink fallback requires a PNG with valid dimensions');
  const pkg = slide[INTERNAL_PACKAGE];
  const tree = requireSpTree(slide);
  const id = nextShapeId(slide);
  let index = 1;
  while (
    pkg.getPart(partName(`/ppt/ink/ink${index}.xml`)) ||
    pkg.getPart(partName(`/ppt/media/ink${index}.png`))
  )
    index++;
  const rels = pkg.getRels(slide[SLIDE_PART_NAME]) ?? emptyRels();
  const inkId = nextRelId(rels.items.map((rel) => rel.id));
  const pictureId = nextRelId([...rels.items.map((rel) => rel.id), inkId]);
  const content = buildInkContent({
    id,
    inkRelId: inkId,
    pictureRelId: pictureId,
    bounds: ink.bounds,
  });
  const inkBytes = encode(serializeFragment(ink.root));
  // Validate/build everything before mutating the package.
  pkg.addPart(partName(`/ppt/ink/ink${index}.xml`), 'application/inkml+xml', inkBytes);
  pkg.addPart(partName(`/ppt/media/ink${index}.png`), 'image/png', fallbackPng.slice());
  rels.items.push(
    {
      id: inkId,
      type: REL_TYPES.customXml,
      target: `../ink/ink${index}.xml`,
      targetMode: 'Internal',
    },
    {
      id: pictureId,
      type: REL_TYPES.image,
      target: `../media/ink${index}.png`,
      targetMode: 'Internal',
    },
  );
  pkg.setRels(slide[SLIDE_PART_NAME], rels);
  tree.children.push(content);
  commitSlideData(slide);
  rebuildShapesFromDocument(slide);
  return id;
};
