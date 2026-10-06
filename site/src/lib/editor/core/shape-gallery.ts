// Insert ▸ Shapes and Home ▸ Shapes: PowerPoint's shape gallery, by category,
// limited to the presets the library writes. Icons come from one render of
// every preset on a scratch slide, used as a sprite.

import {
  addSlide,
  addSlideLine,
  addSlideShape,
  createPresentation,
  findSlideLayout,
  inches,
  setShapeNoFill,
  setShapeStroke,
  setSlideSize,
  type PresetShape,
} from '@office-kit/pptx';
import { renderSlideToSvg } from '@office-kit/pptx-preview';

export type GalleryShape = PresetShape | 'line';

export const SHAPE_GALLERY: readonly {
  readonly title: string;
  readonly shapes: readonly GalleryShape[];
}[] = [
  { title: 'Lines', shapes: ['line'] },
  { title: 'Rectangles', shapes: ['rect', 'roundRect'] },
  {
    title: 'Basic Shapes',
    shapes: [
      'ellipse',
      'triangle',
      'rtTriangle',
      'parallelogram',
      'trapezoid',
      'diamond',
      'pentagon',
      'hexagon',
      'heptagon',
      'octagon',
      'decagon',
      'plus',
      'can',
      'cube',
      'donut',
      'noSmoking',
      'heart',
      'lightningBolt',
      'sun',
      'moon',
      'cloud',
      'bracketPair',
      'bracePair',
      'leftBracket',
      'rightBracket',
      'leftBrace',
      'rightBrace',
    ],
  },
  {
    title: 'Block Arrows',
    shapes: [
      'rightArrow',
      'leftArrow',
      'upArrow',
      'downArrow',
      'leftRightArrow',
      'upDownArrow',
      'bentArrow',
      'curvedRightArrow',
      'curvedLeftArrow',
      'curvedUpArrow',
      'curvedDownArrow',
    ],
  },
  {
    title: 'Equation Shapes',
    shapes: ['mathPlus', 'mathMinus', 'mathMultiply', 'mathDivide', 'mathEqual', 'mathNotEqual'],
  },
  {
    title: 'Stars and Banners',
    shapes: [
      'star4',
      'star5',
      'star6',
      'star7',
      'star8',
      'star10',
      'star12',
      'star16',
      'star24',
      'star32',
    ],
  },
];

const SPRITE_COLUMNS = 10;
const CELL_IN = 1;
const INSET_IN = 0.15;
const ICON_STROKE = '#595959';

export interface ShapeSprite {
  readonly url: string;
  readonly columns: number;
  readonly rows: number;
  /** Cell index of each shape in the sprite. */
  readonly cells: ReadonlyMap<GalleryShape, number>;
}

let sprite: ShapeSprite | null = null;

/** The gallery icons, rendered once on first use. */
export function shapeSprite(): ShapeSprite {
  if (sprite) return sprite;
  const all = SHAPE_GALLERY.flatMap((group) => group.shapes);
  const rows = Math.ceil(all.length / SPRITE_COLUMNS);
  const pres = createPresentation();
  setSlideSize(pres, { width: inches(SPRITE_COLUMNS * CELL_IN), height: inches(rows * CELL_IN) });
  const slide = addSlide(pres, { layout: findSlideLayout(pres, 'Blank')! });
  const cells = new Map<GalleryShape, number>();
  all.forEach((preset, index) => {
    cells.set(preset, index);
    const x = (index % SPRITE_COLUMNS) * CELL_IN + INSET_IN;
    const y = Math.floor(index / SPRITE_COLUMNS) * CELL_IN + INSET_IN;
    const size = CELL_IN - 2 * INSET_IN;
    if (preset === 'line') {
      addSlideLine(slide, {
        from: { x: inches(x), y: inches(y + size) },
        to: { x: inches(x + size), y: inches(y) },
        color: ICON_STROKE,
        widthEmu: 19050,
      });
      return;
    }
    const shape = addSlideShape(slide, {
      preset,
      x: inches(x),
      y: inches(y),
      w: inches(size),
      h: inches(size),
    });
    setShapeNoFill(shape);
    setShapeStroke(shape, { color: ICON_STROKE, widthEmu: 19050 });
  });
  sprite = {
    url: `data:image/svg+xml,${encodeURIComponent(renderSlideToSvg(pres, slide, { textLayout: 'svg' }))}`,
    columns: SPRITE_COLUMNS,
    rows,
    cells,
  };
  return sprite;
}
