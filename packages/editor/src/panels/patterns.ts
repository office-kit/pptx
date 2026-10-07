import {
  addBlankSlide,
  addSlideShape,
  createPresentation,
  inches,
  setShapePatternFill,
  setShapeNoStroke,
  setSlideSize,
  type PatternPreset,
} from '@office-kit/pptx';
import { renderSlideToSvg } from '@office-kit/pptx-preview';

// Order and labels from the Mac PowerPoint Format Shape pattern gallery.
export const patterns = [
  ['pct5', 'Dotted: 5%'],
  ['pct10', 'Dotted: 10%'],
  ['pct20', 'Dotted: 20%'],
  ['pct25', 'Dotted: 25%'],
  ['pct30', 'Dotted: 30%'],
  ['pct40', 'Dotted: 40%'],
  ['pct50', 'Dotted: 50%'],
  ['pct60', 'Dotted: 60%'],
  ['pct70', 'Dotted: 70%'],
  ['pct75', 'Dotted: 75%'],
  ['pct80', 'Dotted: 80%'],
  ['pct90', 'Dotted: 90%'],
  ['ltDnDiag', 'Diagonal stripes: Light downward'],
  ['ltUpDiag', 'Diagonal stripes: Light upward'],
  ['dkDnDiag', 'Diagonal stripes: Dark downward'],
  ['dkUpDiag', 'Diagonal stripes: Dark upward'],
  ['wdDnDiag', 'Diagonal stripes: Wide downward'],
  ['wdUpDiag', 'Diagonal stripes: Wide upward'],
  ['ltVert', 'Vertical stripes: light'],
  ['ltHorz', 'Horizontal stripes: Light'],
  ['narVert', 'Vertical stripes: Narrow'],
  ['narHorz', 'Horizontal stripes: Narrow'],
  ['dkVert', 'Vertical stripes: Dark'],
  ['dkHorz', 'Horizontal stripes: Dark'],
  ['dashDnDiag', 'Stripes: Backslashes'],
  ['dashUpDiag', 'Stripes: Slashes'],
  ['dashHorz', 'Horizontal stripes: Alternating horizontal lines'],
  ['dashVert', 'Vertical stripes: Alternating vertical lines'],
  ['smConfetti', 'Small confetti'],
  ['lgConfetti', 'Large confetti'],
  ['zigZag', 'Zig zag'],
  ['wave', 'Wave'],
  ['diagBrick', 'Diagonal brick'],
  ['horzBrick', 'Horizontal brick'],
  ['weave', 'Weave'],
  ['plaid', 'Plaid'],
  ['divot', 'Divot'],
  ['dotGrid', 'Dotted grid'],
  ['dotDmnd', 'Dotted diamond grid'],
  ['shingle', 'Shingle'],
  ['trellis', 'Trellis'],
  ['sphere', 'Sphere'],
  ['smGrid', 'Small grid'],
  ['lgGrid', 'Large grid'],
  ['smCheck', 'Small checker board'],
  ['lgCheck', 'Large checker board'],
  ['openDmnd', 'Outlined diamond grid'],
  ['solidDmnd', 'Solid diamond grid'],
] as const satisfies ReadonlyArray<readonly [PatternPreset, string]>;

let swatches: Map<PatternPreset, string> | undefined;
export function patternSwatches(): ReadonlyMap<PatternPreset, string> {
  if (swatches) return swatches;
  const pres = createPresentation();
  const width = inches(1.5);
  const height = inches(1);
  setSlideSize(pres, { width, height });
  const slide = addBlankSlide(pres);
  const shape = addSlideShape(slide, {
    preset: 'rect',
    x: inches(0),
    y: inches(0),
    w: width,
    h: height,
  });
  setShapeNoStroke(shape);
  swatches = new Map(
    patterns.map(([preset]) => {
      setShapePatternFill(shape, { preset, foreground: 'accent1', background: '#FFFFFF' });
      return [preset, `data:image/svg+xml,${encodeURIComponent(renderSlideToSvg(pres, slide))}`];
    }),
  );
  return swatches;
}
