import * as pptx from '@office-kit/pptx';
import { renderInkToPng } from '@office-kit/pptx-preview/node';

export interface ShowInkAnnotation {
  slide: string;
  strokes: pptx.InkStroke[];
}

/** Prepare every target and raster before changing any slide. */
export function keepShowInk(
  presentation: pptx.PresentationData,
  annotations: ShowInkAnnotation[] | undefined,
) {
  if (!Array.isArray(annotations) || !annotations.length || annotations.length > 10000)
    throw new Error('Invalid ink annotations.');
  const slides = new Map(
    pptx.getSlides(presentation).map((slide) => [pptx.getSlidePartName(slide), slide]),
  );
  let strokeCount = 0,
    pointCount = 0;
  const validated = annotations.map((annotation) => {
    if (
      !annotation ||
      typeof annotation.slide !== 'string' ||
      !Array.isArray(annotation.strokes) ||
      !annotation.strokes.length
    )
      throw new Error('Invalid ink annotations.');
    const slide = slides.get(annotation.slide);
    if (!slide) throw new Error('Annotated slide no longer exists.');
    strokeCount += annotation.strokes.length;
    if (strokeCount > 10000) throw new Error('Too many ink strokes.');
    for (const stroke of annotation.strokes) {
      if (
        !stroke ||
        !Array.isArray(stroke.points) ||
        typeof stroke.color !== 'string' ||
        typeof stroke.widthEmu !== 'number'
      )
        throw new Error('Invalid ink stroke.');
      pointCount += stroke.points.length;
      if (pointCount > 200000) throw new Error('Too many ink points.');
      for (const point of stroke.points)
        if (
          !point ||
          typeof point.x !== 'number' ||
          typeof point.y !== 'number' ||
          Math.abs(point.x) > 1e9 ||
          Math.abs(point.y) > 1e9
        )
          throw new Error('Invalid ink point.');
      pptx.getInkBounds([stroke]);
    }
    return { slide, strokes: annotation.strokes };
  });
  const prepared = validated.flatMap(({ slide, strokes }) =>
    strokes.map((stroke) => ({ slide, stroke, png: renderInkToPng([stroke]) })),
  );
  for (const { slide, stroke, png } of prepared) pptx.addSlideInk(slide, [stroke], png);
}
