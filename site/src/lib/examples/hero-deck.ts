// The deck shown on the landing page. The page renders this exact code next
// to the slide it produces, and the "Download .pptx" button runs it in the
// visitor's browser — so what is on screen is always what the library does.

import {
  addBlankSlide,
  addSlideChart,
  addSlideShape,
  addSlideTextBox,
  createPresentation,
  inches,
  setShapeFill,
  setShapeNoStroke,
  setShapeTextFormat,
  type PresentationData,
} from '@office-kit/pptx';

export function buildHeroDeck(): PresentationData {
  const pres = createPresentation();
  const slide = addBlankSlide(pres);

  const bar = addSlideShape(slide, {
    preset: 'rect',
    x: inches(0.6),
    y: inches(0.62),
    w: inches(0.12),
    h: inches(0.95),
  });
  setShapeFill(bar, '#E5481F');
  setShapeNoStroke(bar);

  const title = addSlideTextBox(slide, {
    x: inches(0.85),
    y: inches(0.5),
    w: inches(9),
    h: inches(0.7),
    text: 'Revenue grew 2.5× in four quarters',
  });
  setShapeTextFormat(title, { size: 30, bold: true, color: '#15171C' });

  const subtitle = addSlideTextBox(slide, {
    x: inches(0.85),
    y: inches(1.15),
    w: inches(9),
    h: inches(0.45),
    text: 'FY26 plan, USD thousands',
  });
  setShapeTextFormat(subtitle, { size: 16, color: '#5B616E' });

  addSlideChart(slide, {
    x: inches(0.6),
    y: inches(1.9),
    w: inches(12.1),
    h: inches(5),
    spec: {
      kind: 'column',
      categories: ['Q1', 'Q2', 'Q3', 'Q4'],
      series: [
        { name: 'Revenue', values: [120, 180, 240, 300], color: '#E5481F' },
        { name: 'Cost', values: [80, 90, 130, 160], color: '#C9CDD6' },
      ],
      legend: { position: 'b', textStyle: { sizePt: 14, color: '#5B616E' } },
      categoryAxisLabelStyle: { sizePt: 14, color: '#5B616E' },
      valueAxisLabelStyle: { sizePt: 14, color: '#5B616E' },
      valueAxisMajorGridlines: true,
      valueAxisMajorGridlineColor: '#E2E4E9',
      gapWidthPct: 80,
    },
  });

  return pres;
}
