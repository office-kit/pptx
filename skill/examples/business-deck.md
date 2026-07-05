# Worked example — a 7-slide business deck from scratch

This is the canonical end-to-end example referenced by [`../SKILL.md`](../SKILL.md).
The same code is exercised by `test/skill-example.test.ts`, which asserts the
result is structurally valid, schema-valid, and round-trips — so this is known
to produce a clean `.pptx`, not just plausible-looking code.

It demonstrates the rules from the skill: a single two-color palette reused
across slides, varied layouts (title → bulleted agenda → section divider → stat
cards → table → chart → image), numbers shown as a chart, and speaker notes.

```ts
import {
  createPresentation,
  addTitleSlide,
  addContentSlide,
  addSectionHeaderSlide,
  addBlankSlide,
  addSlideShape,
  addSlideTextBox,
  addSlideTable,
  addSlideChart,
  addSlideImage,
  findSlidePlaceholder,
  setShapeText,
  setShapeRunFormat,
  setShapeFill,
  setShapeGradientFill,
  setShapeShadow,
  setShapeStroke,
  setSlideNotes,
  setSlideTransition,
  validatePresentation,
  savePresentation,
  inches,
  pt,
} from '@office-kit/pptx';

// Define the palette once and reuse it everywhere.
const BRAND = '#2563EB';
const BRAND_DARK = '#1E3A8A';
const INK = '#1F2937';

const pres = createPresentation(); // 16:9 by default

// 1. Title slide.
const cover = addTitleSlide(pres, 'FY26 Business Review');
const subtitle = findSlidePlaceholder(cover, 'subTitle');
if (subtitle) setShapeText(subtitle, 'Strategy, results, and the road ahead');

// 2. Agenda — bulleted body (multi-line text + a bullet style).
const agenda = addContentSlide(pres, { title: 'Agenda' });
const body = findSlidePlaceholder(agenda, 'body');
if (body) setShapeText(body, 'Highlights\nFinancials\nRoadmap\nRisks', { bullets: 'bullet' });

// 3. Section divider.
addSectionHeaderSlide(pres, 'Highlights');

// 4. Stat cards on a free-form blank slide.
const cards = addBlankSlide(pres);
const heading = addSlideTextBox(cards, {
  x: inches(0.6),
  y: inches(0.4),
  w: inches(8),
  h: inches(0.9),
  text: 'A strong year',
});
setShapeRunFormat(heading, 0, 0, { bold: true, size: 32, color: INK });

const statCard = (x: ReturnType<typeof inches>, text: string) =>
  addSlideShape(cards, {
    preset: 'roundRect',
    x,
    y: inches(1.6),
    w: inches(3.6),
    h: inches(2.2),
    text,
    textAnchor: 'ctr',
  });

const cardA = statCard(inches(0.6), 'Revenue +38%');
setShapeGradientFill(cardA, {
  stops: [
    { offset: 0, color: BRAND },
    { offset: 1, color: BRAND_DARK },
  ],
  angleDeg: 90,
});
setShapeShadow(cardA, {
  color: '#000000',
  blurEmu: pt(8),
  offsetEmu: pt(3),
  angleDeg: 90,
  opacity: 0.35,
});

const cardB = statCard(inches(4.6), 'NPS 62');
setShapeFill(cardB, '#059669');
setShapeStroke(cardB, { color: '#065F46', widthEmu: pt(1.5) });

// 5. Financials table (banded header).
const fin = addContentSlide(pres, { title: 'Financials' });
addSlideTable(fin, {
  x: inches(0.6),
  y: inches(1.6),
  w: inches(8.2),
  h: inches(2.5),
  rows: [
    ['Metric', 'FY25', 'FY26', 'Δ'],
    ['Revenue', '$120M', '$166M', '+38%'],
    ['Gross margin', '64%', '67%', '+3pt'],
  ],
  firstRow: true,
  bandRow: true,
});

// 6. Chart-led slide — numbers as a chart, not sentences.
const chartSlide = addBlankSlide(pres);
addSlideTextBox(chartSlide, {
  x: inches(0.6),
  y: inches(0.3),
  w: inches(8),
  h: inches(0.7),
  text: 'Quarterly revenue',
});
addSlideChart(chartSlide, {
  x: inches(0.6),
  y: inches(1.1),
  w: inches(8.2),
  h: inches(4.2),
  spec: {
    kind: 'column',
    categories: ['Q1', 'Q2', 'Q3', 'Q4'],
    series: [
      { name: 'Revenue', values: [120, 138, 152, 166], color: BRAND },
      { name: 'Plan', values: [115, 130, 148, 160] },
    ],
    title: 'Revenue vs plan ($M)',
    valueAxisTitle: 'USD (millions)',
    dataLabels: {
      showValue: true,
      showCategory: false,
      showSeriesName: false,
      showPercent: false,
    },
  },
});

// 7. Image slide with a closing transition + speaker notes.
const imgSlide = addBlankSlide(pres);
addSlideImage(imgSlide, pngBytes, {
  x: inches(0.6),
  y: inches(1.1),
  w: inches(8.2),
  h: inches(4.2),
  fit: 'contain',
});
setSlideTransition(imgSlide, { effect: 'fade' });
setSlideNotes(imgSlide, 'Wrap in under two minutes, then open Q&A.');

// QA before shipping.
const errors = validatePresentation(pres).filter((i) => i.severity === 'error');
if (errors.length) throw new Error(`invalid deck: ${errors.map((e) => e.message).join('; ')}`);

const bytes = await savePresentation(pres);
// Node:    await fs.writeFile('review.pptx', bytes)
// Browser: new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' })
```

`pngBytes` is any `Uint8Array` of PNG/JPEG/GIF image data — the format is
detected from the bytes. Chart series `color` accepts the same hex forms as
shape colors (`#RGB` / `#RRGGBB` / bare `RRGGBB`) but, unlike shape colors, not
theme/scheme tokens — a chart series must resolve to a concrete sRGB value.
