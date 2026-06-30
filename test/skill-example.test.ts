// Keeps the agent skill's worked example honest: builds the same business deck
// the skill documents (skill/examples/business-deck.md), then asserts it is
// structurally valid, schema-valid, and round-trips. If this fails, the skill's
// example is wrong and must be updated alongside it.

import { describe, expect, it } from 'vitest';
import {
  _internalPackageOf,
  addBlankSlide,
  addContentSlide,
  addSectionHeaderSlide,
  addSlideShape,
  addSlideChart,
  addSlideImage,
  addSlideTable,
  addSlideTextBox,
  addTitleSlide,
  createPresentation,
  findSlidePlaceholder,
  getPresentationText,
  getSlides,
  inches,
  loadPresentation,
  pt,
  savePresentation,
  setShapeFill,
  setShapeGradientFill,
  setShapeRunFormat,
  setShapeShadow,
  setShapeStroke,
  setShapeText,
  setSlideNotes,
  setSlideTransition,
  validatePresentation,
  type PresentationData,
} from '../src/api/index.ts';
import { buildPng } from './lib/build-png.ts';
import {
  expectSchemaValid,
  isSchemaValidationAvailable,
  type SchemaKind,
} from './lib/expect-schema-valid.ts';

const skipIfNoXmllint = isSchemaValidationAvailable() ? it : it.skip;

// A consistent two-color palette, defined once and reused — exactly what the
// design rules in the skill ask for.
const BRAND = '#2563EB';
const BRAND_DARK = '#1E3A8A';
const INK = '#1F2937';

// Builds the deck documented in skill/examples/business-deck.md.
const buildBusinessDeck = async (): Promise<PresentationData> => {
  const pres = createPresentation();

  // 1. Title slide.
  const cover = addTitleSlide(pres, 'FY26 Business Review');
  const subtitle = findSlidePlaceholder(cover, 'subTitle');
  if (subtitle) setShapeText(subtitle, 'Strategy, results, and the road ahead');

  // 2. Agenda — bulleted body.
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

  // 5. Financials table.
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

  // 6. Chart-led slide.
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

  // 7. Image slide + closing notes/transition.
  const imgSlide = addBlankSlide(pres);
  addSlideImage(imgSlide, buildPng(640, 360, [37, 99, 235]), {
    x: inches(0.6),
    y: inches(1.1),
    w: inches(8.2),
    h: inches(4.2),
    fit: 'contain',
  });
  setSlideTransition(imgSlide, { effect: 'fade' });
  setSlideNotes(imgSlide, 'Wrap in under two minutes, then open Q&A.');

  return pres;
};

const kindFor = (name: string): SchemaKind | null => {
  if (/\/(slides|notesSlides)\/[^/]*\.xml$/.test(name)) return 'pml';
  if (/\/charts\/chart\d+\.xml$/.test(name)) return 'chart';
  return null;
};

describe('skill worked example: business deck', () => {
  it('builds, validates structurally, and round-trips', async () => {
    const pres = await buildBusinessDeck();
    expect(getSlides(pres).length).toBe(7);

    const errors = validatePresentation(pres).filter((i) => i.severity === 'error');
    expect(errors).toEqual([]);

    const bytes = await savePresentation(pres);
    const reloaded = await loadPresentation(bytes);
    expect(getSlides(reloaded).length).toBe(7);

    // No leftover placeholder tokens.
    expect(getPresentationText(reloaded)).not.toMatch(/\{\{|lorem|TODO/i);
  });

  skipIfNoXmllint('every authored part is schema-valid', async () => {
    const pres = await buildBusinessDeck();
    const bytes = await savePresentation(pres);
    const pkg = _internalPackageOf(await loadPresentation(bytes));
    const decode = (b: Uint8Array): string => new TextDecoder().decode(b);
    for (const part of pkg.parts) {
      const kind = kindFor(part.name);
      if (kind) expectSchemaValid(decode(part.data), kind);
    }
  });
});
