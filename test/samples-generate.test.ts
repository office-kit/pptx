// Manual-inspection sample generator.
//
// Produces one `.pptx` per major feature area under `samples/out/`. Open them
// in PowerPoint, Keynote, Google Slides, or LibreOffice Impress to confirm
// the output renders as intended.
//
// Gated on `GENERATE_SAMPLES=1` so it doesn't churn artifacts in CI. Run:
//
//     pnpm samples
//
// The samples cover every public-API feature category. Each sample is small
// and self-explanatory by file name.

import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, it } from 'vitest';
import {
  addSlide,
  addSlideChart,
  addSlideComment,
  addSlideImage,
  addSlideLine,
  addSlideShape,
  addSlideTable,
  addSlideTextBox,
  bringShapeToFront,
  cm,
  duplicateSlide,
  findSlideLayout,
  findSlidePlaceholder,
  getSlides,
  inches,
  loadPresentation,
  moveSlide,
  pt,
  replaceTokensInPresentation,
  savePresentation,
  setCoreProperties,
  setParagraphAlignment,
  setParagraphBullet,
  setShapeAnimation,
  setShapeBullets,
  setShapeFill,
  setShapeFlip,
  setShapeGlow,
  setShapeGradientFill,
  setShapeHyperlink,
  setShapePatternFill,
  setShapeRotation,
  setShapeRunFormat,
  setShapeShadow,
  setShapeStroke,
  setShapeStrokeArrow,
  setShapeStrokeDash,
  setShapeText,
  setShapeTextColumns,
  setShapeTextDirection,
  setShapeTextFormat,
  setSlideBackground,
  setSlideNotes,
  setSlideTitle,
  setSlideTransition,
} from '../src/api/index.ts';
import { buildPng } from './lib/build-png.ts';

const ENABLED = process.env.GENERATE_SAMPLES === '1';

const OUT_DIR = fileURLToPath(new URL('../samples/out/', import.meta.url));
const fixturePath = fileURLToPath(new URL('./fixtures/minimal/blank.pptx', import.meta.url));

const loadBlank = async () => loadPresentation(await readFile(fixturePath));

// blank.pptx ships with layouts/masters but zero slides — every sample
// has to add at least one before it can do anything visible.
const freshSlide = (
  pres: Awaited<ReturnType<typeof loadPresentation>>,
  layoutName = 'Title and Content',
) => {
  const layout =
    findSlideLayout(pres, layoutName) ??
    findSlideLayout(pres, 'Blank') ??
    findSlideLayout(pres, 'Title Slide');
  if (!layout) throw new Error(`no layout found (looked for "${layoutName}")`);
  return addSlide(pres, { layout });
};

const writeSample = async (name: string, bytes: Uint8Array): Promise<void> => {
  await mkdir(OUT_DIR, { recursive: true });
  await writeFile(`${OUT_DIR}${name}`, bytes);
};

describe.skipIf(!ENABLED)('manual-inspection sample generation', () => {
  it('01 — blank deck with title', async () => {
    const pres = await loadBlank();
    const slide = freshSlide(pres);
    setSlideTitle(slide, 'pptx-kit sample 01 — blank deck');
    const bytes = await savePresentation(pres);
    await writeSample('01-title-only.pptx', bytes);
  });

  it('02 — multiple layouts (title, content, section header, blank)', async () => {
    const pres = await loadBlank();
    const blankSlide = freshSlide(pres);
    setSlideTitle(blankSlide, 'Layouts');

    const titleLayout = findSlideLayout(pres, 'Title Slide');
    if (titleLayout) {
      const s = addSlide(pres, { layout: titleLayout });
      const t = findSlidePlaceholder(s, 'ctrTitle') ?? findSlidePlaceholder(s, 'title');
      if (t) setShapeText(t, 'Title Slide layout');
      const sub = findSlidePlaceholder(s, 'subTitle');
      if (sub) setShapeText(sub, 'with subtitle placeholder');
    }
    const sectionLayout =
      findSlideLayout(pres, 'Section Header') ?? findSlideLayout(pres, 'Section Header Slide');
    if (sectionLayout) {
      const s = addSlide(pres, { layout: sectionLayout });
      const t = findSlidePlaceholder(s, 'title');
      if (t) setShapeText(t, 'Section Header');
    }
    const contentLayout = findSlideLayout(pres, 'Title and Content');
    if (contentLayout) {
      const s = addSlide(pres, { layout: contentLayout });
      const t = findSlidePlaceholder(s, 'title');
      if (t) setShapeText(t, 'Title and Content layout');
      const body = findSlidePlaceholder(s, 'body');
      if (body) setShapeText(body, 'A body placeholder for free-form content.');
    }

    const bytes = await savePresentation(pres);
    await writeSample('02-layouts.pptx', bytes);
  });

  it('03 — text formatting (font, size, color, bold, italic, underline)', async () => {
    const pres = await loadBlank();
    const slide = freshSlide(pres);
    setSlideTitle(slide, 'Text formatting');

    const box = addSlideTextBox(slide, {
      x: inches(0.7),
      y: inches(1.5),
      w: inches(9),
      h: inches(4),
      text: 'Default text\nBold red 24pt Calibri\nItalic underline 18pt\nLarge centered headline',
    });
    setShapeRunFormat(box, 1, 0, {
      bold: true,
      size: 24,
      color: '#C00000',
      font: 'Calibri',
    });
    setShapeRunFormat(box, 2, 0, {
      italic: true,
      underline: true,
      size: 18,
      color: '#1F4E79',
    });
    setShapeRunFormat(box, 3, 0, {
      bold: true,
      size: 32,
      color: '#2E75B6',
    });
    setParagraphAlignment(box, 3, 'ctr');

    const bytes = await savePresentation(pres);
    await writeSample('03-text-formatting.pptx', bytes);
  });

  it('04 — bullets and paragraph alignment', async () => {
    const pres = await loadBlank();
    const slide = freshSlide(pres);
    setSlideTitle(slide, 'Bullets & alignment');

    const box = addSlideTextBox(slide, {
      x: inches(0.7),
      y: inches(1.5),
      w: inches(9),
      h: inches(5),
      text: [
        'Bulleted list',
        'First item',
        'Second item',
        'Sub-item',
        'Third item',
        'Right-aligned line',
        'Centered line',
      ].join('\n'),
    });
    setShapeBullets(box, 'bullet');
    setParagraphBullet(box, 0, 'none');
    setParagraphBullet(box, 3, { char: '◦' });
    setParagraphAlignment(box, 5, 'r');
    setParagraphAlignment(box, 6, 'ctr');
    setShapeRunFormat(box, 0, 0, { bold: true, size: 20 });

    const bytes = await savePresentation(pres);
    await writeSample('04-bullets-alignment.pptx', bytes);
  });

  it('05 — preset shapes gallery', async () => {
    const pres = await loadBlank();
    const slide = freshSlide(pres);
    setSlideTitle(slide, 'Preset shapes');

    const presets = [
      'rect',
      'roundRect',
      'ellipse',
      'triangle',
      'diamond',
      'pentagon',
      'hexagon',
      'octagon',
      'star5',
      'star8',
      'rightArrow',
      'leftRightArrow',
    ];
    const cols = 4;
    const cellW = inches(2);
    const cellH = inches(1.3);
    const startX = inches(0.7);
    const startY = inches(1.5);
    presets.forEach((preset, i) => {
      const r = Math.floor(i / cols);
      const c = i % cols;
      addSlideShape(slide, {
        preset,
        x: (startX + c * cellW) as ReturnType<typeof inches>,
        y: (startY + r * cellH) as ReturnType<typeof inches>,
        w: inches(1.7),
        h: inches(1.1),
        text: preset,
      });
    });

    const bytes = await savePresentation(pres);
    await writeSample('05-preset-shapes.pptx', bytes);
  });

  it('06 — lines, connectors, arrows, dashes', async () => {
    const pres = await loadBlank();
    const slide = freshSlide(pres);
    setSlideTitle(slide, 'Lines & arrows');

    const plain = addSlideLine(slide, {
      from: { x: inches(1), y: inches(2) },
      to: { x: inches(6), y: inches(2) },
    });
    setShapeStroke(plain, { color: '#222222', widthEmu: pt(2) });

    const arrow = addSlideLine(slide, {
      from: { x: inches(1), y: inches(3) },
      to: { x: inches(6), y: inches(3) },
    });
    setShapeStroke(arrow, { color: '#C00000', widthEmu: pt(3) });
    setShapeStrokeArrow(arrow, 'head', { type: 'triangle' });

    const dashed = addSlideLine(slide, {
      from: { x: inches(1), y: inches(4) },
      to: { x: inches(6), y: inches(4) },
    });
    setShapeStroke(dashed, { color: '#1F4E79', widthEmu: pt(2) });
    setShapeStrokeDash(dashed, 'dash');

    const diag = addSlideLine(slide, {
      from: { x: inches(7), y: inches(2) },
      to: { x: inches(9), y: inches(5) },
    });
    setShapeStroke(diag, { color: '#2E75B6', widthEmu: pt(4) });
    setShapeStrokeArrow(diag, 'head', { type: 'oval' });
    setShapeStrokeArrow(diag, 'tail', { type: 'triangle' });

    const bytes = await savePresentation(pres);
    await writeSample('06-lines-arrows.pptx', bytes);
  });

  it('07 — fills: solid, gradient, pattern', async () => {
    const pres = await loadBlank();
    const slide = freshSlide(pres);
    setSlideTitle(slide, 'Fills');

    const solid = addSlideShape(slide, {
      preset: 'rect',
      x: inches(0.7),
      y: inches(1.8),
      w: inches(2.5),
      h: inches(2),
      text: 'Solid',
    });
    setShapeFill(solid, '#C00000');

    const grad = addSlideShape(slide, {
      preset: 'rect',
      x: inches(3.5),
      y: inches(1.8),
      w: inches(2.5),
      h: inches(2),
      text: 'Gradient',
    });
    setShapeGradientFill(grad, {
      stops: [
        { offset: 0, color: '#FFD966' },
        { offset: 1, color: '#C00000' },
      ],
      angleDeg: 45,
    });

    const pat = addSlideShape(slide, {
      preset: 'rect',
      x: inches(6.3),
      y: inches(1.8),
      w: inches(2.5),
      h: inches(2),
      text: 'Pattern',
    });
    setShapePatternFill(pat, {
      preset: 'pct50',
      foreground: '#1F4E79',
      background: '#FFFFFF',
    });

    const bytes = await savePresentation(pres);
    await writeSample('07-fills.pptx', bytes);
  });

  it('08 — shadow and glow effects', async () => {
    const pres = await loadBlank();
    const slide = freshSlide(pres);
    setSlideTitle(slide, 'Effects');

    const shadow = addSlideShape(slide, {
      preset: 'roundRect',
      x: inches(1),
      y: inches(2),
      w: inches(3),
      h: inches(2),
      text: 'Shadow',
    });
    setShapeFill(shadow, '#FFFFFF');
    setShapeShadow(shadow, {
      blurEmu: pt(8),
      offsetEmu: pt(4),
      angleDeg: 45,
      color: '#000000',
      opacity: 0.5,
    });

    const glow = addSlideShape(slide, {
      preset: 'ellipse',
      x: inches(5),
      y: inches(2),
      w: inches(3),
      h: inches(2),
      text: 'Glow',
    });
    setShapeFill(glow, '#FFFFFF');
    setShapeGlow(glow, { radiusEmu: pt(12), color: '#2E75B6' });

    const bytes = await savePresentation(pres);
    await writeSample('08-shadow-glow.pptx', bytes);
  });

  it('09 — embedded images (PNG)', async () => {
    const pres = await loadBlank();
    const slide = freshSlide(pres);
    setSlideTitle(slide, 'Embedded images');

    const red = buildPng(120, 80, [200, 50, 50]);
    const blue = buildPng(120, 80, [40, 90, 180]);
    const green = buildPng(120, 80, [70, 160, 80]);

    addSlideImage(slide, red, {
      x: inches(0.7),
      y: inches(1.8),
      w: inches(2.5),
      h: inches(2),
    });
    addSlideImage(slide, blue, {
      x: inches(3.5),
      y: inches(1.8),
      w: inches(2.5),
      h: inches(2),
    });
    addSlideImage(slide, green, {
      x: inches(6.3),
      y: inches(1.8),
      w: inches(2.5),
      h: inches(2),
    });

    const bytes = await savePresentation(pres);
    await writeSample('09-images.pptx', bytes);
  });

  it('10 — tables', async () => {
    const pres = await loadBlank();
    const slide = freshSlide(pres);
    setSlideTitle(slide, 'Tables');

    addSlideTable(slide, {
      x: inches(0.7),
      y: inches(1.8),
      w: inches(9),
      h: inches(3),
      rows: [
        ['Quarter', 'Revenue', 'Cost', 'Margin'],
        ['Q1', '$1.2M', '$0.8M', '33%'],
        ['Q2', '$1.8M', '$0.9M', '50%'],
        ['Q3', '$2.4M', '$1.3M', '46%'],
        ['Q4', '$3.0M', '$1.6M', '47%'],
      ],
      firstRow: true,
      bandRow: true,
    });

    const bytes = await savePresentation(pres);
    await writeSample('10-tables.pptx', bytes);
  });

  it('11 — charts (column, line, pie)', async () => {
    const pres = await loadBlank();
    const first = freshSlide(pres);
    setSlideTitle(first, 'Charts');

    addSlideChart(first, {
      x: inches(0.5),
      y: inches(1.5),
      w: inches(4.5),
      h: inches(3),
      spec: {
        kind: 'column',
        categories: ['Q1', 'Q2', 'Q3', 'Q4'],
        series: [
          { name: 'Revenue', values: [120, 180, 240, 300] },
          { name: 'Cost', values: [80, 90, 130, 160] },
        ],
        title: 'Column',
      },
    });
    addSlideChart(first, {
      x: inches(5.2),
      y: inches(1.5),
      w: inches(4.5),
      h: inches(3),
      spec: {
        kind: 'line',
        categories: ['Mon', 'Tue', 'Wed', 'Thu', 'Fri'],
        series: [{ name: 'Visits', values: [10, 22, 18, 30, 25] }],
        title: 'Line',
      },
    });

    const pieLayout = findSlideLayout(pres, 'Blank') ?? findSlideLayout(pres, 'Title and Content');
    if (pieLayout) {
      const s = addSlide(pres, { layout: pieLayout });
      addSlideChart(s, {
        x: inches(2),
        y: inches(1.5),
        w: inches(6),
        h: inches(4.5),
        spec: {
          kind: 'pie',
          categories: ['Web', 'Mobile', 'Desktop'],
          series: [{ name: 'Share', values: [55, 30, 15] }],
          title: 'Pie',
        },
      });
    }

    const bytes = await savePresentation(pres);
    await writeSample('11-charts.pptx', bytes);
  });

  it('12 — geometry: rotation, flip, translation', async () => {
    const pres = await loadBlank();
    const slide = freshSlide(pres);
    setSlideTitle(slide, 'Geometry');

    const a = addSlideShape(slide, {
      preset: 'rightArrow',
      x: inches(1),
      y: inches(2),
      w: inches(2),
      h: inches(1),
      text: 'rot 30',
    });
    setShapeFill(a, '#2E75B6');
    setShapeRotation(a, 30);

    const b = addSlideShape(slide, {
      preset: 'rightArrow',
      x: inches(4),
      y: inches(2),
      w: inches(2),
      h: inches(1),
      text: 'flipH',
    });
    setShapeFill(b, '#C00000');
    setShapeFlip(b, { horizontal: true });

    const c = addSlideShape(slide, {
      preset: 'rightArrow',
      x: inches(7),
      y: inches(2),
      w: inches(2),
      h: inches(1),
      text: 'flipV',
    });
    setShapeFill(c, '#548235');
    setShapeFlip(c, { vertical: true });

    const bytes = await savePresentation(pres);
    await writeSample('12-geometry.pptx', bytes);
  });

  it('13 — z-order overlap', async () => {
    const pres = await loadBlank();
    const slide = freshSlide(pres);
    setSlideTitle(slide, 'Z-order');

    const back = addSlideShape(slide, {
      preset: 'rect',
      x: inches(1),
      y: inches(2),
      w: inches(4),
      h: inches(3),
      text: 'back',
    });
    setShapeFill(back, '#BDD7EE');
    const front = addSlideShape(slide, {
      preset: 'rect',
      x: inches(2),
      y: inches(3),
      w: inches(4),
      h: inches(3),
      text: 'front (raised)',
    });
    setShapeFill(front, '#F4B084');
    bringShapeToFront(front);

    const bytes = await savePresentation(pres);
    await writeSample('13-zorder.pptx', bytes);
  });

  it('14 — slide background', async () => {
    const pres = await loadBlank();
    const slide = freshSlide(pres);
    setSlideTitle(slide, 'Background fill');
    setSlideBackground(slide, '#FFF2CC');

    const bytes = await savePresentation(pres);
    await writeSample('14-background.pptx', bytes);
  });

  it('15 — notes and comments', async () => {
    const pres = await loadBlank();
    const slide = freshSlide(pres);
    setSlideTitle(slide, 'Notes and comments');
    setSlideNotes(slide, 'Speaker note: open the comments pane to see the review thread.');
    addSlideComment(slide, {
      author: { name: 'Reviewer A', initials: 'RA' },
      text: 'Tighten the headline.',
      position: { x: cm(2), y: cm(2) },
    });
    addSlideComment(slide, {
      author: { name: 'Reviewer B', initials: 'RB' },
      text: 'Numbers look strong.',
      position: { x: cm(10), y: cm(6) },
    });

    const bytes = await savePresentation(pres);
    await writeSample('15-notes-comments.pptx', bytes);
  });

  it('16 — slide transitions', async () => {
    const pres = await loadBlank();
    const first = freshSlide(pres);
    setSlideTitle(first, 'Transitions — fade');
    setSlideTransition(first, { effect: 'fade', speed: 'med' });

    const s2 = freshSlide(pres);
    setSlideTitle(s2, 'Transitions — push');
    setSlideTransition(s2, { effect: 'push', speed: 'med' });

    const s3 = freshSlide(pres);
    setSlideTitle(s3, 'Transitions — wipe');
    setSlideTransition(s3, { effect: 'wipe', speed: 'med' });

    const bytes = await savePresentation(pres);
    await writeSample('16-transitions.pptx', bytes);
  });

  it('17 — animations (entrance / exit)', async () => {
    const pres = await loadBlank();
    const slide = freshSlide(pres);
    setSlideTitle(slide, 'Animations');

    const a = addSlideShape(slide, {
      preset: 'roundRect',
      x: inches(1),
      y: inches(2),
      w: inches(3),
      h: inches(1.5),
      text: 'fadeIn',
    });
    setShapeFill(a, '#2E75B6');
    setShapeAnimation(a, { effect: 'fadeIn', durationMs: 700 });

    const b = addSlideShape(slide, {
      preset: 'roundRect',
      x: inches(5),
      y: inches(2),
      w: inches(3),
      h: inches(1.5),
      text: 'appear',
    });
    setShapeFill(b, '#548235');
    setShapeAnimation(b, { effect: 'appear' });

    const c = addSlideShape(slide, {
      preset: 'roundRect',
      x: inches(3),
      y: inches(4),
      w: inches(3),
      h: inches(1.5),
      text: 'fadeOut',
    });
    setShapeFill(c, '#C00000');
    setShapeAnimation(c, { effect: 'fadeOut', durationMs: 700 });

    const bytes = await savePresentation(pres);
    await writeSample('17-animations.pptx', bytes);
  });

  it('18 — hyperlinks and click actions', async () => {
    const pres = await loadBlank();
    const slide = freshSlide(pres);
    setSlideTitle(slide, 'Hyperlinks');

    const link = addSlideTextBox(slide, {
      x: inches(1),
      y: inches(2),
      w: inches(6),
      h: inches(0.6),
      text: 'Click to open the pptx-kit repository',
    });
    setShapeHyperlink(link, 'https://github.com/baseballyama/pptx-kit');
    setShapeTextFormat(link, { color: '#0563C1', underline: true });

    const bytes = await savePresentation(pres);
    await writeSample('18-hyperlinks.pptx', bytes);
  });

  it('19 — token-based template fill', async () => {
    const pres = await loadBlank();
    const slide = freshSlide(pres);
    setSlideTitle(slide, '{{title}} — {{date}}');
    addSlideTextBox(slide, {
      x: inches(1),
      y: inches(2),
      w: inches(8),
      h: inches(1),
      text: 'Prepared by {{author}} for {{audience}}.',
    });
    replaceTokensInPresentation(pres, {
      title: 'Q3 review',
      date: '2026-05-16',
      author: 'Yamashita',
      audience: 'Board',
    });

    const bytes = await savePresentation(pres);
    await writeSample('19-token-fill.pptx', bytes);
  });

  it('20 — core properties (Author / Title / Subject)', async () => {
    const pres = await loadBlank();
    setCoreProperties(pres, {
      title: 'pptx-kit sample',
      creator: 'pptx-kit',
      subject: 'manual inspection deck',
      keywords: 'pptx-kit, sample, ooxml',
      description: 'Generated by pptx-kit samples-generate script.',
    });
    const slide = freshSlide(pres);
    // '>' rather than '\u25B8': the bundled Carlito/Liberation faces carry no
    // glyph for the triangle, so it rasterizes as tofu and shifts the rest of
    // the line — a bundled-font coverage limit, not a metadata feature.
    setSlideTitle(slide, 'Check File > Info for metadata');
    const bytes = await savePresentation(pres);
    await writeSample('20-core-properties.pptx', bytes);
  });

  it('21 — combined showcase (the everything deck)', async () => {
    const pres = await loadBlank();

    // Slide 1: title.
    const titleSlide = freshSlide(pres);
    setSlideTitle(titleSlide, 'pptx-kit — feature showcase');
    const intro = addSlideTextBox(titleSlide, {
      x: inches(0.7),
      y: inches(2.4),
      w: inches(9),
      h: inches(1),
      text: 'Slides 2–6 cover shapes, charts, tables, notes, transitions.',
    });
    setShapeRunFormat(intro, 0, 0, { size: 16, color: '#404040' });

    // Slide 2: shapes + effects.
    const s = freshSlide(pres);
    setSlideTitle(s, 'Shapes');
    const star = addSlideShape(s, {
      preset: 'star5',
      x: inches(1),
      y: inches(2),
      w: inches(2),
      h: inches(2),
      text: '★',
    });
    setShapeFill(star, '#FFD966');
    setShapeShadow(star, {
      blurEmu: pt(6),
      offsetEmu: pt(3),
      angleDeg: 45,
      color: '#000000',
      opacity: 0.4,
    });
    const callout = addSlideShape(s, {
      preset: 'roundRect',
      x: inches(4),
      y: inches(2),
      w: inches(5),
      h: inches(1.5),
      text: 'Rounded callout',
    });
    setShapeGradientFill(callout, {
      stops: [
        { offset: 0, color: '#BDD7EE' },
        { offset: 1, color: '#2E75B6' },
      ],
      angleDeg: 0,
    });

    // Slide 3: chart.
    const s3 = freshSlide(pres);
    setSlideTitle(s3, 'Chart');
    addSlideChart(s3, {
      x: inches(1),
      y: inches(1.5),
      w: inches(8),
      h: inches(4.5),
      spec: {
        kind: 'column',
        categories: ['Mon', 'Tue', 'Wed', 'Thu', 'Fri'],
        series: [
          { name: 'A', values: [4, 8, 5, 10, 7] },
          { name: 'B', values: [2, 5, 9, 3, 6] },
        ],
        title: 'Weekly',
      },
    });

    // Slide 4: table.
    const s4 = freshSlide(pres);
    setSlideTitle(s4, 'Table');
    addSlideTable(s4, {
      x: inches(0.7),
      y: inches(1.5),
      w: inches(9),
      h: inches(3),
      rows: [
        ['Feature', 'Status'],
        ['Read PPTX', 'OK'],
        ['Edit PPTX', 'OK'],
        ['Author from scratch', 'OK'],
        ['Charts', 'OK'],
      ],
      firstRow: true,
      bandRow: true,
    });

    // Slide 5: image.
    const s5 = freshSlide(pres);
    setSlideTitle(s5, 'Image');
    addSlideImage(s5, buildPng(160, 100, [40, 90, 180]), {
      x: inches(3),
      y: inches(2),
      w: inches(4),
      h: inches(2.5),
    });
    setSlideTransition(s5, { effect: 'fade', speed: 'med' });

    // Slide 6: notes + comment.
    const s6 = freshSlide(pres);
    setSlideTitle(s6, 'Notes & comments');
    setSlideNotes(s6, 'This slide demonstrates speaker notes and a review comment.');
    addSlideComment(s6, {
      author: { name: 'pptx-kit', initials: 'PK' },
      text: 'Generated automatically.',
      position: { x: cm(2), y: cm(2) },
    });

    // Duplicate the title slide at the end to verify duplicate + reorder.
    const dup = duplicateSlide(pres, titleSlide);
    moveSlide(pres, dup, getSlides(pres).length - 1);
    const dupTitle = findSlidePlaceholder(dup, 'title');
    if (dupTitle) setShapeText(dupTitle, 'Thank you');

    const bytes = await savePresentation(pres);
    await writeSample('21-showcase.pptx', bytes);
  });

  it('22 — vertical text (vert + vert270)', async () => {
    const pres = await loadBlank();
    const slide = freshSlide(pres);
    setSlideTitle(slide, 'Vertical text');

    // vert: rotated 90° clockwise — reads top-to-bottom, columns right-to-left.
    const vert = addSlideTextBox(slide, {
      x: inches(1),
      y: inches(1.5),
      w: inches(2),
      h: inches(4),
      text: 'Vertical text rotated ninety degrees clockwise, wrapping across several columns.',
    });
    setShapeTextDirection(vert, 'vert');
    setShapeRunFormat(vert, 0, 0, { size: 18 });

    // vert270: rotated 270° — reads bottom-to-top, columns left-to-right.
    const vert270 = addSlideTextBox(slide, {
      x: inches(5),
      y: inches(1.5),
      w: inches(2),
      h: inches(4),
      text: 'Vertical text rotated two hundred seventy degrees, the opposite reading direction.',
    });
    setShapeTextDirection(vert270, 'vert270');
    setShapeRunFormat(vert270, 0, 0, { size: 18 });

    const bytes = await savePresentation(pres);
    await writeSample('22-vertical-text.pptx', bytes);
  });

  it('23 — multi-column text (2 and 3 columns)', async () => {
    const pres = await loadBlank();
    const slide = freshSlide(pres);
    setSlideTitle(slide, 'Multi-column text');

    const para =
      'PowerPoint fills multi-column text bodies sequentially: the first column ' +
      'is filled down to the box height, then the overflow continues in the next ' +
      'column, and so on across the body. ';

    const two = addSlideTextBox(slide, {
      x: inches(0.7),
      y: inches(1.5),
      w: inches(8.5),
      h: inches(1.8),
      text: para.repeat(3),
    });
    setShapeTextColumns(two, { count: 2, gapEmu: 228600 }); // 0.25in gap
    setShapeRunFormat(two, 0, 0, { size: 14 });

    const three = addSlideTextBox(slide, {
      x: inches(0.7),
      y: inches(3.7),
      w: inches(8.5),
      h: inches(2.5),
      text: para.repeat(5),
    });
    setShapeTextColumns(three, { count: 3, gapEmu: 182880 }); // 0.2in gap
    setShapeRunFormat(three, 0, 0, { size: 14 });

    const bytes = await savePresentation(pres);
    await writeSample('23-columns.pptx', bytes);
  });
});
