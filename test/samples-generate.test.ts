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
  createPresentation,
  duplicateSlide,
  emu,
  findSlideLayout,
  findSlideLayoutByType,
  findSlidePlaceholder,
  getSlides,
  groupShapes,
  inches,
  loadPresentation,
  mergeTableCells,
  moveSlide,
  pt,
  replaceTokensInPresentation,
  savePresentation,
  setCoreProperties,
  setParagraphAlignment,
  setParagraphBullet,
  setParagraphLevel,
  setParagraphSpacing,
  setPresentationFonts,
  setPresentationTheme,
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
  setShapeStrokeCap,
  setShapeStrokeDash,
  setShapeStrokeJoin,
  setShapeText,
  setShapeTextAutoFit,
  setShapeTextColumns,
  setShapeTextDirection,
  setShapeTextFormat,
  setSlideBackground,
  setSlideNotes,
  setSlideTitle,
  setSlideTransition,
  type Emu,
  type SlideData,
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

  it('24 — gradient fills (multi-stop linear at multiple angles)', async () => {
    const pres = await loadBlank();
    const slide = freshSlide(pres);
    setSlideTitle(slide, 'Gradient fills');

    // Each cell is a distinct linear gradient: stop count and axis angle vary
    // independently so SSIM separates interpolation fidelity from axis angle.
    // (radial/path gradients are omitted: setShapeGradientFill only emits
    // <a:lin> today, so a `path` case would render identically to linear.)
    const cases: Array<{
      label: string;
      stops: Array<{ offset: number; color: string }>;
      angleDeg: number;
    }> = [
      {
        label: '2 stops · 0°',
        stops: [
          { offset: 0, color: '#FFD966' },
          { offset: 1, color: '#C00000' },
        ],
        angleDeg: 0,
      },
      {
        label: '3 stops · 90°',
        stops: [
          { offset: 0, color: '#00B0F0' },
          { offset: 0.5, color: '#FFFFFF' },
          { offset: 1, color: '#C00000' },
        ],
        angleDeg: 90,
      },
      {
        label: '4 stops · 45°',
        stops: [
          { offset: 0, color: '#C00000' },
          { offset: 0.33, color: '#FFC000' },
          { offset: 0.66, color: '#70AD47' },
          { offset: 1, color: '#2E75B6' },
        ],
        angleDeg: 45,
      },
      {
        label: '2 stops · 135°',
        stops: [
          { offset: 0, color: '#2E75B6' },
          { offset: 1, color: '#FFD966' },
        ],
        angleDeg: 135,
      },
      {
        label: '3 stops · 270°',
        stops: [
          { offset: 0, color: '#548235' },
          { offset: 0.5, color: '#FFFF00' },
          { offset: 1, color: '#C00000' },
        ],
        angleDeg: 270,
      },
      {
        label: '5 stops · 90°',
        stops: [
          { offset: 0, color: '#FF0000' },
          { offset: 0.25, color: '#FF9900' },
          { offset: 0.5, color: '#FFFF00' },
          { offset: 0.75, color: '#00B050' },
          { offset: 1, color: '#0070C0' },
        ],
        angleDeg: 90,
      },
    ];

    const cols = 3;
    const cellW = inches(3);
    const cellH = inches(2.4);
    const startX = inches(0.5);
    const startY = inches(1.6);
    cases.forEach((cfg, i) => {
      const r = Math.floor(i / cols);
      const c = i % cols;
      const rect = addSlideShape(slide, {
        preset: 'rect',
        x: (startX + c * cellW) as ReturnType<typeof inches>,
        y: (startY + r * cellH) as ReturnType<typeof inches>,
        w: inches(2.7),
        h: inches(2),
        text: cfg.label,
      });
      setShapeGradientFill(rect, { stops: cfg.stops, angleDeg: cfg.angleDeg });
    });

    const bytes = await savePresentation(pres);
    await writeSample('24-gradient-fills.pptx', bytes);
  });

  it('25 — preset pattern fills gallery', async () => {
    const pres = await loadBlank();
    const slide = freshSlide(pres);
    setSlideTitle(slide, 'Pattern fills');

    // A spread of ST_PresetPatternVal tokens (§20.1.10.49). High-contrast
    // navy-on-white so each tiling reads distinctly per cell.
    const presets = [
      'pct5',
      'pct50',
      'ltHorz',
      'ltVert',
      'horz',
      'vert',
      'ltUpDiag',
      'ltDnDiag',
      'dkUpDiag',
      'dkDnDiag',
      'cross',
      'diagCross',
      'smGrid',
      'lgGrid',
      'wave',
      'dotGrid',
    ] as const;
    const cols = 4;
    const cellW = inches(2.2);
    const cellH = inches(1.25);
    const startX = inches(0.6);
    const startY = inches(1.5);
    presets.forEach((preset, i) => {
      const r = Math.floor(i / cols);
      const c = i % cols;
      const shape = addSlideShape(slide, {
        preset: 'rect',
        x: (startX + c * cellW) as ReturnType<typeof inches>,
        y: (startY + r * cellH) as ReturnType<typeof inches>,
        w: inches(2),
        h: inches(1.05),
        text: preset,
      });
      setShapePatternFill(shape, {
        preset,
        foreground: '#1F4E79',
        background: '#FFFFFF',
      });
    });

    const bytes = await savePresentation(pres);
    await writeSample('25-pattern-fills.pptx', bytes);
  });

  it('26 — stroke styles: dashes, widths, caps, joins, arrowheads', async () => {
    const pres = await loadBlank();

    // Slide 1: dash gallery + width ramp.
    const s1 = freshSlide(pres);
    setSlideTitle(s1, 'Stroke dashes & widths');

    const dashes = [
      'solid',
      'dot',
      'dash',
      'lgDash',
      'dashDot',
      'lgDashDot',
      'sysDash',
      'sysDashDot',
    ] as const;
    dashes.forEach((dash, i) => {
      const y = (inches(1.3) + i * inches(0.45)) as ReturnType<typeof inches>;
      const label = addSlideTextBox(s1, {
        x: inches(0.7),
        y: (y - inches(0.12)) as ReturnType<typeof inches>,
        w: inches(1.8),
        h: inches(0.4),
        text: dash,
      });
      setShapeRunFormat(label, 0, 0, { size: 12 });
      const line = addSlideLine(s1, {
        from: { x: inches(2.6), y },
        to: { x: inches(9.3), y },
      });
      setShapeStroke(line, { color: '#1F4E79', widthEmu: pt(2.5) });
      setShapeStrokeDash(line, dash);
    });

    const widths = [1, 3, 6, 10];
    widths.forEach((w, i) => {
      const y = (inches(5.4) + i * inches(0.45)) as ReturnType<typeof inches>;
      const label = addSlideTextBox(s1, {
        x: inches(0.7),
        y: (y - inches(0.12)) as ReturnType<typeof inches>,
        w: inches(1.8),
        h: inches(0.4),
        text: `${w}pt`,
      });
      setShapeRunFormat(label, 0, 0, { size: 12 });
      const line = addSlideLine(s1, {
        from: { x: inches(2.6), y },
        to: { x: inches(9.3), y },
      });
      setShapeStroke(line, { color: '#C00000', widthEmu: pt(w) });
    });

    // Slide 2: caps, joins, arrowheads.
    const s2 = freshSlide(pres);
    setSlideTitle(s2, 'Caps, joins & arrowheads');

    // Thick lines make the cap shape visible at each endpoint.
    const caps = [
      { cap: 'flat', label: 'cap flat' },
      { cap: 'rnd', label: 'cap rnd' },
      { cap: 'sq', label: 'cap sq' },
    ] as const;
    caps.forEach(({ cap, label }, i) => {
      const y = (inches(1.4) + i * inches(0.6)) as ReturnType<typeof inches>;
      const t = addSlideTextBox(s2, {
        x: inches(0.7),
        y: (y - inches(0.15)) as ReturnType<typeof inches>,
        w: inches(1.6),
        h: inches(0.4),
        text: label,
      });
      setShapeRunFormat(t, 0, 0, { size: 12 });
      const line = addSlideLine(s2, {
        from: { x: inches(2.5), y },
        to: { x: inches(5), y },
      });
      setShapeStroke(line, { color: '#404040', widthEmu: pt(14) });
      setShapeStrokeCap(line, cap);
    });

    // Thick-bordered rectangles make the corner join shape visible.
    const joins = [
      { join: 'miter', label: 'miter' },
      { join: 'round', label: 'round' },
      { join: 'bevel', label: 'bevel' },
    ] as const;
    joins.forEach(({ join, label }, i) => {
      const x = (inches(6) + i * inches(1.3)) as ReturnType<typeof inches>;
      const rect = addSlideShape(s2, {
        preset: 'rect',
        x,
        y: inches(1.4),
        w: inches(1),
        h: inches(1),
        text: label,
      });
      setShapeFill(rect, '#FFF2CC');
      setShapeStroke(rect, { color: '#2E75B6', widthEmu: pt(10) });
      setShapeStrokeJoin(rect, join);
      setShapeRunFormat(rect, 0, 0, { size: 11 });
    });

    // Arrowheads on both ends, at large width/length so each type is distinct.
    const arrows = [
      { head: 'triangle', tail: 'triangle' },
      { head: 'stealth', tail: 'oval' },
      { head: 'diamond', tail: 'arrow' },
    ] as const;
    arrows.forEach(({ head, tail }, i) => {
      const y = (inches(3.6) + i * inches(0.7)) as ReturnType<typeof inches>;
      const t = addSlideTextBox(s2, {
        x: inches(0.7),
        y: (y - inches(0.15)) as ReturnType<typeof inches>,
        w: inches(2.3),
        h: inches(0.4),
        text: `${tail} - ${head}`,
      });
      setShapeRunFormat(t, 0, 0, { size: 12 });
      const line = addSlideLine(s2, {
        from: { x: inches(3.2), y },
        to: { x: inches(9.3), y },
      });
      setShapeStroke(line, { color: '#548235', widthEmu: pt(3) });
      setShapeStrokeArrow(line, 'head', { type: head, width: 'lg', length: 'lg' });
      setShapeStrokeArrow(line, 'tail', { type: tail, width: 'lg', length: 'lg' });
    });

    const bytes = await savePresentation(pres);
    await writeSample('26-stroke-styles.pptx', bytes);
  });

  it('27 — rotation combined with flip', async () => {
    const pres = await loadBlank();
    const slide = freshSlide(pres);
    setSlideTitle(slide, 'Rotation + flip');

    // rightArrow is asymmetric on both axes, so each rotation/flip combo
    // points in a visibly distinct direction. PowerPoint flips about the
    // shape centre first, then rotates — the combined cells below differ
    // from the isolated transforms in deck 12.
    const combos: {
      label: string;
      rotation: number;
      flip?: { horizontal?: boolean; vertical?: boolean };
      color: string;
    }[] = [
      { label: 'base', rotation: 0, color: '#2E75B6' },
      { label: 'rot 45', rotation: 45, color: '#2E75B6' },
      { label: 'flipH only', rotation: 0, flip: { horizontal: true }, color: '#C00000' },
      { label: 'rot 45 + flipH', rotation: 45, flip: { horizontal: true }, color: '#C00000' },
      { label: 'rot 90 + flipV', rotation: 90, flip: { vertical: true }, color: '#548235' },
      { label: 'rot 135 + flipH', rotation: 135, flip: { horizontal: true }, color: '#C00000' },
      { label: 'rot 200 + flipV', rotation: 200, flip: { vertical: true }, color: '#548235' },
      {
        label: 'rot 225 + both',
        rotation: 225,
        flip: { horizontal: true, vertical: true },
        color: '#7030A0',
      },
      {
        label: 'rot 315 + both',
        rotation: 315,
        flip: { horizontal: true, vertical: true },
        color: '#7030A0',
      },
    ];

    const cols = 3;
    const cellW = inches(3);
    const cellH = inches(1.6);
    const startX = inches(0.7);
    const startY = inches(1.6);
    combos.forEach((combo, i) => {
      const r = Math.floor(i / cols);
      const c = i % cols;
      const shape = addSlideShape(slide, {
        preset: 'rightArrow',
        x: (startX + c * cellW) as ReturnType<typeof inches>,
        y: (startY + r * cellH) as ReturnType<typeof inches>,
        w: inches(2.4),
        h: inches(1.1),
        text: combo.label,
      });
      setShapeFill(shape, combo.color);
      if (combo.flip) setShapeFlip(shape, combo.flip);
      setShapeRotation(shape, combo.rotation);
    });

    const bytes = await savePresentation(pres);
    await writeSample('27-rotation-flip.pptx', bytes);
  });

  it('28 — multilevel bullets and autonumbered indent levels', async () => {
    const pres = await loadBlank();

    // Slide 1: a 3-deep nested bulleted list. The `lvl` attribute (via
    // setParagraphLevel) pulls each deeper level in using the deck's
    // defaultTextStyle indents; per-level bullet chars keep the nesting
    // legible (•, ◦, – are all covered by the bundled faces).
    const s1 = freshSlide(pres);
    setSlideTitle(s1, 'Nested bullets');
    const nested = addSlideTextBox(s1, {
      x: inches(0.7),
      y: inches(1.5),
      w: inches(9),
      h: inches(5),
      text: [
        'Product roadmap',
        'Q1 — Foundation',
        'Editor core',
        'Undo / redo',
        'Selection model',
        'File I/O',
        'Q2 — Polish',
        'Theming',
        'Dark mode',
      ].join('\n'),
    });
    setShapeBullets(nested, 'bullet');
    const nestedLevels = [0, 0, 1, 2, 2, 1, 0, 1, 2];
    nestedLevels.forEach((lvl, i) => {
      setParagraphLevel(nested, i, lvl);
      if (lvl === 1) setParagraphBullet(nested, i, { char: '◦' });
      else if (lvl === 2) setParagraphBullet(nested, i, { char: '–' });
      if (lvl === 0) setShapeRunFormat(nested, i, 0, { bold: true });
    });

    // Slide 2: an autonumbered list whose scheme changes per level —
    // 1. / a) / i) — so the numbering itself signals the indent depth.
    const s2 = freshSlide(pres);
    setSlideTitle(s2, 'Autonumbered steps');
    const numbered = addSlideTextBox(s2, {
      x: inches(0.7),
      y: inches(1.5),
      w: inches(9),
      h: inches(5),
      text: [
        'Install the CLI',
        'Verify the Node version',
        'Run npm install',
        'Configure the project',
        'Create the config file',
        'Set the API key',
        'Set the region',
        'Build and deploy',
      ].join('\n'),
    });
    setShapeBullets(numbered, { autoNum: 'arabicPeriod' });
    const numberedLevels = [0, 1, 1, 0, 1, 2, 2, 0];
    numberedLevels.forEach((lvl, i) => {
      setParagraphLevel(numbered, i, lvl);
      if (lvl === 1) setParagraphBullet(numbered, i, { autoNum: 'alphaLcParenR' });
      else if (lvl === 2) setParagraphBullet(numbered, i, { autoNum: 'romanLcParenR' });
    });

    const bytes = await savePresentation(pres);
    await writeSample('28-multilevel-bullets.pptx', bytes);
  });

  it('29 — text autofit in multi-column bodies (none / shrink / grow)', async () => {
    const pres = await loadBlank();
    const slide = freshSlide(pres);
    setSlideTitle(slide, 'Text autofit');

    // Identical 2-column overflowing content in three equally sized boxes;
    // the only variable is the <a:bodyPr> autofit child, so any rendered
    // difference is the autofit behaviour itself. Thin borders mark the box
    // bounds so overflow (none) vs. fit (normal) vs. grow (shape) is visible.
    const overflow =
      'This text body is deliberately taller than its box once it wraps ' +
      'across two columns, so the autofit mode decides what the reader sees. ';

    const boxW = inches(2.9);
    const boxH = inches(2.4);
    const top = inches(1.8);

    const none = addSlideTextBox(slide, {
      x: inches(0.4),
      y: top,
      w: boxW,
      h: boxH,
      text: `none — overflow\n${overflow.repeat(3)}`,
    });
    setShapeTextColumns(none, { count: 2, gapEmu: 137160 }); // 0.15in gap
    setShapeTextAutoFit(none, 'none');
    setShapeStroke(none, { color: '#C00000', widthEmu: pt(1) });
    setShapeRunFormat(none, 0, 0, { bold: true, size: 14, color: '#C00000' });

    const normal = addSlideTextBox(slide, {
      x: inches(3.55),
      y: top,
      w: boxW,
      h: boxH,
      text: `normal — shrink to fit\n${overflow.repeat(3)}`,
    });
    setShapeTextColumns(normal, { count: 2, gapEmu: 137160 });
    setShapeTextAutoFit(normal, 'normal');
    setShapeStroke(normal, { color: '#1F4E79', widthEmu: pt(1) });
    setShapeRunFormat(normal, 0, 0, { bold: true, size: 14, color: '#1F4E79' });

    const shape = addSlideTextBox(slide, {
      x: inches(6.7),
      y: top,
      w: boxW,
      h: boxH,
      text: `shape — grow box\n${overflow.repeat(3)}`,
    });
    setShapeTextColumns(shape, { count: 2, gapEmu: 137160 });
    setShapeTextAutoFit(shape, 'shape');
    setShapeStroke(shape, { color: '#548235', widthEmu: pt(1) });
    setShapeRunFormat(shape, 0, 0, { bold: true, size: 14, color: '#548235' });

    const bytes = await savePresentation(pres);
    await writeSample('29-text-columns-autofit.pptx', bytes);
  });

  it('30 — pie & doughnut with data labels and legend', async () => {
    const pres = await loadBlank();

    // Slide 1: pie with per-slice colors, a right-side legend, and
    // category + percent data labels.
    const pieSlide = freshSlide(pres);
    setSlideTitle(pieSlide, 'Pie — legend + percent labels');
    addSlideChart(pieSlide, {
      x: inches(1.5),
      y: inches(1.5),
      w: inches(7),
      h: inches(5),
      spec: {
        kind: 'pie',
        categories: ['Web', 'Mobile', 'Desktop', 'Other'],
        series: [
          {
            name: 'Traffic share',
            values: [48, 27, 18, 7],
            pointColors: ['#2E75B6', '#C00000', '#548235', '#FFC000'],
          },
        ],
        title: 'Traffic by platform',
        legend: { position: 'r' },
        dataLabels: {
          showValue: false,
          showCategory: true,
          showSeriesName: false,
          showPercent: true,
          separator: ' — ',
          position: 'outEnd',
        },
      },
    });

    // Slide 2: doughnut with a 60% hole, first slice rotated to 90°, a
    // bottom legend, and value data labels.
    const doughSlide = freshSlide(pres);
    setSlideTitle(doughSlide, 'Doughnut — hole + value labels');
    addSlideChart(doughSlide, {
      x: inches(2),
      y: inches(1.5),
      w: inches(6),
      h: inches(5),
      spec: {
        kind: 'doughnut',
        categories: ['Engineering', 'Sales', 'Support', 'Ops'],
        series: [
          {
            name: 'Headcount',
            values: [42, 30, 16, 12],
            pointColors: ['#4472C4', '#ED7D31', '#A5A5A5', '#70AD47'],
          },
        ],
        title: 'Headcount by team',
        holeSizePct: 60,
        firstSliceAngleDeg: 90,
        legend: { position: 'b' },
        dataLabels: {
          showValue: true,
          showCategory: false,
          showSeriesName: false,
          showPercent: false,
          position: 'ctr',
        },
      },
    });

    const bytes = await savePresentation(pres);
    await writeSample('30-pie-doughnut-chart.pptx', bytes);
  });

  it('31 — area charts with axis titles (standard + stacked)', async () => {
    const pres = await loadBlank();

    // Slide 1: standard (overlapping) area chart with both axis titles.
    const first = freshSlide(pres);
    setSlideTitle(first, 'Area chart — axis titles');
    addSlideChart(first, {
      x: inches(0.7),
      y: inches(1.5),
      w: inches(8.6),
      h: inches(4.5),
      spec: {
        kind: 'area',
        categories: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'],
        series: [
          { name: 'Organic', values: [120, 150, 170, 210, 260, 300], color: '#2E75B6' },
          { name: 'Paid', values: [60, 90, 110, 130, 160, 180], color: '#C00000' },
        ],
        title: 'Monthly traffic',
        categoryAxisTitle: 'Month',
        // -90° so the value-axis title reads bottom-to-top alongside the axis,
        // the orientation PowerPoint emits by default.
        valueAxisTitle: 'Sessions (thousands)',
        valueAxisTitleRotationDeg: -90,
        valueAxisMajorGridlines: true,
      },
    });

    // Slide 2: stacked area — the same series summed, to show grouping differs
    // from the overlapping default on slide 1.
    const second = freshSlide(pres);
    setSlideTitle(second, 'Stacked area — axis titles');
    addSlideChart(second, {
      x: inches(0.7),
      y: inches(1.5),
      w: inches(8.6),
      h: inches(4.5),
      spec: {
        kind: 'area',
        grouping: 'stacked',
        categories: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'],
        series: [
          { name: 'Organic', values: [120, 150, 170, 210, 260, 300], color: '#2E75B6' },
          { name: 'Paid', values: [60, 90, 110, 130, 160, 180], color: '#ED7D31' },
        ],
        title: 'Cumulative traffic',
        categoryAxisTitle: 'Month',
        valueAxisTitle: 'Total sessions (thousands)',
        valueAxisTitleRotationDeg: -90,
        valueAxisMajorGridlines: true,
      },
    });

    const bytes = await savePresentation(pres);
    await writeSample('31-scatter-area-chart.pptx', bytes);
  });

  it('32 — merged table (horizontal + vertical merges with banding)', async () => {
    const pres = await loadBlank();
    const slide = freshSlide(pres);
    setSlideTitle(slide, 'Merged table');

    // 6×4 grid: a full-width banner header, a styled header row, then four
    // data rows where the leftmost column groups two rows each ("H1", "H2").
    const table = addSlideTable(slide, {
      x: inches(0.7),
      y: inches(1.8),
      w: inches(9),
      h: inches(3.4),
      rows: [
        ['Regional Sales — 2026', '', '', ''],
        ['Half', 'Region', 'Revenue', 'Margin'],
        ['H1', 'North', '$1.2M', '34%'],
        ['H1', 'South', '$0.9M', '29%'],
        ['H2', 'North', '$1.5M', '36%'],
        ['H2', 'South', '$1.1M', '31%'],
      ],
      firstRow: true,
      bandRow: true,
    });

    // Horizontal merge: banner spans all four columns of the top row.
    mergeTableCells(table, { row: 0, col: 0, rowSpan: 1, colSpan: 4 });
    // Vertical merges: each half-year label spans its two data rows. These
    // cross the bandRow shading boundary, which is the point of the deck.
    mergeTableCells(table, { row: 2, col: 0, rowSpan: 2, colSpan: 1 });
    mergeTableCells(table, { row: 4, col: 0, rowSpan: 2, colSpan: 1 });

    const bytes = await savePresentation(pres);
    await writeSample('32-merged-table.pptx', bytes);
  });

  it('33 — run formatting (strike, super/subscript, per-run color & size)', async () => {
    const pres = await loadBlank();
    const slide = freshSlide(pres);
    setSlideTitle(slide, 'Run formatting');

    const box = addSlideTextBox(slide, {
      x: inches(0.7),
      y: inches(1.5),
      w: inches(9),
      h: inches(5),
      text: [
        'Bold weight',
        'Italic slant',
        'Single underline',
        'Wavy underline',
        'Single strikethrough',
        'Double strikethrough',
        'Superscript raised',
        'Subscript lowered',
        'Large crimson 32pt',
        'Small teal 12pt',
      ].join('\n'),
    });

    // Base format applied to every run, then per-line overrides compose on top.
    setShapeTextFormat(box, { font: 'Calibri', size: 18, color: '#333333' });

    setShapeRunFormat(box, 0, 0, { bold: true });
    setShapeRunFormat(box, 1, 0, { italic: true });
    setShapeRunFormat(box, 2, 0, { underline: true });
    setShapeRunFormat(box, 3, 0, { underline: 'wavy' });
    setShapeRunFormat(box, 4, 0, { strike: true });
    setShapeRunFormat(box, 5, 0, { strike: 'dblStrike' });
    setShapeRunFormat(box, 6, 0, { baseline: 0.3, color: '#2E75B6' });
    setShapeRunFormat(box, 7, 0, { baseline: -0.25, color: '#2E75B6' });
    setShapeRunFormat(box, 8, 0, { size: 32, color: '#C00000', bold: true });
    setShapeRunFormat(box, 9, 0, { size: 12, color: '#1F7A6B', italic: true });

    const bytes = await savePresentation(pres);
    await writeSample('33-run-formatting.pptx', bytes);
  });

  it('34 — shadow variants (angle, blur, offset, opacity, color)', async () => {
    const pres = await loadBlank();
    const slide = freshSlide(pres);
    setSlideTitle(slide, 'Shadow variants');

    // Each cell is a mid-blue tile on the default white slide; only the cast
    // shadow differs, so the rasterized foreground is dominated by shadow
    // direction/blur/offset — keeps fg-SSIM sensitive to shadow rendering.
    const variants: ReadonlyArray<{
      readonly label: string;
      readonly shadow: Parameters<typeof setShapeShadow>[1];
    }> = [
      { label: 'angle 0', shadow: { angleDeg: 0, blurEmu: pt(4), offsetEmu: pt(6) } },
      { label: 'angle 90', shadow: { angleDeg: 90, blurEmu: pt(4), offsetEmu: pt(6) } },
      { label: 'angle 180', shadow: { angleDeg: 180, blurEmu: pt(4), offsetEmu: pt(6) } },
      { label: 'angle 270', shadow: { angleDeg: 270, blurEmu: pt(4), offsetEmu: pt(6) } },
      { label: 'tight blur', shadow: { angleDeg: 45, blurEmu: pt(1), offsetEmu: pt(6) } },
      { label: 'wide blur', shadow: { angleDeg: 45, blurEmu: pt(20), offsetEmu: pt(6) } },
      { label: 'far offset', shadow: { angleDeg: 45, blurEmu: pt(6), offsetEmu: pt(18) } },
      {
        label: 'soft blue',
        shadow: { angleDeg: 45, blurEmu: pt(10), offsetEmu: pt(8), color: '#1F4E79', opacity: 0.4 },
      },
    ];

    const cols = 4;
    const cellW = inches(2.3);
    const cellH = inches(2.4);
    const startX = inches(0.6);
    const startY = inches(1.6);
    variants.forEach((v, i) => {
      const r = Math.floor(i / cols);
      const c = i % cols;
      const tile = addSlideShape(slide, {
        preset: 'roundRect',
        x: (startX + c * cellW) as ReturnType<typeof inches>,
        y: (startY + r * cellH) as ReturnType<typeof inches>,
        w: inches(1.8),
        h: inches(1.4),
        text: v.label,
      });
      setShapeFill(tile, '#9DC3E6');
      setShapeShadow(tile, v.shadow);
    });

    const bytes = await savePresentation(pres);
    await writeSample('34-shadow-variants.pptx', bytes);
  });

  it('35 — text direction (wordArtVert / eaVert / mongolianVert / wordArtVertRtl)', async () => {
    const pres = await loadBlank();
    const slide = freshSlide(pres);
    setSlideTitle(slide, 'Text direction');

    // wordArtVert: characters upright, stacked top-to-bottom (not rotated).
    const wordArt = addSlideTextBox(slide, {
      x: inches(0.6),
      y: inches(1.5),
      w: inches(1.6),
      h: inches(4.5),
      text: 'wordArtVert upright stacked characters',
    });
    setShapeTextDirection(wordArt, 'wordArtVert');
    setShapeRunFormat(wordArt, 0, 0, { size: 18 });

    // eaVert: East-Asian upright, columns flow right-to-left.
    const eaVert = addSlideTextBox(slide, {
      x: inches(2.7),
      y: inches(1.5),
      w: inches(1.6),
      h: inches(4.5),
      text: 'eaVert upright with columns right to left',
    });
    setShapeTextDirection(eaVert, 'eaVert');
    setShapeRunFormat(eaVert, 0, 0, { size: 18 });

    // mongolianVert: rotated ninety degrees, columns flow left-to-right.
    const mongolian = addSlideTextBox(slide, {
      x: inches(4.8),
      y: inches(1.5),
      w: inches(1.6),
      h: inches(4.5),
      text: 'mongolianVert rotated columns left to right',
    });
    setShapeTextDirection(mongolian, 'mongolianVert');
    setShapeRunFormat(mongolian, 0, 0, { size: 18 });

    // wordArtVertRtl: upright stacked characters, right-to-left column order.
    const rtl = addSlideTextBox(slide, {
      x: inches(6.9),
      y: inches(1.5),
      w: inches(1.6),
      h: inches(4.5),
      text: 'wordArtVertRtl upright stacked right to left',
    });
    setShapeTextDirection(rtl, 'wordArtVertRtl');
    setShapeRunFormat(rtl, 0, 0, { size: 18 });

    const bytes = await savePresentation(pres);
    await writeSample('35-text-direction.pptx', bytes);
  });

  it('36 — consulting-grade strategy deck (branded theme, KPI cards, roadmap, data appendix)', async () => {
    // The most demanding deck the library can currently produce: a
    // from-scratch, custom-branded 4-slide "board deck" combining every
    // authoring capability — theme branding, grouped multi-shape
    // components (KPI cards, roadmap nodes), a connector-linked process
    // flow, a banded table, a chart, and dense per-run typography.
    const NAVY = '#0B1F3A';
    const TEAL = '#00A9E0'; // accent1
    const GOLD = '#FDB913'; // accent2
    const SLATE = '#6E7B8B'; // accent3
    const INK = '#1A1A1A';
    const MUTED = '#5B6673';
    const PAGE_W = inches(13.333);
    const MARGIN_X = inches(0.6);
    const CONTENT_W = inches(12.13);

    const pres = createPresentation();
    setPresentationTheme(pres, {
      name: 'Northstar Consulting',
      dark1: NAVY,
      accent1: TEAL,
      accent2: GOLD,
      accent3: SLATE,
    });
    setPresentationFonts(pres, { majorLatin: 'Arial', minorLatin: 'Calibri' });

    const contentSlide = (): SlideData => {
      const layout = findSlideLayoutByType(pres, 'blank') ?? findSlideLayout(pres, 'Blank');
      if (!layout) throw new Error('sample 36: deck has no Blank layout');
      return addSlide(pres, { layout });
    };

    // Shared chrome: a kicker label, an "action title" (the McKinsey-style
    // headline that states the slide's takeaway, not just its topic), and
    // a gold rule underneath.
    const addHeader = (slide: SlideData, kicker: string, headline: string): void => {
      const kickerBox = addSlideTextBox(slide, {
        x: MARGIN_X,
        y: inches(0.35),
        w: CONTENT_W,
        h: inches(0.3),
        text: kicker,
      });
      setShapeRunFormat(kickerBox, 0, 0, {
        font: 'Calibri',
        size: 11,
        bold: true,
        color: 'accent2',
        spc: 150,
      });

      const headlineBox = addSlideTextBox(slide, {
        x: MARGIN_X,
        y: inches(0.68),
        w: CONTENT_W,
        h: inches(1.05),
        text: headline,
      });
      setShapeRunFormat(headlineBox, 0, 0, { font: 'Arial', size: 22, bold: true, color: 'dk1' });

      addSlideLine(slide, {
        from: { x: MARGIN_X, y: inches(1.78) },
        to: { x: emu(MARGIN_X + CONTENT_W), y: inches(1.78) },
        color: 'accent2',
        widthEmu: pt(2),
      });
    };

    const addFooter = (slide: SlideData, pageNumber: number, source: string): void => {
      addSlideLine(slide, {
        from: { x: MARGIN_X, y: inches(7.02) },
        to: { x: emu(MARGIN_X + CONTENT_W), y: inches(7.02) },
        color: '#D9DCE1',
        widthEmu: pt(0.75),
      });
      const sourceBox = addSlideTextBox(slide, {
        x: MARGIN_X,
        y: inches(7.1),
        w: inches(9.5),
        h: inches(0.3),
        text: source,
      });
      setShapeRunFormat(sourceBox, 0, 0, { font: 'Calibri', size: 8, italic: true, color: MUTED });
      const pageBox = addSlideTextBox(slide, {
        x: inches(12.4),
        y: inches(7.1),
        w: inches(0.5),
        h: inches(0.3),
        text: String(pageNumber),
      });
      setShapeRunFormat(pageBox, 0, 0, { font: 'Calibri', size: 9, color: MUTED });
      setParagraphAlignment(pageBox, 0, 'r');
    };

    // A single "$340M — Incremental revenue" style dashboard tile: a
    // rounded card, a thin accent tab on top, a big number, and a label —
    // composed with groupShapes so the whole tile moves/resizes as one.
    const addKpiCard = (
      slide: SlideData,
      x: Emu,
      y: Emu,
      w: Emu,
      value: string,
      label: string,
      accent: string,
    ) => {
      const h = inches(1.5);
      const card = addSlideShape(slide, { preset: 'roundRect', x, y, w, h });
      setShapeFill(card, '#FFFFFF');
      setShapeStroke(card, { color: '#E3E6EA', widthEmu: pt(0.75) });
      setShapeShadow(card, {
        blurEmu: pt(8),
        offsetEmu: pt(2),
        angleDeg: 90,
        color: '#000000',
        opacity: 0.12,
      });

      const tab = addSlideShape(slide, { x, y, w, h: inches(0.09), preset: 'rect' });
      setShapeFill(tab, accent);

      const valueBox = addSlideTextBox(slide, {
        x: emu(x + inches(0.2)),
        y: emu(y + inches(0.28)),
        w: emu(w - inches(0.4)),
        h: inches(0.55),
        text: value,
      });
      setShapeRunFormat(valueBox, 0, 0, { font: 'Arial', size: 26, bold: true, color: 'dk1' });

      const labelBox = addSlideTextBox(slide, {
        x: emu(x + inches(0.2)),
        y: emu(y + inches(0.92)),
        w: emu(w - inches(0.4)),
        h: inches(0.5),
        text: label,
      });
      setShapeRunFormat(labelBox, 0, 0, { font: 'Calibri', size: 10.5, color: MUTED });

      return groupShapes(slide, [card, tab, valueBox, labelBox], { name: `KPI: ${label}` });
    };

    // A roadmap "horizon" node: a numbered badge, a title bar, and an
    // initiative list — grouped so a reader (or a future edit) can treat
    // one phase of the plan as a single object.
    const addHorizonNode = (
      slide: SlideData,
      x: Emu,
      w: Emu,
      index: number,
      title: string,
      initiatives: ReadonlyArray<string>,
      accent: string,
    ) => {
      const y = inches(2.35);
      const headerH = inches(0.85);
      const header = addSlideShape(slide, { preset: 'roundRect', x, y, w, h: headerH });
      setShapeFill(header, NAVY);

      const badge = addSlideShape(slide, {
        preset: 'ellipse',
        x: emu(x + inches(0.18)),
        y: emu(y + inches(0.16)),
        w: inches(0.52),
        h: inches(0.52),
        text: String(index),
      });
      setShapeFill(badge, accent);
      setShapeRunFormat(badge, 0, 0, { font: 'Arial', size: 16, bold: true, color: '#FFFFFF' });

      const titleBox = addSlideTextBox(slide, {
        x: emu(x + inches(0.82)),
        y: emu(y + inches(0.14)),
        w: emu(w - inches(1.0)),
        h: inches(0.58),
        text: title,
      });
      setShapeRunFormat(titleBox, 0, 0, { font: 'Arial', size: 13, bold: true, color: '#FFFFFF' });

      const node = groupShapes(slide, [header, badge, titleBox], {
        name: `Horizon ${index}: ${title}`,
      });

      const listBox = addSlideTextBox(slide, {
        x,
        y: emu(y + headerH + inches(0.18)),
        w,
        h: inches(1.7),
        text: initiatives.join('\n'),
      });
      setShapeBullets(listBox, { char: '—' });
      for (let i = 0; i < initiatives.length; i++) {
        setShapeRunFormat(listBox, i, 0, { font: 'Calibri', size: 11, color: INK });
        setParagraphSpacing(listBox, i, { afterPts: 6 });
      }

      return node;
    };

    // ── Slide 1: cover ──────────────────────────────────────────────────
    const cover = contentSlide();
    setSlideBackground(cover, NAVY);
    const goldBar = addSlideShape(cover, {
      preset: 'rect',
      x: inches(0),
      y: inches(7.3),
      w: PAGE_W,
      h: inches(0.2),
    });
    setShapeFill(goldBar, GOLD);

    // Firm monogram: a circular badge + initials, grouped as one mark.
    const mark = addSlideShape(cover, {
      preset: 'ellipse',
      x: inches(0.6),
      y: inches(0.5),
      w: inches(0.55),
      h: inches(0.55),
      text: 'PK',
    });
    setShapeFill(mark, GOLD);
    setShapeRunFormat(mark, 0, 0, { font: 'Arial', size: 14, bold: true, color: 'dk1' });
    const markLabel = addSlideTextBox(cover, {
      x: inches(1.3),
      y: inches(0.58),
      w: inches(3),
      h: inches(0.4),
      text: 'NORTHSTAR STRATEGY',
    });
    setShapeRunFormat(markLabel, 0, 0, {
      font: 'Calibri',
      size: 11,
      bold: true,
      color: '#FFFFFF',
      spc: 100,
    });
    groupShapes(cover, [mark, markLabel], { name: 'Firm mark' });

    const kicker = addSlideTextBox(cover, {
      x: inches(0.7),
      y: inches(2.6),
      w: inches(11.5),
      h: inches(0.4),
      text: 'STRICTLY CONFIDENTIAL — DRAFT FOR DISCUSSION',
    });
    setShapeRunFormat(kicker, 0, 0, {
      font: 'Calibri',
      size: 12,
      bold: true,
      color: GOLD,
      spc: 150,
    });

    const coverTitle = addSlideTextBox(cover, {
      x: inches(0.7),
      y: inches(3.05),
      w: inches(11.5),
      h: inches(1.5),
      text: 'Project Northstar: Market-Entry Strategy for Southeast Asia',
    });
    setShapeRunFormat(coverTitle, 0, 0, { font: 'Arial', size: 36, bold: true, color: '#FFFFFF' });

    const coverSub = addSlideTextBox(cover, {
      x: inches(0.7),
      y: inches(4.5),
      w: inches(11.5),
      h: inches(0.5),
      text: 'Board Steering Committee Review',
    });
    setShapeRunFormat(coverSub, 0, 0, { font: 'Calibri', size: 18, color: '#C9D3E0' });

    const coverMeta = addSlideTextBox(cover, {
      x: inches(0.7),
      y: inches(6.5),
      w: inches(11.5),
      h: inches(0.4),
      text: 'March 2026  |  Prepared by Strategy & Corporate Development',
    });
    setShapeRunFormat(coverMeta, 0, 0, { font: 'Calibri', size: 11, color: '#8FA0B8' });
    setSlideNotes(
      cover,
      'Opening cover slide. Confirm attendee list and confidentiality reminder before presenting.',
    );

    // ── Slide 2: executive summary (KPI dashboard) ──────────────────────
    const summary = contentSlide();
    addHeader(
      summary,
      'EXECUTIVE SUMMARY',
      'The Southeast Asian expansion can add $340M in incremental revenue by FY28 if we act within the next two quarters',
    );

    const kpis: Array<{ value: string; label: string; accent: string }> = [
      { value: '$340M', label: 'Incremental revenue by FY28', accent: TEAL },
      { value: '18%', label: 'Projected market share, Year 3', accent: GOLD },
      { value: '6', label: 'Priority markets identified', accent: SLATE },
      { value: '24 mo', label: 'Time to cash-flow breakeven', accent: TEAL },
    ];
    const cardW = inches(2.85);
    const cardGap = inches(0.19);
    kpis.forEach((k, i) => {
      const x = emu(MARGIN_X + i * (cardW + cardGap));
      addKpiCard(summary, x, inches(2.05), cardW, k.value, k.label, k.accent);
    });

    const takeawaysTitle = addSlideTextBox(summary, {
      x: MARGIN_X,
      y: inches(3.85),
      w: CONTENT_W,
      h: inches(0.35),
      text: 'Key considerations for the steering committee',
    });
    setShapeRunFormat(takeawaysTitle, 0, 0, { font: 'Arial', size: 13, bold: true, color: 'dk1' });

    const takeawayLines = [
      'Regulatory approval timelines in Vietnam and Indonesia are the primary schedule risk — a parallel filing track is recommended.',
      'Local manufacturing partnership reduces landed cost by an estimated 14% versus direct export, but requires a $22M upfront commitment.',
      'Competitive response is likely within two quarters of a public launch announcement; first-mover advantage is time-limited.',
    ];
    const takeaways = addSlideTextBox(summary, {
      x: MARGIN_X,
      y: inches(4.25),
      w: CONTENT_W,
      h: inches(2.6),
      text: takeawayLines.join('\n'),
    });
    setShapeBullets(takeaways, { char: '●' });
    for (let i = 0; i < takeawayLines.length; i++) {
      setShapeRunFormat(takeaways, i, 0, { font: 'Calibri', size: 13, color: INK });
      setParagraphSpacing(takeaways, i, { afterPts: 10 });
    }
    addFooter(summary, 2, 'SOURCE: Internal analysis; IMF World Economic Outlook, April 2026.');

    // ── Slide 3: strategic roadmap (connected process flow) ─────────────
    const roadmap = contentSlide();
    addHeader(
      roadmap,
      'STRATEGIC ROADMAP',
      'A phased three-horizon plan sequences market entry to manage capital risk',
    );

    const nodeW = inches(3.7);
    const nodeGap = inches(0.55);
    const nodeXs = [0, 1, 2].map((i) => emu(MARGIN_X + i * (nodeW + nodeGap)));
    addHorizonNode(
      roadmap,
      nodeXs[0]!,
      nodeW,
      1,
      'Foundation (Q1–Q2 FY27)',
      [
        'Secure regulatory approval in Vietnam',
        'Finalize manufacturing partnership terms',
        'Build local commercial team',
      ],
      TEAL,
    );
    addHorizonNode(
      roadmap,
      nodeXs[1]!,
      nodeW,
      2,
      'Launch (Q3 FY27–Q1 FY28)',
      [
        'Commercial launch in 2 lead markets',
        'Scale distribution partnerships',
        'Launch localized marketing campaign',
      ],
      GOLD,
    );
    addHorizonNode(
      roadmap,
      nodeXs[2]!,
      nodeW,
      3,
      'Scale (FY28+)',
      [
        'Expand to remaining 4 priority markets',
        'Evaluate local manufacturing footprint',
        'Reassess pricing against competitive response',
      ],
      SLATE,
    );

    // Arrows linking each horizon to the next.
    for (let i = 0; i < nodeXs.length - 1; i++) {
      const from = { x: emu(nodeXs[i]! + nodeW), y: inches(2.775) };
      const to = { x: nodeXs[i + 1]!, y: inches(2.775) };
      const arrow = addSlideLine(roadmap, { from, to, color: NAVY, widthEmu: pt(2) });
      setShapeStrokeArrow(arrow, 'tail', { type: 'triangle' });
    }
    addFooter(roadmap, 3, 'SOURCE: Strategy & Corporate Development analysis.');

    // ── Slide 4: data appendix (table + chart) ───────────────────────────
    const appendix = contentSlide();
    addHeader(
      appendix,
      'APPENDIX A — FINANCIAL DETAIL',
      'Revenue and margin projections vary meaningfully by market-entry scenario',
    );

    addSlideTable(appendix, {
      x: MARGIN_X,
      y: inches(2.05),
      w: inches(5.9),
      h: inches(3.2),
      colWidths: [inches(2.6), inches(1.1), inches(1.1), inches(1.1)],
      rows: [
        ['Scenario', 'FY26', 'FY27', 'FY28'],
        ['Organic only', '$40M', '$95M', '$150M'],
        ['Partnership (recommended)', '$60M', '$210M', '$390M'],
        ['Acquisition-led', '$120M', '$260M', '$430M'],
      ],
      firstRow: true,
      bandRow: true,
    });

    addSlideChart(appendix, {
      x: inches(6.7),
      y: inches(2.05),
      w: inches(5.43),
      h: inches(3.2),
      spec: {
        kind: 'column',
        categories: ['FY26', 'FY27', 'FY28'],
        series: [
          { name: 'Organic only', values: [40, 95, 150] },
          { name: 'Partnership', values: [60, 210, 390] },
          { name: 'Acquisition-led', values: [120, 260, 430] },
        ],
        title: 'Revenue by scenario ($M)',
      },
    });

    const footnote = addSlideTextBox(appendix, {
      x: MARGIN_X,
      y: inches(5.45),
      w: CONTENT_W,
      h: inches(0.8),
      text: 'Note: figures are unaudited management estimates in USD millions. The partnership scenario assumes regulatory approval is secured by Q2 FY27; a delay of two quarters reduces FY28 revenue by an estimated 15–20%.',
    });
    setShapeRunFormat(footnote, 0, 0, { font: 'Calibri', size: 10, italic: true, color: MUTED });
    addFooter(appendix, 4, 'SOURCE: Finance business partnering team; scenario model v4.2.');

    const bytes = await savePresentation(pres);
    await writeSample('36-consulting-deck.pptx', bytes);
  });
});
