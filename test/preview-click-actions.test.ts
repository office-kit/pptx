import { expect, it } from 'vitest';
import { buildPng } from './lib/build-png.ts';
import * as pptx from '../src/api/index.ts';
import { renderSlideToSvg } from '../packages/preview/src/index.ts';

it('renders object and run navigation actions and keeps object tooltips separate', async () => {
  const p = pptx.createPresentation();
  const slide = pptx.addBlankSlide(p);
  const destination = pptx.addBlankSlide(p);
  const shape = pptx.addSlideTextBox(slide, {
    x: pptx.inches(1),
    y: pptx.inches(1),
    w: pptx.inches(4),
    h: pptx.inches(1),
    text: 'Linked text',
  });
  pptx.setShapeRunHyperlink(shape, 0, 0, 'https://example.com/run', 'Run tip');
  pptx.setShapeClickAction(shape, { kind: 'slide', slide: destination }, 'Object tip');
  let svg = await renderSlideToSvg(p, slide);
  expect(svg).toContain('href="#slide-2"');
  expect(svg).toContain('<title>Object tip</title>');
  expect(svg).toContain('href="https://example.com/run"');
  // Text hyperlinks must not turn the whole shape into the same external link.
  expect(svg.match(/href="https:\/\/example.com\/run"/g)).toHaveLength(1);
  pptx.setShapeTextRangeClickAction(shape, 0, 6, { kind: 'slide', slide: destination });
  svg = await renderSlideToSvg(p, slide);
  expect(svg.match(/href="#slide-2"/g)).toHaveLength(2);
  for (const kind of [
    'nextSlide',
    'prevSlide',
    'firstSlide',
    'lastSlide',
    'lastSlideViewed',
    'endShow',
  ] as const) {
    pptx.setShapeClickAction(shape, { kind });
    pptx.setShapeTextRangeClickAction(shape, 0, 6, { kind });
    svg = await renderSlideToSvg(p, slide);
    expect(svg.match(new RegExp(`href="#slide-${kind}"`, 'g'))).toHaveLength(2);
  }
});

it('renders navigation on pictures, tables and groups after save/reload', async () => {
  const p = pptx.createPresentation();
  const slide = pptx.addBlankSlide(p);
  const destination = pptx.addBlankSlide(p);
  const box = { x: pptx.inches(1), y: pptx.inches(1), w: pptx.inches(2), h: pptx.inches(1) };
  const picture = pptx.addSlideImage(slide, buildPng(2, 1, [0, 255, 0]), box);
  const table = pptx.addSlideTable(slide, { ...box, rows: [['cell']] });
  const first = pptx.addSlideTextBox(slide, { ...box, text: 'first' });
  const second = pptx.addSlideTextBox(slide, { ...box, text: 'second' });
  const group = pptx.groupShapes([first, second]);
  for (const shape of [picture, table, group]) {
    pptx.setShapeClickAction(shape, { kind: 'slide', slide: destination }, 'Go');
  }
  const saved = await pptx.loadPresentation(await pptx.savePresentation(p));
  const svg = await renderSlideToSvg(saved, pptx.getSlides(saved)[0]!);
  expect(svg.match(/href="#slide-2"/g)).toHaveLength(3);
  expect(svg.match(/<title>Go<\/title>/g)).toHaveLength(3);
});

it.each(['svg', 'foreignObject'] as const)(
  'renders cell links across Unicode and paragraph boundaries in %s',
  async (textLayout) => {
    const p = pptx.createPresentation();
    const slide = pptx.addBlankSlide(p);
    const destination = pptx.addBlankSlide(p);
    const table = pptx.addSlideTable(slide, {
      x: pptx.inches(1),
      y: pptx.inches(1),
      w: pptx.inches(6),
      h: pptx.inches(2),
      rows: [['A😀B\nNext', 'Unlinked']],
    });
    const cell = pptx.getTableCell(table, 0, 0);
    pptx.setTableCellTextRangeClickAction(
      cell,
      1,
      3,
      { kind: 'url', url: 'https://example.com/?a=1&b=2' },
      'Emoji <link>',
    );
    pptx.setTableCellTextRangeClickAction(
      cell,
      5,
      9,
      { kind: 'slide', slide: destination },
      'Next slide',
    );
    const loaded = await pptx.loadPresentation(await pptx.savePresentation(p));
    const svg = renderSlideToSvg(loaded, pptx.getSlides(loaded)[0]!, { textLayout });
    expect(svg.match(/href="https:\/\/example.com\/\?a=1&amp;b=2"/g)).toHaveLength(1);
    expect(svg.match(/href="#slide-2"/g)).toHaveLength(1);
    expect(svg).toContain('Emoji &lt;link&gt;');
    expect(svg).toContain('Next slide');
    expect(svg).toContain('Unlinked');
    expect(svg).toContain('underline');
  },
);

it.each(['lastSlideViewed', 'endShow'] as const)(
  'round-trips %s on objects, text and cells',
  async (kind) => {
    const p = pptx.createPresentation();
    const slide = pptx.addBlankSlide(p);
    const box = { x: pptx.inches(1), y: pptx.inches(1), w: pptx.inches(4), h: pptx.inches(1) };
    const shape = pptx.addSlideTextBox(slide, { ...box, text: 'Jump' });
    const table = pptx.addSlideTable(slide, { ...box, rows: [['Jump']] });
    pptx.setShapeClickAction(shape, { kind }, 'Object');
    pptx.setShapeTextRangeClickAction(shape, 0, 4, { kind }, 'Text');
    pptx.setTableCellTextRangeClickAction(pptx.getTableCell(table, 0, 0), 0, 4, { kind }, 'Cell');
    const restored = await pptx.loadPresentation(await pptx.savePresentation(p));
    const loaded = pptx.getSlides(restored)[0]!;
    const shapes = pptx.getSlideShapes(loaded);
    expect(pptx.getShapeClickAction(shapes[0]!)).toEqual({ kind });
    expect(pptx.getShapeTextRangeClickActions(shapes[0]!)[0]).toMatchObject({
      start: 0,
      end: 4,
      action: { kind },
      tooltip: 'Text',
    });
    expect(
      pptx.getTableCellTextRangeClickActions(pptx.getTableCell(shapes[1]!, 0, 0))[0],
    ).toMatchObject({ start: 0, end: 4, action: { kind }, tooltip: 'Cell' });
    expect(
      pptx
        .getSlideXmlString(loaded)
        .match(new RegExp('ppaction://hlinkshowjump\\?jump=' + kind.toLowerCase(), 'g')),
    ).toHaveLength(3);
    for (const textLayout of ['svg', 'foreignObject'] as const) {
      expect(
        renderSlideToSvg(restored, loaded, { textLayout }).match(
          new RegExp('href="#slide-' + kind + '"', 'g'),
        ),
      ).toHaveLength(3);
    }
  },
);

it.each(['svg', 'foreignObject'] as const)(
  'renders hover-only and independent click/hover actions in %s',
  async (textLayout) => {
    const p = pptx.createPresentation();
    const slide = pptx.addBlankSlide(p);
    const next = pptx.addBlankSlide(p);
    const shape = pptx.addSlideTextBox(slide, {
      x: pptx.inches(1),
      y: pptx.inches(1),
      w: pptx.inches(4),
      h: pptx.inches(1),
      text: 'Hover',
    });
    pptx.setShapeHoverAction(shape, { kind: 'slide', slide: next }, 'Hover <tip>');
    const saved = await pptx.loadPresentation(await pptx.savePresentation(p));
    let svg = await renderSlideToSvg(saved, pptx.getSlides(saved)[0]!, { textLayout });
    expect(svg).toContain('<g data-hover-href="#slide-2" pointer-events="bounding-box">');
    expect(svg).toContain('<title>Hover &lt;tip&gt;</title>');
    expect(svg).not.toContain('<a href=');
    pptx.setShapeClickAction(shape, { kind: 'endShow' }, 'Click tip');
    pptx.setShapeHoverAction(shape, { kind: 'url', url: 'https://example.com/?a=1&b=2' });
    svg = await renderSlideToSvg(p, slide, { textLayout });
    expect(svg).toContain(
      '<a href="#slide-endShow" data-hover-href="https://example.com/?a=1&amp;b=2" pointer-events="bounding-box">',
    );
    expect(svg).toContain('<title>Click tip</title>');
  },
);

it.each(['svg', 'foreignObject'] as const)(
  'renders sound-only triggers after reload in %s',
  async (textLayout) => {
    const p = pptx.createPresentation();
    const slide = pptx.addBlankSlide(p);
    const shape = pptx.addSlideTextBox(slide, {
      x: pptx.inches(1),
      y: pptx.inches(1),
      w: pptx.inches(2),
      h: pptx.inches(1),
      text: 'Sound',
    });
    const bytes = new Uint8Array(12);
    bytes.set(new TextEncoder().encode('RIFF'), 0);
    bytes.set(new TextEncoder().encode('WAVE'), 8);
    pptx.setShapeActionSound(shape, 'click', {
      sound: { name: 'Sound', bytes },
      stopPrevious: false,
    });
    pptx.setShapeActionSound(shape, 'hover', { sound: null, stopPrevious: true });
    const loaded = await pptx.loadPresentation(await pptx.savePresentation(p));
    const svg = renderSlideToSvg(loaded, pptx.getSlides(loaded)[0]!, { textLayout });
    expect(svg).toContain('data-click-sound="data:audio/wav;base64,UklGRgAAAABXQVZF"');
    expect(svg).toContain('data-hover-stop-sound="true"');
    expect(svg).not.toContain('<a href=');
  },
);

it('renders custom show ids and return flags in object, hover and run links', async () => {
  const p = pptx.createPresentation();
  const slide = pptx.addBlankSlide(p);
  pptx.setCustomShows(p, [{ id: 3, name: 'Details', slides: [slide] }]);
  const shape = pptx.addSlideTextBox(slide, {
    x: pptx.inches(1),
    y: pptx.inches(1),
    w: pptx.inches(4),
    h: pptx.inches(1),
    text: 'Details',
  });
  pptx.setShapeClickAction(shape, { kind: 'customShow', id: 3, showAndReturn: true });
  pptx.setShapeHoverAction(shape, { kind: 'customShow', id: 3, showAndReturn: false });
  pptx.setShapeTextRangeClickAction(shape, 0, 7, {
    kind: 'customShow',
    id: 3,
    showAndReturn: true,
  });
  for (const textLayout of ['svg', 'foreignObject'] as const) {
    const svg = await renderSlideToSvg(p, slide, { textLayout });
    expect(svg.match(/href="#slide-customShow-3-return"/g)).toHaveLength(2);
    expect(svg).toContain('data-hover-href="#slide-customShow-3-exit"');
  }
});

it.each(['svg', 'foreignObject'] as const)(
  'hidden objects and groups omit all action hotspots and restore them on show in %s',
  async (textLayout) => {
    const p = pptx.createPresentation();
    const slide = pptx.addBlankSlide(p);
    const box = { x: pptx.inches(1), y: pptx.inches(1), w: pptx.inches(2), h: pptx.inches(1) };
    const linked = pptx.addSlideTextBox(slide, { ...box, text: 'Linked content' });
    const sibling = pptx.addSlideTextBox(slide, { ...box, text: 'Sibling content' });
    pptx.setShapeRunHyperlink(linked, 0, 0, 'https://example.com/hidden-run');
    pptx.setShapeClickAction(linked, { kind: 'endShow' });
    pptx.setShapeHoverAction(linked, { kind: 'nextSlide' });
    pptx.setShapeActionSound(linked, 'click', { sound: null, stopPrevious: true });
    pptx.setShapeActionSound(linked, 'hover', { sound: null, stopPrevious: true });
    const group = pptx.groupShapes([linked, sibling]);
    const image = pptx.addSlideImage(slide, buildPng(2, 1, [0, 255, 0]), box);
    pptx.setShapeClickAction(image, { kind: 'lastSlide' });
    const render = () =>
      renderSlideToSvg(p, slide, { textLayout }).replace(/pkdef-[a-z0-9]+/g, 'pkdef-normalized');
    const shown = render();
    expect(shown).toContain('href="https://example.com/hidden-run"');
    expect(shown).toContain('data-hover-href=');
    expect(shown).toContain('data-click-stop-sound=');
    pptx.setShapeHidden(linked, true);
    const hiddenChild = render();
    expect(hiddenChild).not.toContain('Linked');
    expect(hiddenChild).not.toContain('hidden-run');
    expect(hiddenChild).not.toContain('data-hover-href=');
    expect(hiddenChild).not.toContain('data-click-stop-sound=');
    expect(hiddenChild).not.toContain('data-hover-stop-sound=');
    expect(hiddenChild).toContain('Sibling');
    expect(hiddenChild).toContain('href="#slide-lastSlide"');
    pptx.setShapeHidden(linked, false);
    expect(render()).toBe(shown);
    pptx.setShapeHidden(group, true);
    pptx.setShapeHidden(image, true);
    const loaded = await pptx.loadPresentation(await pptx.savePresentation(p));
    const hiddenAll = renderSlideToSvg(loaded, pptx.getSlides(loaded)[0]!, { textLayout });
    expect(hiddenAll).not.toContain('<a ');
    expect(hiddenAll).not.toContain('data-hover-href=');
    expect(hiddenAll).not.toContain('data-click-stop-sound=');
    expect(hiddenAll).not.toContain('Linked');
    expect(hiddenAll).not.toContain('Sibling');
    pptx.setShapeHidden(group, false);
    pptx.setShapeHidden(image, false);
    expect(render()).toBe(shown);
  },
);
