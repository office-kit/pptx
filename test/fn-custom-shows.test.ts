import { expect, it } from 'vitest';
import * as pptx from '../src/api/index.ts';
import { INTERNAL_PACKAGE, SLIDE_PART_NAME } from '../src/api/_internal-symbols.ts';
import { PRES_PART_NAME, decode, encode } from '../src/api/fn/_helpers.ts';
import { expectSchemaValid, isSchemaValidationAvailable } from './lib/expect-schema-valid.ts';
import { REL_TYPES } from '../src/internal/presentationml/relationship-types.ts';

it('round-trips named sequences with repetitions independently from deck order', async () => {
  const p = pptx.createPresentation();
  const a = pptx.addBlankSlide(p),
    b = pptx.addBlankSlide(p),
    c = pptx.addBlankSlide(p);
  expect(pptx.getCustomShows(p)).toEqual([]);
  pptx.setCustomShows(p, [
    { id: 0, name: 'Short & <focused>', slides: [c, a, c] },
    { id: 7, name: 'Empty', slides: [] },
  ]);
  const xml = decode(p[INTERNAL_PACKAGE].getPart(PRES_PART_NAME)!.data);
  if (isSchemaValidationAvailable()) expectSchemaValid(xml, 'pml');
  pptx.sortSlides(
    p,
    (left, right) => left[SLIDE_PART_NAME].localeCompare(right[SLIDE_PART_NAME]) * -1,
  );
  const loaded = await pptx.loadPresentation(await pptx.savePresentation(p));
  expect(
    pptx.getCustomShows(loaded).map((show) => ({
      ...show,
      slides: show.slides.map((slide) => slide[SLIDE_PART_NAME]),
    })),
  ).toEqual([
    { id: 0, name: 'Short & <focused>', slides: [c, a, c].map((slide) => slide[SLIDE_PART_NAME]) },
    { id: 7, name: 'Empty', slides: [] },
  ]);
  expect(pptx.validatePresentation(loaded).filter((issue) => issue.severity === 'error')).toEqual(
    [],
  );
  expect(pptx.getSlides(p)).toHaveLength(3);
  expect(b).toBeDefined();
});

it('updates surviving metadata and rejects invalid replacements without partial changes', async () => {
  const p = pptx.createPresentation();
  const slide = pptx.addBlankSlide(p);
  pptx.setCustomShows(p, [{ id: 2, name: 'Original', slides: [slide] }]);
  const part = p[INTERNAL_PACKAGE].getPart(PRES_PART_NAME)!;
  part.data = encode(
    decode(part.data).replace(
      '</p:custShow>',
      '<p:extLst><p:ext uri="test"/></p:extLst></p:custShow>',
    ),
  );
  pptx.setCustomShows(p, [{ id: 2, name: 'Renamed', slides: [slide, slide] }]);
  expect(decode(part.data)).toContain('<p:ext uri="test"');
  const before = await pptx.savePresentation(p);
  const other = pptx.addBlankSlide(pptx.createPresentation());
  for (const shows of [
    [{ id: -1, name: 'Invalid', slides: [slide] }],
    [
      { id: 2, name: 'Valid', slides: [slide] },
      { id: 2, name: 'Duplicate ID', slides: [] },
    ],
    [{ id: 3, name: 'Foreign', slides: [other] }],
  ]) {
    expect(() => pptx.setCustomShows(p, shows)).toThrow();
    expect(await pptx.savePresentation(p)).toEqual(before);
  }
  pptx.setCustomShows(p, []);
  expect(pptx.getCustomShows(p)).toEqual([]);
});

it('removes every deleted-slide occurrence using canonical relationship targets', async () => {
  const p = pptx.createPresentation();
  const a = pptx.addBlankSlide(p),
    b = pptx.addBlankSlide(p);
  pptx.setCustomShows(p, [
    { id: 0, name: 'Repeated', slides: [a, b, a] },
    { id: 1, name: 'Only A', slides: [a] },
  ]);
  const pkg = p[INTERNAL_PACKAGE];
  const rels = pkg.getRels(PRES_PART_NAME)!;
  rels.items
    .filter((rel) => rel.type === REL_TYPES.slide)
    .forEach((rel) => {
      rel.target = '/ppt/' + rel.target;
    });
  pkg.setRels(PRES_PART_NAME, rels);
  const foreign = pptx.addBlankSlide(pptx.createPresentation());
  expect(() => pptx.removeSlide(p, foreign)).toThrow();
  pptx.removeSlide(p, a);
  expect(pptx.getCustomShows(p).map((show) => [show.id, show.slides.length])).toEqual([
    [0, 1],
    [1, 0],
  ]);
  const loaded = await pptx.loadPresentation(await pptx.savePresentation(p));
  expect(pptx.getCustomShows(loaded)[0]!.slides[0]![SLIDE_PART_NAME]).toBe(b[SLIDE_PART_NAME]);
  expect(pptx.validatePresentation(loaded).filter((issue) => issue.severity === 'error')).toEqual(
    [],
  );
  const xml = decode(pkg.getPart(PRES_PART_NAME)!.data);
  if (isSchemaValidationAvailable()) expectSchemaValid(xml, 'pml');
});

it('round-trips custom-show object, hover and text actions and rejects absent destinations atomically', async () => {
  const p = pptx.createPresentation();
  const slide = pptx.addBlankSlide(p);
  const destination = pptx.addBlankSlide(p);
  pptx.setCustomShows(p, [{ id: 7, name: 'Detail', slides: [destination] }]);
  const shape = pptx.addSlideTextBox(slide, {
    x: pptx.inches(1),
    y: pptx.inches(1),
    w: pptx.inches(4),
    h: pptx.inches(1),
    text: 'Details',
  });
  const returning = { kind: 'customShow', id: 7, showAndReturn: true } as const;
  const exiting = { ...returning, showAndReturn: false };
  pptx.setShapeClickAction(shape, returning, 'Open details');
  pptx.setShapeHoverAction(shape, exiting);
  pptx.setShapeTextRangeClickAction(shape, 0, 7, returning);
  const xml = pptx.getSlideXmlString(slide);
  expect(xml).toContain('ppaction://customshow?id=7&amp;return=true');
  if (isSchemaValidationAvailable()) expectSchemaValid(xml, 'pml');
  const loaded = await pptx.loadPresentation(await pptx.savePresentation(p));
  const result = pptx.getSlideShapes(pptx.getSlides(loaded)[0]!)[0]!;
  expect(pptx.getShapeClickAction(result)).toEqual(returning);
  expect(pptx.getShapeHoverAction(result)).toEqual(exiting);
  expect(pptx.getShapeTextRangeClickActions(result)[0]?.action).toEqual(returning);
  expect(pptx.getShapeClickActionTooltip(result)).toBe('Open details');
  const before = await pptx.savePresentation(p);
  for (const id of [-1, 8, 1.2, 0x100000000]) {
    expect(() => pptx.setShapeClickAction(shape, { ...returning, id })).toThrow(/custom show/);
    expect(() => pptx.setShapeTextRangeClickAction(shape, 0, 3, { ...returning, id })).toThrow(
      /custom show/,
    );
    expect(await pptx.savePresentation(p)).toEqual(before);
  }
});
