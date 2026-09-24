import { expect, it } from 'vitest';
import * as pptx from '../src/api/index.ts';
import { INTERNAL_PACKAGE } from '../src/api/_internal-symbols.ts';
import { PRES_PART_NAME, decode, encode } from '../src/api/fn/_helpers.ts';
import { partName } from '../src/internal/opc/index.ts';
import { REL_TYPES } from '../src/internal/presentationml/relationship-types.ts';
import { NS } from '../src/internal/xml/index.ts';
import { expectSchemaValid, isSchemaValidationAvailable } from './lib/expect-schema-valid.ts';

it('round-trips slideshow choices and flags with valid OOXML', async () => {
  const p = pptx.createPresentation();
  const slide = pptx.addBlankSlide(p);
  pptx.addBlankSlide(p);
  const defaults = pptx.getSlideShowProperties(p);
  expect(defaults).toEqual({
    mode: { kind: 'present' },
    slides: { kind: 'all' },
    loop: false,
    showNarration: false,
    showAnimation: true,
    useTimings: true,
  });
  pptx.setCustomShows(p, [{ id: 7, name: 'Details', slides: [slide] }]);
  const choices: pptx.SlideShowProperties[] = [
    {
      ...defaults,
      mode: { kind: 'browse', showScrollbar: false },
      slides: { kind: 'range', start: 1, end: 2 },
      loop: true,
      showNarration: true,
    },
    {
      ...defaults,
      mode: { kind: 'kiosk', restart: 5000 },
      slides: { kind: 'customShow', id: 7 },
      useTimings: false,
      showAnimation: false,
    },
    defaults,
  ];
  for (const settings of choices) {
    pptx.setSlideShowProperties(p, settings);
    const xml = decode(p[INTERNAL_PACKAGE].getPart(partName('/ppt/presProps.xml'))!.data);
    if (isSchemaValidationAvailable()) expectSchemaValid(xml, 'pml');
    expect(
      pptx.getSlideShowProperties(await pptx.loadPresentation(await pptx.savePresentation(p))),
    ).toEqual(settings);
  }
});

it('preserves unrelated properties and follows custom relationship targets', () => {
  const p = pptx.createPresentation();
  const pkg = p[INTERNAL_PACKAGE];
  const original = pkg.getPart(partName('/ppt/presProps.xml'))!;
  const related = pkg.addPart(
    partName('/custom/settings.xml'),
    original.contentType,
    encode(
      `<p:presentationPr xmlns:p="${NS.pml}" xmlns:a="${NS.dml}"><p:prnPr/><p:showPr><p:browse/><p:sldAll/><p:penClr><a:srgbClr val="123456"/></p:penClr><p:extLst><p:ext uri="keep"/></p:extLst></p:showPr><p:clrMru><a:srgbClr val="ABCDEF"/></p:clrMru></p:presentationPr>`,
    ),
  );
  const rels = pkg.getRels(PRES_PART_NAME)!;
  rels.items.find((r) => r.type === REL_TYPES.presProps)!.target = '../custom/settings.xml';
  pkg.setRels(PRES_PART_NAME, rels);
  const before = original.data;
  const settings = pptx.getSlideShowProperties(p);
  expect(settings.mode).toEqual({ kind: 'browse', showScrollbar: true });
  pptx.setSlideShowProperties(p, { ...settings, mode: { kind: 'present' } });
  const xml = decode(related.data);
  expect(original.data).toEqual(before);
  for (const preserved of ['<p:prnPr', '123456', 'uri="keep"', 'ABCDEF'])
    expect(xml).toContain(preserved);
  expect(xml).not.toContain('<p:browse');
  if (isSchemaValidationAvailable()) expectSchemaValid(xml, 'pml');
});

it('creates missing properties without overwriting an occupied part', async () => {
  const p = pptx.createPresentation();
  const pkg = p[INTERNAL_PACKAGE];
  const original = pkg.getPart(partName('/ppt/presProps.xml'))!;
  const before = original.data;
  const settings = { ...pptx.getSlideShowProperties(p), loop: true };
  const rels = pkg.getRels(PRES_PART_NAME)!;
  rels.items = rels.items.filter((r) => r.type !== REL_TYPES.presProps);
  pkg.setRels(PRES_PART_NAME, rels);
  pptx.setSlideShowProperties(p, settings);
  expect(original.data).toEqual(before);
  expect(pkg.getPart(partName('/ppt/presProps1.xml'))).toBeTruthy();
  expect(
    pptx.getSlideShowProperties(await pptx.loadPresentation(await pptx.savePresentation(p))),
  ).toEqual(settings);
});

it('rejects invalid settings and unusable parts before mutation', () => {
  const p = pptx.createPresentation();
  pptx.addBlankSlide(p);
  const pkg = p[INTERNAL_PACKAGE];
  const part = pkg.getPart(partName('/ppt/presProps.xml'))!;
  const defaults = pptx.getSlideShowProperties(p);
  const before = part.data;
  for (const patch of [
    { slides: { kind: 'range', start: 0, end: 1 } },
    { slides: { kind: 'range', start: 1, end: 2 } },
    { slides: { kind: 'customShow', id: 99 } },
    { mode: { kind: 'kiosk', restart: -1 } },
    { mode: { kind: 'browse', showScrollbar: 'yes' } },
    { loop: 1 },
  ]) {
    expect(() =>
      pptx.setSlideShowProperties(p, { ...defaults, ...patch } as pptx.SlideShowProperties),
    ).toThrow();
    expect(part.data).toEqual(before);
  }
  for (const xml of ['<invalid', '<wrong/>']) {
    part.data = encode(xml);
    expect(() => pptx.setSlideShowProperties(p, defaults)).toThrow();
    expect(decode(part.data)).toBe(xml);
  }
  part.data = before;
  const rels = pkg.getRels(PRES_PART_NAME)!;
  rels.items.find((r) => r.type === REL_TYPES.presProps)!.targetMode = 'External';
  pkg.setRels(PRES_PART_NAME, rels);
  expect(() => pptx.setSlideShowProperties(p, defaults)).toThrow('internal');
  expect(part.data).toEqual(before);
});
