import { describe, expect, it } from 'vitest';
import { partName } from '../src/internal/opc/index.ts';
import { REL_TYPES } from '../src/internal/presentationml/relationship-types.ts';
import { expectSchemaValid, isSchemaValidationAvailable } from './lib/expect-schema-valid.ts';
import {
  createPresentation,
  getGridSpacing,
  getSnapToGrid,
  setSnapToGrid,
  setGridSpacing,
  getDrawingGuides,
  getDrawingGuidesVisible,
  setDrawingGuidesVisible,
  setDrawingGuides,
  loadPresentation,
  savePresentation,
} from '../src/api/index.ts';
import { INTERNAL_PACKAGE } from '../src/api/_internal-symbols.ts';
import { PRES_PART_NAME, decode, encode } from '../src/api/fn/_helpers.ts';

describe('presentation drawing guides', () => {
  it('round-trips document snapping independently and reads native implicit true', async () => {
    const first = createPresentation();
    const second = createPresentation();
    expect(getSnapToGrid(first)).toBeNull();
    setSnapToGrid(first, true);
    setGridSpacing(first, { x: 72008, y: 72008 });
    setDrawingGuidesVisible(first, true);
    expect(getSnapToGrid(second)).toBeNull();
    expect(getSnapToGrid(await loadPresentation(await savePresentation(first)))).toBe(true);
    setSnapToGrid(first, false);
    const restored = await loadPresentation(await savePresentation(first));
    expect(getSnapToGrid(restored)).toBe(false);
    expect(getGridSpacing(restored)).toEqual({ x: 72008, y: 72008 });
    expect(getDrawingGuidesVisible(restored)).toBe(true);
    const part = restored[INTERNAL_PACKAGE].getPart(partName('/ppt/viewProps.xml'))!;
    if (isSchemaValidationAvailable()) expectSchemaValid(decode(part.data), 'pml');
    part.data = encode(decode(part.data).replace(' snapToGrid="0"', ''));
    expect(getSnapToGrid(restored)).toBe(true);
    const before = part.data;
    expect(() => setSnapToGrid(restored, 1 as unknown as boolean)).toThrow('Invalid');
    expect(part.data).toEqual(before);
  });

  it('round-trips independent grid axes while preserving view settings and guide metadata', async () => {
    const presentation = createPresentation();
    expect(getGridSpacing(presentation)).toBeNull();
    setDrawingGuidesVisible(presentation, true);
    setDrawingGuides(presentation, [{ id: 1, axis: 'x', position: 4572000, color: '#888888' }]);
    setGridSpacing(presentation, { x: 72008, y: 180000 });
    const restored = await loadPresentation(await savePresentation(presentation));
    expect(getGridSpacing(restored)).toEqual({ x: 72008, y: 180000 });
    expect(getDrawingGuidesVisible(restored)).toBe(true);
    expect(getDrawingGuides(restored)).toEqual(getDrawingGuides(presentation));
    const part = restored[INTERNAL_PACKAGE].getPart(partName('/ppt/viewProps.xml'))!;
    if (isSchemaValidationAvailable()) expectSchemaValid(decode(part.data), 'pml');
    const before = part.data;
    expect(() => setGridSpacing(restored, { x: 0, y: 72000 })).toThrow('Invalid');
    expect(() => setGridSpacing(restored, { x: NaN, y: 72000 })).toThrow('Invalid');
    expect(part.data).toEqual(before);
    setDrawingGuidesVisible(restored, false);
    expect(getGridSpacing(restored)).toEqual({ x: 72008, y: 180000 });
  });
  it('writes native master units and round-trips colors, axes, and explicit deletion', async () => {
    const presentation = createPresentation();
    expect(getDrawingGuides(presentation)).toBeNull();
    const guides = [
      { id: 1, axis: 'y' as const, position: 3429000, color: '#2873c4' },
      { id: 2, axis: 'x' as const, position: 4572000, color: '#888888' },
    ];
    setDrawingGuides(presentation, guides);
    const xml = decode(presentation[INTERNAL_PACKAGE].getPart(PRES_PART_NAME)!.data);
    if (isSchemaValidationAvailable()) expectSchemaValid(xml, 'pml');
    expect(xml).toContain('pos="2160"');
    expect(xml).toContain('pos="2880"');
    expect(xml).toContain('EFAFB233-063F-42B5-8137-9DF3F51BA10A');
    const restored = await loadPresentation(await savePresentation(presentation));
    expect(getDrawingGuides(restored)).toEqual(guides);
    setDrawingGuides(restored, []);
    expect(getDrawingGuides(await loadPresentation(await savePresentation(restored)))).toEqual([]);
  });
  it('imports related legacy guides and synchronizes them without changing notes settings', async () => {
    const presentation = createPresentation();
    const pkg = presentation[INTERNAL_PACKAGE];
    const original = pkg.getPart(partName('/ppt/viewProps.xml'))!;
    const common =
      '<p:cViewPr><p:scale><a:sx n="100" d="100"/><a:sy n="100" d="100"/></p:scale><p:origin x="0" y="0"/></p:cViewPr>';
    const notes = `<p:notesViewPr><p:cSldViewPr>${common}<p:guideLst><p:guide pos="777"/></p:guideLst></p:cSldViewPr></p:notesViewPr>`;
    const legacy = pkg.addPart(
      partName('/custom/views.xml'),
      original.contentType,
      encode(
        `<p:viewPr xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:slideViewPr><p:cSldViewPr showGuides="1">${common}<p:guideLst><p:guide orient="horz" pos="2160"/><p:guide/></p:guideLst></p:cSldViewPr></p:slideViewPr>${notes}<p:gridSpacing cx="91440" cy="91440"/></p:viewPr>`,
      ),
    );
    const rels = pkg.getRels(PRES_PART_NAME)!;
    rels.items.find((rel) => rel.type === REL_TYPES.viewProps)!.target = '../custom/views.xml';
    pkg.setRels(PRES_PART_NAME, rels);
    expect(getDrawingGuides(presentation)).toEqual([
      { id: 1, axis: 'y', position: 3429000, color: '#888888' },
      { id: 2, axis: 'x', position: 0, color: '#888888' },
    ]);
    const moved = [{ id: 2, axis: 'x' as const, position: -15875, color: '#2873c4' }];
    setDrawingGuides(presentation, moved);
    const xml = decode(legacy.data);
    expect(xml).toContain('pos="-10"');
    expect(xml).not.toContain('pos="2160"');
    expect(xml).toContain(notes);
    expect(xml).toContain('showGuides="1"');
    expect(xml).toContain('cx="91440"');
    if (isSchemaValidationAvailable()) expectSchemaValid(xml, 'pml');
    expect(getDrawingGuides(await loadPresentation(await savePresentation(presentation)))).toEqual(
      moved,
    );
    setDrawingGuides(presentation, []);
    expect(decode(legacy.data)).not.toContain('pos="-10"');
    // An explicit extended empty list overrides a stale legacy list.
    legacy.data = encode(xml);
    expect(getDrawingGuides(presentation)).toEqual([]);
    const before = pkg.getPart(PRES_PART_NAME)!.data;
    legacy.data = encode('<invalid');
    expect(() => setDrawingGuides(presentation, moved)).toThrow();
    expect(pkg.getPart(PRES_PART_NAME)!.data).toEqual(before);
  });
  it('persists visibility, creates missing view properties and preserves existing settings', async () => {
    const presentation = createPresentation();
    const pkg = presentation[INTERNAL_PACKAGE];
    const name = partName('/ppt/viewProps.xml');
    expect(getDrawingGuidesVisible(presentation)).toBe(false);
    setDrawingGuidesVisible(presentation, true);
    const view = pkg.getPart(name)!;
    if (isSchemaValidationAvailable()) expectSchemaValid(decode(view.data), 'pml');
    view.data = encode(
      decode(view.data).replace('showGuides="1"', 'showGuides="true"').replace('n="100"', 'n="75"'),
    );
    expect(getDrawingGuidesVisible(presentation)).toBe(true);
    setDrawingGuidesVisible(presentation, false);
    expect(decode(view.data)).toContain('snapToGrid="0"');
    expect(decode(view.data)).toContain('n="75"');
    expect(
      getDrawingGuidesVisible(await loadPresentation(await savePresentation(presentation))),
    ).toBe(false);
    const rels = pkg.getRels(PRES_PART_NAME)!;
    rels.items = rels.items.filter((rel) => rel.type !== REL_TYPES.viewProps);
    pkg.setRels(PRES_PART_NAME, rels);
    // Do not overwrite an unrelated part with the conventional name.
    const orphan = view.data;
    setDrawingGuidesVisible(presentation, true);
    expect(view.data).toEqual(orphan);
    const created = pkg.getPart(partName('/ppt/viewProps1.xml'))!;
    expect(created).not.toBeNull();
    if (isSchemaValidationAvailable()) expectSchemaValid(decode(created.data), 'pml');
    expect(
      getDrawingGuidesVisible(await loadPresentation(await savePresentation(presentation))),
    ).toBe(true);
    const before = created.data;
    expect(() => setDrawingGuidesVisible(presentation, 'yes' as unknown as boolean)).toThrow(
      'Invalid',
    );
    expect(created.data).toEqual(before);
  });
  it('retains guide metadata and unrelated extensions, and validates before mutation', () => {
    const presentation = createPresentation();
    setDrawingGuides(presentation, [{ id: 7, axis: 'x', position: 12345, color: '#888888' }]);
    const part = presentation[INTERNAL_PACKAGE].getPart(PRES_PART_NAME)!;
    part.data = encode(
      decode(part.data)
        .replace('userDrawn="1"', 'userDrawn="0" name="Custom guide"')
        .replace('</p15:guide>', '<p15:extLst><p:ext uri="custom"/></p15:extLst></p15:guide>'),
    );
    const guides = getDrawingGuides(presentation)!;
    expect(guides[0]!.position).toBe(12700);
    setDrawingGuides(presentation, [{ ...guides[0]!, color: '#c43c3c' }]);
    const xml = decode(part.data);
    expect(xml).toContain('name="Custom guide"');
    expect(xml).toContain('userDrawn="0"');
    expect(xml).toContain('uri="custom"');
    expect(xml).toContain('val="C43C3C"');
    expect(() => setDrawingGuides(presentation, [guides[0]!, guides[0]!])).toThrow('Invalid');
    expect(() => setDrawingGuides(presentation, [{ ...guides[0]!, position: NaN }])).toThrow(
      'Invalid',
    );
    expect(decode(part.data)).toBe(xml);
  });
});
