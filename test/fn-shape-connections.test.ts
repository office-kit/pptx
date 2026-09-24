import { describe, it, expect } from 'vitest';
import {
  createPresentation,
  addBlankSlide,
  addSlideShape,
  addSlideLine,
  emu,
  getShapeConnectionSites,
  getShapeConnection,
  setShapeConnection,
  getShapeId,
  getSlideXmlString,
  savePresentation,
  loadPresentation,
  getSlides,
  getSlideShapes,
  setShapeAdjustValues,
} from '../src/api/index.ts';
import { presetConnectionSites } from '../src/internal/drawingml/preset-connection-sites.ts';
import { expectSchemaValid, isSchemaValidationAvailable } from './lib/expect-schema-valid.ts';

const setup = () => {
  const presentation = createPresentation();
  const slide = addBlankSlide(presentation);
  const target = addSlideShape(slide, {
    x: emu(0),
    y: emu(0),
    w: emu(2000000),
    h: emu(1000000),
    preset: 'rect',
  });
  const line = addSlideLine(slide, {
    from: { x: emu(0), y: emu(0) },
    to: { x: emu(3000000), y: emu(1000000) },
  });
  return { presentation, slide, target, line };
};
describe('native connector attachments', () => {
  it('uses preset-specific indexed sites and authored adjustment guides', () => {
    const { slide, target } = setup();
    expect(getShapeConnectionSites(target)).toEqual([
      { x: 1000000, y: 0, angle: 270 },
      { x: 0, y: 500000, angle: 180 },
      { x: 1000000, y: 1000000, angle: 90 },
      { x: 2000000, y: 500000, angle: 0 },
    ]);
    for (const [preset, xml] of Object.entries(presetConnectionSites)) {
      const shape = addSlideShape(slide, {
        x: emu(0),
        y: emu(0),
        w: emu(2000000),
        h: emu(1000000),
        preset,
      });
      const sites = getShapeConnectionSites(shape);
      expect(sites.length, preset).toBe((xml.match(/<a:cxn /g) ?? []).length);
      expect(
        sites.every(
          (s) => Number.isFinite(s.x) && Number.isFinite(s.y) && Number.isFinite(s.angle),
        ),
        preset,
      ).toBe(true);
    }
    const triangle = addSlideShape(slide, {
      x: emu(0),
      y: emu(0),
      w: emu(2000000),
      h: emu(1000000),
      preset: 'triangle',
    });
    setShapeAdjustValues(triangle, { adj: 25000 });
    expect(getShapeConnectionSites(triangle)[0]?.x).toBe(500000);
  });
  it('round-trips, preserves schema order, detaches and rejects invalid targets', async () => {
    const { presentation, slide, target, line } = setup();
    const connection = { shapeId: getShapeId(target), siteIndex: 3 };
    setShapeConnection(line, 'end', connection);
    setShapeConnection(line, 'start', { ...connection, siteIndex: 1 });
    expect(getSlideXmlString(slide)).toMatch(/<a:stCxn[^>]+\/>\s*<a:endCxn/);
    const saved = await loadPresentation(await savePresentation(presentation));
    const reloaded = getSlideShapes(getSlides(saved)[0]!)[1]!;
    expect(getShapeConnection(reloaded, 'end')).toEqual(connection);
    setShapeConnection(reloaded, 'start', null);
    expect(getShapeConnection(reloaded, 'start')).toBeNull();
    expect(getShapeConnection(reloaded, 'end')).toEqual(connection);
    expect(() =>
      setShapeConnection(line, 'start', { shapeId: getShapeId(line), siteIndex: 0 }),
    ).toThrow();
    expect(() => setShapeConnection(line, 'start', { ...connection, siteIndex: -1 })).toThrow();
    expect(() => setShapeConnection(target, 'start', connection)).toThrow();
    if (isSchemaValidationAvailable()) expectSchemaValid(getSlideXmlString(slide), 'pml');
  });
});
