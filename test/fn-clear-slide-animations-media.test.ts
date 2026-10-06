import { describe, expect, it } from 'vitest';
import { SLIDE_DOCUMENT } from '../src/api/_internal-symbols.ts';
import {
  addBlankSlide,
  addSlideMedia,
  addSlideShape,
  clearSlideAnimations,
  createPresentation,
  getShapeId,
  getSlides,
  inches,
  loadPresentation,
  savePresentation,
  validatePresentation,
} from '../src/api/index.ts';
import { parseXml, serializeXml } from '../src/internal/xml/index.ts';
import { expectSchemaValid } from './lib/expect-schema-valid.ts';

const nestedMediaSlide = (list = 'childTnLst', reference = '1') => {
  const pres = createPresentation();
  const slide = addBlankSlide(pres);
  const box = { x: inches(1), y: inches(1), w: inches(2), h: inches(1) };
  const clip = addSlideMedia(slide, {
    kind: 'audio',
    data: new Uint8Array([73, 68, 51, 3, 0, 0, 0, 0, 0, 0, 255, 251]),
    ...box,
  });
  const shape = addSlideShape(slide, { preset: 'rect', ...box });
  const conditions =
    '<p:stCondLst><p:cond evt="onClick" delay="500"/></p:stCondLst>' +
    `<p:endSync evt="onEnd" delay="0"><p:tn val="${reference}"/></p:endSync>` +
    '<p:iterate type="el"><p:tmAbs val="250"/></p:iterate>';
  const media =
    '<p:audio><p:cMediaNode vol="25000" mute="1">' +
    `<p:cTn id="4" fill="hold"/><p:tgtEl><p:spTgt spid="${getShapeId(clip)}"/></p:tgtEl>` +
    '</p:cMediaNode></p:audio>';
  const timing = parseXml(
    '<p:timing xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
      '<p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">' +
      '<p:childTnLst><p:seq><p:cTn id="2" dur="indefinite" nodeType="mainSeq">' +
      conditions +
      '<p:childTnLst>' +
      `<p:par><p:cTn id="3" fill="hold"><p:${list}>${media}</p:${list}></p:cTn></p:par>` +
      '<p:par><p:cTn id="5" presetID="1" presetClass="entr"><p:childTnLst>' +
      `<p:set><p:cBhvr><p:cTn id="6"/><p:tgtEl><p:spTgt spid="${getShapeId(shape)}"/></p:tgtEl></p:cBhvr></p:set>` +
      '</p:childTnLst></p:cTn></p:par>' +
      '</p:childTnLst></p:cTn></p:seq></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>',
  ).root;
  const root = slide[SLIDE_DOCUMENT].root;
  root.children = root.children.filter(
    (c) => !(c.kind === 'element' && c.name.localName === 'timing'),
  );
  root.children.push(timing);
  return { pres, slide, conditions, media };
};

describe('clearSlideAnimations: nested media timing', () => {
  it.each(['childTnLst', 'subTnLst'])(
    'preserves %s media, ancestry and metadata through save/reload',
    async (list) => {
      const { pres, slide, conditions, media } = nestedMediaSlide(list);
      expectSchemaValid(serializeXml(slide[SLIDE_DOCUMENT]), 'pml');
      clearSlideAnimations(slide);
      const xml = serializeXml(slide[SLIDE_DOCUMENT]);
      expect(xml).toContain(`<p:${list}>${media}</p:${list}>`);
      expect(xml).toContain(conditions);
      expect(xml).not.toContain('presetID="1"');
      expect(xml).not.toContain('<p:set>');
      expectSchemaValid(xml, 'pml');
      const loaded = await loadPresentation(await savePresentation(pres));
      expect(serializeXml(getSlides(loaded)[0]![SLIDE_DOCUMENT])).toBe(xml);
      expect(validatePresentation(loaded).filter((issue) => issue.severity === 'error')).toEqual(
        [],
      );
    },
  );

  it('refuses a dangling media timing dependency before modifying the slide', () => {
    const { pres, slide } = nestedMediaSlide('childTnLst', '6');
    const before = serializeXml(slide[SLIDE_DOCUMENT]);
    expect(() => clearSlideAnimations(slide)).toThrow(/tim|refer|depend/i);
    expect(serializeXml(slide[SLIDE_DOCUMENT])).toBe(before);
    expect(getSlides(pres)[0]).toBe(slide);
  });
});
