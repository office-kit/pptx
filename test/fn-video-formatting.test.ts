import { describe, expect, it } from 'vitest';
import {
  addBlankSlide,
  addSlideMedia,
  createPresentation,
  getShapeMedia,
  getShapePreset,
  getSlideShapes,
  getSlides,
  getSlideXmlString,
  inches,
  resetShapeVideoFormatting,
  setShapePreset,
  loadPresentation,
  savePresentation,
} from '../src/api/index.ts';
import { SHAPE_ELEMENT } from '../src/api/_internal-symbols.ts';
import { NS, attr, elem, firstChildElement, qname } from '../src/internal/xml/index.ts';

const mp4 = new Uint8Array([
  0,
  0,
  0,
  0x18,
  ...Array.from('ftypmp42', (c) => c.charCodeAt(0)),
  0,
  0,
  0,
  0,
]);
const box = { x: inches(1), y: inches(1), w: inches(4), h: inches(2.25) };

describe('resetShapeVideoFormatting', () => {
  it('matches Video Format Reset while preserving media, crop, and unknown XML', async () => {
    const pres = createPresentation();
    const slide = addBlankSlide(pres);
    const video = addSlideMedia(slide, { kind: 'video', data: mp4, ...box });
    setShapePreset(video, 'ellipse');
    const spPr = firstChildElement(video[SHAPE_ELEMENT], qname('p', 'spPr', NS.pml))!;
    const blipFill = firstChildElement(video[SHAPE_ELEMENT], qname('p', 'blipFill', NS.pml))!;
    const blip = firstChildElement(blipFill, qname('a', 'blip', NS.dml))!;

    spPr.children.push(
      elem(qname('a', 'solidFill', NS.dml)),
      elem(qname('a', 'ln', NS.dml)),
      elem(qname('a', 'effectLst', NS.dml)),
      elem(qname('a', 'effectDag', NS.dml)),
      elem(qname('a', 'scene3d', NS.dml)),
      elem(qname('a', 'sp3d', NS.dml)),
      elem(qname('a', 'extLst', NS.dml), {
        children: [
          elem(qname('a', 'ext', NS.dml), { attrs: [attr(qname('', 'uri', ''), 'unknown')] }),
        ],
      }),
    );
    blip.children.push(elem(qname('a', 'grayscl', NS.dml)));
    const srcRect = elem(qname('a', 'srcRect', NS.dml), {
      attrs: [attr(qname('', 'l', ''), '12000')],
    });
    blipFill.children.splice(blipFill.children.indexOf(blip) + 1, 0, srcRect);

    resetShapeVideoFormatting(video);

    expect(getShapeMedia(video)?.kind).toBe('video');
    expect(getShapePreset(video)).toBe('rect');
    const xml = getSlideXmlString(slide);
    expect(xml).not.toContain('<a:solidFill');
    expect(xml).not.toContain('<a:ln');
    expect(xml).not.toContain('<a:effectLst');
    expect(xml).not.toContain('<a:effectDag');
    expect(xml).not.toContain('<a:scene3d');
    expect(xml).not.toContain('<a:sp3d');
    expect(xml).not.toContain('<a:grayscl');
    expect(xml).toContain('uri="unknown"');
    expect(xml).toContain('<a:srcRect l="12000"');

    const restored = await loadPresentation(await savePresentation(pres));
    const restoredSlide = getSlideShapes(getSlides(restored)[0]!)[0]!;
    const restoredXml = getSlideXmlString(getSlides(restored)[0]!);
    expect(getShapePreset(restoredSlide)).toBe('rect');
    expect(restoredXml).toContain('uri="unknown"');
    expect(restoredXml).toContain('<a:srcRect l="12000"');
  });

  it('rejects audio without changing its XML', () => {
    const pres = createPresentation();
    const slide = addBlankSlide(pres);
    const audio = addSlideMedia(slide, {
      kind: 'audio',
      data: new Uint8Array([0x49, 0x44, 0x33, 3, 0]),
      format: 'mp3',
      ...box,
    });
    const before = getSlideXmlString(slide);
    expect(getShapeMedia(audio)?.kind).toBe('audio');
    expect(() => resetShapeVideoFormatting(audio)).toThrow(/video shape/);
    expect(getSlideXmlString(slide)).toBe(before);
    expect(getSlideShapes(slide)).toHaveLength(1);
  });
});
