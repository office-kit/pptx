import { describe, expect, it } from 'vitest';
import { SLIDE_DOCUMENT } from '../src/api/_internal-symbols.ts';
import { addBlankSlide, createPresentation } from '../src/api/index.ts';
import { removeMediaTimingNodes } from '../src/api/fn/_media-timing.ts';
import { parseXml, serializeXml } from '../src/internal/xml/index.ts';
import { expectSchemaValid } from './lib/expect-schema-valid.ts';

const timingSlide = (nodes: string) => {
  const slide = addBlankSlide(createPresentation());
  const timing = parseXml(
    `<p:timing xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:tnLst>${nodes}</p:tnLst></p:timing>`,
  ).root;
  slide[SLIDE_DOCUMENT].root.children.push(timing);
  expectSchemaValid(serializeXml(slide[SLIDE_DOCUMENT]), 'pml');
  return slide;
};

const media = (kind: 'audio' | 'video', id: number): string =>
  `<p:${kind}><p:cMediaNode><p:cTn id="${id + 10}"/>` +
  `<p:tgtEl><p:spTgt spid="${id}"/></p:tgtEl></p:cMediaNode></p:${kind}>`;

const condition = '<p:stCondLst><p:cond evt="onClick" delay="0"/></p:stCondLst>';

describe('nested media timing removal', () => {
  it.each(['childTnLst', 'subTnLst'])(
    'removes multiple media-only %s lists while retaining their parent conditions',
    (list) => {
      const slide = timingSlide(
        `<p:par><p:cTn id="1"><p:childTnLst>` +
          `<p:seq><p:cTn id="2">${condition}<p:${list}>${media('audio', 21)}</p:${list}></p:cTn></p:seq>` +
          `<p:par><p:cTn id="3"><p:${list}>${media('video', 22)}</p:${list}></p:cTn></p:par>` +
          `</p:childTnLst></p:cTn></p:par>`,
      );
      removeMediaTimingNodes(slide, new Set([21, 22]));
      const xml = serializeXml(slide[SLIDE_DOCUMENT]);
      expect(xml).not.toContain('<p:audio');
      expect(xml).not.toContain('<p:video');
      expect(xml).toContain(condition);
      expect(xml).toContain('<p:cTn id="3"/>');
      expectSchemaValid(xml, 'pml');
    },
  );

  it('preserves other root nodes when the standard root media list is emptied', () => {
    const otherRoot = `<p:seq><p:cTn id="2">${condition}</p:cTn></p:seq>`;
    const slide = timingSlide(
      `<p:par><p:cTn id="1"><p:childTnLst>${media('audio', 21)}</p:childTnLst></p:cTn></p:par>` +
        otherRoot,
    );
    removeMediaTimingNodes(slide, new Set([21]));
    const xml = serializeXml(slide[SLIDE_DOCUMENT]);
    expect(xml).toContain(otherRoot);
    expect(xml).not.toContain('<p:audio');
    expectSchemaValid(xml, 'pml');
  });

  it('removes a directly rooted media node without deleting unrelated timing', () => {
    const otherRoot = `<p:seq><p:cTn id="2">${condition}</p:cTn></p:seq>`;
    const slide = timingSlide(media('video', 21) + otherRoot);
    removeMediaTimingNodes(slide, new Set([21]));
    const xml = serializeXml(slide[SLIDE_DOCUMENT]);
    expect(xml).toContain(otherRoot);
    expect(xml).not.toContain('<p:video');
    expectSchemaValid(xml, 'pml');
  });

  it('removes an emptied top-level node list and leaves schema-valid timing', () => {
    const slide = timingSlide(media('audio', 21));
    removeMediaTimingNodes(slide, new Set([21]));
    const xml = serializeXml(slide[SLIDE_DOCUMENT]);
    expect(xml).not.toContain('<p:tnLst');
    expect(xml).not.toContain('<p:audio');
    expectSchemaValid(xml, 'pml');
  });

  it('leaves a tree unchanged when no media target matches', () => {
    const slide = timingSlide(
      `<p:seq><p:cTn id="1">${condition}<p:childTnLst>${media('video', 21)}</p:childTnLst></p:cTn></p:seq>`,
    );
    const before = serializeXml(slide[SLIDE_DOCUMENT]);
    removeMediaTimingNodes(slide, new Set([22]));
    expect(serializeXml(slide[SLIDE_DOCUMENT])).toBe(before);
  });
});
