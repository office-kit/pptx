import { describe, expect, it } from 'vitest';
import {
  addBlankSlide,
  addSlideTextBox,
  createPresentation,
  getParagraphLineSpacing,
  getParagraphPropertiesEffective,
  getSlideShapes,
  getSlides,
  inches,
  loadPresentation,
  savePresentation,
  setParagraphLineSpacing,
} from '../src/api/index.ts';
import { readZip, writeZip } from '../src/internal/opc/index.ts';
import { renderSlideToSvg } from '../packages/preview/src/index.ts';

describe.each(['paragraph', 'list style'] as const)('imported %s line spacing', (layer) => {
  it.each([
    ['100%', 1],
    ['150%', 1.5],
    ['250%', 2.5],
    ['150000', 1.5],
    ['1', 0.00001],
  ])('reads %s before and after saving', async (value, expected) => {
    const pres = createPresentation();
    const shape = addSlideTextBox(addBlankSlide(pres), {
      x: inches(1),
      y: inches(1),
      w: inches(5),
      h: inches(2),
      text: 'First\nSecond',
    });
    setParagraphLineSpacing(shape, 0, { kind: 'pct', value: 1.5 });
    const zip = readZip(await savePresentation(pres));
    const entry = zip.entries.find((part) => part.name === 'ppt/slides/slide1.xml')!;
    let xml = new TextDecoder().decode(entry.data);
    const spacing = `<a:lnSpc><a:spcPct val="${value}"/></a:lnSpc>`;
    expect(xml).toContain('<a:lnSpc><a:spcPct val="150000"/></a:lnSpc>');
    xml = xml.replace(
      '<a:lnSpc><a:spcPct val="150000"/></a:lnSpc>',
      layer === 'paragraph' ? spacing : '',
    );
    if (layer === 'list style') {
      expect(xml).toContain('<a:lstStyle/>');
      xml = xml.replace(
        '<a:lstStyle/>',
        `<a:lstStyle><a:lvl1pPr>${spacing}</a:lvl1pPr></a:lstStyle>`,
      );
    }
    const data = new TextEncoder().encode(xml);
    const imported = await loadPresentation(
      writeZip(zip.entries.map((part) => (part === entry ? { ...part, data } : part))),
    );
    const reloaded = await loadPresentation(await savePresentation(imported));
    const referenceData = new TextEncoder().encode(
      xml.replace(spacing, `<a:lnSpc><a:spcPct val="${expected * 100000}"/></a:lnSpc>`),
    );
    const reference = await loadPresentation(
      writeZip(
        zip.entries.map((part) => (part === entry ? { ...part, data: referenceData } : part)),
      ),
    );
    for (const current of [imported, reloaded]) {
      const box = getSlideShapes(getSlides(current)[0]!)[0]!;
      expect(getParagraphLineSpacing(box, 0)).toEqual(
        layer === 'paragraph' ? { kind: 'pct', value: expected } : null,
      );
      expect(getParagraphPropertiesEffective(current, box, 0).lineSpacing).toEqual({
        kind: 'pct',
        value: expected,
      });
      for (const textLayout of ['svg', 'foreignObject'] as const) {
        expect(renderSlideToSvg(current, getSlides(current)[0]!, { textLayout })).toBe(
          renderSlideToSvg(reference, getSlides(reference)[0]!, { textLayout }),
        );
      }
    }
  });
});
