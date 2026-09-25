import {
  addSlide,
  createPresentation,
  findSlideLayout,
  findSlidePlaceholder,
  getShapeText,
  getSlides,
  loadPresentation,
  savePresentation,
  setShapeText,
} from '../src/api/index.ts';
import { LAYOUT_PART, SHAPE_ELEMENT } from '../src/api/_internal-symbols.ts';
import { describe, expect, it } from 'vitest';
import { buildSlideFromLayout } from '../src/internal/presentationml/slide-builder.ts';
import {
  NS,
  attr,
  firstChildElement,
  getAttrValue,
  parseXml,
  qname,
  serializeXml,
  walkElements,
} from '../src/internal/xml/index.ts';

function elements(root: ReturnType<typeof parseXml>['root']) {
  const result: (typeof root)[] = [];
  walkElements(root, (element) => {
    result.push(element);
  });
  return result;
}
function placeholders(root: ReturnType<typeof parseXml>['root']) {
  return elements(root).filter(
    (element) => element.name.namespaceURI === NS.pml && element.name.localName === 'ph',
  );
}

describe('new slide placeholder metadata', () => {
  it('preserves layout binding metadata through public add/edit/save/load operations', async () => {
    const pres = createPresentation();
    const layout = findSlideLayout(pres, 'Title and Content')!;
    const body = placeholders(layout[LAYOUT_PART].root).find(
      (ph) => getAttrValue(ph, qname('', 'idx', '')) === '1',
    )!;
    body.attrs.push(attr(qname('', 'orient', ''), 'vert'), attr(qname('', 'sz', ''), 'quarter'));
    const slide = addSlide(pres, { layout });
    const shape = findSlidePlaceholder(slide, 'body')!;
    setShapeText(shape, '日本語 / English');
    const loaded = await loadPresentation(await savePresentation(pres));
    const restored = findSlidePlaceholder(getSlides(loaded)[0]!, 'body')!;
    expect(getShapeText(restored)).toBe('日本語 / English');
    const ph = placeholders(restored[SHAPE_ELEMENT])[0]!;
    expect(getAttrValue(ph, qname('', 'orient', ''))).toBe('vert');
    expect(getAttrValue(ph, qname('', 'sz', ''))).toBe('quarter');
  });

  it.each(['half', 'quarter', 'full'])(
    'retains vertical orientation and %s size without freezing layout content',
    (size) => {
      const layout = parseXml(
        `<p:spTree xmlns:p="${NS.pml}" xmlns:a="${NS.dml}"><p:sp><p:nvSpPr><p:cNvPr id="8" name="縦書き / Vertical"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="7" orient="vert" sz="${size}" hasCustomPrompt="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Layout prompt / 入力案内</a:t></a:r></a:p></p:txBody></p:sp></p:spTree>`,
      );
      const before = serializeXml(layout);
      const built = buildSlideFromLayout(layout.root);
      const result = parseXml(serializeXml(built));
      const ph = placeholders(result.root)[0]!;
      for (const [name, value] of Object.entries({
        type: 'body',
        idx: '7',
        orient: 'vert',
        sz: size,
      })) {
        expect(getAttrValue(ph, qname('', name, ''))).toBe(value);
      }
      expect(getAttrValue(ph, qname('', 'hasCustomPrompt', ''))).toBeNull();
      expect(serializeXml(result)).not.toContain('Layout prompt');
      const spPr = elements(result.root).find((element) => element.name.localName === 'spPr')!;
      expect(firstChildElement(spPr, qname('a', 'xfrm', NS.dml))).toBeNull();
      ph.attrs.length = 0;
      expect(serializeXml(layout)).toBe(before);
    },
  );

  it('keeps omitted placeholder defaults omitted', () => {
    const layout = parseXml(
      `<p:spTree xmlns:p="${NS.pml}"><p:sp><p:nvSpPr><p:nvPr><p:ph/></p:nvPr></p:nvSpPr></p:sp></p:spTree>`,
    );
    expect(placeholders(buildSlideFromLayout(layout.root).root)[0]!.attrs).toEqual([]);
  });
});
