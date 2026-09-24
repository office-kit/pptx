import { expect, it } from 'vitest';
import * as pptx from '../src/api/index.ts';
import { findCNvPr } from '../src/api/fn/shape-click-action.ts';
import { attr, elem, firstChildElement, NS, qname } from '../src/internal/xml/index.ts';

const bounds = { x: pptx.inches(1), y: pptx.inches(1), w: pptx.inches(2), h: pptx.inches(1) };

it('keeps hover and click destinations independent through save, replacement and removal', async () => {
  const p = pptx.createPresentation();
  const slide = pptx.addBlankSlide(p);
  const destination = pptx.addBlankSlide(p);
  const shape = pptx.addSlideShape(slide, { ...bounds, preset: 'actionButtonForwardNext' });
  findCNvPr(shape)!.children.push(elem(qname('a', 'extLst', NS.dml)));
  pptx.setShapeHoverAction(shape, { kind: 'slide', slide: destination }, 'Hover destination');
  pptx.setShapeClickAction(
    shape,
    { kind: 'url', url: 'https://example.com/click' },
    'Click destination',
  );
  const saved = await pptx.loadPresentation(await pptx.savePresentation(p));
  const loadedSlide = pptx.getSlides(saved)[0]!;
  const loaded = pptx.getSlideShapes(loadedSlide)[0]!;
  expect(pptx.getShapeClickAction(loaded)).toEqual({
    kind: 'url',
    url: 'https://example.com/click',
  });
  const hover = pptx.getShapeHoverAction(loaded);
  expect(hover?.kind).toBe('slide');
  if (hover?.kind === 'slide')
    expect(pptx.getSlidePartName(hover.slide)).toBe(
      pptx.getSlidePartName(pptx.getSlides(saved)[1]!),
    );
  expect(pptx.getShapeHoverActionTooltip(loaded)).toBe('Hover destination');
  expect(pptx.getShapeClickActionTooltip(loaded)).toBe('Click destination');
  let xml = pptx.getSlideXmlString(loadedSlide);
  expect(xml.indexOf('hlinkClick')).toBeLessThan(xml.indexOf('hlinkHover'));
  expect(xml.indexOf('hlinkHover')).toBeLessThan(xml.indexOf('a:extLst'));
  pptx.setShapeHoverAction(loaded, { kind: 'endShow' });
  expect(pptx.getShapeHoverAction(loaded)).toEqual({ kind: 'endShow' });
  expect(pptx.getShapeHoverActionTooltip(loaded)).toBeNull();
  pptx.setShapeClickAction(loaded, null);
  expect(pptx.getShapeHoverAction(loaded)).toEqual({ kind: 'endShow' });
  pptx.setShapeHoverAction(loaded, null);
  xml = pptx.getSlideXmlString(loadedSlide);
  expect(xml).not.toContain('hlinkHover');
  expect(xml).not.toContain('hlinkClick');
  expect(xml).toContain('a:extLst');
});

it.each([
  'nextSlide',
  'prevSlide',
  'firstSlide',
  'lastSlide',
  'lastSlideViewed',
  'endShow',
] as const)(
  'round-trips hover preset %s on groups and tables without changing child actions',
  async (kind) => {
    const p = pptx.createPresentation();
    const slide = pptx.addBlankSlide(p);
    const first = pptx.addSlideShape(slide, { ...bounds, preset: 'rect' });
    const second = pptx.addSlideShape(slide, { ...bounds, preset: 'ellipse' });
    pptx.setShapeClickAction(first, { kind: 'nextSlide' });
    const group = pptx.groupShapes([first, second]);
    const table = pptx.addSlideTable(slide, { ...bounds, rows: [['Cell']] });
    for (const shape of [group, table]) pptx.setShapeHoverAction(shape, { kind }, 'Tip');
    const saved = await pptx.loadPresentation(await pptx.savePresentation(p));
    const shapes = pptx.getSlideShapes(pptx.getSlides(saved)[0]!);
    const targets = shapes.filter(
      (shape) => pptx.getShapeKind(shape) === 'group' || pptx.isTableShape(shape),
    );
    expect(targets).toHaveLength(2);
    for (const shape of targets) {
      expect(pptx.getShapeHoverAction(shape)).toEqual({ kind });
      expect(pptx.getShapeClickAction(shape)).toBeNull();
      expect(pptx.getShapeHoverActionTooltip(shape)).toBe('Tip');
    }
    const loadedGroup = targets.find((shape) => pptx.getShapeKind(shape) === 'group')!;
    const children = pptx.getGroupChildren(loadedGroup);
    expect(children.every((shape) => pptx.getShapeHoverAction(shape) === null)).toBe(true);
    expect(pptx.getShapeClickAction(children[0]!)).toEqual({
      kind: 'nextSlide',
    });
  },
);

it('rejects foreign slide destinations before changing hover metadata or relationships', async () => {
  const p = pptx.createPresentation();
  const slide = pptx.addBlankSlide(p);
  const shape = pptx.addSlideShape(slide, { ...bounds, preset: 'rect' });
  pptx.setShapeHoverAction(shape, { kind: 'url', url: 'https://example.com/hover' }, 'Keep');
  const other = pptx.addBlankSlide(pptx.createPresentation());
  const before = pptx.getSlideXmlString(slide);
  const rels = pptx.readPackagePart(p, '/ppt/slides/_rels/slide1.xml.rels');
  expect(() => pptx.setShapeHoverAction(shape, { kind: 'slide', slide: other })).toThrow(
    'target slide must belong',
  );
  expect(pptx.getSlideXmlString(slide)).toBe(before);
  expect(pptx.readPackagePart(p, '/ppt/slides/_rels/slide1.xml.rels')).toEqual(rels);
  const loaded = await pptx.loadPresentation(await pptx.savePresentation(p));
  expect(pptx.getShapeHoverAction(pptx.getSlideShapes(pptx.getSlides(loaded)[0]!)[0]!)).toEqual({
    kind: 'url',
    url: 'https://example.com/hover',
  });
});

it.each(['click', 'hover'] as const)(
  'retains sound and extension metadata when changing the %s destination',
  async (trigger) => {
    const p = pptx.createPresentation();
    const slide = pptx.addBlankSlide(p);
    const shape = pptx.addSlideShape(slide, { ...bounds, preset: 'rect' });
    const set = trigger === 'click' ? pptx.setShapeClickAction : pptx.setShapeHoverAction;
    set(shape, { kind: 'url', url: 'https://example.com/old' }, 'Old');
    const link = firstChildElement(
      findCNvPr(shape)!,
      qname('a', trigger === 'click' ? 'hlinkClick' : 'hlinkHover', NS.dml),
    )!;
    link.attrs.push(
      attr(qname('', 'endSnd', ''), '1'),
      attr(qname('', 'history', ''), '0'),
      attr(qname('', 'invalidUrl', ''), 'old broken link'),
    );
    link.children.push(
      elem(qname('a', 'snd', NS.dml), {
        attrs: [
          attr(qname('r', 'embed', NS.officeDocRels), 'rIdAudio'),
          attr(qname('', 'name', ''), 'Chime'),
        ],
      }),
    );
    link.children.push(
      elem(qname('a', 'extLst', NS.dml), {
        children: [
          elem(qname('a', 'ext', NS.dml), {
            attrs: [attr(qname('', 'uri', ''), 'test-extension')],
          }),
        ],
      }),
    );
    set(shape, { kind: 'nextSlide' }, 'New');
    const saved = await pptx.loadPresentation(await pptx.savePresentation(p));
    const loadedSlide = pptx.getSlides(saved)[0]!;
    const xml = pptx.getSlideXmlString(loadedSlide);
    expect(xml).toContain('endSnd="1"');
    expect(xml).toContain('history="0"');
    expect(xml).toContain('r:embed="rIdAudio"');
    expect(xml).toContain('test-extension');
    expect(xml).toContain('tooltip="New"');
    expect(xml).not.toContain('invalidUrl=');
    set(pptx.getSlideShapes(loadedSlide)[0]!, null);
    expect(pptx.getSlideXmlString(loadedSlide)).not.toContain('rIdAudio');
  },
);
