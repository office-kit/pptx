import { readFile } from 'node:fs/promises';
import { INTERNAL_PACKAGE, LAYOUT_PART_NAME } from '../src/api/_internal-symbols.ts';
import {
  NS,
  firstChildElement,
  qname,
  parseXml,
  serializeXml,
  attr,
} from '../src/internal/xml/index.ts';
import { readShapeTreeFromCsldRoot, REL_TYPES } from '../src/internal/presentationml/index.ts';
import { resolveTarget } from '../src/internal/opc/index.ts';
import { setPosition, setSize, setRotation, setFlip } from '../src/internal/drawingml/index.ts';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  addSlideTextBox,
  createPresentation,
  findSlideLayout,
  getShapeBoundsResolved,
  getShapeFlip,
  getShapeId,
  getShapeRotation,
  getShapeText,
  getShapeXmlString,
  getSlideShapes,
  getSlides,
  inches,
  isShapePlaceholder,
  loadPresentation,
  resetSlidePlaceholderGeometry,
  savePresentation,
  setShapeFlip,
  setShapePosition,
  setShapeRotation,
  setShapeSize,
  setShapeText,
  setShapeTextFormat,
  setShapeHyperlink,
  getShapeHyperlink,
  getShapeParagraphElements,
  setSlideLayout,
  groupShapes,
} from '../src/api/index.ts';

describe('resetSlidePlaceholderGeometry', () => {
  it('restores layout geometry while preserving content and unrelated shapes through save/load', async () => {
    const pres = createPresentation();
    const layout = findSlideLayout(pres, 'Title and Content')!;
    const slide = addSlide(pres, { layout });
    const placeholders = getSlideShapes(slide).filter(isShapePlaceholder);
    const expected = placeholders.map((shape) => getShapeBoundsResolved(pres, shape));
    const ids = placeholders.map(getShapeId);
    for (const shape of placeholders) {
      setShapeText(shape, '日本語 / English');
      setShapeTextFormat(shape, { bold: true, size: 22, color: '#125678' });
      setShapeHyperlink(shape, 'https://example.com/slides');
      setShapePosition(shape, inches(3), inches(4));
      setShapeSize(shape, inches(2), inches(1));
      setShapeRotation(shape, 35);
      setShapeFlip(shape, { horizontal: true, vertical: true });
    }
    const textBefore = placeholders.map((shape) => getShapeParagraphElements(shape, 0));
    const decoration = addSlideTextBox(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(2),
      h: inches(1),
      text: 'Keep me',
    });
    const untouched = getShapeXmlString(decoration);
    expect(resetSlidePlaceholderGeometry(slide)).toBe(placeholders.length);
    expect(getShapeXmlString(decoration)).toBe(untouched);
    const loaded = await loadPresentation(await savePresentation(pres));
    const restored = getSlideShapes(getSlides(loaded)[0]!).filter(isShapePlaceholder);
    expect(restored.map(getShapeId)).toEqual(ids);
    expect(restored.map((shape) => getShapeBoundsResolved(loaded, shape))).toEqual(expected);
    expect(restored.map(getShapeRotation)).toEqual(placeholders.map(() => 0));
    expect(restored.map(getShapeFlip)).toEqual(
      placeholders.map(() => ({ horizontal: false, vertical: false })),
    );
    expect(restored.map(getShapeText)).toEqual(placeholders.map(() => '日本語 / English'));
    expect(restored.map((shape) => getShapeParagraphElements(shape, 0))).toEqual(textBefore);
    expect(restored.map(getShapeHyperlink)).toEqual(
      placeholders.map(() => 'https://example.com/slides'),
    );
  });

  it.each(['layout', 'master'])(
    'restores rotation and flips from the %s, matching each content slot by index',
    async (source) => {
      const pres = await loadPresentation(
        await readFile(new URL('./fixtures/minimal/blank.pptx', import.meta.url)),
      );
      const layout = findSlideLayout(pres, 'Two Content')!;
      const pkg = pres[INTERNAL_PACKAGE];
      const layoutPart = pkg.getPart(layout[LAYOUT_PART_NAME])!;
      const layoutDoc = parseXml(new TextDecoder().decode(layoutPart.data));
      const slots = readShapeTreeFromCsldRoot(layoutDoc.root, 'sldLayout').shapes.filter(
        (shape) => shape.placeholderIdx === 1 || shape.placeholderIdx === 2,
      );
      expect(slots).toHaveLength(2);
      const expected = slots.map((_, index) => ({
        x: inches(index + 1),
        y: inches(2),
        w: inches(3),
        h: inches(4),
      }));
      for (const [index, slot] of slots.entries()) {
        if (source === 'layout') {
          setPosition(slot.element, slot.kind, expected[index]!.x, expected[index]!.y);
          setSize(slot.element, slot.kind, expected[index]!.w, expected[index]!.h);
          setRotation(slot.element, slot.kind, 20 + index * 10);
          setFlip(slot.element, slot.kind, { horizontal: true, vertical: false });
        } else {
          const spPr = firstChildElement(slot.element, qname('p', 'spPr', NS.pml))!;
          spPr.children = spPr.children.filter(
            (node) => !(node.kind === 'element' && node.name.localName === 'xfrm'),
          );
        }
      }
      layoutPart.data = new TextEncoder().encode(serializeXml(layoutDoc));
      if (source === 'master') {
        const rel = pkg
          .getRels(layout[LAYOUT_PART_NAME])!
          .items.find((item) => item.type === REL_TYPES.slideMaster)!;
        const part = pkg.getPart(resolveTarget(layout[LAYOUT_PART_NAME], rel.target))!;
        const doc = parseXml(new TextDecoder().decode(part.data));
        const body = readShapeTreeFromCsldRoot(doc.root, 'sldMaster').shapes.find(
          (shape) => shape.placeholderType === 'body',
        )!;
        setPosition(body.element, body.kind, inches(5), inches(1));
        setSize(body.element, body.kind, inches(2), inches(3));
        setRotation(body.element, body.kind, 45);
        setFlip(body.element, body.kind, { horizontal: false, vertical: true });
        // Master matching uses type, not the layout's unrelated slot indices.
        const nv = firstChildElement(body.element, qname('p', 'nvSpPr', NS.pml))!;
        const ph = firstChildElement(
          firstChildElement(nv, qname('p', 'nvPr', NS.pml))!,
          qname('p', 'ph', NS.pml),
        )!;
        ph.attrs = ph.attrs.filter((item) => item.name.localName !== 'idx');
        ph.attrs.push(attr(qname('', 'idx', ''), '99'));
        part.data = new TextEncoder().encode(serializeXml(doc));
      }
      const slide = addSlide(pres, { layout: findSlideLayout(pres, 'Two Content')! });
      // Use slot order from the freshly loaded layout to avoid relying on cloned IDs.
      const content = getSlideShapes(slide).filter(
        (shape) =>
          getShapeXmlString(shape).includes('idx="1"') ||
          getShapeXmlString(shape).includes('idx="2"'),
      );
      expect(content).toHaveLength(2);
      for (const shape of content) setShapePosition(shape, inches(8), inches(6));
      resetSlidePlaceholderGeometry(slide);
      for (const [index, shape] of content.entries()) {
        expect(getShapeBoundsResolved(pres, shape)).toEqual(
          source === 'layout'
            ? expected[index]
            : { x: inches(5), y: inches(1), w: inches(2), h: inches(3) },
        );
        expect(getShapeRotation(shape)).toBe(source === 'layout' ? 20 + index * 10 : 45);
        expect(getShapeFlip(shape)).toEqual(
          source === 'layout'
            ? { horizontal: true, vertical: false }
            : { horizontal: false, vertical: true },
        );
      }
      const loaded = await loadPresentation(await savePresentation(pres));
      expect(getSlideShapes(getSlides(loaded)[0]!).map(getShapeRotation)).toEqual(
        getSlideShapes(slide).map(getShapeRotation),
      );
    },
  );

  it('preserves grouped placeholders in their local coordinate system', () => {
    const pres = createPresentation();
    const slide = addSlide(pres, { layout: findSlideLayout(pres, 'Title and Content')! });
    const shapes = [...getSlideShapes(slide)];
    for (const shape of shapes) {
      const bounds = getShapeBoundsResolved(pres, shape)!;
      setShapePosition(shape, bounds.x, bounds.y);
      setShapeSize(shape, bounds.w, bounds.h);
    }
    groupShapes(shapes);
    const before = getSlideShapes(slide).map(getShapeXmlString);
    expect(resetSlidePlaceholderGeometry(slide)).toBe(0);
    expect(getSlideShapes(slide).map(getShapeXmlString)).toEqual(before);
  });

  it('leaves unmatched placeholders intact after switching to a blank layout', () => {
    const pres = createPresentation();
    const slide = addSlide(pres, { layout: findSlideLayout(pres, 'Title and Content')! });
    setSlideLayout(slide, findSlideLayout(pres, 'Blank')!);
    const before = getSlideShapes(slide).map(getShapeXmlString);
    expect(resetSlidePlaceholderGeometry(slide)).toBe(0);
    expect(getSlideShapes(slide).map(getShapeXmlString)).toEqual(before);
  });
});
