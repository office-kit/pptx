import { readFile } from 'node:fs/promises';
import { describe, expect, it } from 'vitest';
import {
  addBlankSlide,
  addSlideShape,
  createPresentation,
  getGroupChildren,
  getShapeFillColorResolved,
  getShapeFillEffective,
  getShapeFillOpacity,
  getShapeGradientFillEffective,
  getSlideShapes,
  getSlides,
  groupShapes,
  inches,
  loadPresentation,
  savePresentation,
  ungroupShapes,
} from '../src/api/index.ts';
import { SHAPE_ELEMENT, SHAPE_SLIDE, type SlideShapeData } from '../src/api/_internal-symbols.ts';
import { commitSlideData } from '../src/api/fn/_helpers.ts';
import { NS, firstChildElement, parseXml, qname } from '../src/internal/xml/index.ts';

const fillNames = new Set(['noFill', 'solidFill', 'gradFill', 'pattFill', 'blipFill', 'grpFill']);
const setFill = (shape: SlideShapeData, xml: string | null) => {
  const element = shape[SHAPE_ELEMENT];
  const props = firstChildElement(
    element,
    qname('p', element.name.localName === 'grpSp' ? 'grpSpPr' : 'spPr', NS.pml),
  )!;
  props.children = props.children.filter(
    (child) => child.kind !== 'element' || !fillNames.has(child.name.localName),
  );
  if (xml) {
    const fill = parseXml(xml.replace('>', ` xmlns:a="${NS.dml}">`)).root;
    const position = props.children.findIndex(
      (child) =>
        child.kind === 'element' &&
        ['ln', 'effectLst', 'effectDag', 'scene3d', 'sp3d', 'extLst'].includes(
          child.name.localName,
        ),
    );
    props.children.splice(position < 0 ? props.children.length : position, 0, fill);
  }
};
const solid = (rgb: string) =>
  `<a:solidFill><a:srgbClr val="${rgb}"><a:alpha val="42000"/></a:srgbClr></a:solidFill>`;
const groupFill = `<a:grpFill xmlns:a="${NS.dml}"/>`;
const noFill = `<a:noFill xmlns:a="${NS.dml}"/>`;
const makeDeck = async () => {
  const pres = createPresentation();
  const slide = addBlankSlide(pres);
  const box = (name: string) =>
    addSlideShape(slide, {
      preset: 'rect',
      name,
      x: inches(1),
      y: inches(1),
      w: inches(2),
      h: inches(1),
    });
  const child = box('child');
  const style = parseXml(
    await readFile(
      new URL('./fixtures/native/quickstyle-colored-fill-accent1-style.xml', import.meta.url),
      'utf8',
    ),
  ).root;
  child[SHAPE_ELEMENT].children.push(style);
  const group = groupShapes([child, box('sibling')]);
  // Replace the fill choice, rather than constructing invalid solidFill + grpFill XML.
  const props = firstChildElement(child[SHAPE_ELEMENT], qname('p', 'spPr', NS.pml))!;
  props.children = props.children.filter(
    (node) => node.kind !== 'element' || !fillNames.has(node.name.localName),
  );
  props.children.push(parseXml(groupFill).root);
  return { pres, slide, child: getGroupChildren(group)[0]!, group, box };
};

describe('group inherited fills', () => {
  it('uses group paint before the child style and reads replaced paint after an earlier lookup', async () => {
    const { pres, child, group } = await makeDeck();
    setFill(group, solid('112233'));
    expect(getShapeFillEffective(pres, child)).toEqual({ kind: 'solid', color: '#112233' });
    expect(getShapeFillColorResolved(pres, child)).toBe('#112233');
    expect(getShapeFillOpacity(child, pres)).toBeCloseTo(0.42);
    setFill(group, solid('445566'));
    expect(getShapeFillColorResolved(pres, child)).toBe('#445566');
    commitSlideData(child[SHAPE_SLIDE]);
    const restored = await loadPresentation(await savePresentation(pres));
    const restoredChild = getGroupChildren(getSlideShapes(getSlides(restored)[0]!)[0]!)[0]!;
    expect(getShapeFillColorResolved(restored, restoredChild)).toBe('#445566');
    expect(getShapeFillOpacity(restoredChild, restored)).toBeCloseTo(0.42);
  });

  it('resolves gradient details through group inheritance and save/load', async () => {
    const { pres, child, group } = await makeDeck();
    setFill(
      group,
      '<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="112233"/></a:gs><a:gs pos="100000"><a:srgbClr val="445566"/></a:gs></a:gsLst><a:lin ang="5400000"/></a:gradFill>',
    );
    expect(getShapeFillEffective(pres, child)).toEqual({ kind: 'gradient' });
    expect(
      getShapeGradientFillEffective(pres, child)?.stops.map((stop) => stop.resolvedColor),
    ).toEqual(['#112233', '#445566']);
    commitSlideData(child[SHAPE_SLIDE]);
    const restored = await loadPresentation(await savePresentation(pres));
    const restoredChild = getGroupChildren(getSlideShapes(getSlides(restored)[0]!)[0]!)[0]!;
    expect(
      getShapeGradientFillEffective(restored, restoredChild)?.stops.map(
        (stop) => stop.resolvedColor,
      ),
    ).toEqual(['#112233', '#445566']);
  });

  it('only inherits a group gradient when the child explicitly selects grpFill', async () => {
    const { pres, child, group } = await makeDeck();
    setFill(
      group,
      '<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="112233"/></a:gs><a:gs pos="100000"><a:srgbClr val="445566"/></a:gs></a:gsLst><a:lin ang="0"/></a:gradFill>',
    );
    setFill(child, null);
    expect(getShapeFillEffective(pres, child).kind).toBe('solid');
    expect(getShapeGradientFillEffective(pres, child)).toBeNull();
    const style = firstChildElement(child[SHAPE_ELEMENT], qname('p', 'style', NS.pml))!;
    const fillRef = firstChildElement(style, qname('a', 'fillRef', NS.dml))!;
    fillRef.attrs[0] = { ...fillRef.attrs[0]!, value: '2' };
    setFill(group, solid('112233'));
    const props = firstChildElement(child[SHAPE_ELEMENT], qname('p', 'spPr', NS.pml))!;
    props.children.push(parseXml(groupFill).root);
    expect(getShapeGradientFillEffective(pres, child)).toBeNull();
  });

  it('follows nested grpFill, stops at noFill, and invalidates ancestry after ungrouping', async () => {
    const { pres, child, group, box } = await makeDeck();
    setFill(group, null);
    const props = firstChildElement(group[SHAPE_ELEMENT], qname('p', 'grpSpPr', NS.pml))!;
    props.children.push(parseXml(groupFill).root);
    const outer = groupShapes([group, box('outer sibling')]);
    setFill(outer, solid('778899'));
    expect(getShapeFillColorResolved(pres, child)).toBe('#778899');
    commitSlideData(child[SHAPE_SLIDE]);
    const restored = await loadPresentation(await savePresentation(pres));
    const restoredInner = getGroupChildren(getSlideShapes(getSlides(restored)[0]!)[0]!)[0]!;
    expect(getShapeFillColorResolved(restored, getGroupChildren(restoredInner)[0]!)).toBe(
      '#778899',
    );
    setFill(group, null);
    props.children.push(parseXml(noFill).root);
    expect(getShapeFillEffective(pres, child)).toEqual({ kind: 'none' });
    expect(getShapeFillColorResolved(pres, child)).toBeNull();
    ungroupShapes(group);
    expect(getShapeFillColorResolved(pres, child)).toBe('#778899');
    ungroupShapes(outer);
    expect(getShapeFillColorResolved(pres, child)).toBeNull();
  });
});
