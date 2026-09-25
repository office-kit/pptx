import { describe, expect, it } from 'vitest';
import {
  addSlide,
  addSlideTextBox,
  createPresentation,
  findSlideLayout,
  getSlideShapes,
  getShapeText,
  getShapeXmlString,
  resetSlidePlaceholderTextFormatting,
  savePresentation,
  loadPresentation,
  getSlides,
  setShapeText,
  setShapeTextFormat,
  inches,
  groupShapes,
  setShapePosition,
  setShapeSize,
  getShapeBoundsResolved,
} from '../src/api/index.ts';
import { SHAPE_ELEMENT } from '../src/api/_internal-symbols.ts';
import { NS, firstChildElement, qname, parseXml } from '../src/internal/xml/index.ts';

describe('resetSlidePlaceholderTextFormatting', () => {
  it('restores inherited text formatting while preserving content, list level, links, language and extensions', async () => {
    const pres = createPresentation();
    const slide = addSlide(pres, { layout: findSlideLayout(pres, 'Title and Content')! });
    const body = getSlideShapes(slide)[1]!;
    const element = body[SHAPE_ELEMENT];
    const old = firstChildElement(element, qname('p', 'txBody', NS.pml))!;
    element.children[element.children.indexOf(old)] = parseXml(
      `<p:txBody xmlns:p="${NS.pml}" xmlns:a="${NS.dml}" xmlns:z="urn:test"><a:bodyPr anchor="b" lIns="1000"><a:normAutofit fontScale="80000"/><a:extLst><a:ext uri="body"><z:keep/></a:ext></a:extLst></a:bodyPr><a:lstStyle><a:lvl2pPr marL="999" algn="r"><a:defRPr sz="5000" lang="ja-JP"/></a:lvl2pPr></a:lstStyle><a:p><a:pPr lvl="1" algn="ctr" marL="888"><a:buNone/><a:defRPr b="1" lang="ja-JP"/><a:extLst><a:ext uri="para"><z:keep/></a:ext></a:extLst></a:pPr><a:r><a:rPr sz="4400" b="1" lang="ja-JP"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:hlinkClick action="ppaction://nextslide"/><a:extLst><a:ext uri="run"><z:keep/></a:ext></a:extLst></a:rPr><a:t>日本語 / English</a:t></a:r><a:br/><a:fld id="{00000000-0000-0000-0000-000000000001}" type="slidenum"><a:rPr sz="5000"/><a:t>1</a:t></a:fld><a:endParaRPr sz="5000"/></a:p></p:txBody>`,
    ).root;
    const decoration = addSlideTextBox(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(2),
      h: inches(1),
      text: 'Decoration',
    });
    setShapeTextFormat(decoration, { bold: true });
    const untouched = getShapeXmlString(decoration);
    const bounds = getShapeBoundsResolved(pres, body);
    expect(resetSlidePlaceholderTextFormatting(slide)).toBe(2);
    const xml = getShapeXmlString(body);
    for (const removed of [
      'anchor="b"',
      'lIns="1000"',
      'normAutofit',
      'marL=',
      'algn=',
      'buNone',
      'sz=',
      'b="1"',
      'solidFill',
    ])
      expect(xml).not.toContain(removed);
    for (const kept of [
      'lvl="1"',
      'lang="ja-JP"',
      'hlinkClick',
      'ppaction://nextslide',
      'urn:test',
      'uri="body"',
      'uri="para"',
      'uri="run"',
      'type="slidenum"',
    ])
      expect(xml).toContain(kept);
    expect(getShapeBoundsResolved(pres, body)).toEqual(bounds);
    expect(getShapeXmlString(decoration)).toBe(untouched);
    resetSlidePlaceholderTextFormatting(slide);
    expect(getShapeXmlString(body)).toBe(xml);
    const loaded = await loadPresentation(await savePresentation(pres));
    expect(getSlideShapes(getSlides(loaded)[0]!).map(getShapeText)).toEqual([
      '',
      '日本語 / English\n1',
      'Decoration',
    ]);
  });

  // A group scales and turns what is inside it; it does not give the text a
  // font. So a placeholder keeps its layout's formatting wherever it sits.
  it('resets a placeholder inside a group, leaving its place in the group alone', () => {
    const pres = createPresentation();
    const slide = addSlide(pres, { layout: findSlideLayout(pres, 'Title and Content')! });
    for (const shape of getSlideShapes(slide)) {
      setShapeText(shape, 'Keep');
      setShapeTextFormat(shape, { bold: true });
      const bounds = getShapeBoundsResolved(pres, shape)!;
      setShapePosition(shape, bounds.x, bounds.y);
      setShapeSize(shape, bounds.w, bounds.h);
    }
    const decoration = addSlideTextBox(slide, {
      x: inches(1),
      y: inches(4),
      w: inches(2),
      h: inches(1),
      text: 'Decoration',
    });
    setShapeTextFormat(decoration, { bold: true });
    groupShapes([...getSlideShapes(slide)]);
    const placed = getSlideShapes(slide).map((shape) => getShapeBoundsResolved(pres, shape));

    expect(resetSlidePlaceholderTextFormatting(slide)).toBe(2);

    const shapes = getSlideShapes(slide);
    const [group, title, body, plain] = shapes;
    expect(getShapeXmlString(title!)).not.toContain('b="1"');
    expect(getShapeXmlString(body!)).not.toContain('b="1"');
    // The text box beside them is not a placeholder and is not touched.
    expect(getShapeXmlString(plain!)).toContain('b="1"');
    expect(shapes.map(getShapeText)).toEqual(['', 'Keep', 'Keep', 'Decoration']);
    // Nothing moved: the group's own frame, and every child's place in it.
    expect(shapes.map((shape) => getShapeBoundsResolved(pres, shape))).toEqual(placed);
    void group;
  });
});
