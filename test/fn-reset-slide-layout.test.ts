import { describe, expect, it } from 'vitest';
import {
  addSlide,
  addSlideTextBox,
  addSlideImage,
  createPresentation,
  findSlideLayout,
  getSlideShapes,
  getShapeText,
  getShapeXmlString,
  getShapeBoundsResolved,
  getShapeFill,
  getShapeImageBytes,
  setShapeImageCrop,
  getShapeHyperlink,
  getShapeId,
  resetSlideLayout,
  savePresentation,
  loadPresentation,
  getSlides,
  setShapeText,
  setShapeTextFormat,
  setShapeFill,
  setShapeStroke,
  setShapePosition,
  setShapeSize,
  setShapeRotation,
  setShapeHyperlink,
  removeShape,
  inches,
} from '../src/api/index.ts';
import { commitSlideData } from '../src/api/fn/_helpers.ts';
import { SHAPE_ELEMENT } from '../src/api/_internal-symbols.ts';
import {
  NS,
  firstChildElement,
  qname,
  parseXml,
  serializeFragment,
} from '../src/internal/xml/index.ts';

const child = (shape: ReturnType<typeof getSlideShapes>[number], local: string) =>
  firstChildElement(shape[SHAPE_ELEMENT], qname('p', local, NS.pml))!;

describe('resetSlideLayout', () => {
  it('restores missing slots, geometry and formatting together while retaining content and other shapes', async () => {
    const pres = createPresentation();
    const slide = addSlide(pres, { layout: findSlideLayout(pres, 'Title and Content')! });
    const [title, body] = getSlideShapes(slide);
    const bounds = getShapeBoundsResolved(pres, title!);
    const id = getShapeId(title!);
    setShapeText(title!, '日本語 / English');
    setShapeTextFormat(title!, { size: 48, bold: true, color: '#FF0000' });
    setShapeFill(title!, '#00FF00');
    setShapeStroke(title!, { color: '#0000FF', widthEmu: 30000 });
    setShapePosition(title!, inches(3), inches(4));
    setShapeSize(title!, inches(2), inches(1));
    setShapeRotation(title!, 25);
    setShapeHyperlink(title!, 'https://example.com');
    child(title!, 'spPr').children.push(
      parseXml(
        `<a:extLst xmlns:a="${NS.dml}" xmlns:z="urn:test"><a:ext uri="keep"><z:metadata/></a:ext></a:extLst>`,
      ).root,
    );
    const element = title![SHAPE_ELEMENT];
    element.children.splice(
      element.children.indexOf(child(title!, 'txBody')),
      0,
      parseXml(
        `<p:style xmlns:p="${NS.pml}" xmlns:a="${NS.dml}"><a:lnRef idx="2"/><a:fillRef idx="2"/><a:effectRef idx="2"/><a:fontRef idx="minor"/></p:style>`,
      ).root,
    );
    removeShape(body!);
    const decoration = addSlideTextBox(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(2),
      h: inches(1),
      text: 'Keep',
    });
    setShapeFill(decoration, '#123456');
    const untouched = getShapeXmlString(decoration);
    expect(resetSlideLayout(slide)).toBe(2);
    const shapes = getSlideShapes(slide);
    const restored = shapes.find((shape) => getShapeId(shape) === id)!;
    expect(getShapeBoundsResolved(pres, restored)).toEqual(bounds);
    expect(getShapeFill(restored)).toEqual({ kind: 'inherit' });
    expect(getShapeHyperlink(restored)).toBe('https://example.com');
    expect(getShapeXmlString(restored)).toContain('urn:test');
    expect(getShapeXmlString(restored)).not.toContain('fontRef');
    expect(getShapeXmlString(restored)).not.toContain('sz="4800"');
    expect(getShapeXmlString(decoration)).toBe(untouched);
    const once = shapes.map(getShapeXmlString);
    expect(resetSlideLayout(slide)).toBe(2);
    expect(getSlideShapes(slide).map(getShapeXmlString)).toEqual(once);
    const loaded = await loadPresentation(await savePresentation(pres));
    expect(getSlideShapes(getSlides(loaded)[0]!).map(getShapeText)).toEqual([
      '日本語 / English',
      'Keep',
      '',
    ]);
  });

  it('retains a populated picture placeholder and its image relationship while resetting its bounds', async () => {
    const pres = createPresentation();
    const slide = addSlide(pres, { layout: findSlideLayout(pres, 'Title and Content')! });
    const body = getSlideShapes(slide)[1]!;
    const expected = getShapeBoundsResolved(pres, body);
    removeShape(body);
    const picture = addSlideImage(
      slide,
      Uint8Array.from(
        Buffer.from(
          'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVQIHWP4z8DwHwAFgAI/ScLttAAAAABJRU5ErkJggg==',
          'base64',
        ),
      ),
      { x: inches(3), y: inches(3), w: inches(1), h: inches(1) },
    );
    setShapeImageCrop(picture, { left: 0.1, top: 0.2, right: 0.15, bottom: 0.05 });
    const nv = firstChildElement(child(picture, 'nvPicPr'), qname('p', 'nvPr', NS.pml))!;
    nv.children.push(parseXml(`<p:ph xmlns:p="${NS.pml}" idx="1"/>`).root);
    // Commit the imported XML before reloading its placeholder binding.
    commitSlideData(slide);
    const loaded = await loadPresentation(await savePresentation(pres));
    const target = getSlides(loaded)[0]!;
    const image = getSlideShapes(target)[1]!;
    const content = serializeFragment(child(image, 'blipFill'));
    const imageBytes = getShapeImageBytes(image);
    expect(imageBytes).not.toBeNull();
    expect(content).toContain('l="10000"');
    expect(resetSlideLayout(target)).toBe(2);
    expect(getSlideShapes(target)).toHaveLength(2);
    expect(getShapeBoundsResolved(loaded, image)).toEqual(expected);
    expect(serializeFragment(child(image, 'blipFill'))).toBe(content);
    expect(getShapeImageBytes(image)).toEqual(imageBytes);
    const roundtrip = await loadPresentation(await savePresentation(loaded));
    expect(serializeFragment(child(getSlideShapes(getSlides(roundtrip)[0]!)[1]!, 'blipFill'))).toBe(
      content,
    );
  });
});
