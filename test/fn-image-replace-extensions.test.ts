import { expect, it } from 'vitest';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import * as pptx from '../src/api/index.ts';
import { buildPng } from './lib/build-png.ts';

const svgNs = 'http://schemas.microsoft.com/office/drawing/2016/SVG/main';
const sourceSvg = '<svg xmlns="http://www.w3.org/2000/svg"><rect width="1" height="1"/></svg>';

it.each([true, false])(
  'replacement clears stale SVG sources (isolated=%s) and preserves other extensions',
  async (isolated) => {
    const original = pptx.createPresentation();
    const slide = pptx.addBlankSlide(original);
    const png = buildPng(2, 1, [255, 0, 0]);
    const picture = pptx.addSlideImage(slide, png, {
      x: pptx.inches(1),
      y: pptx.inches(1),
      w: pptx.inches(2),
      h: pptx.inches(1),
    });
    pptx.setShapeImageCrop(picture, { left: 0.1 });
    pptx.copyShape(slide, picture);
    const zip = unzipSync(await pptx.savePresentation(original));
    // A nonstandard prefix ensures matching uses the namespace, not its spelling.
    const extensions = `<a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}"><vector:svgBlip xmlns:vector="${svgNs}" r:embed="rIdSVG"/></a:ext><a:ext uri="dpi"><a14:useLocalDpi xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" val="0"/></a:ext><a:ext uri="private"/></a:extLst>`;
    zip['ppt/slides/slide1.xml'] = strToU8(
      strFromU8(zip['ppt/slides/slide1.xml']!).replace(
        /<a:blip([^>]*?)\/>/g,
        `<a:blip$1>${extensions}</a:blip>`,
      ),
    );
    zip['ppt/slides/_rels/slide1.xml.rels'] = strToU8(
      strFromU8(zip['ppt/slides/_rels/slide1.xml.rels']!).replace(
        '</Relationships>',
        '<Relationship Id="rIdSVG" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/original.svg"/></Relationships>',
      ),
    );
    zip['ppt/media/original.svg'] = strToU8(sourceSvg);
    zip['[Content_Types].xml'] = strToU8(
      strFromU8(zip['[Content_Types].xml']!).replace(
        '</Types>',
        '<Default Extension="svg" ContentType="image/svg+xml"/></Types>',
      ),
    );
    const pres = await pptx.loadPresentation(zipSync(zip));
    const currentSlide = pptx.getSlides(pres)[0]!;
    const target = pptx.getSlideShapes(currentSlide)[0]!;
    expect(pptx.getSlideXmlString(currentSlide).match(/vector:svgBlip/g)).toHaveLength(2);
    const replacement = buildPng(1, 2, [0, 255, 0]);
    pptx.setShapeImage(target, replacement, { isolated });
    const saved = await pptx.savePresentation(pres);
    const output = unzipSync(saved);
    const xml = strFromU8(output['ppt/slides/slide1.xml']!);
    expect(xml.match(/vector:svgBlip/g)).toHaveLength(1);
    expect(xml.match(/useLocalDpi/g)).toHaveLength(2);
    expect(xml.match(/uri="private"/g)).toHaveLength(2);
    expect(strFromU8(output['ppt/media/original.svg']!)).toBe(sourceSvg);
    const loaded = await pptx.loadPresentation(saved);
    const [first, second] = pptx.getSlideShapes(pptx.getSlides(loaded)[0]!);
    expect(pptx.getShapeImageBytes(first!)).toEqual(replacement);
    expect(pptx.getShapeImageCrop(first!)).toEqual(pptx.getShapeImageCrop(second!));
    if (isolated) expect(pptx.getShapeImageBytes(second!)).toEqual(png);
  },
);
