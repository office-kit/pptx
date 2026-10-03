import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideShape,
  getShapeFillEffective,
  getShapeStrokeEffective,
  getShapeStrokeColorResolved,
  getShapeGradientFillEffective,
  getSlideShapes,
  getSlides,
  inches,
  loadPresentation,
  savePresentation,
} from '../src/api/index.ts';
import { readZip, writeZip } from '../src/internal/opc/index.ts';

const fixturePath = fileURLToPath(new URL('./fixtures/minimal/two-slides.pptx', import.meta.url));
const decoder = new TextDecoder();
const encoder = new TextEncoder();

const secondMasterStyle =
  '<p:style xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
  '<a:lnRef idx="2"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef>' +
  '<a:effectRef idx="0"/><a:fontRef idx="minor"/></p:style>';
const secondMasterGradientStyle = secondMasterStyle.replace(
  '<a:fillRef idx="1">',
  '<a:fillRef idx="2">',
);

const secondThemeFillStyles =
  '<a:fillStyleLst>' +
  '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' +
  '<a:gradFill><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"/></a:gs>' +
  '<a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="50000"/></a:schemeClr></a:gs>' +
  '</a:gsLst><a:lin ang="0"/></a:gradFill>' +
  '<a:gradFill><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"/></a:gs>' +
  '<a:gs pos="100000"><a:schemeClr val="phClr"/></a:gs></a:gsLst><a:lin ang="0"/></a:gradFill>' +
  '</a:fillStyleLst>';

const addContentType = (xml: string, partName: string, contentType: string): string =>
  xml.replace(
    '</Types>',
    `<Override PartName="/${partName}" ContentType="${contentType}"/></Types>`,
  );

const makeMultiMasterDeck = async () => {
  const pres = await loadPresentation(await readFile(fixturePath));
  const secondSlide = getSlides(pres)[1]!;
  addSlideShape(secondSlide, {
    preset: 'rect',
    x: inches(1),
    y: inches(1),
    w: inches(2),
    h: inches(1),
  });
  addSlideShape(secondSlide, {
    preset: 'rect',
    x: inches(1),
    y: inches(2.5),
    w: inches(2),
    h: inches(1),
  });

  const { entries } = readZip(await savePresentation(pres));
  const byName = new Map(entries.map((entry) => [entry.name, entry]));
  const text = (name: string): string => decoder.decode(byName.get(name)!.data);
  const replacement = (name: string, value: string) =>
    byName.set(name, { ...byName.get(name)!, data: encoder.encode(value) });

  let shapeIndex = 0;
  const secondSlideXml = text('ppt/slides/slide2.xml').replace(
    /<p:sp>[\s\S]*?<\/p:sp>/g,
    (shape) => {
      if (shape.includes('<p:txBody>')) return shape;
      const style = shapeIndex++ === 0 ? secondMasterStyle : secondMasterGradientStyle;
      return shape.replace('</p:sp>', `${style}</p:sp>`);
    },
  );
  replacement('ppt/slides/slide2.xml', secondSlideXml);
  replacement(
    'ppt/slides/_rels/slide2.xml.rels',
    text('ppt/slides/_rels/slide2.xml.rels').replace('slideLayout2.xml', 'slideLayout12.xml'),
  );

  byName.set('ppt/slideLayouts/slideLayout12.xml', {
    ...byName.get('ppt/slideLayouts/slideLayout2.xml')!,
    name: 'ppt/slideLayouts/slideLayout12.xml',
  });
  byName.set('ppt/slideLayouts/_rels/slideLayout12.xml.rels', {
    ...byName.get('ppt/slideLayouts/_rels/slideLayout2.xml.rels')!,
    name: 'ppt/slideLayouts/_rels/slideLayout12.xml.rels',
    data: encoder.encode(
      text('ppt/slideLayouts/_rels/slideLayout2.xml.rels').replace(
        'slideMaster1.xml',
        'slideMaster2.xml',
      ),
    ),
  });
  byName.set('ppt/slideMasters/slideMaster2.xml', {
    ...byName.get('ppt/slideMasters/slideMaster1.xml')!,
    name: 'ppt/slideMasters/slideMaster2.xml',
    data: encoder.encode(
      text('ppt/slideMasters/slideMaster1.xml').replace(
        /<p:sldLayoutIdLst>[\s\S]*?<\/p:sldLayoutIdLst>/,
        '<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>',
      ),
    ),
  });
  byName.set('ppt/slideMasters/_rels/slideMaster2.xml.rels', {
    ...byName.get('ppt/slideMasters/_rels/slideMaster1.xml.rels')!,
    name: 'ppt/slideMasters/_rels/slideMaster2.xml.rels',
    data: encoder.encode(
      text('ppt/slideMasters/_rels/slideMaster1.xml.rels')
        .replace('slideLayout1.xml', 'slideLayout12.xml')
        .replace('theme1.xml', 'theme2.xml'),
    ),
  });
  const secondTheme = text('ppt/theme/theme1.xml')
    .replace('<a:accent1><a:srgbClr val="4F81BD"', '<a:accent1><a:srgbClr val="112233"')
    .replace(/<a:fillStyleLst>[\s\S]*?<\/a:fillStyleLst>/, secondThemeFillStyles)
    .replace('<a:ln w="25400"', '<a:ln w="50800"');
  byName.set('ppt/theme/theme2.xml', {
    ...byName.get('ppt/theme/theme1.xml')!,
    name: 'ppt/theme/theme2.xml',
    data: encoder.encode(secondTheme),
  });

  replacement(
    'ppt/presentation.xml',
    text('ppt/presentation.xml').replace(
      '</p:sldMasterIdLst>',
      '<p:sldMasterId id="2147483649" r:id="rId9"/></p:sldMasterIdLst>',
    ),
  );
  replacement(
    'ppt/_rels/presentation.xml.rels',
    text('ppt/_rels/presentation.xml.rels').replace(
      '</Relationships>',
      '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster2.xml"/></Relationships>',
    ),
  );
  let contentTypes = text('[Content_Types].xml');
  contentTypes = addContentType(
    contentTypes,
    'ppt/slideLayouts/slideLayout12.xml',
    'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml',
  );
  contentTypes = addContentType(
    contentTypes,
    'ppt/slideMasters/slideMaster2.xml',
    'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml',
  );
  contentTypes = addContentType(
    contentTypes,
    'ppt/theme/theme2.xml',
    'application/vnd.openxmlformats-officedocument.theme+xml',
  );
  replacement('[Content_Types].xml', contentTypes);

  return loadPresentation(writeZip([...byName.values()]));
};

describe('shape style matrix references use the owning master theme', () => {
  it('resolves solid and gradient styles from a second slide master', async () => {
    const pres = await makeMultiMasterDeck();
    const assertColors = (loaded: Awaited<ReturnType<typeof loadPresentation>>) => {
      const shapes = getSlideShapes(getSlides(loaded)[1]!);
      const solid = shapes.at(-2)!;
      const gradient = shapes.at(-1)!;
      expect(getShapeStrokeColorResolved(loaded, solid)).toBe('#112233');
      expect(getShapeStrokeEffective(loaded, solid)).toMatchObject({
        kind: 'solid',
        widthEmu: 50800,
      });
      expect(getShapeFillEffective(loaded, solid)).toEqual({ kind: 'solid', color: '#112233' });
      expect(getShapeGradientFillEffective(loaded, gradient)?.stops[0]?.resolvedColor).toBe(
        '#112233',
      );
    };

    assertColors(pres);
    assertColors(await loadPresentation(await savePresentation(pres)));
  });
});
