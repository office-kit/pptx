import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideTextBox,
  getShapeRunFormatEffective,
  getSlideShapes,
  getSlides,
  inches,
  loadPresentation,
  savePresentation,
} from '../src/api/index.ts';
import { expectSchemaValid, isSchemaValidationAvailable } from './lib/expect-schema-valid.ts';
import { readZip, writeZip } from '../src/internal/opc/index.ts';

const fixturePath = fileURLToPath(new URL('./fixtures/minimal/two-slides.pptx', import.meta.url));
const decoder = new TextDecoder();
const encoder = new TextEncoder();
const fontStyle =
  '<p:style xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:lnRef idx="1"/><a:fillRef idx="1"/><a:effectRef idx="0"/><a:fontRef idx="minor"/></p:style>';

const addContentType = (xml: string, partName: string, contentType: string): string =>
  xml.replace(
    '</Types>',
    `<Override PartName="/${partName}" ContentType="${contentType}"/></Types>`,
  );

const makeMultiMasterFontDeck = async () => {
  const pres = await loadPresentation(await readFile(fixturePath));
  const secondSlide = getSlides(pres)[1]!;
  addSlideTextBox(secondSlide, {
    x: inches(1),
    y: inches(1),
    w: inches(4),
    h: inches(1),
    text: 'Master font probe',
    name: 'Master font probe',
  });

  const { entries } = readZip(await savePresentation(pres));
  const byName = new Map(entries.map((entry) => [entry.name, entry]));
  const text = (name: string): string => decoder.decode(byName.get(name)!.data);
  const replacement = (name: string, value: string) =>
    byName.set(name, { ...byName.get(name)!, data: encoder.encode(value) });

  replacement(
    'ppt/slides/slide2.xml',
    text('ppt/slides/slide2.xml').replace(
      /<p:sp>[\s\S]*?name="Master font probe"[\s\S]*?<\/p:sp>/,
      (shape) => shape.replace('<p:txBody>', `${fontStyle}<p:txBody>`),
    ),
  );
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
  const originalTheme = text('ppt/theme/theme1.xml');
  replacement(
    'ppt/theme/theme1.xml',
    originalTheme.replace(
      /<a:latin typeface="Calibri"\/>/g,
      '<a:latin typeface="First Master Sans"/>',
    ),
  );
  byName.set('ppt/theme/theme2.xml', {
    ...byName.get('ppt/theme/theme1.xml')!,
    name: 'ppt/theme/theme2.xml',
    data: encoder.encode(
      originalTheme.replace(
        /<a:latin typeface="Calibri"\/>/g,
        '<a:latin typeface="Second Master Sans"/>',
      ),
    ),
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
    addContentType(
      addContentType(
        contentTypes,
        'ppt/slideLayouts/slideLayout12.xml',
        'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml',
      ),
      'ppt/slideMasters/slideMaster2.xml',
      'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml',
    ),
    'ppt/theme/theme2.xml',
    'application/vnd.openxmlformats-officedocument.theme+xml',
  );
  replacement('[Content_Types].xml', contentTypes);
  return loadPresentation(writeZip([...byName.values()]));
};

const makeTitleQuickStyleDeck = async () => {
  const titleColoredFontStyle = await readFile(
    new URL('./fixtures/native/quickstyle-colored-fill-accent1-style.xml', import.meta.url),
    'utf8',
  );
  const pres = await loadPresentation(await readFile(fixturePath));
  const { entries } = readZip(await savePresentation(pres));
  const byName = new Map(entries.map((entry) => [entry.name, entry]));
  const text = (name: string): string => decoder.decode(byName.get(name)!.data);
  const replacement = (name: string, value: string) =>
    byName.set(name, { ...byName.get(name)!, data: encoder.encode(value) });

  replacement(
    'ppt/slides/slide1.xml',
    text('ppt/slides/slide1.xml').replace(
      /(<p:sp>[\s\S]*?<p:ph type="title"\/>[\s\S]*?)(<p:txBody>)/,
      `$1${titleColoredFontStyle}$2`,
    ),
  );
  replacement(
    'ppt/theme/theme1.xml',
    (() => {
      let latinIndex = 0;
      return text('ppt/theme/theme1.xml').replace(
        /<a:latin typeface="Calibri"\/>/g,
        () =>
          `<a:latin typeface="${latinIndex++ === 0 ? 'Title Master Headings' : 'Title Master Body'}"/>`,
      );
    })(),
  );
  return loadPresentation(writeZip([...byName.values()]));
};

describe('shape style fontRef resolves the owning master theme', () => {
  it.runIf(isSchemaValidationAvailable())(
    'uses schema-valid slide and master fixtures',
    async () => {
      for (const pres of [await makeMultiMasterFontDeck(), await makeTitleQuickStyleDeck()]) {
        for (const entry of readZip(await savePresentation(pres)).entries) {
          if (/^ppt\/(slides|slideMasters)\/[^/]+\.xml$/.test(entry.name))
            expectSchemaValid(decoder.decode(entry.data), 'pml');
        }
      }
    },
  );

  it('uses the second master font scheme and preserves it through roundtrip', async () => {
    const pres = await makeMultiMasterFontDeck();
    const slideXml = decoder.decode(
      readZip(await savePresentation(pres)).entries.find(
        (entry) => entry.name === 'ppt/slides/slide2.xml',
      )!.data,
    );
    expect(slideXml).toContain('<a:fontRef idx="minor"/>');
    const shape = getSlideShapes(getSlides(pres)[1]!).at(-1)!;
    expect(getShapeRunFormatEffective(pres, shape, 0, 0).font).toBe('Second Master Sans');

    const restored = await loadPresentation(await savePresentation(pres));
    const restoredShape = getSlideShapes(getSlides(restored)[1]!).at(-1)!;
    expect(getShapeRunFormatEffective(restored, restoredShape, 0, 0).font).toBe(
      'Second Master Sans',
    );
  });

  it('lets a title Quick Style override inherited master font and color', async () => {
    const pres = await makeTitleQuickStyleDeck();
    const title = getSlideShapes(getSlides(pres)[0]!)[0]!;
    const format = getShapeRunFormatEffective(pres, title, 0, 0);
    expect(format.font).toBe('Title Master Body');
    expect(format.color).toBe('#FFFFFF');

    const restored = await loadPresentation(await savePresentation(pres));
    const restoredTitle = getSlideShapes(getSlides(restored)[0]!)[0]!;
    expect(getShapeRunFormatEffective(restored, restoredTitle, 0, 0)).toMatchObject({
      font: 'Title Master Body',
      color: '#FFFFFF',
    });
  });
});
