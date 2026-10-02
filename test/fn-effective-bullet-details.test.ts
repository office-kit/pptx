import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  addSlideTable,
  addSlideTextBox,
  findSlideLayout,
  findSlidePlaceholder,
  getParagraphPropertiesEffective,
  getSlideShapes,
  isTableShape,
  getSlides,
  getTableCells,
  getShapeText,
  inches,
  loadPresentation,
  savePresentation,
  setParagraphBullet,
  setShapeText,
} from '../src/api/index.ts';
import { readZip, writeZip } from '../src/internal/opc/index.ts';
import { expectSchemaValid, isSchemaValidationAvailable } from './lib/expect-schema-valid.ts';

const fixturePath = fileURLToPath(new URL('./fixtures/minimal/blank.pptx', import.meta.url));
const decoder = new TextDecoder();
const encoder = new TextEncoder();

const blankSlide = async () => {
  const pres = await loadPresentation(await readFile(fixturePath));
  const layout = findSlideLayout(pres, 'Blank');
  if (!layout) throw new Error('blank layout missing');
  return { pres, slide: addSlide(pres, { layout }) };
};

const rewrite = async (
  bytes: Uint8Array,
  mutate: (name: string, xml: string) => string,
): Promise<Uint8Array> => {
  const { entries } = readZip(bytes);
  return writeZip(
    entries.map((entry) => {
      if (!entry.name.endsWith('.xml')) return entry;
      const xml = decoder.decode(entry.data);
      const rewritten = mutate(entry.name, xml);
      if (rewritten === xml) return entry;
      if (isSchemaValidationAvailable()) expectSchemaValid(rewritten, 'pml');
      return { ...entry, data: encoder.encode(rewritten) };
    }),
  );
};

describe('effective bullet details', () => {
  it('merges imported text-body defaults and preserves explicit paragraph overrides after reload', async () => {
    const { pres, slide } = await blankSlide();
    const box = addSlideTextBox(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(4),
      h: inches(1),
      text: 'Inherited',
    });
    setParagraphBullet(box, 0, 'bullet');
    const zip = await rewrite(await savePresentation(pres), (name, xml) =>
      name === 'ppt/slides/slide1.xml'
        ? xml.replace(
            '<a:lstStyle/>',
            '<a:lstStyle><a:lvl1pPr><a:buClr><a:srgbClr val="00FF00"/></a:buClr><a:buSzPct val="50%"/><a:buFont typeface="Courier New"/></a:lvl1pPr></a:lstStyle>',
          )
        : xml,
    );
    const imported = await loadPresentation(zip);
    const importedBox = getSlideShapes(getSlides(imported).at(-1)!)[0]!;
    expect(getParagraphPropertiesEffective(imported, importedBox, 0).bulletDetail).toMatchObject({
      color: '#00FF00',
      sizePct: 0.5,
      sizePts: null,
      font: 'Courier New',
      colorFollowText: false,
      sizeFollowText: false,
      fontFollowText: false,
    });

    const reloaded = await loadPresentation(await savePresentation(imported));
    const reloadedBox = getSlideShapes(getSlides(reloaded).at(-1)!)[0]!;
    expect(getParagraphPropertiesEffective(reloaded, reloadedBox, 0).bulletDetail).toMatchObject({
      color: '#00FF00',
      sizePct: 0.5,
      font: 'Courier New',
    });
  });

  it.each(['values', 'follow-text'])(
    'lets explicit paragraph properties override inherited %s',
    async (inherited) => {
      const { pres, slide } = await blankSlide();
      const box = addSlideTextBox(slide, {
        x: inches(1),
        y: inches(1),
        w: inches(4),
        h: inches(1),
        text: 'Override',
      });
      setParagraphBullet(box, 0, 'bullet');
      const zip = await rewrite(await savePresentation(pres), (name, xml) => {
        if (name !== 'ppt/slides/slide1.xml') return xml;
        const withInherited = xml.replace(
          '<a:lstStyle/>',
          inherited === 'follow-text'
            ? '<a:lstStyle><a:lvl1pPr><a:buClrTx/><a:buSzTx/><a:buFontTx/></a:lvl1pPr></a:lstStyle>'
            : '<a:lstStyle><a:lvl1pPr><a:buClr><a:srgbClr val="00FF00"/></a:buClr><a:buSzPct val="125%"/><a:buFont typeface="Courier New"/></a:lvl1pPr></a:lstStyle>',
        );
        const rewritten = withInherited.replace(
          /<a:pPr([^>]*)>([\s\S]*?)<\/a:pPr>/,
          (_match, attrs, body) => {
            const withoutBulletChoices = body
              .replace(/<a:bu(?:Clr|SzPct|SzPts|Font|Char|AutoNum|None)\b[^>]*\/>/g, '')
              .replace(/<a:buClr\b[^>]*>[\s\S]*?<\/a:buClr>/g, '');
            return `<a:pPr${attrs}><a:buClr><a:srgbClr val="FF0000"/></a:buClr><a:buSzPts val="1800"/><a:buFont typeface="Arial"/><a:buChar char="•"/>${withoutBulletChoices}</a:pPr>`;
          },
        );
        expect(rewritten).not.toBe(withInherited);
        return rewritten;
      });
      const imported = await loadPresentation(zip);
      const importedShape = getSlideShapes(getSlides(imported).at(-1)!).find(
        (candidate) => getShapeText(candidate) === 'Override',
      )!;
      const detail = getParagraphPropertiesEffective(imported, importedShape, 0).bulletDetail;
      expect(detail).toEqual({
        color: '#FF0000',
        colorFollowText: false,
        sizePct: null,
        sizePts: 18,
        sizeFollowText: false,
        font: 'Arial',
        fontFollowText: false,
      });
    },
  );

  it('keeps bu*Tx as explicit follow-text sentinels and resolves table-cell details', async () => {
    const { pres, slide } = await blankSlide();
    const table = addSlideTable(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(4),
      h: inches(1),
      rows: [['Cell']],
    });
    const cell = getTableCells(table)[0]![0]!;
    setParagraphBullet(cell, 0, 'bullet');
    const zip = await rewrite(await savePresentation(pres), (name, xml) => {
      if (name !== 'ppt/slides/slide1.xml') return xml;
      const withInherited = xml.replace(
        '<a:lstStyle/>',
        '<a:lstStyle><a:lvl1pPr><a:buClr><a:srgbClr val="00FF00"/></a:buClr><a:buSzPts val="2400"/><a:buFont typeface="Courier New"/></a:lvl1pPr></a:lstStyle>',
      );
      const rewritten = withInherited.replace(
        '<a:buChar',
        '<a:buClrTx/><a:buSzTx/><a:buFontTx/><a:buChar',
      );
      expect(rewritten).not.toBe(xml);
      return rewritten;
    });
    const imported = await loadPresentation(zip);
    const importedTable = getSlideShapes(getSlides(imported).at(-1)!).find(isTableShape)!;
    const importedCell = getTableCells(importedTable)[0]![0]!;
    expect(getParagraphPropertiesEffective(imported, importedCell, 0).bulletDetail).toMatchObject({
      color: null,
      colorFollowText: true,
      sizePct: null,
      sizePts: null,
      sizeFollowText: true,
      font: null,
      fontFollowText: true,
    });
  });

  it('reads all inherited master bodyStyle details from the imported fixture', async () => {
    const pres = await loadPresentation(await readFile(fixturePath));
    const layout = findSlideLayout(pres, 'Title and Content');
    if (!layout) throw new Error('layout missing');
    const slide = addSlide(pres, { layout });
    const body = findSlidePlaceholder(slide, 'body');
    if (!body) throw new Error('body missing');
    setShapeText(body, 'Master');
    const zip = await rewrite(await savePresentation(pres), (name, xml) => {
      if (name !== 'ppt/slideMasters/slideMaster1.xml') return xml;
      const rewritten = xml.replace(
        /(<p:bodyStyle>[\s\S]*?<a:lvl1pPr[^>]*>)([\s\S]*?)(<\/a:lvl1pPr>)/,
        (_match, start, body, end) => {
          const withoutBulletChoices = body
            .replace(/<a:bu(?:Clr|SzPct|SzPts|Font|Char|AutoNum|None)\b[^>]*\/>/g, '')
            .replace(/<a:buClr\b[^>]*>[\s\S]*?<\/a:buClr>/g, '');
          const bullets =
            '<a:buClr><a:srgbClr val="112233"/></a:buClr><a:buSzPts val="2400"/><a:buFont typeface="Verdana"/>';
          const insertion = withoutBulletChoices.search(/<a:(?:tabLst|defRPr)\b/);
          const orderedBody =
            insertion < 0
              ? `${withoutBulletChoices}${bullets}`
              : `${withoutBulletChoices.slice(0, insertion)}${bullets}${withoutBulletChoices.slice(insertion)}`;
          return `${start}${orderedBody}${end}`;
        },
      );
      expect(rewritten).not.toBe(xml);
      return rewritten;
    });
    const imported = await loadPresentation(zip);
    const importedBody = findSlidePlaceholder(getSlides(imported).at(-1)!, 'body');
    if (!importedBody) throw new Error('reloaded body missing');
    expect(getParagraphPropertiesEffective(imported, importedBody, 0).bulletDetail).toMatchObject({
      color: '#112233',
      sizePct: null,
      sizePts: 24,
      font: 'Verdana',
    });
  });

  it('reads bullet details inherited from a layout body placeholder', async () => {
    const pres = await loadPresentation(await readFile(fixturePath));
    const layout = findSlideLayout(pres, 'Title and Content');
    if (!layout) throw new Error('layout missing');
    const slide = addSlide(pres, { layout });
    const body = findSlidePlaceholder(slide, 'body');
    if (!body) throw new Error('body missing');
    setShapeText(body, 'Layout');
    const zip = await rewrite(await savePresentation(pres), (name, xml) => {
      if (name !== 'ppt/slideLayouts/slideLayout2.xml') return xml;
      const style =
        '<a:lstStyle><a:lvl1pPr><a:buClr><a:srgbClr val="445566"/></a:buClr><a:buSzPts val="2000"/><a:buFont typeface="Tahoma"/></a:lvl1pPr></a:lstStyle>';
      const rewritten = xml.replace(/<p:sp>[\s\S]*?<\/p:sp>/g, (shape) =>
        shape.includes('<p:ph idx="1"/>') ? shape.replace('<a:lstStyle/>', style) : shape,
      );
      expect(rewritten).not.toBe(xml);
      return rewritten;
    });
    const imported = await loadPresentation(zip);
    const importedBody = findSlidePlaceholder(getSlides(imported).at(-1)!, 'body');
    if (!importedBody) throw new Error('reloaded body missing');
    expect(getParagraphPropertiesEffective(imported, importedBody, 0).bulletDetail).toMatchObject({
      color: '#445566',
      sizePts: 20,
      font: 'Tahoma',
    });
  });
});
