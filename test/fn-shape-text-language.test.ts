// Review ▸ Language: the proofing language on a shape's runs.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addBlankSlide,
  addSlideTable,
  createPresentation,
  getParagraphElementLanguages,
  getTableCells,
  addSlide,
  addSlideShape,
  addSlideTextBox,
  findSlideLayout,
  getShapeId,
  getShapeRunFormat,
  getShapeTextLanguage,
  getShapeXmlString,
  getSlideShapes,
  getSlideXmlString,
  getSlides,
  inches,
  loadPresentation,
  savePresentation,
  setShapeTextField,
  setShapeTextLanguage,
} from '../src/api/index.ts';
import { readZip, writeZip } from '../src/internal/opc/index.ts';
import { expectSchemaValid, isSchemaValidationAvailable } from './lib/expect-schema-valid.ts';

const fixture = fileURLToPath(new URL('./fixtures/minimal/blank.pptx', import.meta.url));
const skipIfNoXmllint = isSchemaValidationAvailable() ? it : it.skip;

const deck = async () => {
  const pres = await loadPresentation(await readFile(fixture));
  const slide = addSlide(pres, { layout: findSlideLayout(pres, 'Blank')! });
  return { pres, slide };
};

describe('setShapeTextLanguage', () => {
  it('marks every run and paragraph end, keeps the format and survives save', async () => {
    const { pres, slide } = await deck();
    const box = addSlideTextBox(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(3),
      h: inches(1),
      text: 'Hello\nWorld',
    });
    const format = getShapeRunFormat(box, 0, 0);
    setShapeTextLanguage(box, 'ja-JP');
    expect(getShapeTextLanguage(box)).toBe('ja-JP');
    expect(getShapeXmlString(box).match(/lang="ja-JP"/g)!.length).toBeGreaterThanOrEqual(2);
    expect(getShapeRunFormat(box, 0, 0)).toEqual(format);
    const reloaded = getSlideShapes(
      getSlides(await loadPresentation(await savePresentation(pres))).at(-1)!,
    ).find((shape) => getShapeId(shape) === getShapeId(box))!;
    expect(getShapeTextLanguage(reloaded)).toBe('ja-JP');
  });

  skipIfNoXmllint('gives fields and empty shapes a schema-valid rPr', async () => {
    const { slide } = await deck();
    const field = addSlideTextBox(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(3),
      h: inches(1),
      text: '',
    });
    setShapeTextField(field, 'slidenum');
    setShapeTextLanguage(field, 'en-GB');
    const empty = addSlideShape(slide, {
      preset: 'rect',
      x: inches(1),
      y: inches(3),
      w: inches(1),
      h: inches(1),
    });
    setShapeTextLanguage(empty, 'fr');
    expect(getShapeTextLanguage(field)).toBe('en-GB');
    expectSchemaValid(getSlideXmlString(slide), 'pml');
  });

  it('rejects a value that is not a language tag', async () => {
    const { slide } = await deck();
    const box = addSlideTextBox(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(3),
      h: inches(1),
      text: 'A',
    });
    const before = getShapeXmlString(box);
    expect(() => setShapeTextLanguage(box, 'en US')).toThrow(/language tag/);
    expect(getShapeXmlString(box)).toBe(before);
  });
});

describe('getParagraphElementLanguages', () => {
  it.each(['shape', 'table'] as const)(
    'reads each run, field and break language in a %s and survives save',
    async (kind) => {
      const original = createPresentation();
      const slide = addBlankSlide(original);
      const bounds = { x: inches(1), y: inches(1), w: inches(4), h: inches(2) };
      if (kind === 'shape') addSlideTextBox(slide, { ...bounds, text: 'Text' });
      else addSlideTable(slide, { ...bounds, rows: [['Text']] });
      const { entries } = readZip(await savePresentation(original));
      const paragraph =
        '<a:p><a:r><a:rPr lang="de-DE"/><a:t>1,5</a:t></a:r><a:r><a:t>x</a:t></a:r>' +
        '<a:br><a:rPr lang="fr-FR"/></a:br><a:fld id="{00000000-0000-0000-0000-000000000001}" type="slidenum"><a:rPr lang="en-US"/><a:t>1</a:t></a:fld>' +
        '<a:endParaRPr lang="ja-JP"/></a:p>';
      const edited = writeZip(
        entries.map((entry) =>
          entry.name !== 'ppt/slides/slide1.xml'
            ? entry
            : {
                ...entry,
                data: new TextEncoder().encode(
                  new TextDecoder().decode(entry.data).replace(/<a:p>[\s\S]*?<\/a:p>/, paragraph),
                ),
              },
        ),
      );
      const read = async (bytes: Uint8Array) => {
        const shape = getSlideShapes(getSlides(await loadPresentation(bytes))[0]!)[0]!;
        return getParagraphElementLanguages(
          kind === 'table' ? getTableCells(shape)[0]![0]! : shape,
          0,
        );
      };
      const expected = ['de-DE', null, 'fr-FR', 'en-US'];
      expect(await read(edited)).toEqual(expected);
      const pres = await loadPresentation(edited);
      expect(await read(await savePresentation(pres))).toEqual(expected);
    },
  );

  it('follows setShapeTextLanguage', async () => {
    const { slide } = await deck();
    const box = addSlideTextBox(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(3),
      h: inches(1),
      text: 'A\nB',
    });
    setShapeTextLanguage(box, 'de-DE');
    expect(getParagraphElementLanguages(box, 0)).toEqual(['de-DE']);
    expect(() => getParagraphElementLanguages(box, 5)).toThrow(RangeError);
  });
});
