// The reference desktop app's built-in table styles: the generated data reproduces
// the reference desktop app's serialization, and applying a style writes its definition into
// ppt/tableStyles.xml the way the reference desktop app does.

import { readFile } from 'node:fs/promises';
import { describe, expect, it } from 'vitest';
import {
  BUILTIN_TABLE_STYLES,
  addBlankSlide,
  addSlideTable,
  createPresentation,
  getSlides,
  getSlideShapes,
  getSlideXmlString,
  getTableStyleId,
  inches,
  loadPresentation,
  savePresentation,
  setTableStyleId,
} from '../src/api/index.ts';
import { builtinTableStyleXml } from '../src/internal/presentationml/index.ts';
import { readZip, writeZip } from '../src/internal/opc/index.ts';
import { expectSchemaValid, isSchemaValidationAvailable } from './lib/expect-schema-valid.ts';

const FIXTURE = new URL('./fixtures/native/builtin-table-styles.xml', import.meta.url);
const decoder = new TextDecoder();
const encoder = new TextEncoder();

const entryText = (bytes: Uint8Array, name: string): string | null => {
  const entry = readZip(bytes).entries.find((e) => e.name === name);
  return entry ? decoder.decode(entry.data) : null;
};

const styleIds = (tableStylesXml: string): string[] =>
  [...tableStylesXml.matchAll(/<a:tblStyle [^>]*styleId="([^"]+)"/g)].map((m) => m[1]!);

const deckWithTable = () => {
  const pres = createPresentation();
  const table = addSlideTable(addBlankSlide(pres), {
    x: inches(1),
    y: inches(1),
    w: inches(6),
    h: inches(2),
    rows: [
      ['a', 'b', 'c'],
      ['d', 'e', 'f'],
    ],
  });
  return { pres, table };
};

describe('built-in table style data', () => {
  it("expands every style to the reference desktop app's own bytes, in gallery order", async () => {
    const fixture = await readFile(FIXTURE, 'utf8');
    const native = [...fixture.matchAll(/<a:tblStyle [\s\S]*?<\/a:tblStyle>/g)].map((m) => m[0]);
    expect(native).toHaveLength(74);
    expect(BUILTIN_TABLE_STYLES).toHaveLength(74);
    expect(BUILTIN_TABLE_STYLES.map((style) => builtinTableStyleXml(style.id))).toEqual(native);
    expect(new Set(BUILTIN_TABLE_STYLES.map((style) => style.id)).size).toBe(74);
  });

  it('groups the gallery like the reference desktop app', () => {
    const count = (category: string) =>
      BUILTIN_TABLE_STYLES.filter((style) => style.category === category).length;
    expect([count('bestMatch'), count('light'), count('medium'), count('dark')]).toEqual([
      14, 21, 28, 11,
    ]);
    expect(BUILTIN_TABLE_STYLES[43]).toEqual({
      id: '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}',
      name: 'Medium Style 2 - Accent 1',
      category: 'medium',
    });
  });
});

describe('setTableStyleId with built-in styles', () => {
  it('applies a style by name and writes its definition once', async () => {
    const { pres, table } = deckWithTable();
    setTableStyleId(table, 'Dark Style 2 - Accent 3/Accent 4');
    setTableStyleId(table, 'Light Style 1 - Accent 2');
    setTableStyleId(table, '{91EBBBCC-DAD2-459C-BE2E-F6DE35CF9A28}');
    const bytes = await savePresentation(pres);
    const tableStyles = entryText(bytes, 'ppt/tableStyles.xml')!;
    expect(styleIds(tableStyles)).toEqual([
      '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}',
      '{91EBBBCC-DAD2-459C-BE2E-F6DE35CF9A28}',
      '{0E3FDE45-AF77-4B5C-9715-49D594BDF05E}',
    ]);
    expect(tableStyles).toContain(builtinTableStyleXml('{0E3FDE45-AF77-4B5C-9715-49D594BDF05E}'));
    const reloaded = await loadPresentation(bytes);
    const reloadedTable = getSlideShapes(getSlides(reloaded)[0]!)[0]!;
    expect(getTableStyleId(reloadedTable)).toBe('{91EBBBCC-DAD2-459C-BE2E-F6DE35CF9A28}');
    // A second save of the reloaded deck is byte-stable for this part.
    expect(entryText(await savePresentation(reloaded), 'ppt/tableStyles.xml')).toBe(tableStyles);
    if (isSchemaValidationAvailable()) {
      expectSchemaValid(tableStyles, 'dml');
      expectSchemaValid(getSlideXmlString(getSlides(reloaded)[0]!), 'pml');
    }
  });

  it('writes a schema-valid list for every built-in style', async () => {
    const { pres, table } = deckWithTable();
    for (const style of BUILTIN_TABLE_STYLES) setTableStyleId(table, style.name);
    const tableStyles = entryText(await savePresentation(pres), 'ppt/tableStyles.xml')!;
    expect(new Set(styleIds(tableStyles))).toEqual(
      new Set(BUILTIN_TABLE_STYLES.map((style) => style.id)),
    );
    if (isSchemaValidationAvailable()) expectSchemaValid(tableStyles, 'dml');
  });

  it("keeps a deck's own definition of a built-in GUID", async () => {
    const { pres } = deckWithTable();
    const own =
      '<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}">' +
      '<a:tblStyle styleId="{073A0DAA-6AF3-43AB-8588-CEC1D06C72B9}" styleName="Edited"><a:wholeTbl><a:tcStyle><a:tcBdr/></a:tcStyle></a:wholeTbl></a:tblStyle></a:tblStyleLst>';
    const { entries } = readZip(await savePresentation(pres));
    const edited = await loadPresentation(
      writeZip(
        entries.map((entry) =>
          entry.name === 'ppt/tableStyles.xml' ? { ...entry, data: encoder.encode(own) } : entry,
        ),
      ),
    );
    setTableStyleId(getSlideShapes(getSlides(edited)[0]!)[0]!, 'Medium Style 2');
    const tableStyles = entryText(await savePresentation(edited), 'ppt/tableStyles.xml')!;
    expect(tableStyles).toContain('styleName="Edited"');
    expect(styleIds(tableStyles)).toEqual(['{073A0DAA-6AF3-43AB-8588-CEC1D06C72B9}']);
  });

  it('creates ppt/tableStyles.xml with its relationship and content type when missing', async () => {
    const { pres } = deckWithTable();
    const { entries } = readZip(await savePresentation(pres));
    const stripped = entries
      .filter((entry) => entry.name !== 'ppt/tableStyles.xml')
      .map((entry) => {
        const text = decoder.decode(entry.data);
        if (entry.name === 'ppt/_rels/presentation.xml.rels')
          return {
            ...entry,
            data: encoder.encode(text.replace(/<Relationship [^>]*tableStyles"[^>]*\/>/, '')),
          };
        if (entry.name === '[Content_Types].xml')
          return {
            ...entry,
            data: encoder.encode(
              text.replace(/<Override PartName="\/ppt\/tableStyles.xml"[^>]*\/>/, ''),
            ),
          };
        return entry;
      });
    const deck = await loadPresentation(writeZip(stripped));
    setTableStyleId(getSlideShapes(getSlides(deck)[0]!)[0]!, 'Themed Style 1 - Accent 4');
    const bytes = await savePresentation(deck);
    const tableStyles = entryText(bytes, 'ppt/tableStyles.xml')!;
    expect(styleIds(tableStyles)).toEqual(['{775DCB02-9BB8-47FD-8907-85C794F793BA}']);
    expect(tableStyles).toContain('def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"');
    expect(entryText(bytes, 'ppt/_rels/presentation.xml.rels')).toMatch(/tableStyles" Target="/);
    const contentTypes = entryText(bytes, '[Content_Types].xml')!;
    expect(contentTypes).toContain('presentationml.tableStyles+xml');
    if (isSchemaValidationAvailable()) {
      expectSchemaValid(tableStyles, 'dml');
      expectSchemaValid(contentTypes, 'contentTypes');
      expectSchemaValid(entryText(bytes, 'ppt/_rels/presentation.xml.rels')!, 'rels');
    }
  });

  it('leaves tableStyles.xml alone for a custom GUID', async () => {
    const { pres, table } = deckWithTable();
    const before = entryText(await savePresentation(pres), 'ppt/tableStyles.xml');
    setTableStyleId(table, '{12345678-1234-1234-1234-123456789ABC}');
    expect(entryText(await savePresentation(pres), 'ppt/tableStyles.xml')).toBe(before);
  });

  it('rejects a name that is neither a GUID nor a built-in style', () => {
    const { table } = deckWithTable();
    expect(() => setTableStyleId(table, 'Medium Style 9')).toThrow(/not a GUID/);
  });

  it('addSlideTable writes the default style definition', async () => {
    const { pres } = deckWithTable();
    const tableStyles = entryText(await savePresentation(pres), 'ppt/tableStyles.xml')!;
    expect(styleIds(tableStyles)).toEqual(['{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}']);
  });
});
