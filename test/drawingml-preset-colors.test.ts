import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { describe, expect, it } from 'vitest';
import {
  DRAWINGML_PRESET_COLORS,
  resolveDrawingMLPresetColor,
} from '../src/internal/drawingml/preset-colors';

const XSD_PATH = resolve(
  process.cwd(),
  'references/ecma-376-5th/ECMA-376/OfficeOpenXML-XMLSchema-Strict/dml-main.xsd',
);

const xsd = readFileSync(XSD_PATH, 'utf8');
const presetType = xsd.slice(
  xsd.indexOf('name="ST_PresetColorVal"'),
  xsd.indexOf('</xsd:simpleType>', xsd.indexOf('name="ST_PresetColorVal"')),
);
const xsdPresetNames = [...presetType.matchAll(/enumeration value="([^"]+)"/g)].map(
  (match) => match[1]!,
);

describe('DrawingML preset colors', () => {
  it('covers every strict ECMA-376 ST_PresetColorVal enumeration value', () => {
    expect(xsdPresetNames).toHaveLength(190);
    expect(Object.keys(DRAWINGML_PRESET_COLORS)).toHaveLength(190);
    expect(Object.keys(DRAWINGML_PRESET_COLORS)).toEqual(expect.arrayContaining(xsdPresetNames));
    for (const name of xsdPresetNames) {
      expect(resolveDrawingMLPresetColor(name)).toMatch(/^#[0-9A-F]{6}$/);
    }
  });

  it('keeps DrawingML short aliases equal to their long names', () => {
    expect(resolveDrawingMLPresetColor('red')).toBe('#FF0000');
    expect(resolveDrawingMLPresetColor('dkBlue')).toBe(resolveDrawingMLPresetColor('darkBlue'));
    expect(resolveDrawingMLPresetColor('dkSeaGreen')).toBe('#8FBC8B');
    expect(resolveDrawingMLPresetColor('ltGoldenrodYellow')).toBe('#FAFA78');
    expect(resolveDrawingMLPresetColor('ltSlateGrey')).toBe(
      resolveDrawingMLPresetColor('lightSlateGrey'),
    );
    expect(resolveDrawingMLPresetColor('medTurquoise')).toBe(
      resolveDrawingMLPresetColor('mediumTurquoise'),
    );
  });

  it('rejects unknown values and inherited object-property names', () => {
    expect(resolveDrawingMLPresetColor('notAColor')).toBeNull();
    expect(resolveDrawingMLPresetColor('__proto__')).toBeNull();
    expect(resolveDrawingMLPresetColor('constructor')).toBeNull();
  });
});
