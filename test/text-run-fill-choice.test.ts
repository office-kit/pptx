import { describe, expect, it } from 'vitest';
import { readFile } from 'node:fs/promises';
import {
  getShapeRunFormat,
  getSlideShapes,
  getSlideXmlString,
  getSlides,
  loadPresentation,
  savePresentation,
  setShapeRunFormat,
} from '../src/api/index.ts';
import { readZip, writeZip } from '../src/internal/opc/index.ts';
import { expectSchemaValid, isSchemaValidationAvailable } from './lib/expect-schema-valid.ts';
import { applyRunFormat } from '../src/internal/drawingml/text-format.ts';
import { parseRPrLikeElement } from '../src/api/fn/shape-color.ts';
import type { PatternPreset } from '../src/internal/drawingml/fill.ts';
import { NS, parseXml, serializeXml } from '../src/internal/xml/index.ts';

const fills = [
  '<a:noFill/>',
  '<a:solidFill><a:srgbClr val="112233"/></a:solidFill>',
  '<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="112233"/></a:gs><a:gs pos="100000"><a:srgbClr val="FFFFFF"/></a:gs></a:gsLst><a:lin ang="0" scaled="1"/></a:gradFill>',
  '<a:pattFill prst="pct5"><a:fgClr><a:srgbClr val="112233"/></a:fgClr><a:bgClr><a:srgbClr val="FFFFFF"/></a:bgClr></a:pattFill>',
  '<a:blipFill><a:blip/><a:stretch><a:fillRect/></a:stretch></a:blipFill>',
  '<a:grpFill/>',
];
const fillNames = new Set(['noFill', 'solidFill', 'gradFill', 'pattFill', 'blipFill', 'grpFill']);

describe('text run fill choice', () => {
  it('writes and reads gradient and pattern text fills without flattening them', () => {
    for (const fill of [
      {
        kind: 'gradient' as const,
        stops: [
          { offset: 0, color: '#112233' as const },
          { offset: 1, color: 'accent2' as const, brightness: 0.2 },
        ],
        angleDeg: 45,
      },
      {
        kind: 'pattern' as const,
        preset: 'dkUpDiag' as const,
        foreground: 'accent1' as const,
        background: '#FFFFFF' as const,
      },
    ]) {
      const doc = parseXml(`<a:rPr xmlns:a="${NS.dml}"/>`);
      applyRunFormat(doc.root, { textFill: fill });
      const xml = serializeXml(doc);
      expect(xml).toContain(fill.kind === 'gradient' ? '<a:gradFill' : '<a:pattFill');
      expect(xml).not.toContain('<a:solidFill');
    }
  });

  it.each(['wordart-gray-gradient-shape.xml', 'wordart-white-pattern-shadow-shape.xml'])(
    'round-trips a font color change on native %s',
    async (name) => {
      const native = await readFile(new URL(`./fixtures/native/${name}`, import.meta.url), 'utf8');
      const bytes = await readFile(
        new URL('./fixtures/minimal/one-text-slide.pptx', import.meta.url),
      );
      const entries = readZip(bytes).entries.map((entry) => {
        if (entry.name !== 'ppt/slides/slide1.xml') return entry;
        const xml = new TextDecoder().decode(entry.data);
        return {
          ...entry,
          data: new TextEncoder().encode(xml.replace(/<p:sp[ >][\s\S]*?<\/p:sp>/, native)),
        };
      });
      const pres = await loadPresentation(writeZip(entries));
      const shape = getSlideShapes(getSlides(pres)[0]!)[0]!;
      setShapeRunFormat(shape, 0, 0, { color: '#ABCDEF' });
      const restored = await loadPresentation(await savePresentation(pres));
      const slide = getSlides(restored)[0]!;
      expect(getShapeRunFormat(getSlideShapes(slide)[0]!, 0, 0)?.color).toBe('#ABCDEF');
      if (isSchemaValidationAvailable()) expectSchemaValid(getSlideXmlString(slide), 'pml');
    },
  );

  it.each([
    ['wordart-gray-gradient-shape.xml', 'gradient'],
    ['wordart-white-pattern-shadow-shape.xml', 'pattern'],
  ] as const)('reads native WordArt %s as a text fill', async (name, kind) => {
    const native = await readFile(new URL(`./fixtures/native/${name}`, import.meta.url), 'utf8');
    const bytes = await readFile(
      new URL('./fixtures/minimal/one-text-slide.pptx', import.meta.url),
    );
    const entries = readZip(bytes).entries.map((entry) =>
      entry.name !== 'ppt/slides/slide1.xml'
        ? entry
        : {
            ...entry,
            data: new TextEncoder().encode(
              new TextDecoder().decode(entry.data).replace(/<p:sp[ >][\s\S]*?<\/p:sp>/, native),
            ),
          },
    );
    const pres = await loadPresentation(writeZip(entries));
    const format = getShapeRunFormat(getSlideShapes(getSlides(pres)[0]!)[0]!, 0, 0);
    expect(format?.textFill?.kind).toBe(kind);
    if (kind === 'pattern' && format?.textFill?.kind === 'pattern') {
      expect(format.textFill.foregroundTransforms?.[0]?.kind).toBe('lumMod');
      expect(format.textFill.backgroundTransforms?.[0]?.kind).toBe('lumMod');
    }
  });

  it('round-trips an authored gradient text fill through PPTX XML', async () => {
    const native = await readFile(
      new URL('./fixtures/native/wordart-gray-gradient-shape.xml', import.meta.url),
      'utf8',
    );
    const bytes = await readFile(
      new URL('./fixtures/minimal/one-text-slide.pptx', import.meta.url),
    );
    const entries = readZip(bytes).entries.map((entry) =>
      entry.name !== 'ppt/slides/slide1.xml'
        ? entry
        : {
            ...entry,
            data: new TextEncoder().encode(
              new TextDecoder().decode(entry.data).replace(/<p:sp[ >][\s\S]*?<\/p:sp>/, native),
            ),
          },
    );
    const pres = await loadPresentation(writeZip(entries));
    const shape = getSlideShapes(getSlides(pres)[0]!)[0]!;
    setShapeRunFormat(shape, 0, 0, {
      textFill: {
        kind: 'gradient',
        stops: [
          { offset: 0, color: '#112233', colorTransforms: [{ kind: 'lumMod', value: 0.5 }] },
          { offset: 1, color: 'accent2' },
        ],
        angleDeg: 270,
      },
    });
    const restored = await loadPresentation(await savePresentation(pres));
    const slide = getSlides(restored)[0]!;
    const fill = getShapeRunFormat(getSlideShapes(slide)[0]!, 0, 0)?.textFill;
    expect(fill?.kind).toBe('gradient');
    if (fill?.kind === 'gradient') {
      expect(fill.stops).toHaveLength(2);
      expect(fill.stops[0]?.color).toBe('#112233');
      expect(fill.stops[0]?.colorTransforms?.[0]?.kind).toBe('lumMod');
    }
    if (isSchemaValidationAvailable()) expectSchemaValid(getSlideXmlString(slide), 'pml');
  });

  it('preserves the existing fill when an invalid color is rejected', () => {
    const doc = parseXml(`<a:rPr xmlns:a="${NS.dml}">${fills[2]}</a:rPr>`);
    const before = serializeXml(doc);
    expect(() => applyRunFormat(doc.root, { color: 'invalid' as '#FFFFFF' })).toThrow();
    expect(serializeXml(doc)).toBe(before);
  });

  it('rejects invalid text fills without mutating the existing choice', () => {
    const doc = parseXml(
      `<a:rPr xmlns:a="${NS.dml}"><a:ln/><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="112233"/></a:gs><a:gs pos="100000"><a:srgbClr val="FFFFFF"/></a:gs></a:gsLst></a:gradFill></a:rPr>`,
    );
    const before = serializeXml(doc);
    expect(() =>
      applyRunFormat(doc.root, {
        textFill: {
          kind: 'pattern',
          preset: 'not-a-pattern' as PatternPreset,
          foreground: '#112233',
          background: '#FFFFFF',
        },
      }),
    ).toThrow();
    expect(serializeXml(doc)).toBe(before);
  });

  it('updates pattern transforms while preserving the existing base colors', () => {
    const doc = parseXml(
      `<a:rPr xmlns:a="${NS.dml}"><a:pattFill prst="pct5"><a:fgClr><a:schemeClr val="accent1"/></a:fgClr><a:bgClr><a:schemeClr val="tx1"/></a:bgClr></a:pattFill></a:rPr>`,
    );
    applyRunFormat(doc.root, {
      textFill: {
        kind: 'pattern',
        preset: 'pct5',
        foreground: 'accent1',
        background: 'tx1',
        foregroundTransforms: [{ kind: 'lumMod', value: 0.5 }],
      },
    });
    const xml = serializeXml(doc);
    expect(xml).toContain('val="accent1"');
    expect(xml).toContain('val="tx1"');
    expect(xml).toContain('<a:lumMod val="50000"');
  });

  it('reads percent gradient positions and boolean gradient attributes', () => {
    const doc = parseXml(
      `<a:rPr xmlns:a="${NS.dml}"><a:gradFill rotWithShape="false"><a:gsLst><a:gs pos="21%"><a:srgbClr val="112233"/></a:gs><a:gs pos="100000"><a:srgbClr val="FFFFFF"/></a:gs></a:gsLst><a:lin ang="0" scaled="true"/></a:gradFill></a:rPr>`,
    );
    const format = parseRPrLikeElement(doc.root);
    expect(format.textFill?.kind).toBe('gradient');
    if (format.textFill?.kind === 'gradient') {
      expect(format.textFill.stops.map((stop) => stop.offset)).toEqual([0.21, 1]);
      expect(format.textFill.scaled).toBe(true);
      expect(format.textFill.rotateWithShape).toBe(false);
    }
  });

  it('rejects an invalid text fill before changing other run properties', () => {
    const doc = parseXml(`<a:rPr xmlns:a="${NS.dml}" sz="1200"/>`);
    const before = serializeXml(doc);
    expect(() =>
      applyRunFormat(doc.root, {
        size: 24,
        textFill: {
          kind: 'gradient',
          stops: [
            { offset: 0, color: '#112233' },
            { offset: 1, color: 'bad-color' as '#FFFFFF' },
          ],
        },
      }),
    ).toThrow();
    expect(serializeXml(doc)).toBe(before);
  });

  it('rejects invalid gradient opacity before changing other run properties', () => {
    const doc = parseXml(`<a:rPr xmlns:a="${NS.dml}" sz="1200"/>`);
    const before = serializeXml(doc);
    expect(() =>
      applyRunFormat(doc.root, {
        size: 24,
        textFill: {
          kind: 'gradient',
          stops: [
            { offset: 0, color: '#112233' },
            { offset: 1, color: '#FFFFFF', opacity: 2 },
          ],
        },
      }),
    ).toThrow();
    expect(serializeXml(doc)).toBe(before);
  });

  it.each(fills)('replaces an existing fill when setting a font color: %s', (fill) => {
    const doc = parseXml(
      `<a:rPr xmlns:a="${NS.dml}" b="1"><a:ln w="12700"/>${fill}<a:latin typeface="Arial"/></a:rPr>`,
    );
    applyRunFormat(doc.root, { color: '#ABCDEF' });
    const restored = parseXml(serializeXml(doc)).root;
    expect(
      restored.children
        .filter((node) => node.kind === 'element' && fillNames.has(node.name.localName))
        .map((node) => node.kind === 'element' && node.name.localName),
    ).toEqual(['solidFill']);
    expect(serializeXml(doc)).toContain('val="ABCDEF"');
    expect(serializeXml(doc)).toContain('b="1"');
    expect(serializeXml(doc)).toContain('<a:ln w="12700"/>');
    expect(serializeXml(doc)).toContain('<a:latin typeface="Arial"/>');
  });

  it.each(fills)('clears the local fill for inherited font color: %s', (fill) => {
    const doc = parseXml(`<a:rPr xmlns:a="${NS.dml}">${fill}</a:rPr>`);
    applyRunFormat(doc.root, { color: null });
    expect(
      doc.root.children.filter(
        (node) => node.kind === 'element' && fillNames.has(node.name.localName),
      ),
    ).toEqual([]);
  });
});
