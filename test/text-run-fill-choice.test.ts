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

  it('preserves the existing fill when an invalid color is rejected', () => {
    const doc = parseXml(`<a:rPr xmlns:a="${NS.dml}">${fills[2]}</a:rPr>`);
    const before = serializeXml(doc);
    expect(() => applyRunFormat(doc.root, { color: 'invalid' as '#FFFFFF' })).toThrow();
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
