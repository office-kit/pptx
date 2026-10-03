// `getShapeEffects` — returns every effect on the shape's `<a:effectLst>`
// in document order, not just the first one. Renderers need the full
// list because PowerPoint composes shadow + glow + softEdge into a
// single filter stack.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideShape,
  getShapeEffect,
  getShapeEffects,
  getShapeEffectsEffective,
  getSlideShapes,
  getSlides,
  inches,
  loadPresentation,
  savePresentation,
  setShapeGlow,
  setShapeShadow,
} from '../src/api/index.ts';
import { readZip, writeZip } from '../src/internal/opc/index.ts';
import { getShapePlaceholderIdx } from '../src/api/fn/shape-read-base.ts';
import { LAYOUT_PART, SHAPE_ELEMENT } from '../src/api/_internal-symbols.ts';
import { getSlideLayout } from '../src/api/fn/shape-slide-read.ts';
import { firstChildElement, NS, parseXml, qname } from '../src/internal/xml/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getShapeEffects', () => {
  it('does not treat an authored effectDag as inherited effects', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = getSlideShapes(slide).find(
      (candidate) => getShapePlaceholderIdx(candidate) !== null,
    );
    expect(shape).toBeDefined();
    const layout = getSlideLayout(slide);
    expect(layout).not.toBeNull();
    const layoutShape = layout![LAYOUT_PART].shapes.find(
      (candidate) => candidate.placeholderIdx === getShapePlaceholderIdx(shape!),
    );
    expect(layoutShape).toBeDefined();
    const layoutSpPr = firstChildElement(layoutShape!.element, qname('p', 'spPr', NS.pml));
    layoutSpPr!.children.push(
      parseXml(
        '<a:effectLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:outerShdw blurRad="40000" dist="20000" dir="5400000"><a:srgbClr val="112233"/></a:outerShdw></a:effectLst>',
      ).root,
    );
    const spPr = firstChildElement(shape![SHAPE_ELEMENT], qname('p', 'spPr', NS.pml));
    expect(spPr).not.toBeNull();
    spPr!.children.push(
      parseXml('<a:effectDag xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>')
        .root,
    );
    expect(getShapeEffectsEffective(pres, shape!)).toEqual([]);
  });

  it('does not inherit through an unresolved authored effectRef', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const placeholder = getSlideShapes(slide).find(
      (shape) => getShapePlaceholderIdx(shape) !== null,
    );
    expect(placeholder).toBeDefined();
    const layout = getSlideLayout(slide);
    expect(layout).not.toBeNull();
    const layoutShape = layout![LAYOUT_PART].shapes.find(
      (candidate) => candidate.placeholderIdx === getShapePlaceholderIdx(placeholder!),
    );
    const layoutSpPr = firstChildElement(layoutShape!.element, qname('p', 'spPr', NS.pml));
    layoutSpPr!.children.push(
      parseXml(
        '<a:effectLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:outerShdw blurRad="40000" dist="20000" dir="5400000"><a:srgbClr val="112233"/></a:outerShdw></a:effectLst>',
      ).root,
    );
    placeholder![SHAPE_ELEMENT].children.push(
      parseXml(
        '<p:style xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:effectRef idx="999999"/></p:style>',
      ).root,
    );
    expect(getShapeEffectsEffective(pres, placeholder!)).toEqual([]);
  });
  it('stops at an explicit empty layout effect list before reaching the master', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideShape(slide, {
      preset: 'rect',
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(2),
    });
    const { entries } = readZip(await savePresentation(pres));
    const decoder = new TextDecoder();
    const encoder = new TextEncoder();
    const layoutEffect =
      '<a:effectLst><a:outerShdw blurRad="40000" dist="20000" dir="5400000">' +
      '<a:srgbClr val="112233"/></a:outerShdw></a:effectLst>';
    const masterEffect =
      '<a:effectLst><a:outerShdw blurRad="40000" dist="20000" dir="5400000">' +
      '<a:srgbClr val="445566"/></a:outerShdw></a:effectLst>';
    const insertIntoPlaceholder = (xml: string, idx: string, effect: string): string => {
      const re = /<p:sp>[\s\S]*?<\/p:sp>/g;
      return xml.replace(re, (sp) => {
        if (!new RegExp(`<p:ph[^>]*idx="${idx}"[^>]*>`).test(sp)) return sp;
        return sp.includes('<p:spPr/>')
          ? sp.replace('<p:spPr/>', `<p:spPr>${effect}</p:spPr>`)
          : sp.replace('</p:spPr>', `${effect}</p:spPr>`);
      });
    };
    const patched = entries.map((entry) => {
      if (entry.name === 'ppt/slideLayouts/slideLayout2.xml') {
        return {
          ...entry,
          data: encoder.encode(
            insertIntoPlaceholder(decoder.decode(entry.data), '1', layoutEffect),
          ),
        };
      }
      if (entry.name === 'ppt/slideMasters/slideMaster1.xml') {
        return {
          ...entry,
          data: encoder.encode(
            insertIntoPlaceholder(decoder.decode(entry.data), '1', masterEffect),
          ),
        };
      }
      return entry;
    });
    expect(
      decoder.decode(
        patched.find((entry) => entry.name === 'ppt/slideLayouts/slideLayout2.xml')!.data,
      ),
    ).toContain(layoutEffect);
    expect(
      decoder.decode(
        patched.find((entry) => entry.name === 'ppt/slideLayouts/slideLayout2.xml')!.data,
      ),
    ).toContain(`<p:spPr>${layoutEffect}</p:spPr>`);
    const imported = await loadPresentation(writeZip(patched));
    const importedShape = getSlideShapes(getSlides(imported)[0]!)[1]!;
    expect(getShapePlaceholderIdx(importedShape)).toBe(1);
    expect(getShapeEffectsEffective(imported, importedShape)).toMatchObject([
      { kind: 'outerShdw', color: '#112233' },
    ]);

    const withEmptyLayout = patched.map((entry) => {
      if (entry.name !== 'ppt/slideLayouts/slideLayout2.xml') return entry;
      const xml = decoder.decode(entry.data).replace(layoutEffect, '<a:effectLst/>');
      return { ...entry, data: encoder.encode(xml) };
    });
    const emptyLayout = await loadPresentation(writeZip(withEmptyLayout));
    const emptyShape = getSlideShapes(getSlides(emptyLayout)[0]!)[1]!;
    expect(getShapeEffectsEffective(emptyLayout, emptyShape)).toEqual([]);
  });

  it('returns an empty array when no effects are set', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = addSlideShape(slide, {
      preset: 'rect',
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(2),
    });
    expect(getShapeEffects(pres, shape)).toEqual([]);
  });

  it('resolves a theme effectRef, including its phClr reference color', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideShape(slide, {
      preset: 'rect',
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(2),
      name: 'theme effect reference',
    });
    const { entries } = readZip(await savePresentation(pres));
    const decoder = new TextDecoder();
    const encoder = new TextEncoder();
    const style =
      '<p:style><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef>' +
      '<a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef>' +
      '<a:effectRef idx="1"><a:schemeClr val="accent1"/></a:effectRef>' +
      '<a:fontRef idx="minor"><a:schemeClr val="accent1"/></a:fontRef></p:style>';
    const patched = entries.map((entry) => {
      if (entry.name === 'ppt/slides/slide1.xml') {
        const xml = decoder.decode(entry.data);
        const end = xml.lastIndexOf('</p:sp>');
        expect(end).toBeGreaterThan(0);
        return { ...entry, data: encoder.encode(xml.slice(0, end) + style + xml.slice(end)) };
      }
      if (entry.name === 'ppt/theme/theme1.xml') {
        const xml = decoder.decode(entry.data);
        return {
          ...entry,
          data: encoder.encode(
            xml.replace(
              '<a:srgbClr val="000000"><a:alpha val="38000"/></a:srgbClr>',
              '<a:schemeClr val="phClr"><a:alpha val="38000"/></a:schemeClr>',
            ),
          ),
        };
      }
      return entry;
    });
    const imported = await loadPresentation(writeZip(patched));
    const importedShape = getSlideShapes(getSlides(imported)[0]!).at(-1)!;
    expect(getShapeEffects(imported, importedShape)).toMatchObject([
      { kind: 'outerShdw', color: '#4F81BD', blurEmu: 40000, distEmu: 20000 },
    ]);
  });

  it('treats effectRef idx 0 as an explicit empty effect list', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideShape(slide, {
      preset: 'rect',
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(2),
    });
    const { entries } = readZip(await savePresentation(pres));
    const decoder = new TextDecoder();
    const encoder = new TextEncoder();
    const patched = entries.map((entry) => {
      if (entry.name !== 'ppt/slides/slide1.xml') return entry;
      const xml = decoder.decode(entry.data);
      const end = xml.lastIndexOf('</p:sp>');
      const style =
        '<p:style><a:lnRef idx="1"/><a:fillRef idx="1"/>' +
        '<a:effectRef idx="0"/><a:fontRef idx="minor"/></p:style>';
      return { ...entry, data: encoder.encode(xml.slice(0, end) + style + xml.slice(end)) };
    });
    const imported = await loadPresentation(writeZip(patched));
    const importedShape = getSlideShapes(getSlides(imported)[0]!).at(-1)!;
    expect(getShapeEffects(imported, importedShape)).toEqual([]);
  });

  it('reads the outer shadow set by setShapeShadow with all numeric fields', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = addSlideShape(slide, {
      preset: 'rect',
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(2),
    });
    setShapeShadow(shape, { color: '#000000', angleDeg: 45, opacity: 0.5 });
    const effects = getShapeEffects(pres, shape);
    expect(effects).toHaveLength(1);
    expect(effects[0]!.kind).toBe('outerShdw');
    if (effects[0]!.kind === 'outerShdw') {
      expect(effects[0]!.color).toBe('#000000');
      expect(effects[0]!.angleDeg).toBeCloseTo(45);
      expect(effects[0]!.opacity).toBeCloseTo(0.5, 3);
      expect(effects[0]!.blurEmu).toBeGreaterThan(0);
    }
  });

  it('reads the glow set by setShapeGlow', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = addSlideShape(slide, {
      preset: 'rect',
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(2),
    });
    setShapeGlow(shape, { color: '#FF0000', radiusEmu: 63500 });
    const effects = getShapeEffects(pres, shape);
    expect(effects).toHaveLength(1);
    expect(effects[0]!.kind).toBe('glow');
    if (effects[0]!.kind === 'glow') {
      expect(effects[0]!.color).toBe('#FF0000');
      expect(effects[0]!.radiusEmu).toBe(63500);
    }
  });

  it('reads scRGB and HSL effect colors through the public reader', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = addSlideShape(slide, {
      preset: 'rect',
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(2),
    });
    setShapeGlow(shape, { color: '#FF0000', radiusEmu: 100 });
    setShapeShadow(shape, { color: '#0000FF', angleDeg: 0, opacity: 1 });
    const zip = readZip(await savePresentation(pres));
    const slideEntry = zip.entries.find((entry) => entry.name === 'ppt/slides/slide1.xml')!;
    let xml = new TextDecoder().decode(slideEntry.data);
    xml = xml.replace('<a:srgbClr val="FF0000"/>', '<a:scrgbClr r="50%" g="50%" b="50%"/>');
    xml = xml.replace(
      '<a:srgbClr val="0000FF"/>',
      '<a:hslClr hue="7200000" sat="100%" lum="50%"/>',
    );
    const imported = await loadPresentation(
      writeZip(
        zip.entries.map((entry) =>
          entry.name === slideEntry.name
            ? { ...entry, data: new TextEncoder().encode(xml) }
            : entry,
        ),
      ),
    );
    const importedShape = getSlideShapes(getSlides(imported)[0]!).find(
      (candidate) => getShapeEffects(imported, candidate).length > 0,
    )!;
    const effects = getShapeEffects(imported, importedShape);
    expect(effects).toMatchObject([
      { kind: 'glow', color: '#BCBCBC' },
      { kind: 'outerShdw', color: '#00FF00' },
    ]);
    expect(getShapeEffect(importedShape)).toMatchObject({ kind: 'shadow', color: '#00FF00' });

    const reloaded = await loadPresentation(await savePresentation(imported));
    const restoredShape = getSlideShapes(getSlides(reloaded)[0]!).find(
      (candidate) => getShapeEffects(reloaded, candidate).length > 0,
    )!;
    expect(getShapeEffects(reloaded, restoredShape)).toMatchObject([
      { kind: 'glow', color: '#BCBCBC' },
      { kind: 'outerShdw', color: '#00FF00' },
    ]);
    expect(getShapeEffect(restoredShape)).toMatchObject({ kind: 'shadow', color: '#00FF00' });
  });
});
