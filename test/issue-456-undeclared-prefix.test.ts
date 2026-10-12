import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import {
  addBlankSlide,
  addSlideImage,
  createPresentation,
  duplicateSlide,
  emu,
  getSlides,
  getSlideShapes,
  loadPresentation,
  removeSlide,
  savePresentation,
} from '../src/api/index.ts';
import { attr, elem, parseXml, qname, serializeXml } from '../src/internal/xml/index.ts';

const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PNG = Uint8Array.from(
  atob(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==',
  ),
  (c) => c.charCodeAt(0),
);

const editPart = (bytes: Uint8Array, name: string, edit: (xml: string) => string) => {
  const files = unzipSync(bytes);
  files[name] = strToU8(edit(strFromU8(files[name]!)));
  return zipSync(files);
};

const deckWithOneSlide = async () => {
  const pres = createPresentation();
  addBlankSlide(pres);
  return savePresentation(pres);
};

// Namespaces in XML lets a producer declare a prefix on the element that uses
// it rather than on the part root.
const declareRLocally = (xml: string) =>
  xml.replace(` xmlns:r="${R}"`, '').replace(/<p:(sldMasterId|sldId) /g, `<p:$1 xmlns:r="${R}" `);

describe('issue-456: undeclared r: prefix on new elements', () => {
  it('duplicateSlide declares r: when presentation.xml declares it per element', async () => {
    const bytes = editPart(await deckWithOneSlide(), 'ppt/presentation.xml', declareRLocally);
    const pres = await loadPresentation(bytes);
    duplicateSlide(pres, getSlides(pres)[0]!);
    const reloaded = await loadPresentation(await savePresentation(pres));
    expect(getSlides(reloaded)).toHaveLength(2);
  });

  it('removeSlide after duplicateSlide re-parses presentation.xml cleanly', async () => {
    const bytes = editPart(await deckWithOneSlide(), 'ppt/presentation.xml', declareRLocally);
    const pres = await loadPresentation(bytes);
    duplicateSlide(pres, getSlides(pres)[0]!);
    removeSlide(pres, getSlides(pres)[0]!);
    expect(getSlides(pres)).toHaveLength(1);
  });

  it('addSlideImage declares r: on a slide whose root omits it', async () => {
    const bytes = editPart(await deckWithOneSlide(), 'ppt/slides/slide1.xml', (xml) =>
      xml.replace(` xmlns:r="${R}"`, ''),
    );
    const pres = await loadPresentation(bytes);
    addSlideImage(getSlides(pres)[0]!, PNG, {
      x: emu(0),
      y: emu(0),
      w: emu(914400),
      h: emu(914400),
    });
    const reloaded = await loadPresentation(await savePresentation(pres));
    expect(getSlideShapes(getSlides(reloaded)[0]!)).toHaveLength(1);
  });
});

describe('serializeXml namespace scope', () => {
  const doc = (root: ReturnType<typeof elem>) => ({
    kind: 'document' as const,
    decl: null,
    prolog: [],
    root,
    epilog: [],
  });

  it('declares a prefix only where it is not already in scope', () => {
    const child = elem(qname('p', 'c', 'urn:p'), {
      attrs: [attr(qname('r', 'id', 'urn:r'), 'rId1')],
    });
    const root = elem(qname('p', 'root', 'urn:p'), {
      prefixDecls: new Map([['p', 'urn:p']]),
      children: [child, child],
    });
    const out = serializeXml(doc(root), { standaloneDeclaration: false });
    expect(out).toBe(
      '<p:root xmlns:p="urn:p"><p:c xmlns:r="urn:r" r:id="rId1"/><p:c xmlns:r="urn:r" r:id="rId1"/></p:root>',
    );
    expect(() => parseXml(out)).not.toThrow();
  });

  it('redeclares a prefix that is bound to a different URI in scope', () => {
    const root = elem(qname('x', 'root', 'urn:one'), {
      prefixDecls: new Map([['x', 'urn:one']]),
      children: [elem(qname('x', 'c', 'urn:two'))],
    });
    expect(serializeXml(doc(root), { standaloneDeclaration: false })).toBe(
      '<x:root xmlns:x="urn:one"><x:c xmlns:x="urn:two"/></x:root>',
    );
  });

  it('declares the default namespace for an unprefixed element in a namespace', () => {
    const root = elem(qname('', 'Types', 'urn:ct'), {
      children: [elem(qname('', 'Default', 'urn:ct'))],
    });
    expect(serializeXml(doc(root), { standaloneDeclaration: false })).toBe(
      '<Types xmlns="urn:ct"><Default/></Types>',
    );
  });

  it('rejects one prefix bound to two URIs on the same element', () => {
    const root = elem(qname('x', 'root', 'urn:one'), {
      attrs: [attr(qname('x', 'a', 'urn:two'), 'v')],
    });
    expect(() => serializeXml(doc(root))).toThrow(/prefix "x"/);
  });
});
