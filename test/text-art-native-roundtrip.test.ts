// The reference desktop app's twenty text-art presets (test/fixtures/native/text-art-*-shape.xml)
// read through the run and text-body readers and written back through the
// public setters: every color — theme tints included — and the bevel 3-D come
// out as the reference desktop app wrote them.

import { readdirSync, readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import {
  addBlankSlide,
  addSlideTextBox,
  asColor,
  createPresentation,
  getShapeText3D,
  getShapeXmlString,
  inches,
  setShapeText3D,
  setShapeTextFormat,
  toWritableTextFormat,
  type ReadText3D,
  type Text3D,
} from '../src/api/index.ts';
import { parseRPrLikeElement } from '../src/api/fn/shape-color.ts';
import { readText3D } from '../src/internal/drawingml/index.ts';
import {
  NS,
  type XmlElement,
  firstChildElement,
  parseXml,
  qname,
} from '../src/internal/xml/index.ts';

const NATIVE = new URL('./fixtures/native/', import.meta.url);
const FIXTURES = readdirSync(NATIVE).filter((name) => /^text-art-.*-shape\.xml$/.test(name));

const compact = (xml: string) =>
  xml
    .replace(/>\s+</g, '><')
    .replace(/\s+\/>/g, '/>')
    .replace(/ xmlns:\w+="[^"]*"/g, '');

const find = (root: XmlElement, local: string): XmlElement => {
  const stack = [root];
  while (stack.length) {
    const node = stack.pop()!;
    if (node.name.localName === local) return node;
    for (const child of node.children) if (child.kind === 'element') stack.push(child);
  }
  throw new Error(`no a:${local}`);
};

// Every color element, transforms included, in document order.
const colors = (xml: string) =>
  [...compact(xml).matchAll(/<a:(srgbClr|schemeClr)\b[^>]*?(?:\/>|>.*?<\/a:\1>)/g)].map(
    (match) => match[0],
  );
const element = (xml: string, tag: string) =>
  compact(xml).match(new RegExp(`<a:${tag}\\b[^>]*?(?:/>|>.*?</a:${tag}>)`))?.[0] ?? '';

const writable = (read: ReadText3D): Text3D => {
  const { contourColor, extrusionColor, ...rest } = read;
  return {
    ...rest,
    ...(contourColor === undefined ? {} : { contourColor: asColor(contourColor)! }),
    ...(extrusionColor === undefined ? {} : { extrusionColor: asColor(extrusionColor)! }),
  };
};

describe('native text-art payloads round trip', () => {
  it('covers all twenty gallery captures', () => {
    expect(FIXTURES).toHaveLength(20);
  });

  for (const name of FIXTURES) {
    it(name, () => {
      const native = readFileSync(new URL(name, NATIVE), 'utf8');
      const root = parseXml(native).root;
      const nativeRPr = find(root, 'rPr');
      const nativeBodyPr = firstChildElement(
        firstChildElement(root, qname('p', 'txBody', NS.pml))!,
        qname('a', 'bodyPr', NS.dml),
      )!;

      const shape = addSlideTextBox(addBlankSlide(createPresentation()), {
        x: inches(1),
        y: inches(1),
        w: inches(6),
        h: inches(1),
        text: 'Outline title',
      });
      setShapeTextFormat(shape, toWritableTextFormat(parseRPrLikeElement(nativeRPr)));
      const native3D = readText3D(nativeBodyPr);
      if (native3D) setShapeText3D(shape, writable(native3D));

      const written = getShapeXmlString(shape);
      expect(colors(element(written, 'rPr'))).toEqual(colors(element(native, 'rPr')));
      expect(getShapeText3D(shape)).toEqual(native3D);
      for (const tag of ['scene3d', 'sp3d'])
        expect(element(written, tag)).toBe(element(native, tag));
    });
  }
});
