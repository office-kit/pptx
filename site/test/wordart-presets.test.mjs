// The WordArt gallery's presets against the run XML Mac PowerPoint wrote for
// each one (test/fixtures/native/wordart-*-shape.xml).

import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import test from 'node:test';
import {
  addBlankSlide,
  addSlideTextBox,
  createPresentation,
  getShapeXmlString,
  inches,
} from '@office-kit/pptx';
import { applyWordArtPreset, WORDART_PRESETS } from '../src/lib/editor/core/wordart-presets.ts';

const fixture = (name) =>
  readFile(
    new URL(`../../test/fixtures/native/wordart-${name}-shape.xml`, import.meta.url),
    'utf8',
  );

// Gallery order, as wordart-capture.md lists it.
const FIXTURES = [
  'black-shadow',
  'accent1-shadow',
  'accent2-outline',
  'white-accent5-shadow',
  'gray-gradient',
  'accent4-soft-bevel',
  'accent5-gradient-reflection',
  'accent4-gradient-outline',
  'white-accent1-glow',
  'accent3-sharp-bevel',
  'black-white-hard-shadow',
  'black-accent5-hard-shadow',
  'accent5-hard-shadow',
  'white-accent2-hard-shadow',
  'background2-inner-shadow',
  'white-pattern-shadow',
  'accent3-pattern-inner-shadow',
  'accent1-pattern-hard-shadow',
  'accent5-pattern-outline',
  'dark-blue-pattern-hard-shadow',
];

const DML = 'http://schemas.openxmlformats.org/drawingml/2006/main';

// Just enough XML for character properties: elements and attributes, no text.
function parse(xml) {
  const root = { children: [] };
  const stack = [root];
  for (const [, close, name, rawAttrs, selfClosing] of xml.matchAll(
    /<(\/?)([\w:]+)((?:\s+[\w:]+="[^"]*")*)\s*(\/?)>/g,
  )) {
    if (close) {
      stack.pop();
      continue;
    }
    const [prefix, localName] = name.includes(':') ? name.split(':') : ['', name];
    const element = {
      kind: 'element',
      name: { prefix, localName, namespaceURI: prefix === 'a' ? DML : '' },
      attrs: [...rawAttrs.matchAll(/([\w:]+)="([^"]*)"/g)].map(([, key, value]) => ({
        name: { prefix: '', localName: key, namespaceURI: '' },
        value,
      })),
      prefixDecls: new Map(),
      children: [],
    };
    stack.at(-1).children.push(element);
    if (!selfClosing) stack.push(element);
  }
  return root.children[0];
}

const properties = (xml, tag) => {
  const match = xml.match(new RegExp(`<a:${tag}\\b[^>]*?(?:/>|>[\\s\\S]*?</a:${tag}>)`));
  assert.ok(match, `no a:${tag}`);
  return parse(match[0]);
};

const attrsOf = (element, defaults = {}) => {
  const out = { ...defaults };
  for (const a of element.attrs) out[a.name.localName] = a.value;
  return out;
};

// Schema defaults, so an attribute the library spells out and native omits
// (or the reverse) compares equal.
const OUTER_SHADOW = {
  blurRad: '0',
  dist: '0',
  dir: '0',
  sx: '100000',
  sy: '100000',
  kx: '0',
  ky: '0',
  algn: 'b',
  rotWithShape: '1',
};
const INNER_SHADOW = { blurRad: '0', dist: '0', dir: '0' };
const REFLECTION = {
  blurRad: '0',
  stA: '100000',
  stPos: '0',
  endA: '0',
  endPos: '100000',
  dist: '0',
  dir: '0',
  fadeDir: '5400000',
  sx: '100000',
  sy: '100000',
  kx: '0',
  ky: '0',
  algn: 'b',
  rotWithShape: '1',
};
const EFFECT_DEFAULTS = {
  outerShdw: OUTER_SHADOW,
  innerShdw: INNER_SHADOW,
  reflection: REFLECTION,
};

function normalize(rPr) {
  const exact = (color) => ({
    [color.name.localName]: attrsOf(color).val,
    transforms: color.children.map((t) => [t.name.localName, attrsOf(t).val]),
  });
  const child = (element) => {
    const local = element.name.localName;
    if (local === 'ln') {
      const { cmpd, ...attrs } = attrsOf(element);
      // `cmpd="sng"` and `<a:prstDash val="solid"/>` are the schema defaults,
      // and the text outline API has no way to spell them.
      if (cmpd !== undefined) assert.equal(cmpd, 'sng');
      return {
        ln: attrs,
        children: element.children
          .filter((c) => !(c.name.localName === 'prstDash' && attrsOf(c).val === 'solid'))
          .map((c) => ({ [c.name.localName]: exact(c.children[0]) })),
      };
    }
    if (local === 'solidFill') return { solidFill: exact(element.children[0]) };
    if (local === 'gradFill') {
      const [gsLst, lin] = element.children;
      return {
        gradFill: gsLst.children.map((gs) => [attrsOf(gs).pos, exact(gs.children[0])]),
        lin: attrsOf(lin).ang,
      };
    }
    if (local === 'pattFill') {
      return {
        pattFill: attrsOf(element).prst,
        colors: element.children.map((c) => [c.name.localName, exact(c.children[0])]),
      };
    }
    if (local === 'effectLst') {
      return {
        effectLst: element.children.map((effect) => ({
          [effect.name.localName]: attrsOf(effect, EFFECT_DEFAULTS[effect.name.localName]),
          color: effect.children.map(exact),
        })),
      };
    }
    return { [local]: attrsOf(element) };
  };
  const { b, spc } = attrsOf(rPr);
  return { b, spc, children: rPr.children.map(child) };
}

// The text body's 3-D, verbatim: the bevels' `<a:scene3d>` and `<a:sp3d>`.
const body3D = (xml) =>
  ['scene3d', 'sp3d'].map(
    (tag) =>
      xml
        .replace(/>\s+</g, '><')
        .replace(/\s+\/>/g, '/>')
        .match(new RegExp(`<a:${tag}\\b[^>]*?(?:/>|>.*?</a:${tag}>)`))?.[0] ?? null,
  );

function deck() {
  const pres = createPresentation();
  const slide = addBlankSlide(pres);
  const box = () =>
    addSlideTextBox(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(6),
      h: inches(1),
      text: 'Outline title',
    });
  return { box };
}

test('the gallery lists the twenty native presets in order', () => {
  assert.equal(WORDART_PRESETS.length, FIXTURES.length);
});

for (const [index, name] of FIXTURES.entries()) {
  const preset = WORDART_PRESETS[index];
  test(`${preset.label} writes what PowerPoint wrote`, async () => {
    const { box } = deck();
    const shape = box();
    applyWordArtPreset(shape, preset);
    const written = getShapeXmlString(shape);
    const native = await fixture(name);
    const expected = normalize(properties(native, 'rPr'));
    assert.deepEqual(normalize(properties(written, 'rPr')), expected);
    // Native styles the paragraph end the same way, so typing continues in it.
    assert.deepEqual(normalize(properties(written, 'endParaRPr')), expected);
    assert.deepEqual(body3D(written), body3D(native));
  });
}

test('a preset without a bevel removes the previous preset’s', () => {
  const { box } = deck();
  const shape = box();
  applyWordArtPreset(shape, WORDART_PRESETS[FIXTURES.indexOf('accent3-sharp-bevel')]);
  applyWordArtPreset(shape, WORDART_PRESETS[FIXTURES.indexOf('black-shadow')]);
  assert.deepEqual(body3D(getShapeXmlString(shape)), [null, null]);
});

test('a preset replaces the previous one instead of merging with it', async () => {
  const { box } = deck();
  const shape = box();
  applyWordArtPreset(shape, WORDART_PRESETS[FIXTURES.indexOf('white-accent5-shadow')]);
  applyWordArtPreset(shape, WORDART_PRESETS[FIXTURES.indexOf('black-shadow')]);
  const rPr = normalize(properties(getShapeXmlString(shape), 'rPr'));
  // Native removes `b`; the library can only write it off.
  assert.equal(rPr.b, '0');
  assert.deepEqual(
    { ...rPr, b: undefined },
    normalize(properties(await fixture('black-shadow'), 'rPr')),
  );
});
