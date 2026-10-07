import assert from 'node:assert/strict';
import { readdir, readFile } from 'node:fs/promises';
import test from 'node:test';
import {
  addBlankSlide,
  createPresentation,
  getSlideTransition,
  getSlideXmlString,
  setSlideTransition,
} from '@office-kit/pptx';
import {
  shownDurationMs,
  TRANSITION_TILES,
  transitionSpeed,
} from '../src/lib/editor/ribbon/transition-gallery.ts';

// Mac PowerPoint 16.113.3's own save of each gallery tile, in gallery order,
// headed by the Duration its ribbon showed after applying it.
const CAPTURES = new URL('../../test/fixtures/native/transitions/', import.meta.url);
const NO_TRANSITION = '(no transition)';

const captures = async () => {
  const names = (await readdir(CAPTURES)).filter((name) => name.endsWith('.xml')).sort();
  return Promise.all(
    names.map(async (name) => {
      const text = await readFile(new URL(name, CAPTURES), 'utf8');
      const header =
        /^<!-- Transition "([^"]+)".*ribbon Duration after apply: (\d+\.\d+) -->\n/.exec(text);
      assert.ok(header, name);
      return {
        name: header[1],
        ribbonMs: Math.round(Number(header[2]) * 1000),
        xml: text.slice(header[0].length).trim(),
      };
    }),
  );
};

// Where a namespace is declared is the serializer's choice (PowerPoint puts
// p14 on the p:transition beside a p15 / p159 effect, the library on the
// mc:Choice); everything else must match.
const normalise = (xml) =>
  xml
    .replace(/\s+xmlns:[\w]+="[^"]*"/g, '')
    .replace(/\s+/g, ' ')
    .replace(/ (\/?>)/g, '$1');

const transitionXml = (slide) => {
  const xml = getSlideXmlString(slide);
  const found =
    /<mc:AlternateContent\b[^>]*><mc:Choice\b[\s\S]*?<\/mc:AlternateContent>/.exec(xml) ??
    /<p:transition\b[^>]*\/>|<p:transition\b[\s\S]*?<\/p:transition>/.exec(xml);
  return found === null ? NO_TRANSITION : found[0];
};

test('every gallery tile writes what Mac PowerPoint writes for it', async () => {
  const native = await captures();
  assert.equal(native.length, TRANSITION_TILES.length);
  const pres = createPresentation();
  for (const [index, tile] of TRANSITION_TILES.entries()) {
    const capture = native[index];
    assert.equal(tile.en, capture.name, `tile ${index}`);
    const slide = addBlankSlide(pres);
    // What choosing the tile writes (TransitionsRibbon's chooseTile); None on
    // a slide without a transition leaves it without one.
    if (tile.choice.effect !== 'none')
      setSlideTransition(slide, { ...tile.choice, durationMs: tile.durationMs });
    assert.equal(normalise(transitionXml(slide)), normalise(capture.xml), tile.en);
    assert.equal(tile.durationMs, capture.ribbonMs, `${tile.en} duration`);
    assert.equal(shownDurationMs(getSlideTransition(slide)), capture.ribbonMs, `${tile.en} ribbon`);
  }
});

test('a Duration edit writes spd and p14:dur by the rule PowerPoint’s defaults follow', () => {
  const slide = addBlankSlide(createPresentation());
  const written = (durationMs) => {
    setSlideTransition(slide, { effect: 'push', direction: 'u', durationMs });
    return {
      xml: normalise(transitionXml(slide)),
      shown: shownDurationMs(getSlideTransition(slide)),
    };
  };
  // A speed's own duration is the speed alone; fast is the schema default.
  assert.deepEqual(written(1000), {
    xml: '<p:transition spd="slow"><p:push dir="u"/></p:transition>',
    shown: 1000,
  });
  assert.deepEqual(written(750), {
    xml: '<p:transition spd="med"><p:push dir="u"/></p:transition>',
    shown: 750,
  });
  assert.deepEqual(written(500), {
    xml: '<p:transition><p:push dir="u"/></p:transition>',
    shown: 500,
  });
  // Anything else is the next speed up plus p14:dur.
  assert.equal(
    written(760).xml,
    '<mc:AlternateContent><mc:Choice Requires="p14"><p:transition spd="slow" p14:dur="760"><p:push dir="u"/></p:transition></mc:Choice><mc:Fallback><p:transition spd="slow"><p:push dir="u"/></p:transition></mc:Fallback></mc:AlternateContent>',
  );
  assert.equal(written(760).shown, 760);
  assert.match(written(501).xml, /<p:transition spd="med" p14:dur="501">/);
  assert.match(written(250).xml, /<p:transition p14:dur="250">/);
});

test('a transition without spd runs at the schema default, fast', () => {
  // Cut as PowerPoint saves it (p14:dur only), and a bare transition.
  assert.equal(transitionSpeed({ effect: 'cut', durationMs: 100 }), 'fast');
  assert.equal(transitionSpeed({ effect: 'fade' }), 'fast');
  assert.equal(shownDurationMs({ effect: 'fade' }), 500);
  assert.equal(transitionSpeed({ effect: 'pull', speed: 'med' }), 'med');
  assert.equal(shownDurationMs({ effect: 'pull', speed: 'med' }), 750);
  assert.equal(shownDurationMs(null), 2000);
});
