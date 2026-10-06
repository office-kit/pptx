import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdir, mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import {
  addSlideImage,
  addSlideTextBox,
  bringShapeToFront,
  getShapeHyperlink,
  getShapeImagePartName,
  getShapeName,
  getSlides,
  getSlideShapes,
  getShapeText,
  loadPresentation,
  removeShape,
  savePresentation,
  setShapeHyperlink,
  setShapeText,
} from '@office-kit/pptx';
import { buildDeck } from '../dist/index.mjs';
import { mergeDecks } from './helpers/deck-merge.mjs';

const EMU = 914400;
const PNG = Uint8Array.from(
  atob(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==',
  ),
  (c) => c.charCodeAt(0),
);
const deck = (a, b) =>
  `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={8} height={1}>${a}</Text><Text x={1} y={3} width={8} height={1}>${b}</Text></Slide><Slide><Text x={1} y={1} width={8} height={1}>Second</Text></Slide></Presentation>`;

let dir;
let builds = 0;
async function build(source) {
  dir ??= await mkdtemp(join(tmpdir(), 'office-merge-'));
  const folder = join(dir, String(builds++));
  await mkdir(folder);
  await writeFile(join(folder, 'deck.tsx'), source);
  return (await buildDeck(join(folder, 'deck.tsx'))).bytes;
}
const base = await build(deck('A', 'B'));

/** Applies `change` to the first slide's shapes and saves the deck. */
async function edit(bytes, change) {
  const pres = await loadPresentation(bytes);
  await change(getSlideShapes(getSlides(pres)[0]), getSlides(pres)[0]);
  return savePresentation(pres);
}
const texts = async (bytes) =>
  getSlideShapes(getSlides(await loadPresentation(bytes))[0]).map(getShapeText);

test.after(() => rm(dir, { recursive: true, force: true }));

test('merges different shapes changed on each side', async () => {
  const ours = await edit(base, ([a]) => setShapeText(a, 'A edited'));
  const theirs = await build(deck('A', 'B source'));
  const merged = mergeDecks(base, ours, theirs);
  assert.equal(merged.ok, true);
  assert.deepEqual(await texts(merged.bytes), ['A edited', 'B source']);
});

test('reports the same shape changed on both sides', async () => {
  const ours = await edit(base, ([a]) => setShapeText(a, 'A edited'));
  const theirs = await build(deck('A source', 'B'));
  assert.deepEqual(mergeDecks(base, ours, theirs), {
    ok: false,
    conflicts: [
      {
        part: 'ppt/slides/slide1.xml',
        slide: 1,
        shape: { id: '2', name: 'TextBox 2' },
        reason: 'both-changed',
      },
    ],
  });
});

test('keeps a shape added in the editor while the source changes another', async () => {
  const ours = await edit(base, (_, slide) =>
    addSlideTextBox(slide, { x: 0, y: 5 * EMU, w: 4 * EMU, h: EMU, text: 'Added' }),
  );
  const theirs = await build(deck('A', 'B source'));
  const merged = mergeDecks(base, ours, theirs);
  assert.equal(merged.ok, true);
  assert.deepEqual(await texts(merged.bytes), ['A', 'B source', 'Added']);
});

test('keeps the editor stacking order while merging source text', async () => {
  const ours = await edit(base, ([a]) => bringShapeToFront(a));
  const theirs = await build(deck('A source', 'B'));
  const merged = mergeDecks(base, ours, theirs);
  assert.equal(merged.ok, true);
  assert.deepEqual(await texts(merged.bytes), ['B', 'A source']);
});

test('reports a shape deleted on one side and changed on the other', async () => {
  const ours = await edit(base, ([, b]) => removeShape(b));
  const theirs = await build(deck('A', 'B source'));
  const merged = mergeDecks(base, ours, theirs);
  assert.equal(merged.ok, false);
  assert.deepEqual(merged.conflicts, [
    {
      part: 'ppt/slides/slide1.xml',
      slide: 1,
      shape: { id: '3', name: 'TextBox 3' },
      reason: 'deleted-and-changed',
    },
  ]);
});

test('reports different shapes added with the same ID on both sides', async () => {
  const add = (text) => (_, slide) =>
    addSlideTextBox(slide, { x: 0, y: 5 * EMU, w: 4 * EMU, h: EMU, text });
  const merged = mergeDecks(base, await edit(base, add('Ours')), await edit(base, add('Theirs')));
  assert.equal(merged.ok, false);
  assert.deepEqual(
    merged.conflicts.map(({ shape, reason }) => [shape?.id, reason]),
    [['4', 'both-added']],
  );
});

test('keeps media and relationships added in the editor', async () => {
  const linked = await edit(base, ([, b]) => setShapeHyperlink(b, 'https://example.com/a'));
  const ours = await edit(linked, (_, slide) =>
    addSlideImage(slide, PNG, { x: 0, y: 5 * EMU, w: EMU, h: EMU }),
  );
  // The source changes the same slide's relationships.
  const theirs = await edit(linked, ([, b]) => setShapeHyperlink(b, 'https://example.com/b'));
  const merged = mergeDecks(linked, ours, theirs);
  assert.equal(merged.ok, true);
  const shapes = getSlideShapes(getSlides(await loadPresentation(merged.bytes))[0]);
  assert.equal(getShapeHyperlink(shapes[1]), 'https://example.com/b');
  const image = getShapeImagePartName(shapes[2]);
  assert.ok(image);
  const parts = unzipSync(merged.bytes);
  assert.deepEqual(parts[image.replace(/^\//, '')], PNG);
  assert.match(strFromU8(parts['[Content_Types].xml']), /Extension="png"/);
});

test('renames a relationship the editor added under an ID the source also used', async () => {
  const ours = await edit(base, (_, slide) =>
    addSlideImage(slide, PNG, { x: 0, y: 5 * EMU, w: EMU, h: EMU }),
  );
  const theirs = await edit(base, ([, b]) => setShapeHyperlink(b, 'https://example.com/'));
  const merged = mergeDecks(base, ours, theirs);
  assert.equal(merged.ok, true);
  const shapes = getSlideShapes(getSlides(await loadPresentation(merged.bytes))[0]);
  assert.equal(getShapeHyperlink(shapes[1]), 'https://example.com/');
  assert.deepEqual(unzipSync(merged.bytes)[getShapeImagePartName(shapes[2]).slice(1)], PNG);
});

test('ignores generated core property timestamps', async () => {
  const stamp = (bytes, time) => {
    const parts = unzipSync(bytes);
    parts['docProps/core.xml'] = strToU8(
      strFromU8(parts['docProps/core.xml']).replace(
        '</cp:coreProperties>',
        `<dcterms:created xsi:type="dcterms:W3CDTF">${time}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${time}</dcterms:modified></cp:coreProperties>`,
      ),
    );
    return zipSync(parts);
  };
  const ours = stamp(
    await edit(base, ([a]) => setShapeText(a, 'A edited')),
    '2001-01-01T00:00:00Z',
  );
  const theirs = stamp(await build(deck('A', 'B source')), '2002-02-02T00:00:00Z');
  const merged = mergeDecks(stamp(base, '2000-01-01T00:00:00Z'), ours, theirs);
  assert.equal(merged.ok, true);
  assert.deepEqual(await texts(merged.bytes), ['A edited', 'B source']);
});

test('reports slide numbers by presentation order', async () => {
  const pres = await loadPresentation(base);
  const second = getSlides(pres)[1];
  const ours = await (async () => {
    setShapeText(getSlideShapes(second)[0], 'Second edited');
    return savePresentation(pres);
  })();
  const theirsPres = await loadPresentation(base);
  setShapeText(getSlideShapes(getSlides(theirsPres)[1])[0], 'Second source');
  const merged = mergeDecks(base, ours, await savePresentation(theirsPres));
  assert.equal(merged.ok, false);
  assert.equal(merged.conflicts[0].slide, 2);
  assert.equal(merged.conflicts[0].shape.name, getShapeName(getSlideShapes(second)[0]));
});
