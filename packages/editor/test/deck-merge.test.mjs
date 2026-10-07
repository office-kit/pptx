import assert from 'node:assert/strict';
import { fileURLToPath } from 'node:url';
import test from 'node:test';
import { build } from 'esbuild';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import {
  addBlankSlide,
  addSlideImage,
  addSlideTextBox,
  bringShapeToFront,
  createPresentation,
  getShapeHyperlink,
  getShapeImagePartName,
  getShapeName,
  getShapeText,
  getSlideShapes,
  getSlides,
  loadPresentation,
  removeShape,
  savePresentation,
  setShapeHyperlink,
  setShapeText,
} from '@office-kit/pptx';

// The merge bundles the core's package layer, whose TypeScript uses syntax
// Node's type stripping cannot run, so the tests load it the way tsdown ships it.
const { outputFiles } = await build({
  entryPoints: [fileURLToPath(new URL('../src/merge.ts', import.meta.url))],
  bundle: true,
  write: false,
  format: 'esm',
  platform: 'node',
  target: 'es2022',
});
const { mergeDecks, describeConflict } = await import(
  `data:text/javascript;base64,${Buffer.from(outputFiles[0].contents).toString('base64')}`
);

const EMU = 914400;
const PNG = Uint8Array.from(
  atob(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==',
  ),
  (c) => c.charCodeAt(0),
);

/** Builds the whole deck from scratch, as a source rebuild does. */
async function deck(a, b) {
  const pres = createPresentation();
  const first = addBlankSlide(pres);
  addSlideTextBox(first, { x: EMU, y: EMU, w: 8 * EMU, h: EMU, text: a });
  addSlideTextBox(first, { x: EMU, y: 3 * EMU, w: 8 * EMU, h: EMU, text: b });
  const second = addBlankSlide(pres);
  addSlideTextBox(second, { x: EMU, y: EMU, w: 8 * EMU, h: EMU, text: 'Second' });
  return savePresentation(pres);
}
const base = await deck('A', 'B');

/** Applies `change` to the first slide's shapes and saves the deck. */
async function edit(bytes, change) {
  const pres = await loadPresentation(bytes);
  await change(getSlideShapes(getSlides(pres)[0]), getSlides(pres)[0]);
  return savePresentation(pres);
}
const texts = async (bytes) =>
  getSlideShapes(getSlides(await loadPresentation(bytes))[0]).map(getShapeText);

test('merges different shapes changed on each side', async () => {
  const ours = await edit(base, ([a]) => setShapeText(a, 'A edited'));
  const theirs = await deck('A', 'B source');
  const merged = mergeDecks(base, ours, theirs);
  assert.equal(merged.ok, true);
  assert.deepEqual(await texts(merged.bytes), ['A edited', 'B source']);
});

test('reports the same shape changed on both sides', async () => {
  const ours = await edit(base, ([a]) => setShapeText(a, 'A edited'));
  const theirs = await deck('A source', 'B');
  const merged = mergeDecks(base, ours, theirs);
  assert.deepEqual(merged, {
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
  assert.equal(
    describeConflict(merged.conflicts[0]),
    'Slide 1: TextBox 2 was changed both in the editor and in the source.',
  );
});

test('keeps a shape added in the editor while the source changes another', async () => {
  const ours = await edit(base, (_, slide) =>
    addSlideTextBox(slide, { x: 0, y: 5 * EMU, w: 4 * EMU, h: EMU, text: 'Added' }),
  );
  const theirs = await deck('A', 'B source');
  const merged = mergeDecks(base, ours, theirs);
  assert.equal(merged.ok, true);
  assert.deepEqual(await texts(merged.bytes), ['A', 'B source', 'Added']);
});

test('keeps the editor stacking order while merging source text', async () => {
  const ours = await edit(base, ([a]) => bringShapeToFront(a));
  const theirs = await deck('A source', 'B');
  const merged = mergeDecks(base, ours, theirs);
  assert.equal(merged.ok, true);
  assert.deepEqual(await texts(merged.bytes), ['B', 'A source']);
});

test('reports a shape deleted on one side and changed on the other', async () => {
  const ours = await edit(base, ([, b]) => removeShape(b));
  const theirs = await deck('A', 'B source');
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
  assert.equal(
    describeConflict(merged.conflicts[0]),
    'Slide 1: TextBox 3 was deleted on one side and changed on the other.',
  );
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
      strFromU8(parts['docProps/core.xml'])
        .replace(/<dcterms:(created|modified)\b[^>]*>[^<]*<\/dcterms:\1>/g, '')
        .replace(
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
  const theirs = stamp(await deck('A', 'B source'), '2002-02-02T00:00:00Z');
  const merged = mergeDecks(stamp(base, '2000-01-01T00:00:00Z'), ours, theirs);
  assert.equal(merged.ok, true);
  assert.deepEqual(await texts(merged.bytes), ['A edited', 'B source']);
});

test('reports slide numbers by presentation order', async () => {
  const pres = await loadPresentation(base);
  const second = getSlides(pres)[1];
  setShapeText(getSlideShapes(second)[0], 'Second edited');
  const ours = await savePresentation(pres);
  const theirsPres = await loadPresentation(base);
  setShapeText(getSlideShapes(getSlides(theirsPres)[1])[0], 'Second source');
  const merged = mergeDecks(base, ours, await savePresentation(theirsPres));
  assert.equal(merged.ok, false);
  assert.equal(merged.conflicts[0].slide, 2);
  assert.equal(merged.conflicts[0].shape.name, getShapeName(getSlideShapes(second)[0]));
});
