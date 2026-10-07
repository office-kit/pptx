import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import test from 'node:test';
import { build, transform } from 'esbuild';
import { compileModule } from 'svelte/compiler';

// Compile the actual rune-backed model, using the same client runtime as the UI.
const result = await build({
  stdin: {
    contents: `export { EditorDocument } from './src/core/document.svelte.ts';
      export { addSlideImage, addSlideShape, addTitleSlide, createPresentation, emu, getMediaParts, getShapeBoundsResolved, getShapeId, getSlides, getSlideText, loadPresentation, savePresentation, setShapeBounds }
        from '@office-kit/pptx';`,
    resolveDir: fileURLToPath(new URL('..', import.meta.url)),
  },
  bundle: true,
  write: false,
  platform: 'node',
  format: 'esm',
  plugins: [
    {
      name: 'editor-runes',
      setup(builder) {
        builder.onLoad({ filter: /\.svelte\.ts$/ }, async ({ path }) => {
          const source = await transform(await readFile(path, 'utf8'), { loader: 'ts' });
          return {
            contents: compileModule(source.code, { filename: path, generate: 'client' }).js.code,
          };
        });
      },
    },
  ],
});
const {
  EditorDocument,
  addSlideImage,
  addSlideShape,
  addTitleSlide,
  createPresentation,
  emu,
  getMediaParts,
  getShapeBoundsResolved,
  getShapeId,
  getSlides,
  getSlideText,
  loadPresentation,
  savePresentation,
  setShapeBounds,
} = await import(
  `data:text/javascript;base64,${Buffer.from(result.outputFiles[0].text).toString('base64')}`
);

// Minimal 1×1 PNG (transparent).
const PNG = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
  0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
  0x89, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x62, 0x00, 0x01, 0x00, 0x00,
  0x05, 0x00, 0x01, 0x0d, 0x0a, 0x2d, 0xb4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae,
  0x42, 0x60, 0x82,
]);

const titles = (doc) => getSlides(doc.pres).map(getSlideText);
const failure = new Error('command failed midway');
const failAfter = (mutate) => () => {
  mutate();
  throw failure;
};

test('a transaction that throws after mutating is rolled back without an undo step', async () => {
  const doc = new EditorDocument();
  const before = titles(doc);
  assert.throws(
    () =>
      doc.transact(
        'Broken',
        failAfter(() => addTitleSlide(doc.pres, 'Partial')),
      ),
    (error) => error === failure,
  );
  await doc.rollback;
  assert.deepEqual(titles(doc), before);
  assert.equal(doc.canUndo, false);
  assert.equal(doc.dirty, false);
  doc.transact('Next', () => addTitleSlide(doc.pres, 'Next'));
  assert.deepEqual(titles(doc), [...before, 'Next']);
  await doc.undo();
  assert.deepEqual(titles(doc), before);
});

test('a half-inserted media part never reaches the saved package', async () => {
  const doc = new EditorDocument();
  const media = getMediaParts(doc.pres).length;
  const slide = getSlides(doc.pres)[0];
  const box = { x: emu(0), y: emu(0), w: emu(100000), h: emu(100000) };
  assert.throws(() =>
    doc.transact(
      'Insert image',
      failAfter(() => addSlideImage(slide, PNG, box)),
    ),
  );
  await doc.rollback;
  const saved = await loadPresentation(await doc.toBytes());
  assert.equal(getMediaParts(saved).length, media);
});

test('saving right after a failed edit writes the restored state, not the partial one', async () => {
  const doc = new EditorDocument();
  const before = titles(doc);
  assert.throws(() =>
    doc.transact(
      'Broken',
      failAfter(() => addTitleSlide(doc.pres, 'Partial')),
    ),
  );
  const saved = await loadPresentation(await doc.toBytes());
  assert.deepEqual(getSlides(saved).map(getSlideText), before);
});

test('a validation error thrown before any mutation keeps the document content and selection', async () => {
  const doc = new EditorDocument();
  doc.selectSlide(0);
  const selection = doc.selection;
  const before = titles(doc);
  assert.throws(() =>
    doc.transact('Rejected', () => {
      throw new Error('Choose a slide layout.');
    }),
  );
  // Synchronous reads still see the unchanged document before the reload lands.
  assert.deepEqual(titles(doc), before);
  await doc.rollback;
  assert.deepEqual(titles(doc), before);
  assert.deepEqual(doc.selection, selection);
  assert.equal(doc.dirty, false);
  assert.equal(doc.canUndo, false);
});

test('edits are refused until the rollback has replaced the partial document', async () => {
  const doc = new EditorDocument();
  assert.throws(() =>
    doc.transact(
      'Broken',
      failAfter(() => addTitleSlide(doc.pres, 'Partial')),
    ),
  );
  assert.throws(() => doc.transact('Too early', () => {}), /being undone/);
  assert.throws(() => doc.applyLive(() => {}), /being undone/);
  assert.throws(() => doc.setDocumentSetting(() => {}), /being undone/);
  await doc.rollback;
  doc.transact('After', () => addTitleSlide(doc.pres, 'After'));
  assert.equal(titles(doc).at(-1), 'After');
});

test('a failed edit keeps redo, the dirty flag and the selection it started from', async () => {
  const doc = new EditorDocument();
  const id = doc.transact('Add shape', () =>
    getShapeId(
      addSlideShape(getSlides(doc.pres)[0], {
        preset: 'rect',
        x: emu(0),
        y: emu(0),
        w: emu(100000),
        h: emu(100000),
      }),
    ),
  );
  doc.transact('Redo target', () => addTitleSlide(doc.pres, 'Redo target'));
  await doc.undo();
  doc.selectShape(0, id);
  const selection = doc.selection;
  assert.throws(() =>
    doc.transact(
      'Broken',
      failAfter(() => {
        addTitleSlide(doc.pres, 'Partial');
        doc.selectSlide(1);
      }),
    ),
  );
  await doc.rollback;
  assert.deepEqual(doc.selection, selection);
  assert.equal(doc.dirty, true);
  assert.equal(doc.canRedo, true);
  await doc.redo();
  assert.equal(titles(doc).at(-1), 'Redo target');
});

test('a live frame that throws cancels the whole gesture', async () => {
  const doc = new EditorDocument();
  const before = titles(doc);
  doc.applyLive(() => addTitleSlide(doc.pres, 'Frame 1'));
  assert.throws(() => doc.applyLive(failAfter(() => addTitleSlide(doc.pres, 'Frame 2'))));
  await doc.rollback;
  assert.equal(doc.liveEditing, false);
  assert.deepEqual(titles(doc), before);
  assert.equal(doc.canUndo, false);
});

test('New during a rollback wins and leaves the editor usable', async () => {
  const doc = new EditorDocument();
  assert.throws(() =>
    doc.transact(
      'Broken',
      failAfter(() => addTitleSlide(doc.pres, 'Partial')),
    ),
  );
  doc.resetBlank();
  await doc.rollback;
  assert.deepEqual(titles(doc), ['Untitled presentation']);
  doc.transact('After', () => addTitleSlide(doc.pres, 'After'));
  assert.deepEqual(titles(doc), ['Untitled presentation', 'After']);
});

test('a drag that resolves shapes by id each frame continues on the restored document', async () => {
  const doc = new EditorDocument();
  const size = { w: emu(100000), h: emu(100000) };
  const id = doc.transact('Add shape', () =>
    getShapeId(
      addSlideShape(getSlides(doc.pres)[0], { preset: 'rect', x: emu(0), y: emu(0), ...size }),
    ),
  );
  // SlideCanvas and VideoSection re-resolve their targets every frame; a
  // reference held across a rollback belongs to the discarded document.
  const moveTo = (x) => setShapeBounds(doc.shapeById(0, id), { x: emu(x), y: emu(0), ...size });
  const x = () => getShapeBoundsResolved(doc.pres, doc.shapeById(0, id)).x;
  doc.applyLive(() => moveTo(1000));
  const stale = doc.shapeById(0, id);
  assert.throws(() => doc.applyLive(failAfter(() => moveTo(2000))));
  await doc.rollback;
  assert.notEqual(doc.shapeById(0, id), stale);
  assert.equal(x(), 0);
  doc.applyLive(() => moveTo(3000));
  doc.commit('Move');
  assert.equal(x(), 3000);
  await doc.undo();
  assert.equal(x(), 0);
});

// Script callers can catch a failed edit and immediately try another operation
// in the same turn, before the rollback has landed.
test('a failed Open during a rollback still restores the document and keeps it editable', async () => {
  const doc = new EditorDocument();
  doc.transact('Committed', () => addTitleSlide(doc.pres, 'Committed'));
  doc.selectSlide(1);
  const selection = doc.selection;
  const before = titles(doc);
  assert.throws(() =>
    doc.transact(
      'Broken',
      failAfter(() => addTitleSlide(doc.pres, 'Partial')),
    ),
  );
  await assert.rejects(doc.loadBytes(new Uint8Array([1, 2]), 'bad.pptx'));
  await doc.rollback;
  assert.deepEqual(titles(doc), before);
  assert.deepEqual(doc.selection, selection);
  assert.equal(doc.fileName, 'Untitled.pptx');
  const saved = await loadPresentation(await doc.toBytes());
  assert.deepEqual(getSlides(saved).map(getSlideText), before);
  doc.transact('After', () => addTitleSlide(doc.pres, 'After'));
  await doc.undo();
  assert.deepEqual(titles(doc), before);
  await doc.undo();
  assert.deepEqual(titles(doc), before.slice(0, -1));
});

test('a successful Open during a rollback wins', async () => {
  const opened = createPresentation();
  addTitleSlide(opened, 'Opened');
  const bytes = await savePresentation(opened);
  const doc = new EditorDocument();
  assert.throws(() =>
    doc.transact(
      'Broken',
      failAfter(() => addTitleSlide(doc.pres, 'Partial')),
    ),
  );
  await doc.loadBytes(bytes, 'opened.pptx');
  await doc.rollback;
  assert.deepEqual(titles(doc), ['Opened']);
  assert.equal(doc.fileName, 'opened.pptx');
  doc.transact('After', () => addTitleSlide(doc.pres, 'After'));
  assert.deepEqual(titles(doc), ['Opened', 'After']);
});

test('an Undo requested during a rollback wins', async () => {
  const doc = new EditorDocument();
  const initial = titles(doc);
  doc.transact('First', () => addTitleSlide(doc.pres, 'First'));
  assert.throws(() =>
    doc.transact(
      'Broken',
      failAfter(() => addTitleSlide(doc.pres, 'Partial')),
    ),
  );
  await doc.undo();
  await doc.rollback;
  assert.deepEqual(titles(doc), initial);
  assert.equal(doc.canRedo, true);
  doc.transact('After', () => addTitleSlide(doc.pres, 'After'));
  assert.deepEqual(titles(doc), [...initial, 'After']);
});
