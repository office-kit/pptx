import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import test from 'node:test';
import { build, transform } from 'esbuild';
import { compileModule } from 'svelte/compiler';

// Compile the actual rune-backed model, using the same client runtime as the UI.
const result = await build({
  stdin: {
    contents: `export { EditorController } from './src/lib/editor/core/controller.svelte.ts';
      export { EditorDocument } from './src/lib/editor/core/document.svelte.ts';
      export { addTitleSlide, createPresentation, getSlides, getSlideText, savePresentation, getSlideShapes, getShapeId, getShapeText, setShapeText }
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
  EditorController,
  getSlideShapes,
  getShapeId,
  getShapeText,
  setShapeText,
  addTitleSlide,
  createPresentation,
  getSlides,
  getSlideText,
  savePresentation,
} = await import(
  `data:text/javascript;base64,${Buffer.from(result.outputFiles[0].text).toString('base64')}`
);

function append(doc, title) {
  doc.transact(title, () => addTitleSlide(doc.pres, title));
}
function titles(doc) {
  return getSlides(doc.pres).map(getSlideText);
}

test('each rapid edit has its own undo step, including the first edit', async () => {
  const doc = new EditorDocument();
  append(doc, 'A');
  append(doc, 'B');
  append(doc, 'C');
  await doc.undo();
  assert.equal(titles(doc).at(-1), 'B');
  await doc.undo();
  assert.equal(titles(doc).at(-1), 'A');
  await doc.undo();
  assert.equal(titles(doc).length, 1);
  assert.equal(doc.canUndo, false);
  await doc.redo();
  assert.equal(titles(doc).at(-1), 'A');
});

test('rapid undo and redo requests advance through distinct snapshots', async () => {
  const doc = new EditorDocument();
  append(doc, 'A');
  append(doc, 'B');
  append(doc, 'C');
  await Promise.all([doc.undo(), doc.undo()]);
  assert.equal(titles(doc).at(-1), 'A');
  await Promise.all([doc.redo(), doc.redo()]);
  assert.equal(titles(doc).at(-1), 'C');
});

test('new and open invalidate pending restores and older history', async () => {
  const doc = new EditorDocument();
  append(doc, 'A');
  const undo = doc.undo();
  doc.resetBlank();
  await undo;
  assert.equal(titles(doc).length, 1);
  assert.equal(doc.canRedo, false);
  append(doc, 'B');
  const replacement = createPresentation();
  addTitleSlide(replacement, '日本語の文書');
  const opening = doc.loadBytes(await savePresentation(replacement), '日本語.pptx');
  await opening;
  await doc.undo();
  assert.deepEqual(titles(doc), ['日本語の文書']);
  assert.equal(doc.dirty, false);
});

test('a transaction during undo keeps the newer edit and a usable history', async () => {
  const doc = new EditorDocument();
  append(doc, 'A');
  append(doc, 'B');
  const undo = doc.undo();
  append(doc, 'C');
  await undo;
  assert.equal(titles(doc).at(-1), 'C');
  await doc.undo();
  assert.equal(titles(doc).at(-1), 'B');
  append(doc, 'D');
  assert.equal(doc.canRedo, false);
});

test('a whole live gesture commits one undo step and preserves its selection', async () => {
  const doc = new EditorDocument();
  doc.applyLive(() => addTitleSlide(doc.pres, 'A'));
  doc.applyLive(() => addTitleSlide(doc.pres, 'B'));
  doc.selectSlide(2);
  doc.commit('Gesture');
  doc.selectSlide(0);
  await doc.undo();
  assert.equal(titles(doc).length, 1);
  await doc.redo();
  assert.equal(titles(doc).at(-1), 'B');
  assert.equal(doc.selection.slideIndex, 2);
});

test('a completed save only marks the captured document version clean', async () => {
  const doc = new EditorDocument();
  append(doc, 'A');
  const savedVersion = doc.version;
  const saving = doc.toBytes();
  append(doc, 'B');
  await saving;
  doc.markSaved(savedVersion);
  assert.equal(doc.dirty, true);
  doc.markSaved(doc.version);
  assert.equal(doc.dirty, false);
});

test('a failed open cancels pending restores without disabling history', async () => {
  const doc = new EditorDocument();
  append(doc, 'A');
  append(doc, 'B');
  const undo = doc.undo();
  await assert.rejects(doc.loadBytes(new Uint8Array([1, 2]), 'bad.pptx'));
  await undo;
  assert.equal(titles(doc).at(-1), 'B');
  await doc.undo();
  assert.equal(titles(doc).at(-1), 'A');
});

test('new and open end an unfinished live gesture', async () => {
  const doc = new EditorDocument();
  doc.applyLive(() => addTitleSlide(doc.pres, 'A'));
  doc.resetBlank();
  assert.equal(doc.liveEditing, false);
  const bytes = await doc.toBytes();
  doc.applyLive(() => addTitleSlide(doc.pres, 'B'));
  await doc.loadBytes(bytes, 'replacement.pptx');
  assert.equal(doc.liveEditing, false);
  assert.equal(doc.canUndo, false);
});

function selectedTitle(editor) {
  const shape = getSlideShapes(editor.doc.slideAt(0))[0];
  editor.doc.selectShape(0, getShapeId(shape));
  return shape;
}

test('cut retains a fixed copy and paste restores it with redo selection', async () => {
  const editor = new EditorController();
  const source = selectedTitle(editor);
  const text = getShapeText(source);
  const count = getSlideShapes(editor.doc.slideAt(0)).length;
  editor.cutSelection();
  assert.equal(getSlideShapes(editor.doc.slideAt(0)).length, count - 1);
  await editor.paste();
  assert.equal(getShapeText(editor.selectedShapes()[0]), text);
  const selection = editor.doc.selection;
  await editor.doc.undo();
  assert.equal(getSlideShapes(editor.doc.slideAt(0)).length, count - 1);
  await editor.doc.redo();
  assert.deepEqual(editor.doc.selection, selection);
  assert.equal(getShapeText(editor.selectedShapes()[0]), text);
});

test('copy is unaffected by source edits, undo, slide changes or a new document', async () => {
  const editor = new EditorController();
  const source = selectedTitle(editor);
  const text = getShapeText(source);
  editor.copySelection();
  editor.doc.transact('Edit original', () => setShapeText(source, 'Changed'));
  await editor.paste();
  assert.equal(getShapeText(editor.selectedShapes()[0]), text);
  await editor.doc.undo();
  editor.doc.resetBlank();
  const count = getSlideShapes(editor.doc.slideAt(0)).length;
  await editor.paste();
  assert.equal(getShapeText(editor.selectedShapes()[0]), text);
  assert.equal(getSlideShapes(editor.doc.slideAt(0)).length, count + 1);
});
