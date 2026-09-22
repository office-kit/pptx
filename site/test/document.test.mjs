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
      export { setShapeBounds, getShapeTextAnchor, getParagraphPropertiesEffective, addSlideLine, getShapeRotation, setShapeRotation, getShapeFlip, setShapeFlip, getShapeParagraphElements, getShapeFillColor, getShapeStrokeColor, getShapeImageBytes, getShapeImageCrop, getShapeDescription, getSlideSize, addSlideShape, getShapeKind, getGroupChildren, getShapeBoundsResolved, emu, addTitleSlide, createPresentation, getSlides, getSlideText, savePresentation, getSlideShapes, getShapeId, getShapeText, setShapeText }
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
  setShapeBounds,
  getShapeTextAnchor,
  getParagraphPropertiesEffective,
  addSlideLine,
  getShapeRotation,
  setShapeRotation,
  getShapeFlip,
  setShapeFlip,
  getShapeParagraphElements,
  getShapeFillColor,
  getShapeStrokeColor,
  getShapeImageBytes,
  getShapeImageCrop,
  getShapeDescription,
  getSlideSize,
  addSlideShape,
  getShapeKind,
  getGroupChildren,
  getShapeBoundsResolved,
  emu,
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

test('slide commands use the selected slide and keep insertion order through undo and redo', async () => {
  const editor = new EditorController();
  append(editor.doc, 'A');
  append(editor.doc, 'B');
  editor.doc.selectSlide(1);
  assert.equal(editor.command('duplicateSlide').params.length, 0);
  editor.invoke('duplicateSlide');
  assert.deepEqual(titles(editor.doc).slice(1), ['A', 'A', 'B']);
  assert.equal(editor.doc.selection.slideIndex, 2);
  await editor.doc.undo();
  assert.deepEqual(titles(editor.doc).slice(1), ['A', 'B']);
  await editor.doc.redo();
  assert.equal(editor.doc.selection.slideIndex, 2);
  editor.invoke('removeSlide');
  assert.deepEqual(titles(editor.doc).slice(1), ['A', 'B']);
  assert.equal(editor.doc.selection.slideIndex, 2);
});

test('new slides insert after the active slide and moving selects the resulting position', async () => {
  const editor = new EditorController();
  append(editor.doc, 'A');
  append(editor.doc, 'B');
  editor.doc.selectSlide(0);
  editor.invoke('addBlankSlide');
  assert.deepEqual(titles(editor.doc).slice(1), ['', 'A', 'B']);
  assert.equal(editor.doc.selection.slideIndex, 1);
  editor.doc.selectSlide(3);
  editor.invoke('moveSlide', { toIndex: 0 });
  assert.equal(titles(editor.doc)[0], 'B');
  assert.equal(editor.doc.selection.slideIndex, 0);
  await editor.doc.undo();
  await editor.doc.redo();
  assert.equal(editor.doc.selection.slideIndex, 0);
});

test('deleting the final slide leaves an empty deck that can accept a new slide', async () => {
  const editor = new EditorController();
  editor.invoke('removeSlide');
  assert.equal(editor.doc.slides.length, 0);
  assert.equal(editor.canRun('removeSlide'), false);
  assert.equal(editor.canRun('duplicateSlide'), false);
  editor.invoke('addBlankSlide');
  assert.equal(editor.doc.slides.length, 1);
  assert.equal(editor.doc.selection.slideIndex, 0);
  await editor.doc.undo();
  assert.equal(editor.doc.slides.length, 0);
});

function arrangedShapes(editor) {
  editor.invoke('addBlankSlide');
  const slide = editor.doc.slideAt(editor.doc.selection.slideIndex);
  const shapes = editor.doc.transact('Shapes', () => [
    addSlideShape(slide, { preset: 'rect', x: emu(100), y: emu(100), w: emu(100), h: emu(100) }),
    addSlideShape(slide, { preset: 'rect', x: emu(400), y: emu(300), w: emu(200), h: emu(200) }),
    addSlideShape(slide, { preset: 'rect', x: emu(900), y: emu(800), w: emu(300), h: emu(300) }),
  ]);
  editor.doc.select({
    kind: 'shape',
    slideIndex: editor.doc.selection.slideIndex,
    shapeIds: shapes.map(getShapeId).reverse(),
  });
  return shapes;
}

test('group and ungroup bind the live selection in stacking order and restore selection on undo', async () => {
  const editor = new EditorController();
  assert.equal(editor.canRun('groupShapes'), false);
  assert.equal(editor.canRun('ungroupShapes'), false);
  const shapes = arrangedShapes(editor);
  const ids = shapes.map(getShapeId);
  const selected = editor.doc.selection;
  assert.equal(
    editor.command('groupShapes').params.some((p) => p.name === 'shapes'),
    false,
  );
  editor.runOrPrompt('groupShapes');
  assert.equal(editor.activeDialog, null);
  const group = editor.selectedShapes()[0];
  assert.equal(getShapeKind(group), 'group');
  editor.selectAllShapes();
  assert.deepEqual(editor.doc.selection.shapeIds, [getShapeId(group)]);
  assert.deepEqual(getGroupChildren(group).map(getShapeId), ids);
  assert.equal(editor.canRun('groupShapes'), false);
  assert.equal(editor.canRun('ungroupShapes'), true);
  await editor.doc.undo();
  assert.deepEqual(editor.doc.selection, selected);
  await editor.doc.redo();
  editor.invoke('ungroupShapes');
  assert.deepEqual(editor.doc.selection.shapeIds, ids);
  assert.equal(editor.canRun('ungroupShapes'), false);
  await editor.doc.undo();
  assert.equal(getShapeKind(editor.selectedShapes()[0]), 'group');
  await editor.doc.redo();
  assert.deepEqual(editor.doc.selection.shapeIds, ids);
});

test('alignment and equal gaps preserve dimensions and undo as one operation', async () => {
  const editor = new EditorController();
  const shapes = arrangedShapes(editor);
  const before = shapes.map((s) => getShapeBoundsResolved(editor.doc.pres, s));
  editor.alignSelection('left');
  assert.deepEqual(
    shapes.map((s) => getShapeBoundsResolved(editor.doc.pres, s).x),
    [100, 100, 100],
  );
  await editor.doc.undo();
  const live = editor.selectedShapes().reverse();
  assert.deepEqual(
    live.map((s) => getShapeBoundsResolved(editor.doc.pres, s)),
    before,
  );
  editor.distributeSelection('horizontal');
  assert.deepEqual(
    live.map((s) => getShapeBoundsResolved(editor.doc.pres, s).x),
    [100, 450, 900],
  );
  assert.deepEqual(
    live.map((s) => getShapeBoundsResolved(editor.doc.pres, s).w),
    [100, 200, 300],
  );
  await editor.doc.undo();
  assert.deepEqual(
    editor
      .selectedShapes()
      .reverse()
      .map((s) => getShapeBoundsResolved(editor.doc.pres, s)),
    before,
  );
  editor.alignSelection('bottom');
  assert.deepEqual(
    editor.selectedShapes().map((s) => {
      const b = getShapeBoundsResolved(editor.doc.pres, s);
      return b.y + b.h;
    }),
    [1100, 1100, 1100],
  );
});

test('single-object alignment uses slide bounds and vertical distribution keeps the endpoints', async () => {
  const editor = new EditorController();
  const shapes = arrangedShapes(editor);
  editor.distributeSelection('vertical');
  assert.deepEqual(
    shapes.map((s) => getShapeBoundsResolved(editor.doc.pres, s).y),
    [100, 400, 800],
  );
  editor.doc.selectShape(editor.doc.selection.slideIndex, getShapeId(shapes[1]));
  editor.alignSelection('center');
  editor.alignSelection('middle');
  const b = getShapeBoundsResolved(editor.doc.pres, shapes[1]);
  const size = getSlideSize(editor.doc.pres);
  assert.equal(b.x + b.w / 2, size.width / 2);
  assert.equal(b.y + b.h / 2, size.height / 2);
  assert.equal(editor.canRun('groupShapes'), false);
});

test('nested groups remain a single selectable object and ungroup one level at a time', () => {
  const editor = new EditorController();
  const shapes = arrangedShapes(editor);
  const slideIndex = editor.doc.selection.slideIndex;
  editor.doc.select({ kind: 'shape', slideIndex, shapeIds: shapes.slice(0, 2).map(getShapeId) });
  editor.invoke('groupShapes');
  const inner = editor.selectedShapes()[0];
  editor.selectAllShapes();
  assert.deepEqual(editor.doc.selection.shapeIds, [getShapeId(inner), getShapeId(shapes[2])]);
  editor.invoke('groupShapes');
  const outer = editor.selectedShapes()[0];
  editor.selectAllShapes();
  assert.deepEqual(editor.doc.selection.shapeIds, [getShapeId(outer)]);
  editor.invoke('ungroupShapes');
  assert.deepEqual(editor.selectedShapes().map(getShapeKind), ['group', 'shape']);
  editor.invoke('ungroupShapes');
  assert.deepEqual(editor.doc.selection.shapeIds, shapes.map(getShapeId));
});

test('grouped siblings can regroup and ungroup in place with undo and mixed-scope rejection', async () => {
  const editor = new EditorController();
  const shapes = arrangedShapes(editor);
  const ids = shapes.map(getShapeId);
  const slideIndex = editor.doc.selection.slideIndex;
  editor.invoke('groupShapes');
  const outerId = getShapeId(editor.selectedShapes()[0]);
  const outer = () => editor.doc.shapeById(slideIndex, outerId);
  editor.doc.transact('Transform outer', () => {
    setShapeRotation(outer(), 35);
    setShapeFlip(outer(), { horizontal: true, vertical: false });
    setShapeBounds(outer(), { x: emu(500), y: emu(700), w: emu(2200), h: emu(3000) });
  });
  const outerBounds = getShapeBoundsResolved(editor.doc.pres, outer());
  const childBounds = getGroupChildren(outer()).map((s) =>
    getShapeBoundsResolved(editor.doc.pres, s),
  );
  editor.doc.select({ kind: 'shape', slideIndex, shapeIds: [ids[1], ids[0]] });
  assert.equal(editor.canRun('groupShapes'), true);
  editor.invoke('groupShapes');
  const innerId = getShapeId(editor.selectedShapes()[0]);
  assert.deepEqual(getGroupChildren(outer()).map(getShapeId), [innerId, ids[2]]);
  assert.deepEqual(getGroupChildren(editor.selectedShapes()[0]).map(getShapeId), ids.slice(0, 2));
  await editor.doc.undo();
  assert.deepEqual(editor.doc.selection.shapeIds, [ids[1], ids[0]]);
  assert.deepEqual(getGroupChildren(outer()).map(getShapeId), ids);
  await editor.doc.redo();
  assert.equal(editor.canRun('ungroupShapes'), true);
  editor.invoke('ungroupShapes');
  assert.deepEqual(editor.doc.selection.shapeIds, ids.slice(0, 2));
  assert.deepEqual(getGroupChildren(outer()).map(getShapeId), ids);
  assert.deepEqual(
    getGroupChildren(outer()).map((s) => getShapeBoundsResolved(editor.doc.pres, s)),
    childBounds,
  );
  assert.deepEqual(getShapeBoundsResolved(editor.doc.pres, outer()), outerBounds);
  assert.equal(getShapeRotation(outer()), 35);
  assert.deepEqual(getShapeFlip(outer()), { horizontal: true, vertical: false });
  await editor.doc.undo();
  assert.deepEqual(editor.doc.selection.shapeIds, [innerId]);
  await editor.doc.redo();
  editor.doc.select({ kind: 'shape', slideIndex, shapeIds: [outerId, ids[0], ids[1]] });
  assert.equal(editor.canRun('groupShapes'), false);
  assert.equal(editor.canRun('ungroupShapes'), false);
});

test('image insertion preserves aspect ratio, selects the image and restores it through history', async () => {
  const editor = new EditorController();
  const bytes = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]);
  const before = getSlideShapes(editor.doc.currentSlide).length;
  assert.equal(editor.applyImage(bytes, 'Photo.png', 800, 400), true);
  const selection = editor.doc.selection;
  assert.equal(selection.kind, 'shape');
  let image = editor.doc.shapeById(selection.slideIndex, selection.shapeIds[0]);
  const bounds = getShapeBoundsResolved(editor.doc.pres, image);
  const size = getSlideSize(editor.doc.pres);
  assert.equal(bounds.w, bounds.h * 2);
  assert.equal(bounds.x, Math.round((size.width - bounds.w) / 2));
  assert.equal(bounds.y, Math.round((size.height - bounds.h) / 2));
  assert.deepEqual(getShapeImageBytes(image), bytes);
  await editor.doc.undo();
  assert.equal(getSlideShapes(editor.doc.currentSlide).length, before);
  await editor.doc.redo();
  assert.deepEqual(editor.doc.selection, selection);
  editor.invoke('setShapeImageCrop', { crop: { left: 0.2 } });
  editor.invoke('setShapeDescription', { description: '写真 / Photo' });
  const replacement = new Uint8Array([255, 216, 255, 224, 0, 16]);
  assert.equal(editor.applyImage(replacement, 'New.jpg', 400, 800, true), true);
  image = editor.doc.shapeById(selection.slideIndex, selection.shapeIds[0]);
  assert.deepEqual(getShapeBoundsResolved(editor.doc.pres, image), bounds);
  assert.equal(getShapeImageCrop(image).left, 0.2);
  assert.equal(getShapeDescription(image), '写真 / Photo');
  assert.deepEqual(getShapeImageBytes(image), replacement);
  await editor.doc.undo();
  image = editor.doc.shapeById(selection.slideIndex, selection.shapeIds[0]);
  assert.deepEqual(getShapeImageBytes(image), bytes);
});

test('shape appearance commands format every selected object in one undo step', async () => {
  const editor = new EditorController();
  const doc = editor.doc;
  const slide = getSlides(doc.pres)[0];
  const shapes = doc.transact('Add test shapes', () =>
    [0, 1, 2].map((i) =>
      addSlideShape(slide, {
        preset: 'rect',
        x: emu(i * 1000000),
        y: emu(0),
        w: emu(500000),
        h: emu(500000),
      }),
    ),
  );
  const ids = shapes.map(getShapeId);
  doc.select({ kind: 'shape', slideIndex: 0, shapeIds: ids.slice(0, 2) });
  const colors = (reader) => ids.map((id) => reader(doc.shapeById(0, id)));
  const original = colors(getShapeFillColor);
  editor.invoke('setShapeFill', { color: '#123456' });
  assert.deepEqual(colors(getShapeFillColor), ['#123456', '#123456', original[2]]);
  await doc.undo();
  assert.deepEqual(colors(getShapeFillColor), original);
  await doc.redo();
  assert.deepEqual(colors(getShapeFillColor), ['#123456', '#123456', original[2]]);
  const outlines = colors(getShapeStrokeColor);
  editor.invoke('setShapeStroke', { options: { color: '#ABCDEF' } });
  assert.deepEqual(colors(getShapeStrokeColor), ['#ABCDEF', '#ABCDEF', outlines[2]]);
  await doc.undo();
  assert.deepEqual(colors(getShapeStrokeColor), outlines);
  assert.deepEqual(doc.selection.shapeIds, ids.slice(0, 2));
});

test('selected text formatting preserves content and rejects non-text targets before mutation', async () => {
  const editor = new EditorController();
  const doc = editor.doc;
  const ids = doc.transact('Add test shapes', () =>
    [0, 1, 2].map((i) => {
      const shape = addSlideShape(getSlides(doc.pres)[0], {
        preset: 'rect',
        x: emu(i * 1000000),
        y: emu(0),
        w: emu(500000),
        h: emu(500000),
      });
      if (i < 2) setShapeText(shape, ['日本語', 'English'][i]);
      return getShapeId(shape);
    }),
  );
  doc.select({ kind: 'shape', slideIndex: 0, shapeIds: ids.slice(0, 2) });
  const bold = () =>
    ids.slice(0, 2).map((id) => getShapeParagraphElements(doc.shapeById(0, id), 0)[0].format?.bold);
  editor.invoke('setShapeTextFormat', { format: { bold: true } });
  assert.deepEqual(bold(), [true, true]);
  editor.invoke('setShapeTextFormat', { format: {}, options: { reset: true } });
  assert.deepEqual(bold(), [undefined, undefined]);
  await doc.undo();
  assert.deepEqual(bold(), [true, true]);
  assert.deepEqual(
    ids.slice(0, 2).map((id) => getShapeText(doc.shapeById(0, id))),
    ['日本語', 'English'],
  );
  await doc.undo();
  assert.deepEqual(bold(), [undefined, undefined]);
  doc.select({ kind: 'shape', slideIndex: 0, shapeIds: ids });
  editor.invoke('setShapeTextFormat', { format: { bold: true } });
  assert.deepEqual(bold(), [true, true]);
  await doc.undo();
  const lineId = doc.transact('Add non-text line', () =>
    getShapeId(
      addSlideLine(getSlides(doc.pres)[0], {
        from: { x: emu(0), y: emu(0) },
        to: { x: emu(100000), y: emu(100000) },
      }),
    ),
  );
  doc.select({ kind: 'shape', slideIndex: 0, shapeIds: [...ids.slice(0, 2), lineId] });
  editor.invoke('setShapeTextFormat', { format: { bold: true } });
  assert.deepEqual(bold(), [undefined, undefined]);
});

test('multiple objects can align to slide edges and centers in one undo step', async () => {
  const editor = new EditorController();
  arrangedShapes(editor);
  const before = editor.selectedShapes().map((s) => getShapeBoundsResolved(editor.doc.pres, s));
  const size = getSlideSize(editor.doc.pres);
  for (const alignment of ['left', 'center', 'right', 'top', 'middle', 'bottom']) {
    editor.alignSelection(alignment, 'slide');
    const after = editor.selectedShapes().map((s) => getShapeBoundsResolved(editor.doc.pres, s));
    after.forEach((b, i) => {
      assert.equal(b.w, before[i].w);
      assert.equal(b.h, before[i].h);
      if (alignment === 'left') assert.equal(b.x, 0);
      if (alignment === 'right') assert.equal(b.x + b.w, size.width);
      if (alignment === 'center') assert.ok(Math.abs(b.x + b.w / 2 - size.width / 2) <= 0.5);
      if (alignment === 'top') assert.equal(b.y, 0);
      if (alignment === 'bottom') assert.equal(b.y + b.h, size.height);
      if (alignment === 'middle') assert.ok(Math.abs(b.y + b.h / 2 - size.height / 2) <= 0.5);
    });
    await editor.doc.undo();
    assert.deepEqual(
      editor.selectedShapes().map((s) => getShapeBoundsResolved(editor.doc.pres, s)),
      before,
    );
  }
});

test('flip commands preserve the other axis and update only selected objects with history', async () => {
  const editor = new EditorController();
  arrangedShapes(editor);
  const doc = editor.doc;
  const slideIndex = doc.selection.slideIndex;
  const shapes = editor.selectedShapes();
  const ids = shapes.map(getShapeId);
  doc.transact('Seed mixed flips', () => setShapeFlip(shapes[0], { vertical: true }));
  const state = () => ids.map((id) => getShapeFlip(doc.shapeById(slideIndex, id)));
  const before = state();
  doc.select({ kind: 'shape', slideIndex, shapeIds: ids.slice(0, 2) });
  editor.invoke('setShapeFlip', { options: { horizontal: true } });
  assert.deepEqual(state().slice(0, 2), [
    { horizontal: true, vertical: true },
    { horizontal: true, vertical: false },
  ]);
  assert.deepEqual(state()[2], before[2]);
  await doc.undo();
  assert.deepEqual(state(), before);
  await doc.redo();
  assert.deepEqual(
    state()
      .slice(0, 2)
      .map((value) => value.horizontal),
    [true, true],
  );
});

test('rotation commands update selected objects together and restore mixed angles', async () => {
  const editor = new EditorController();
  arrangedShapes(editor);
  const doc = editor.doc;
  const slideIndex = doc.selection.slideIndex;
  const shapes = editor.selectedShapes();
  const ids = shapes.map(getShapeId);
  doc.transact('Seed mixed angles', () => setShapeRotation(shapes[0], 20));
  const state = () =>
    ids.map((id) => ({
      rotation: getShapeRotation(doc.shapeById(slideIndex, id)),
      bounds: getShapeBoundsResolved(doc.pres, doc.shapeById(slideIndex, id)),
    }));
  const before = state();
  doc.select({ kind: 'shape', slideIndex, shapeIds: ids.slice(0, 2) });
  editor.invoke('setShapeRotation', { degrees: -30.5 });
  const expected = before.map((value, index) =>
    index < 2 ? { ...value, rotation: 329.5 } : value,
  );
  assert.deepEqual(state(), expected);
  await doc.undo();
  assert.deepEqual(state(), before);
  await doc.redo();
  assert.deepEqual(state(), expected);
});

test('text alignment updates all selected paragraphs with one undo step', async () => {
  const editor = new EditorController();
  const doc = editor.doc;
  const ids = doc.transact('Add text', () =>
    [0, 1, 2].map((i) => {
      const shape = addSlideShape(getSlides(doc.pres)[0], {
        preset: 'rect',
        x: emu(i * 1000000),
        y: emu(0),
        w: emu(500000),
        h: emu(500000),
      });
      setShapeText(shape, '日本語\nEnglish');
      return getShapeId(shape);
    }),
  );
  const state = () =>
    ids.map((id) => {
      const shape = doc.shapeById(0, id);
      return {
        align: [0, 1].map((i) => getParagraphPropertiesEffective(doc.pres, shape, i).align),
        anchor: getShapeTextAnchor(shape),
      };
    });
  const before = state();
  doc.select({ kind: 'shape', slideIndex: 0, shapeIds: ids.slice(0, 2) });
  editor.invoke('setShapeAlignment', { align: 'right' });
  assert.deepEqual(
    state()
      .slice(0, 2)
      .map((s) => s.align),
    [
      ['right', 'right'],
      ['right', 'right'],
    ],
  );
  assert.deepEqual(state()[2], before[2]);
  await doc.undo();
  assert.deepEqual(state(), before);
  await doc.redo();
  editor.invoke('setShapeTextAnchor', { anchor: 'bottom' });
  assert.deepEqual(
    state().map((s) => s.anchor),
    ['bottom', 'bottom', before[2].anchor],
  );
  await doc.undo();
  assert.deepEqual(
    state().map((s) => s.anchor),
    before.map((s) => s.anchor),
  );
  const unchanged = state();
  const lineId = doc.transact('Add line', () =>
    getShapeId(
      addSlideLine(getSlides(doc.pres)[0], {
        from: { x: emu(0), y: emu(0) },
        to: { x: emu(100000), y: emu(100000) },
      }),
    ),
  );
  doc.select({ kind: 'shape', slideIndex: 0, shapeIds: [...ids.slice(0, 2), lineId] });
  editor.invoke('setShapeAlignment', { align: 'left' });
  assert.deepEqual(state(), unchanged);
  editor.invoke('setShapeTextAnchor', { anchor: 'top' });
  assert.deepEqual(state(), unchanged);
});

test('grouped alignment and distribution use visible edges through rotation reflection and scale', async () => {
  const editor = new EditorController();
  const shapes = arrangedShapes(editor);
  const ids = shapes.map(getShapeId);
  editor.invoke('groupShapes');
  editor.doc.transact('Transform group', () => {
    const group = editor.selectedShapes()[0];
    setShapeRotation(group, 90);
    setShapeFlip(group, { horizontal: true, vertical: false });
    setShapeBounds(group, { x: emu(100), y: emu(100), w: emu(2200), h: emu(3000) });
  });
  editor.doc.select({ kind: 'shape', slideIndex: editor.doc.selection.slideIndex, shapeIds: ids });
  const bounds = () =>
    editor.selectedShapes().map((shape) => getShapeBoundsResolved(editor.doc.pres, shape));
  const visible = () =>
    bounds().map((b) => ({
      x: 3000 - 3 * (b.y + b.h),
      y: 2900 - 2 * (b.x + b.w),
      w: 3 * b.h,
      h: 2 * b.w,
    }));
  const original = bounds();
  const size = getSlideSize(editor.doc.pres);
  for (const alignment of ['left', 'center', 'right', 'top', 'middle', 'bottom']) {
    editor.alignSelection(alignment, 'slide');
    for (const b of visible()) {
      const actual = {
        left: b.x,
        center: b.x + b.w / 2,
        right: b.x + b.w,
        top: b.y,
        middle: b.y + b.h / 2,
        bottom: b.y + b.h,
      }[alignment];
      const expected = {
        left: 0,
        center: size.width / 2,
        right: size.width,
        top: 0,
        middle: size.height / 2,
        bottom: size.height,
      }[alignment];
      assert.ok(Math.abs(actual - expected) <= 2, `${alignment}: ${actual} vs ${expected}`);
    }
    await editor.doc.undo();
    assert.deepEqual(bounds(), original);
  }
  for (const direction of ['horizontal', 'vertical']) {
    const axis = direction === 'horizontal' ? 'x' : 'y';
    const extent = direction === 'horizontal' ? 'w' : 'h';
    const before = visible().sort((a, b) => a[axis] - b[axis]);
    editor.distributeSelection(direction);
    const after = visible().sort((a, b) => a[axis] - b[axis]);
    assert.deepEqual(after[0], before[0]);
    assert.deepEqual(after[2], before[2]);
    const firstGap = after[1][axis] - after[0][axis] - after[0][extent];
    const lastGap = after[2][axis] - after[1][axis] - after[1][extent];
    assert.ok(Math.abs(firstGap - lastGap) <= 3);
    await editor.doc.undo();
    assert.deepEqual(bounds(), original);
  }
});
