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
      export { reconcileSelection, resolveShape, selectionShapeRefs } from './src/core/shape-ref.ts';
      export { addTitleSlide, getShapeId, getShapeName, getSlideShapes, getSlides, removeShape, removeSlide, setShapeText }
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
  reconcileSelection,
  resolveShape,
  selectionShapeRefs,
  addTitleSlide,
  getShapeId,
  getShapeName,
  getSlideShapes,
  getSlides,
  removeShape,
  removeSlide,
  setShapeText,
} = await import(
  `data:text/javascript;base64,${Buffer.from(result.outputFiles[0].text).toString('base64')}`
);

const title = (doc) => getSlideShapes(getSlides(doc.pres)[0])[0];

test('reports committed changes with their source, but not rolled-back ones', async () => {
  const doc = new EditorDocument();
  const changes = [];
  doc.onChange = (source) => changes.push(source);

  doc.transact('Retitle', () => setShapeText(title(doc), 'By the user'));
  doc.transact('Agent: Retitle', () => setShapeText(title(doc), 'By an agent'), 'agent');
  assert.throws(
    () =>
      doc.transact(
        'Agent: Broken',
        () => {
          setShapeText(title(doc), 'partial');
          throw new Error('broken');
        },
        'agent',
      ),
    /broken/,
  );
  await doc.rollback;
  assert.equal(doc.undoLabel, 'Agent: Retitle');
  await doc.undo();
  await doc.redo();
  doc.applyLive(() => setShapeText(title(doc), 'Dragging'));
  assert.deepEqual(changes, ['user', 'agent', 'user', 'user']);
  doc.commit('Drag');
  assert.deepEqual(changes, ['user', 'agent', 'user', 'user', 'user']);
});

test('a listener that throws does not undo the edit it hears about', () => {
  const doc = new EditorDocument();
  doc.onChange = () => {
    throw new Error('listener failed');
  };
  assert.throws(() => doc.transact('Retitle', () => setShapeText(title(doc), 'Kept')), /listener/);
  assert.equal(doc.undoLabel, 'Retitle');
});

test('shape refs resolve by slide part and shape id, and fail once the shape is gone', () => {
  const doc = new EditorDocument();
  addTitleSlide(doc.pres, 'Second');
  const second = getSlides(doc.pres)[1];
  const shape = getSlideShapes(second)[0];
  const [ref] = selectionShapeRefs(doc.pres, {
    kind: 'shape',
    slideIndex: 1,
    shapeIds: [getShapeId(shape)],
  });
  assert.deepEqual(ref, {
    slideIndex: 1,
    slide: '/ppt/slides/slide2.xml',
    shapeId: getShapeId(shape),
    name: getShapeName(shape),
  });
  // Moving slides does not break a ref.
  removeSlide(doc.pres, getSlides(doc.pres)[0]);
  assert.equal(resolveShape(doc.pres, ref), getSlideShapes(getSlides(doc.pres)[0])[0]);

  removeShape(resolveShape(doc.pres, ref));
  assert.throws(() => resolveShape(doc.pres, ref), /deleted from slide \/ppt\/slides\/slide2\.xml/);
  removeSlide(doc.pres, getSlides(doc.pres)[0]);
  assert.throws(() => resolveShape(doc.pres, ref), /its slide was deleted/);
});

test('reconciles a selection with the shapes and slides that are left', () => {
  const doc = new EditorDocument();
  addTitleSlide(doc.pres, 'Second');
  const [first, second] = getSlideShapes(getSlides(doc.pres)[1]).map(getShapeId);
  const selection = { kind: 'shape', slideIndex: 1, shapeIds: [first, second] };
  assert.equal(reconcileSelection(doc.pres, selection), selection);

  removeShape(getSlideShapes(getSlides(doc.pres)[1])[0]);
  assert.deepEqual(reconcileSelection(doc.pres, selection), { ...selection, shapeIds: [second] });
  const table = { kind: 'cell', slideIndex: 1, shapeId: first, row: 0, col: 0 };
  assert.deepEqual(reconcileSelection(doc.pres, table), { kind: 'none', slideIndex: 1 });

  removeSlide(doc.pres, getSlides(doc.pres)[1]);
  assert.deepEqual(reconcileSelection(doc.pres, selection), { kind: 'none', slideIndex: 0 });
  assert.deepEqual(
    reconcileSelection(doc.pres, { kind: 'slide', slideIndex: 0, slideIndices: [0, 1] }),
    { kind: 'slide', slideIndex: 0 },
  );
});
