import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import {
  getSlides,
  getSlideShapes,
  getShapeBounds,
  inches,
  loadPresentation,
  savePresentation,
  setShapeBounds,
  setShapeFill,
  setShapeRotation,
} from '@office-kit/pptx';
import { buildDeck } from '../dist/index.mjs';
import { createSourceSync } from '../src/source-sync.ts';

const DECK = `import { Presentation, Slide, Shape, Text } from '@office-kit/pptx-dsl';
const col = 2;
const rows = ['a', 'b'];
export default (
  <Presentation>
    <Slide>
      <Shape preset="rect" x={1} y={1} width={2} height={1} />
      <Text x={col * 1.5} y={2} width={3} height={1}>Hello</Text>
      {rows.map((row, i) => <Shape preset="ellipse" x={5} y={1 + i} width={1} height={1} />)}
    </Slide>
  </Presentation>
);
`;

async function withProject(run, { failRebuild = false } = {}) {
  const dir = await mkdtemp(join(tmpdir(), 'office-source-sync-'));
  const entry = join(dir, 'deck.tsx');
  await writeFile(entry, DECK);
  let latest = await buildDeck(entry);
  let failNext = failRebuild;
  const sync = createSourceSync(
    entry,
    () => latest,
    async () => {
      if (failNext) {
        failNext = false;
        return 'simulated build failure';
      }
      try {
        latest = await buildDeck(entry);
        return null;
      } catch (error) {
        return String(error);
      }
    },
  );
  try {
    await run({ entry, sync, latest: () => latest });
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
}

async function edit(bytes, change) {
  const deck = await loadPresentation(bytes);
  change(getSlideShapes(getSlides(deck)[0]));
  return savePresentation(deck);
}

test('geometry edits become TSX props and a fully written deck needs no sidecar', async () => {
  await withProject(async ({ entry, sync, latest }) => {
    const edited = await edit(latest().bytes, ([rect, text]) => {
      const bounds = getShapeBounds(rect);
      setShapeBounds(rect, { ...bounds, x: inches(1.5), w: 1234567 });
      setShapeRotation(text, 15);
      const textBounds = getShapeBounds(text);
      setShapeBounds(text, { ...textBounds, x: inches(3.25) });
    });
    const result = await sync.sync(edited);
    assert.equal(result.written.length, 2);
    assert.deepEqual(result.pending, []);
    assert.equal(result.complete, true);
    const source = await readFile(entry, 'utf8');
    assert.match(
      source,
      /<Shape preset="rect" x=\{1\.5\} y=\{1\} width=\{1\.350139\} height=\{1\} \/>/,
    );
    assert.match(
      source,
      /<Text x=\{col \* 1\.5 \+ 0\.25\} y=\{2\} width=\{3\} height=\{1\} rotation=\{15\}>/,
    );
    const [rect] = getSlideShapes(getSlides(await loadPresentation(latest().bytes))[0]);
    assert.equal(getShapeBounds(rect).w, 1234567);
  });
});

test('a repeated element and non-geometry changes stay pending, and nothing else moves', async () => {
  await withProject(async ({ entry, sync, latest }) => {
    const edited = await edit(latest().bytes, (shapes) => {
      const [rect, , first] = shapes;
      setShapeBounds(rect, { ...getShapeBounds(rect), y: inches(1.5) });
      setShapeBounds(first, { ...getShapeBounds(first), x: inches(6) });
      setShapeFill(rect, '#FF0000');
    });
    const result = await sync.sync(edited);
    assert.equal(result.complete, false);
    assert.deepEqual(
      result.written.map((change) => change.kind),
      ['geometry'],
    );
    assert.deepEqual(result.pending.map((change) => change.reason ?? change.description).sort(), [
      'element makes several shapes',
      'properties other than position, size and rotation',
    ]);
    assert.match(await readFile(entry, 'utf8'), /x=\{1\} y=\{1\.5\} width=\{2\}/);
  });
});

test('a write whose rebuild fails restores the source', async () => {
  await withProject(
    async ({ entry, sync, latest }) => {
      const edited = await edit(latest().bytes, ([rect]) => {
        setShapeBounds(rect, { ...getShapeBounds(rect), x: inches(2) });
      });
      await assert.rejects(sync.sync(edited), /did not build/);
      assert.equal(await readFile(entry, 'utf8'), DECK);
    },
    { failRebuild: true },
  );
});
