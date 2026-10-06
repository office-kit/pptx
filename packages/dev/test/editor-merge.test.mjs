import assert from 'node:assert/strict';
import { mkdir, mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import {
  getShapeText,
  getSlideShapes,
  getSlides,
  loadPresentation,
  savePresentation,
  setShapeText,
} from '@office-kit/pptx';
import { buildDeck } from '../dist/index.mjs';
import { startPreview, waitForState } from './helpers/server.mjs';

const source = (a, b) =>
  `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={8} height={1}>${a}</Text><Text x={1} y={3} width={8} height={1}>${b}</Text></Slide></Presentation>`;
const texts = async (bytes) =>
  getSlideShapes(getSlides(await loadPresentation(bytes))[0]).map(getShapeText);

test('source changes merge with saved and unsaved editor edits', { timeout: 60000 }, async () => {
  const dir = await mkdtemp(join(tmpdir(), 'office-editor-merge-'));
  const file = join(dir, 'deck.tsx');
  const sidecar = join(dir, '.office-kit', 'deck.tsx.editor.zip');
  let preview;
  try {
    await writeFile(file, source('A', 'B'));
    preview = await startPreview(file);
    const put = (revision, body) =>
      fetch(preview.url + '/editor/document', {
        method: 'PUT',
        headers: {
          Origin: preview.url,
          'If-Match': revision,
          'Content-Type': 'application/octet-stream',
        },
        body,
      });
    const document = async () =>
      new Uint8Array(await (await fetch(preview.url + '/editor/document')).arrayBuffer());
    const exported = async () =>
      texts(new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()));
    const edit = async (bytes, index, text) => {
      const pres = await loadPresentation(bytes);
      setShapeText(getSlideShapes(getSlides(pres)[0])[index], text);
      return savePresentation(pres);
    };
    const rewrite = async (a, b) => {
      const { revision } = await waitForState(preview.url, () => true);
      await writeFile(file, source(a, b));
      return waitForState(preview.url, (state) => state.revision !== revision);
    };

    let state = await waitForState(preview.url, (state) => state.available);
    assert.equal(
      (await put(state.revision, await edit(await document(), 0, 'A edited'))).status,
      200,
    );
    const stored = unzipSync(await readFile(sidecar));
    assert.deepEqual(JSON.parse(strFromU8(stored['state.json'])).version, 2);
    assert.ok(stored['base.pptx']);

    // Saved edits: the source changes another shape.
    state = await rewrite('A', 'B source');
    assert.equal(state.conflict, false);
    assert.deepEqual(await exported(), ['A edited', 'B source']);
    assert.deepEqual(await texts((await buildDeck(file)).bytes), ['A edited', 'B source']);
    const merged = JSON.parse(strFromU8(unzipSync(await readFile(sidecar))['state.json']));
    assert.deepEqual(merged, { version: 2, sourceHash: state.sourceHash });

    // Unsaved edits made on a deck the source has since replaced.
    const basis = state.revision;
    const local = await edit(await document(), 0, 'A edited twice');
    state = await rewrite('A', 'B source 2');
    let response = await put(basis, local);
    assert.equal(response.status, 200);
    const saved = await response.json();
    assert.match(saved.uploadRevision, /:upload:/);
    assert.equal(saved.conflict, false);
    assert.deepEqual(await exported(), ['A edited twice', 'B source 2']);

    // Further edits on top of the upload still merge.
    const again = await edit(local, 0, 'A edited thrice');
    response = await put(saved.uploadRevision, again);
    assert.equal(response.status, 200);
    assert.deepEqual(await exported(), ['A edited thrice', 'B source 2']);

    // Unsaved edits of the shape the source changed too.
    state = await waitForState(preview.url, () => true);
    const clashing = await edit(await document(), 1, 'B local');
    await rewrite('A', 'B source 3');
    response = await put(state.revision, clashing);
    assert.equal(response.status, 409);
    const refused = await response.json();
    assert.equal(refused.refused, 'conflict');
    assert.deepEqual(refused.mergeConflicts, [
      {
        part: 'ppt/slides/slide1.xml',
        slide: 1,
        shape: { id: '3', name: 'TextBox 3' },
        reason: 'both-changed',
      },
    ]);
    assert.deepEqual(await exported(), ['A edited thrice', 'B source 3']);

    // Saved edits of the shape the source changes.
    state = await rewrite('A source', 'B source 3');
    assert.equal(state.conflict, true);
    assert.deepEqual(
      state.conflicts.map((conflict) => conflict.shape.name),
      ['TextBox 2'],
    );
    await assert.rejects(buildDeck(file), /Slide 1: TextBox 2 was changed both/);
    assert.deepEqual(await exported(), ['A edited thrice', 'B source 3']);
  } finally {
    await preview?.close();
    await rm(dir, { recursive: true, force: true });
  }
});

test('version 1 editor files are read but cannot merge', { timeout: 60000 }, async () => {
  const dir = await mkdtemp(join(tmpdir(), 'office-editor-v1-'));
  const file = join(dir, 'deck.tsx');
  let preview;
  try {
    await writeFile(file, source('A', 'B'));
    preview = await startPreview(file);
    const state = await waitForState(preview.url, (state) => state.available);
    const pres = await loadPresentation(
      new Uint8Array(await (await fetch(preview.url + '/editor/document')).arrayBuffer()),
    );
    setShapeText(getSlideShapes(getSlides(pres)[0])[0], 'A edited');
    await preview.close();
    preview = undefined;
    await mkdir(join(dir, '.office-kit'));
    await writeFile(
      join(dir, '.office-kit', 'deck.tsx.editor.zip'),
      zipSync({
        'state.json': strToU8(JSON.stringify({ version: 1, sourceHash: state.sourceHash })),
        'document.pptx': await savePresentation(pres),
      }),
    );
    assert.deepEqual(await texts((await buildDeck(file)).bytes), ['A edited', 'B']);
    await writeFile(file, source('A', 'B source'));
    await assert.rejects(buildDeck(file), /Source changed/);
    preview = await startPreview(file);
    const conflicted = await waitForState(preview.url, (state) => state.available);
    assert.equal(conflicted.conflict, true);
    assert.deepEqual(conflicted.conflicts, []);
  } finally {
    await preview?.close();
    await rm(dir, { recursive: true, force: true });
  }
});
