import assert from 'node:assert/strict';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import {
  getSlides,
  getSlideShapes,
  getSlideText,
  loadPresentation,
  savePresentation,
  setShapeText,
} from '@office-kit/pptx';
import { buildDeck } from '../dist/index.mjs';
import { startPreview, waitForState } from './helpers/server.mjs';

test(
  'editor saves survive restart/export and source conflicts require an explicit choice',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-editor-'));
    const file = join(dir, 'deck.tsx');
    const source = (text) =>
      `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={8} height={1}>${text}</Text></Slide></Presentation>`;
    await writeFile(file, source('Source'));
    let preview;
    try {
      preview = await startPreview(file);
      let state = await waitForState(preview.url, (state) => state.available);
      const originalRevision = state.revision;
      const request = (path, revision, body, extra = {}) =>
        fetch(preview.url + path, {
          method: 'PUT',
          headers: {
            Origin: preview.url,
            'If-Match': revision,
            'Content-Type': 'application/octet-stream',
            ...extra,
          },
          body,
        });
      assert.equal((await fetch(preview.url + '/editor')).status, 200);
      assert.equal((await fetch(preview.url + '/editor.js')).status, 200);
      const original = new Uint8Array(
        await (await fetch(preview.url + '/editor/document')).arrayBuffer(),
      );
      const pres = await loadPresentation(original);
      setShapeText(getSlideShapes(getSlides(pres)[0])[0], '編集済み / Edited');
      const edited = await savePresentation(pres);
      assert.equal(
        (
          await request('/editor/document', state.revision, edited, {
            Origin: 'https://example.com',
          })
        ).status,
        403,
      );
      assert.equal(
        (await request('/editor/document', state.revision, new Uint8Array([1, 2]))).status,
        400,
      );
      const attempts = await Promise.all([
        request('/editor/document', state.revision, edited),
        request('/editor/document', state.revision, original),
      ]);
      assert.deepEqual(attempts.map((response) => response.status).sort(), [200, 409]);
      // Whichever concurrent request won, explicitly save the edited document next.
      state = await waitForState(preview.url, (state) => state.hasEdits);
      const saved = await request('/editor/document', state.revision, edited);
      assert.equal(saved.status, 200);
      state = await saved.json();
      const title = async (bytes) => getSlideText(getSlides(await loadPresentation(bytes))[0]);
      assert.equal(
        await title(new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer())),
        '編集済み / Edited',
      );
      assert.equal(await title((await buildDeck(file)).bytes), '編集済み / Edited');
      assert.equal(await readFile(file, 'utf8'), source('Source'));
      await preview.close();
      preview = await startPreview(file);
      state = await waitForState(preview.url, (state) => state.hasEdits);
      assert.equal(state.conflict, false);
      assert.equal((await request('/editor/document', originalRevision, original)).status, 409);
      assert.equal(
        await title(new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer())),
        '編集済み / Edited',
      );

      await writeFile(file, source('External change'));
      state = await waitForState(preview.url, (state) => state.conflict);
      assert.equal(
        await title(
          new Uint8Array(await (await fetch(preview.url + '/editor/source')).arrayBuffer()),
        ),
        'External change',
      );
      assert.equal((await request('/editor/document', state.revision, edited)).status, 409);
      await assert.rejects(buildDeck(file), /Source changed/);
      const keep = await request('/editor/document', state.revision, edited, {
        'X-Editor-Resolve': 'edits',
      });
      assert.equal(keep.status, 200);
      state = await keep.json();
      assert.equal(state.conflict, false);
      assert.equal(await title((await buildDeck(file)).bytes), '編集済み / Edited');

      const reset = await fetch(preview.url + '/editor/resolve', {
        method: 'POST',
        headers: { Origin: preview.url, 'If-Match': state.revision },
      });
      assert.equal(reset.status, 200);
      assert.equal((await reset.json()).hasEdits, false);
      assert.equal(await title((await buildDeck(file)).bytes), 'External change');
    } finally {
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
