import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { startPreview, waitForState } from '../helpers/server.mjs';

test('embedded media is deduplicated and served with immutable range support', async () => {
  const directory = await mkdtemp(join(tmpdir(), 'pptx-media-server-'));
  let preview;
  try {
    await writeFile(
      join(directory, 'deck.tsx'),
      `import {Presentation, Slide, Media} from '@office-kit/pptx-dsl';
const clip = Uint8Array.from([73, 68, 51, 3, 0, 0, 0, 0, 0, 0]);
export default <Presentation><Slide><Media kind="audio" data={clip} x={1} y={1} width={2} height={1}/></Slide><Slide><Media kind="audio" data={clip} x={1} y={1} width={2} height={1}/></Slide></Presentation>;`,
    );
    preview = await startPreview(join(directory, 'deck.tsx'));
    await waitForState(preview.url, (value) => value.available);
    const state = await (await fetch(preview.url + '/state')).json();
    assert.equal(state.media.length, 2);
    assert.match(state.media[0].src, /^\/media\/[a-f0-9]{64}$/);
    assert.equal(state.media[0].src, state.media[1].src);

    const mediaUrl = preview.url + state.media[0].src;
    const full = await fetch(mediaUrl);
    assert.equal(full.status, 200);
    assert.equal(full.headers.get('content-type'), 'audio/mpeg');
    assert.equal(full.headers.get('cache-control'), 'public, max-age=31536000, immutable');
    assert.deepEqual(
      [...new Uint8Array(await full.arrayBuffer())],
      [73, 68, 51, 3, 0, 0, 0, 0, 0, 0],
    );

    const ranged = await fetch(mediaUrl, { headers: { Range: 'bytes=2-4' } });
    assert.equal(ranged.status, 206);
    assert.equal(ranged.headers.get('content-range'), 'bytes 2-4/10');
    assert.deepEqual([...new Uint8Array(await ranged.arrayBuffer())], [51, 3, 0]);
    const suffix = await fetch(mediaUrl, { headers: { Range: 'bytes=-2' } });
    assert.equal(suffix.status, 206);
    assert.equal(suffix.headers.get('content-range'), 'bytes 8-9/10');
    assert.deepEqual([...new Uint8Array(await suffix.arrayBuffer())], [0, 0]);
    const invalid = await fetch(mediaUrl, { headers: { Range: 'bytes=100-' } });
    assert.equal(invalid.status, 416);
    assert.equal(invalid.headers.get('content-range'), 'bytes */10');
  } finally {
    await preview?.close();
    await rm(directory, { recursive: true, force: true });
  }
});
