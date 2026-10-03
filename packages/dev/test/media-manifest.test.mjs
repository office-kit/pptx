import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { buildDeck } from '../dist/index.mjs';

test('build manifest exposes embedded media with geometry and playback settings', async (t) => {
  const directory = await mkdtemp(join(tmpdir(), 'pptx-media-manifest-'));
  t.after(() => rm(directory, { recursive: true, force: true }));
  const deck = join(directory, 'deck.tsx');
  await writeFile(
    deck,
    `import { Presentation, Slide, Media } from '@office-kit/pptx-dsl';
const mp3 = Uint8Array.from([0x49, 0x44, 0x33, 3, 0, 0, 0, 0, 0, 0]);
export default <Presentation><Slide><Media kind="audio" data={mp3} x={1} y={2} width={3} height={1} /></Slide></Presentation>;`,
  );
  const result = await buildDeck(deck);
  assert.equal(result.media.length, 1);
  const [media] = result.media;
  assert.equal(media.slideIndex, 0);
  assert.equal(media.kind, 'audio');
  assert.match(media.src, /^data:audio\/mpeg;base64,/);
  assert.deepEqual(media.bounds, { x: 914400, y: 1828800, w: 2743200, h: 914400 });
  assert.deepEqual(media.playback, {
    autoplay: false,
    loop: false,
    volume: 0.8,
    muted: false,
    fullScreen: false,
    hideWhenStopped: false,
  });
});
