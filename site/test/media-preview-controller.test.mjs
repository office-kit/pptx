import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import test from 'node:test';
import { build, transform } from 'esbuild';
import { compileModule } from 'svelte/compiler';

const result = await build({
  stdin: {
    contents: `export { getMediaPreview } from './src/lib/editor/core/media-preview.svelte.ts';`,
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

const { getMediaPreview } = await import(
  `data:text/javascript;base64,${Buffer.from(result.outputFiles[0].text).toString('base64')}`
);

class FakeMediaElement extends EventTarget {
  constructor() {
    super();
    this.currentTime = 0;
    this.duration = 10;
    this.paused = true;
    this.volume = 1;
    this.muted = false;
    this.play = () => {
      this.paused = false;
      return Promise.resolve();
    };
    this.pause = () => {
      this.paused = true;
      this.dispatchEvent(new Event('pause'));
    };
  }
}

test('detach pauses media and removes state updates', () => {
  const preview = getMediaPreview({});
  const media = new FakeMediaElement();
  const detach = preview.attach(10, media);

  assert.equal(preview.state.shapeId, 10);
  detach();
  assert.equal(preview.state.shapeId, null);
  assert.equal(media.paused, true);
  media.currentTime = 4;
  media.dispatchEvent(new Event('timeupdate'));
  assert.equal(preview.state.currentTime, 0);
});

test('stale play rejection does not overwrite replacement state', async () => {
  const preview = getMediaPreview({});
  const first = new FakeMediaElement();
  const second = new FakeMediaElement();
  let rejectPlay;
  first.play = () =>
    new Promise((_resolve, reject) => {
      rejectPlay = reject;
    });

  preview.attach(1, first);
  preview.command('play', 1);
  preview.attach(2, second);
  rejectPlay(new Error('stale play failed'));
  await Promise.resolve();
  await Promise.resolve();

  assert.equal(preview.state.shapeId, 2);
  assert.equal(preview.state.error, '');
  assert.equal(preview.state.playing, false);
});

test('current-time bookmark lookup preserves source order', () => {
  const preview = getMediaPreview({});
  const media = new FakeMediaElement();
  preview.attach(3, media);
  media.currentTime = 2;
  media.dispatchEvent(new Event('timeupdate'));

  assert.equal(
    preview.bookmarkAtCurrent({
      autoplay: false,
      loop: false,
      volume: 1,
      muted: false,
      fullScreen: false,
      hideWhenStopped: false,
      bookmarks: [
        { name: 'Later', timeMs: 4000 },
        { name: 'Earlier', timeMs: 2000 },
      ],
    }),
    1,
  );
});
