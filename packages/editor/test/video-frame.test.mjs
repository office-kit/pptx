import assert from 'node:assert/strict';
import { fileURLToPath } from 'node:url';
import test from 'node:test';
import { build } from 'esbuild';

const result = await build({
  stdin: {
    contents: `export { decodeFirstVideoFrame } from './src/core/video-frame.ts';`,
    resolveDir: fileURLToPath(new URL('..', import.meta.url)),
    loader: 'ts',
  },
  bundle: true,
  write: false,
  platform: 'node',
  format: 'esm',
});

const { decodeFirstVideoFrame } = await import(
  `data:text/javascript;base64,${Buffer.from(result.outputFiles[0].text).toString('base64')}`
);

/**
 * A video whose first frame reaches the paint path only after `paintAfter`
 * draws, as Chromium on Linux x86-64 can do right after `loadeddata`.
 */
class FakeVideo extends EventTarget {
  readyState = 0;
  videoWidth = 2;
  videoHeight = 1;
  draws = 0;
  constructor(paintAfter) {
    super();
    this.paintAfter = paintAfter;
  }
  set src(_value) {}
  removeAttribute() {}
  load() {
    if (this.readyState) return;
    setTimeout(() => {
      this.readyState = 4;
      this.dispatchEvent(new Event('loadeddata'));
    });
  }
}

class FakeCanvas {
  width = 0;
  height = 0;
  alpha = 0;
  getContext() {
    return {
      drawImage: (video) => {
        if (video.draws++ >= video.paintAfter) this.alpha = 255;
      },
      getImageData: (_x, _y, width, height) => ({
        data: new Uint8ClampedArray(width * height * 4).map((_, index) =>
          index % 4 === 3 ? this.alpha : 0,
        ),
      }),
    };
  }
  toBlob(callback) {
    callback(new Blob([Uint8Array.of(this.alpha)]));
  }
}

function withFakeDom(video, run) {
  const previous = globalThis.document;
  globalThis.document = {
    createElement: (tag) => (tag === 'video' ? video : new FakeCanvas()),
  };
  return run().finally(() => {
    globalThis.document = previous;
  });
}

test('Reset poster waits for the first frame to paint instead of keeping a blank draw', async () => {
  const video = new FakeVideo(3);
  const bytes = await withFakeDom(video, () =>
    decodeFirstVideoFrame(new Uint8Array([1]), 'video/webm'),
  );
  assert.deepEqual([...bytes], [255]);
  assert.equal(video.draws, 4);
});

test('Reset poster keeps a first frame that is transparent once the wait runs out', async () => {
  const video = new FakeVideo(Infinity);
  const bytes = await withFakeDom(video, () =>
    decodeFirstVideoFrame(new Uint8Array([1]), 'video/webm'),
  );
  assert.deepEqual([...bytes], [0]);
  assert.ok(video.draws > 1);
});
