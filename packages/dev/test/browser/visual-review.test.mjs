import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, readFile, rm, copyFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { createVisualReviewer } from '../../src/visual-review.ts';

test('visual review captures changed offscreen slides and bounds correction passes', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'office-visual-'));
  const svg = (text) =>
    `<svg xmlns="http://www.w3.org/2000/svg" width="960" height="540" viewBox="0 0 960 540"><rect width="960" height="540" fill="#16182b"/><text x="80" y="250" font-family="Arial" font-size="60" fill="white">${text}</text></svg>`;
  let slides = [svg('Cover'), svg('Before')];
  const reviewer = createVisualReviewer(join(dir, 'deck.tsx'), () => slides);
  try {
    reviewer.begin();
    assert.equal(await reviewer.next(), undefined);
    for (let round = 1; round <= 3; round++) {
      slides = [slides[0], svg('After edit ' + round)];
      const review = await reviewer.next();
      assert.equal(review.images.length, 1);
      assert.match(review.images[0], /slide-2\.png$/);
      assert.match(review.prompt, new RegExp(`round ${round}/3`));
      const bytes = await readFile(review.images[0]);
      assert.equal(bytes.readUInt32BE(16), 1600);
      assert.equal(bytes.readUInt32BE(20), 900);
      if (round === 3) {
        assert.match(review.prompt, /do not edit further/);
        await copyFile(review.images[0], '/tmp/office-kit-slide-review.png');
      }
      assert.equal(await reviewer.next(), undefined);
    }
    slides = [slides[0], svg('Unexpected fourth correction')];
    await assert.rejects(reviewer.next(), /correction limit/);
    reviewer.begin();
    assert.equal(await reviewer.next(), undefined);
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});
