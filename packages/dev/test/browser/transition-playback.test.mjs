import { chromium } from 'playwright';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';

test('transition playback isolates SVGs and cancels stale completion callbacks', async () => {
  const source = await readFile(
    new URL('../../src/transition-playback.ts', import.meta.url),
    'utf8',
  );
  const runtime = source.slice(source.indexOf('`') + 1, source.lastIndexOf('`'));
  const browser = await chromium.launch({ headless: true });
  try {
    const page = await browser.newPage();
    await page.setContent(
      '<div id="slide" style="position:relative;width:640px;height:360px"></div>',
    );
    await page.evaluate((runtime) => {
      const svg =
        '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 640 360"><defs><linearGradient id="paint"><stop stop-color="red"/></linearGradient></defs><rect width="640" height="360" fill="url(#paint)"/></svg>';
      window.fixtureSvg = svg;
      window.completions = [];
      window.eval(
        'const canvas=document.getElementById("slide").attachShadow({mode:"open"});let displayedSvg=window.fixtureSvg;let presenting=false;const index=0;const state={slides:[displayedSvg],editor:{slides:[{transition:{effect:"fade",durationMs:1000}}]}};' +
          runtime +
          ';window.play=playTransition;window.cancel=cancelTransitionPlayback;window.preview=previewTransition;window.run=()=>transitionRun;',
      );
    }, runtime);
    const result = await page.evaluate(async () => {
      window.play(window.fixtureSvg, { effect: 'push', direction: 'l', durationMs: 1000 }, () =>
        window.completions.push('stale'),
      );
      const first = window.run();
      first.animations.forEach((a) => {
        a.pause();
        a.currentTime = 500;
      });
      const layers = [...first.overlay.children];
      const transforms = layers.map((l) => getComputedStyle(l).transform);
      const size = first.overlay.getBoundingClientRect();
      const isolated = layers.every((l) => l.shadowRoot.querySelector('#paint'));
      window.play('', { effect: 'fade', durationMs: 1000 }, () =>
        window.completions.push('current'),
      );
      const removed = !first.overlay.isConnected;
      window.run().animations.forEach((a) => a.finish());
      await new Promise((resolve) => setTimeout(resolve, 30));
      return {
        transforms,
        width: size.width,
        height: size.height,
        isolated,
        removed,
        completions: window.completions,
        cleared: window.run() === null,
      };
    });
    assert.deepEqual(result.transforms, [
      'matrix(1, 0, 0, 1, -320, 0)',
      'matrix(1, 0, 0, 1, 320, 0)',
    ]);
    assert.equal(result.width, 640);
    assert.equal(result.height, 360);
    assert.ok(result.isolated && result.removed && result.cleared);
    assert.deepEqual(result.completions, ['current']);
    assert.deepEqual(
      await page.evaluate(() => {
        window.preview();
        const effect = window.run().overlay.dataset.transitionPlayback;
        window.cancel();
        let instant = false;
        window.play('', { effect: 'unsupported' }, () => {
          instant = true;
        });
        return { effect, instant, cleared: window.run() === null };
      }),
      { effect: 'fade', instant: true, cleared: true },
    );
    for (const [effect, direction] of [
      ['zoom', 'out'],
      ['strips', 'lu'],
    ]) {
      const defaults = await page.evaluate(
        ({ effect, direction }) => {
          const sample = (settings) => {
            window.play(window.fixtureSvg, { durationMs: 1000, effect, ...settings });
            const run = window.run();
            run.animations.forEach((animation) => {
              animation.pause();
              animation.currentTime = 500;
            });
            const layers = [...run.overlay.children].map((layer) => ({
              transform: getComputedStyle(layer).transform,
              clipPath: getComputedStyle(layer).clipPath,
              zIndex: getComputedStyle(layer).zIndex,
            }));
            window.cancel();
            return layers;
          };
          return { omitted: sample({}), explicit: sample({ direction }) };
        },
        { effect, direction },
      );
      assert.deepEqual(defaults.omitted, defaults.explicit);
    }
    for (const direction of ['lu', 'ru', 'ld', 'rd']) {
      const coverage = await page.evaluate((direction) => {
        window.play('', { effect: 'strips', direction, durationMs: 1000 });
        const run = window.run();
        const layers = [...run.overlay.children].slice(1);
        layers.forEach((layer) => {
          layer.style.pointerEvents = 'auto';
        });
        const bounds = run.overlay.getBoundingClientRect();
        const snapshots = [0, 500, 1000].map((time) => {
          run.animations.forEach((animation) => {
            animation.pause();
            animation.currentTime = time;
          });
          return layers.map((layer, row) =>
            [0.1, 0.5, 0.9].map((x) =>
              run.overlay
                .getRootNode()
                .elementsFromPoint(
                  bounds.x + bounds.width * x,
                  bounds.y + (bounds.height * (row + 0.5)) / 8,
                )
                .includes(layer),
            ),
          );
        });
        window.cancel();
        return { snapshots, removed: !run.overlay.isConnected };
      }, direction);
      assert.ok(coverage.removed);
      assert.ok(coverage.snapshots[0].flat().every((value) => !value));
      assert.ok(coverage.snapshots[2].flat().every(Boolean));
      const middle = coverage.snapshots[1];
      const leading = direction.endsWith('u') ? 7 : 0;
      assert.deepEqual(middle[leading], [true, true, true]);
      assert.deepEqual(middle[7 - leading], [false, false, false]);
      assert.deepEqual(
        middle[direction.endsWith('u') ? 4 : 3],
        direction.startsWith('l') ? [false, true, true] : [true, true, false],
      );
    }
    const newsflash = await page.evaluate(() => {
      window.play(window.fixtureSvg, { effect: 'newsflash', durationMs: 1200 });
      const run = window.run();
      const animation = run.animations[0];
      animation.pause();
      const samples = [0, 300, 600, 1200].map((time) => {
        animation.currentTime = time;
        const matrix = new DOMMatrix(getComputedStyle(run.overlay.children[1]).transform);
        return [matrix.a, matrix.b, matrix.c, matrix.d];
      });
      window.cancel();
      return samples;
    });
    const expectedNewsflash = [
      [0, 0, 0, 0],
      [0, -0.25, 0.25, 0],
      [-0.5, 0, 0, -0.5],
      [1, 0, 0, 1],
    ];
    newsflash.forEach((sample, i) =>
      sample.forEach((value, j) => assert.ok(Math.abs(value - expectedNewsflash[i][j]) < 0.00001)),
    );
    const random = await page.evaluate(() => {
      const originalRandom = Math.random;
      const settings = Object.freeze({ effect: 'random', durationMs: 1300 });
      const selected = [];
      try {
        for (const value of [0, 0.999999]) {
          Math.random = () => value;
          window.play('', settings);
          const run = window.run();
          selected.push({
            effect: run.overlay.dataset.transitionPlayback,
            duration: run.animations[0].effect.getTiming().duration,
          });
          window.cancel();
        }
      } finally {
        Math.random = originalRandom;
      }
      return { selected, settings };
    });
    assert.deepEqual(random, {
      selected: [
        { effect: 'fade', duration: 1300 },
        { effect: 'newsflash', duration: 1300 },
      ],
      settings: { effect: 'random', durationMs: 1300 },
    });
    for (const direction of ['horz', 'vert']) {
      const strips = await page.evaluate((direction) => {
        window.play('', { effect: 'blinds', direction, durationMs: 1000 });
        const run = window.run();
        run.animations.forEach((a) => {
          a.pause();
          a.currentTime = 500;
        });
        const clips = [...run.overlay.children]
          .slice(1)
          .map((layer) => getComputedStyle(layer).clipPath);
        const count = run.animations.length;
        window.cancel();
        return { clips, count };
      }, direction);
      assert.equal(strips.count, 8);
      assert.equal(
        strips.clips[0],
        direction === 'horz' ? 'inset(0px 93.75% 0px 0%)' : 'inset(0% 0px 93.75%)',
      );
      assert.equal(
        strips.clips[7],
        direction === 'horz' ? 'inset(0px 6.25% 0px 87.5%)' : 'inset(87.5% 0px 6.25%)',
      );
    }
    const cross = await page.evaluate(() => {
      window.play('', { effect: 'plus', durationMs: 1000 });
      const run = window.run();
      run.animations[0].pause();
      run.animations[0].currentTime = 500;
      return getComputedStyle(run.overlay.children[1]).clipPath;
    });
    assert.equal(
      cross,
      'polygon(25% -25%, 75% -25%, 75% 25%, 125% 25%, 125% 75%, 75% 75%, 75% 125%, 25% 125%, 25% 75%, -25% 75%, -25% 25%, 25% 25%)',
    );
    await page.evaluate(() => window.run().animations.forEach((a) => a.finish()));
    await page.waitForFunction(() => window.run() === null);
    for (const spokes of [1, 2, 3, 4, 8]) {
      const wheel = await page.evaluate((spokes) => {
        window.play('', { effect: 'wheel', spokes, durationMs: 1000 });
        const run = window.run();
        const animation = run.animations[0];
        animation.pause();
        animation.currentTime = 500;
        const frames = animation.effect.getKeyframes();
        const result = {
          middle: getComputedStyle(run.overlay.children[1]).clipPath,
          start: frames[0].clipPath,
          end: frames.at(-1).clipPath,
        };
        return result;
      }, spokes);
      assert.notEqual(wheel.middle, wheel.start);
      assert.notEqual(wheel.middle, wheel.end);
      assert.equal(wheel.middle.split(',').length, spokes * 27);
      await page.evaluate(() => window.run().animations.forEach((a) => a.finish()));
      await page.waitForFunction(() => window.run() === null);
    }
    for (const [effect, direction, columns, rows] of [
      ['dissolve', 'horz', 24, 14],
      ['randomBar', 'horz', 1, 64],
      ['randomBar', 'vert', 64, 1],
      ['checker', 'horz', 8, 6],
      ['checker', 'vert', 8, 6],
    ]) {
      const coverage = await page.evaluate(
        ({ effect, direction, columns, rows }) => {
          window.play('', { effect, direction, durationMs: 1000 });
          const run = window.run();
          const layer = run.overlay.children[1];
          layer.style.pointerEvents = 'auto';
          const root = document.getElementById('slide').shadowRoot;
          const box = run.overlay.getBoundingClientRect();
          const animation = run.animations[0];
          animation.pause();
          const counts = [0, 500, 1000].map((time) => {
            animation.currentTime = time;
            let count = 0;
            for (let row = 0; row < rows; row++)
              for (let column = 0; column < columns; column++) {
                if (
                  root.elementFromPoint(
                    box.x + ((column + 0.413) * box.width) / columns,
                    box.y + ((row + 0.617) * box.height) / rows,
                  ) === layer
                )
                  count++;
              }
            return count;
          });
          window.cancel();
          return counts;
        },
        { effect, direction, columns, rows },
      );
      assert.deepEqual(
        coverage,
        [0, (columns * rows) / 2, columns * rows],
        effect + ' ' + direction,
      );
    }
    for (const direction of ['horz', 'vert']) {
      const halves = await page.evaluate((direction) => {
        window.play('', { effect: 'comb', direction, durationMs: 1000 });
        const run = window.run();
        run.animations.forEach((a) => {
          a.pause();
          a.currentTime = 500;
        });
        const layers = [...run.overlay.children].slice(1);
        layers.forEach((layer) => {
          layer.style.pointerEvents = 'auto';
        });
        const root = document.getElementById('slide').shadowRoot;
        const box = run.overlay.getBoundingClientRect();
        const result = layers.map((layer, i) =>
          [0.25, 0.75].map((part) => {
            const x = direction === 'vert' ? (i + 0.5) / 8 : part;
            const y = direction === 'vert' ? part : (i + 0.5) / 8;
            return root.elementFromPoint(box.x + x * box.width, box.y + y * box.height) === layer;
          }),
        );
        window.cancel();
        return result;
      }, direction);
      assert.deepEqual(
        halves,
        Array.from({ length: 8 }, (_, i) => (i % 2 ? [false, true] : [true, false])),
      );
    }
    for (const effect of ['wheel', 'wheelReverse']) {
      const coverage = await page.evaluate((effect) => {
        window.play('', { effect, spokes: 1, durationMs: 1000 });
        const run = window.run();
        run.animations[0].pause();
        run.animations[0].currentTime = 500;
        const layer = run.overlay.children[1];
        layer.style.pointerEvents = 'auto';
        const root = document.getElementById('slide').shadowRoot;
        const box = run.overlay.getBoundingClientRect();
        const visible = [
          [0.2, 0.25],
          [0.8, 0.25],
          [0.2, 0.75],
          [0.8, 0.75],
        ].map(
          ([x, y]) =>
            root.elementFromPoint(box.x + x * box.width, box.y + y * box.height) === layer,
        );
        window.cancel();
        return visible;
      }, effect);
      assert.deepEqual(
        coverage,
        effect === 'wheel' ? [false, true, false, true] : [true, false, true, false],
      );
    }
    const wedge = await page.evaluate(() => {
      window.play('', { effect: 'wedge', durationMs: 1000 });
      const run = window.run();
      run.animations[0].pause();
      run.animations[0].currentTime = 500;
      const layer = run.overlay.children[1];
      layer.style.pointerEvents = 'auto';
      const root = document.getElementById('slide').shadowRoot;
      const box = run.overlay.getBoundingClientRect();
      const visible = [
        [0.2, 0.25],
        [0.8, 0.25],
        [0.2, 0.75],
        [0.8, 0.75],
      ].map(
        ([x, y]) => root.elementFromPoint(box.x + x * box.width, box.y + y * box.height) === layer,
      );
      window.cancel();
      return visible;
    });
    assert.deepEqual(wedge, [true, true, false, false]);
    await page.evaluate(() => {
      window.play('', { effect: 'checker', durationMs: 1000 });
      window.run().animations[0].pause();
      window.run().animations[0].currentTime = 500;
      document.getElementById('slide').style.width = '320px';
    });
    await page.waitForFunction(() => {
      const frame = window.run().animations[0].effect.getKeyframes().at(-1);
      return frame.clipPath.includes('h 40');
    });
    await page.evaluate(() => window.cancel());
  } finally {
    await browser.close();
  }
});

test('transition sound continues across silent slides and stops on replacement or exit', async () => {
  const source = await readFile(
    new URL('../../src/transition-playback.ts', import.meta.url),
    'utf8',
  );
  const runtime = source.slice(source.indexOf('`') + 1, source.lastIndexOf('`'));
  const browser = await chromium.launch({ headless: true });
  try {
    const page = await browser.newPage();
    const result = await page.evaluate(async (runtime) => {
      const sounds = [];
      window.Audio = class extends EventTarget {
        constructor(src) {
          super();
          this.src = src;
          this.paused = false;
          sounds.push(this);
        }
        play() {
          return this.src.endsWith('reject')
            ? Promise.reject(new Error('Blocked'))
            : Promise.resolve();
        }
        pause() {
          this.paused = true;
        }
        removeAttribute() {
          this.src = '';
        }
        load() {}
      };
      window.eval(
        runtime +
          ';window.sound=playTransitionSound;window.stopSound=stopTransitionSound;window.cancelVisual=cancelTransitionPlayback;',
      );
      window.sound({ kind: 'play', base64: 'first', loop: true });
      const first = sounds[0];
      window.cancelVisual();
      window.sound({ kind: 'none' });
      const continued = !first.paused && first.loop && first.src === 'data:audio/wav;base64,first';
      window.sound({ kind: 'play', base64: 'second', loop: false });
      const replaced = first.paused && !first.src && !sounds[1].loop;
      first.dispatchEvent(new Event('ended'));
      const staleIgnored = !sounds[1].paused;
      window.sound({ kind: 'stop' });
      const stopped = sounds[1].paused;
      window.sound({ kind: 'play', base64: 'third', loop: false });
      sounds[2].dispatchEvent(new Event('ended'));
      const ended = sounds[2].paused && !sounds[2].src;
      window.sound({ kind: 'play', base64: 'reject', loop: true });
      await Promise.resolve();
      const rejected = sounds[3].paused && !sounds[3].src;
      window.sound({ kind: 'play', base64: 'last', loop: true });
      window.stopSound();
      return {
        continued,
        replaced,
        staleIgnored,
        stopped,
        ended,
        rejected,
        exited: sounds[4].paused,
      };
    }, runtime);
    assert.deepEqual(result, {
      continued: true,
      replaced: true,
      staleIgnored: true,
      stopped: true,
      ended: true,
      rejected: true,
      exited: true,
    });
  } finally {
    await browser.close();
  }
});
