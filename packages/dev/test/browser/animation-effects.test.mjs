// Fly, zoom and spin against a real slide, in a real browser.
//
// These drive the shared player over a hand-built SVG rather than a compiled
// deck, because what is being checked is geometry: which edge a shape comes
// from, whether a rotated group turns the direction with it, and whether two
// effects over one shape both survive. The deck-level path — a saved file read
// back and played in the preview — is covered by `animation-playback`.
//
// Every reading is taken with the animations paused at a stated time, so a
// probe cannot pass by arriving a frame early or late.

import assert from 'node:assert/strict';
import test from 'node:test';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { chromium } from 'playwright';
import { compile, Presentation, Slide, Text } from '@office-kit/pptx-dsl';
import { getSlideShapes, getSlides, savePresentation, setShapeAnimation } from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

const DURATION = 2000;

/** A step as `getSlideAnimations` reports one, with the fields a test varies. */
const step = (over) => ({
  id: 1,
  target: { kind: 'shape', shapeId: 10 },
  targetShapeIds: [10],
  effect: 'flyIn',
  direction: null,
  presetId: 2,
  presetClass: 'entr',
  start: 'click',
  durationMs: DURATION,
  delayMs: 0,
  valueAfterEnd: 'held',
  buildByParagraph: false,
  buildLevel: null,
  sequence: 'mainSeq',
  playable: true,
  editable: true,
  ...over,
});

/**
 * A slide 1000×600 user units wide, drawn at half that on screen. The second
 * shape hangs under a quarter turn, which is what a deck rotating a shape or
 * its group leaves above the marker the player animates.
 */
const SLIDE = `
  <svg viewBox="0 0 1000 600" width="1000" height="600" style="display:block">
    <g data-pptx-shape-id="10"><rect x="400" y="250" width="200" height="100" fill="#c33"/></g>
    <g transform="rotate(90 500 300)">
      <g data-pptx-shape-id="11"><rect x="400" y="250" width="200" height="100" fill="#3c3"/></g>
    </g>
    <g data-pptx-shape-id="12">
      <g data-pptx-paragraph="0"><rect x="100" y="60" width="120" height="40" fill="#33c"/></g>
      <g data-pptx-paragraph="1"><rect x="100" y="120" width="120" height="40" fill="#36c"/></g>
    </g>
  </svg>`;

/**
 * Runs one scene: builds the slide, plays it the way `go` says, then pauses
 * every animation at `at` milliseconds and reports what each named element
 * shows. `at` of `null` leaves the animations where they stand.
 */
const scene = (page, { steps, go, at, watch, reduced }) =>
  page.evaluate(
    async ({ steps, go, at, watch, reduced, slide }) => {
      const { createAnimationPlayer } = await import('/animation-player.js');
      document.body.style.margin = '0';
      for (const old of document.querySelectorAll('[data-scene]')) old.remove();
      const root = document.createElement('div');
      root.dataset.scene = '1';
      root.innerHTML = slide;
      document.body.append(root);

      const player = createAnimationPlayer({
        root,
        steps,
        reducedMotion: () => reduced === true,
      });
      player.reset();
      if (go.resume) player.resume(go.resume[0], go.resume[1]);
      else for (let n = 0; n < (go.advance ?? 1); n += 1) player.advance();

      const deadline = Date.now() + 5000;
      // As soon as anything being watched is animating, everything is paused —
      // waiting for *all* of them would sit out a whole effect on a scene where
      // one of the elements is deliberately still waiting for its own click.
      const started = () =>
        at === null ||
        watch.some((selector) => (root.querySelector(selector)?.getAnimations().length ?? 0) > 0);
      while (!started() && Date.now() < deadline) await new Promise((r) => setTimeout(r, 20));

      const svg = root.querySelector('svg').getBoundingClientRect();
      const read = (selector) => {
        const el = root.querySelector(selector);
        if (el === null) return null;
        const animations = el.getAnimations();
        for (const animation of animations) {
          animation.pause();
          // `'hold'` reads the clock the player itself set, which is the only
          // way a resume can be shown to have restored a moment rather than
          // having one handed to it afterwards.
          if (at !== null && at !== 'hold') animation.currentTime = at;
        }
        const box = el.getBoundingClientRect();
        const style = getComputedStyle(el);
        return {
          animations: animations.length,
          timeMs: animations.length === 0 ? null : Number(animations[0].currentTime),
          translate: style.translate,
          rotate: style.rotate,
          scale: style.scale,
          opacity: style.opacity,
          visibility: el.style.visibility || 'visible',
          top: box.top,
          bottom: box.bottom,
          left: box.left,
          right: box.right,
          width: box.width,
          height: box.height,
        };
      };
      const out = {};
      for (const selector of watch) out[selector] = read(selector);
      return { slide: svg, unsupported: player.unsupported.map((u) => u.reason), ...out };
    },
    { steps, go, at, watch, reduced, slide: SLIDE },
  );

const shape = (id) => `[data-pptx-shape-id="${id}"]`;

/**
 * Where shape 10 was drawn, as user units within the slide. The page around the
 * slide has a height of its own, so every place is read against the slide's own
 * box rather than the viewport's.
 */
const DRAWN = { left: 400, top: 250, width: 200, height: 100 };

const assertDrawnPlace = (result, seen, label) => {
  assert.ok(
    Math.abs(seen.left - (result.slide.left + DRAWN.left)) < 1,
    `${label} left ${seen.left} vs ${result.slide.left + DRAWN.left}`,
  );
  assert.ok(
    Math.abs(seen.top - (result.slide.top + DRAWN.top)) < 1,
    `${label} top ${seen.top} vs ${result.slide.top + DRAWN.top}`,
  );
  assert.ok(Math.abs(seen.width - DRAWN.width) < 1, `${label} width ${seen.width}`);
};

/** Whether a computed individual transform actually moves the element. */
const moves = (value) => value !== 'none' && value !== '' && !/^0px 0px$/.test(value);
const turns = (value) => value !== 'none' && value !== '' && value !== '0deg';
const scales = (value) => value !== 'none' && value !== '' && value !== '1';

const withBrowser = async (body) => {
  const dir = await mkdtemp(join(tmpdir(), 'ok-anim-effects-'));
  const deck = await compile(
    Presentation({
      children: Slide({
        children: Text({ x: 1, y: 1, width: 8, height: 2, children: 'effects' }),
      }),
    }),
  );
  await writeFile(join(dir, 'source.pptx'), await savePresentation(deck));
  const file = join(dir, 'deck.tsx');
  await writeFile(
    file,
    `import {readFile} from 'node:fs/promises';import {Presentation} from '@office-kit/pptx-dsl';` +
      `export default <Presentation source={await readFile(${JSON.stringify(join(dir, 'source.pptx'))})} />;`,
  );
  const preview = await startPreview(file);
  const browser = await chromium.launch({ headless: true });
  const page = await browser.newPage({ viewport: { width: 1200, height: 800 } });
  const errors = [];
  page.on('pageerror', (error) => errors.push(error.message));
  await page.goto(preview.url);
  try {
    await body(page);
    assert.deepEqual(errors, []);
  } finally {
    await browser.close();
    await preview.close();
    await rm(dir, { recursive: true, force: true });
  }
};

test('a fly comes from, and leaves by, the edge the deck names', async () => {
  await withBrowser(async (page) => {
    for (const direction of ['bottom', 'top', 'left', 'right']) {
      const start = await scene(page, {
        steps: [step({ effect: 'flyIn', direction })],
        go: { advance: 1 },
        at: 0,
        watch: [shape(10)],
      });
      const at = start[shape(10)];
      const edge = {
        bottom: () => assert.ok(at.top >= start.slide.bottom - 1, `top ${at.top}`),
        top: () => assert.ok(at.bottom <= start.slide.top + 1, `bottom ${at.bottom}`),
        left: () => assert.ok(at.right <= start.slide.left + 1, `right ${at.right}`),
        right: () => assert.ok(at.left >= start.slide.right - 1, `left ${at.left}`),
      };
      edge[direction]();
      assert.equal(at.visibility, 'visible', direction);

      // And it is back where the renderer drew it once the effect is over.
      const end = await scene(page, {
        steps: [step({ effect: 'flyIn', direction })],
        go: { advance: 1 },
        at: DURATION,
        watch: [shape(10)],
      });
      assertDrawnPlace(end, end[shape(10)], direction);
    }

    // An exit runs the same journey the other way.
    const out = await scene(page, {
      steps: [step({ effect: 'flyOut', direction: 'right', presetClass: 'exit' })],
      go: { advance: 1 },
      at: DURATION,
      watch: [shape(10)],
    });
    assert.ok(out[shape(10)].left >= out.slide.right - 1, `left ${out[shape(10)].left}`);
  });
});

test('a fly under a rotated group still comes from the slide’s own edge', async () => {
  await withBrowser(async (page) => {
    const target = { kind: 'shape', shapeId: 11 };
    const rotated = await scene(page, {
      steps: [step({ effect: 'flyIn', direction: 'bottom', target, targetShapeIds: [11] })],
      go: { advance: 1 },
      at: 0,
      watch: [shape(11)],
    });
    // The marker's own coordinates are a quarter turn away from the slide's, so
    // a translation written straight down in them would send the shape out of
    // the side instead.
    assert.ok(
      rotated[shape(11)].top >= rotated.slide.bottom - 1,
      `top ${rotated[shape(11)].top} vs slide bottom ${rotated.slide.bottom}`,
    );

    const landed = await scene(page, {
      steps: [step({ effect: 'flyIn', direction: 'bottom', target, targetShapeIds: [11] })],
      go: { advance: 1 },
      at: DURATION,
      watch: [shape(11)],
    });
    // Back under its group's own quarter turn: the rotated rect is 100 wide.
    assert.ok(Math.abs(landed[shape(11)].width - 100) < 1, `width ${landed[shape(11)].width}`);
    assert.ok(landed[shape(11)].top > landed.slide.top, 'inside the slide again');
  });
});

test('a fly over a build moves one paragraph at a time', async () => {
  await withBrowser(async (page) => {
    const para = (n) => ({
      kind: 'paragraphs',
      shapeId: 12,
      firstParagraph: n,
      lastParagraph: n,
    });
    const steps = [0, 1].map((n) =>
      step({
        id: n + 1,
        effect: 'flyIn',
        direction: 'left',
        target: para(n),
        targetShapeIds: [12],
        buildByParagraph: true,
      }),
    );
    const first = await scene(page, {
      steps,
      go: { advance: 1 },
      at: 0,
      watch: ['[data-pptx-paragraph="0"]', '[data-pptx-paragraph="1"]'],
    });
    assert.deepEqual(first.unsupported, []);
    assert.ok(first['[data-pptx-paragraph="0"]'].right <= first.slide.left + 1, 'first off-slide');
    // The second paragraph is still waiting for its own click, so it is off the
    // slide the way an entrance leaves its target: hidden, not moved.
    assert.equal(first['[data-pptx-paragraph="1"]'].visibility, 'hidden');
    assert.equal(first['[data-pptx-paragraph="1"]'].animations, 0);
  });
});

test('a fly and a spin over one shape both take effect, playing forward and on resume', async () => {
  await withBrowser(async (page) => {
    const steps = [
      step({ id: 1, effect: 'flyIn', direction: 'bottom' }),
      step({
        id: 2,
        effect: 'spin',
        direction: null,
        presetId: 8,
        presetClass: 'emph',
        start: 'withPrevious',
      }),
    ];
    const mid = DURATION / 2;
    const forward = await scene(page, { steps, go: { advance: 1 }, at: mid, watch: [shape(10)] });
    const seen = forward[shape(10)];
    // Written to one property, the effect that begins second would replace the
    // first outright and only the turn would survive.
    assert.equal(seen.animations, 2);
    assert.ok(moves(seen.translate), `translate ${seen.translate}`);
    assert.ok(turns(seen.rotate), `rotate ${seen.rotate}`);

    // What the player itself put on the clock, not a time handed to it after
    // the fact: joining a stop half-way through has to start both effects
    // half-way through.
    const joined = (
      await scene(page, { steps, go: { resume: [1, mid] }, at: 'hold', watch: [shape(10)] })
    )[shape(10)];
    assert.equal(joined.animations, 2);
    assert.ok(Math.abs(joined.timeMs - mid) < 200, `clock ${joined.timeMs}`);
    assert.ok(moves(joined.translate), `resumed translate ${joined.translate}`);
    assert.ok(turns(joined.rotate), `resumed rotate ${joined.rotate}`);
    // Half a turn in, give or take the moment it took to pause it.
    assert.ok(/^17\d(\.\d+)?deg$|^18\d(\.\d+)?deg$/.test(joined.rotate), joined.rotate);

    // And wound to the same moment it shows the same picture as playing there.
    const again = (
      await scene(page, { steps, go: { resume: [1, mid] }, at: mid, watch: [shape(10)] })
    )[shape(10)];
    assert.ok(Math.abs(again.top - seen.top) < 1, `top ${again.top} vs ${seen.top}`);
    assert.equal(again.rotate, seen.rotate);
  });
});

test('a zoom and a spin over one shape both take effect', async () => {
  await withBrowser(async (page) => {
    const steps = [
      step({ id: 1, effect: 'zoomIn', direction: null, presetId: 23 }),
      step({
        id: 2,
        effect: 'spin',
        direction: null,
        presetId: 8,
        presetClass: 'emph',
        start: 'withPrevious',
      }),
    ];
    const seen = (
      await scene(page, {
        steps,
        go: { advance: 1 },
        at: DURATION / 2,
        watch: [shape(10)],
      })
    )[shape(10)];
    assert.equal(seen.animations, 2);
    assert.ok(scales(seen.scale), `scale ${seen.scale}`);
    assert.ok(turns(seen.rotate), `rotate ${seen.rotate}`);

    // Half-way through, the shape covers less than its drawn size but is not
    // gone: the zoom grows it about its own centre rather than the viewport's.
    assert.ok(seen.width > 0 && seen.width < DRAWN.width, `width ${seen.width}`);

    const joined = (
      await scene(page, {
        steps,
        go: { resume: [1, DURATION / 2] },
        at: 'hold',
        watch: [shape(10)],
      })
    )[shape(10)];
    assert.equal(joined.animations, 2);
    assert.ok(Math.abs(joined.timeMs - DURATION / 2) < 200, `clock ${joined.timeMs}`);
    assert.ok(scales(joined.scale), `resumed scale ${joined.scale}`);
    assert.ok(turns(joined.rotate), `resumed rotate ${joined.rotate}`);

    const again = (
      await scene(page, {
        steps,
        go: { resume: [1, DURATION / 2] },
        at: DURATION / 2,
        watch: [shape(10)],
      })
    )[shape(10)];
    assert.equal(again.scale, seen.scale);
    assert.equal(again.rotate, seen.rotate);
  });
});

test('a spin leaves the shape where it stood, and never shows or hides it', async () => {
  await withBrowser(async (page) => {
    const spin = step({
      effect: 'spin',
      direction: null,
      presetId: 8,
      presetClass: 'emph',
    });
    const before = await scene(page, {
      steps: [spin],
      go: { advance: 0 },
      at: null,
      watch: [shape(10)],
    });
    // An emphasis effect animates a shape that is already on the slide, so it
    // is not held back before its turn the way an entrance's target is.
    assert.equal(before[shape(10)].visibility, 'visible');

    const after = await scene(page, {
      steps: [spin],
      go: { advance: 1 },
      at: DURATION,
      watch: [shape(10)],
    });
    assert.equal(after[shape(10)].visibility, 'visible');
    assertDrawnPlace(after, after[shape(10)], 'after a spin');
  });
});

test('a spin never puts a shape on the slide that no entrance has shown yet', async () => {
  await withBrowser(async (page) => {
    const spin = step({
      id: 1,
      effect: 'spin',
      direction: null,
      presetId: 8,
      presetClass: 'emph',
    });
    const entrance = step({ id: 2, effect: 'fadeIn', direction: null, presetId: 10 });
    const steps = [spin, entrance];

    // Playing the spin's own stop: its target is one a later entrance is still
    // holding back, so turning it must not reveal it.
    for (const at of [0, DURATION / 2, DURATION]) {
      const seen = (await scene(page, { steps, go: { advance: 1 }, at, watch: [shape(10)] }))[
        shape(10)
      ];
      assert.equal(seen.visibility, 'hidden', `advance at ${at}`);
    }
    // And joining that stop half-way lands on the same picture.
    const resumed = (
      await scene(page, {
        steps,
        go: { resume: [1, DURATION / 2] },
        at: DURATION / 2,
        watch: [shape(10)],
      })
    )[shape(10)];
    assert.equal(resumed.visibility, 'hidden');
    assert.ok(turns(resumed.rotate), `rotate ${resumed.rotate}`);

    // The entrance's own stop still shows it.
    const shown = (
      await scene(page, { steps, go: { advance: 2 }, at: DURATION, watch: [shape(10)] })
    )[shape(10)];
    assert.equal(shown.visibility, 'visible');
  });
});

test('a spin after an exit leaves the shape off the slide', async () => {
  await withBrowser(async (page) => {
    const steps = [
      step({ id: 1, effect: 'fadeOut', direction: null, presetId: 10, presetClass: 'exit' }),
      step({ id: 2, effect: 'spin', direction: null, presetId: 8, presetClass: 'emph' }),
    ];
    for (const at of [0, DURATION / 2, DURATION]) {
      const seen = (await scene(page, { steps, go: { advance: 2 }, at, watch: [shape(10)] }))[
        shape(10)
      ];
      assert.equal(seen.visibility, 'hidden', `at ${at}`);
    }
    const resumed = (
      await scene(page, {
        steps,
        go: { resume: [2, DURATION / 2] },
        at: DURATION / 2,
        watch: [shape(10)],
      })
    )[shape(10)];
    assert.equal(resumed.visibility, 'hidden');
  });
});

test('an entrance running beside a spin still shows its shape', async () => {
  await withBrowser(async (page) => {
    const steps = [
      step({ id: 1, effect: 'fadeIn', direction: null, presetId: 10 }),
      step({
        id: 2,
        effect: 'spin',
        direction: null,
        presetId: 8,
        presetClass: 'emph',
        start: 'withPrevious',
      }),
    ];
    const during = (
      await scene(page, { steps, go: { advance: 1 }, at: DURATION / 2, watch: [shape(10)] })
    )[shape(10)];
    assert.equal(during.visibility, 'visible');
    assert.ok(turns(during.rotate), `rotate ${during.rotate}`);
    assert.ok(Number(during.opacity) > 0.2 && Number(during.opacity) < 0.8, during.opacity);

    const after = (
      await scene(page, { steps, go: { advance: 1 }, at: DURATION, watch: [shape(10)] })
    )[shape(10)];
    assert.equal(after.visibility, 'visible');
  });
});

test(
  'the preview plays a fly, a zoom and a spin out of the file it read',
  { timeout: 120000 },
  async () => {
    // The scenes above build their own SVG to pin geometry down. This one goes
    // the whole way: the effects are written by the public API, saved, read
    // back from the bytes and played over the renderer's own markers.
    const dir = await mkdtemp(join(tmpdir(), 'ok-anim-file-'));
    let browser;
    let preview;
    try {
      const deck = await compile(
        Presentation({
          children: Slide({
            children: [
              Text({ x: 1, y: 1, width: 4, height: 1, children: 'flying' }),
              Text({ x: 1, y: 3, width: 4, height: 1, children: 'zooming' }),
              Text({ x: 1, y: 5, width: 4, height: 1, children: 'spinning' }),
            ],
          }),
        }),
      );
      const [flier, zoomer, spinner] = getSlideShapes(getSlides(deck)[0]);
      setShapeAnimation(flier, { effect: 'flyIn', direction: 'left', durationMs: 2000 });
      setShapeAnimation(zoomer, { effect: 'zoomIn', durationMs: 2000, start: 'withPrevious' });
      setShapeAnimation(spinner, { effect: 'spin', durationMs: 2000, start: 'withPrevious' });

      await writeFile(join(dir, 'source.pptx'), await savePresentation(deck));
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {readFile} from 'node:fs/promises';` +
          `import {Presentation} from '@office-kit/pptx-dsl';` +
          `export default <Presentation source={await readFile(${JSON.stringify(join(dir, 'source.pptx'))})} />;`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1280, height: 800 } });
      const errors = [];
      page.on('pageerror', (error) => errors.push(error.message));
      await page.goto(preview.url);
      await page.getByRole('button', { name: 'Preview', exact: true }).click();
      await page.waitForFunction(() => state.slides.length === 1);
      await page.getByRole('button', { name: 'Present', exact: true }).click();
      // The spun shape is the one nothing reveals, so it is on the slide from
      // the start; the two entrances are not.
      assert.deepEqual(
        await page.evaluate(() =>
          Array.from(slideRoot().querySelectorAll('[data-pptx-shape-id]')).map(
            (el) => el.style.visibility || 'visible',
          ),
        ),
        ['hidden', 'hidden', 'visible'],
      );
      assert.equal(await page.evaluate(() => animationPlayer.unsupported.length), 0);

      await page.keyboard.press('ArrowRight');
      const seen = await page.evaluate(() => {
        const slide = slideRoot().querySelector('svg').getBoundingClientRect();
        return Array.from(slideRoot().querySelectorAll('[data-pptx-shape-id]')).map((el) => {
          const animations = el.getAnimations();
          for (const animation of animations) {
            animation.pause();
            animation.currentTime = 1000;
          }
          const style = getComputedStyle(el);
          const box = el.getBoundingClientRect();
          return {
            translate: style.translate,
            scale: style.scale,
            rotate: style.rotate,
            fromLeft: box.left - slide.left,
            width: box.width,
            slideWidth: slide.width,
          };
        });
      });
      // It comes in from the left edge the file names, half-way there.
      assert.ok(seen[0].translate !== 'none', `translate ${seen[0].translate}`);
      assert.ok(seen[0].fromLeft < 0, `left of the slide: ${seen[0].fromLeft}`);
      // The zoom is half grown and the spin half turned, each on its own shape.
      assert.ok(seen[1].scale !== 'none' && seen[1].scale !== '1', `scale ${seen[1].scale}`);
      assert.ok(seen[2].rotate !== 'none' && seen[2].rotate !== '0deg', `rotate ${seen[2].rotate}`);

      // Let them run on from where they were paused, so the stop finishes the
      // way it would have without a probe in the middle of it.
      await page.evaluate(() => {
        for (const el of slideRoot().querySelectorAll('[data-pptx-shape-id]'))
          for (const animation of el.getAnimations()) animation.play();
      });
      await page.waitForFunction(() => !animationPlayer || !animationPlayer.running, null, {
        timeout: 15000,
      });
      assert.deepEqual(
        await page.evaluate(() =>
          Array.from(slideRoot().querySelectorAll('[data-pptx-shape-id]')).map(
            (el) => el.style.visibility || 'visible',
          ),
        ),
        ['visible', 'visible', 'visible'],
      );
      assert.deepEqual(errors, []);
    } finally {
      if (browser) await browser.close();
      if (preview) await preview.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

test('reduced motion puts every effect where it ends without moving anything', async () => {
  await withBrowser(async (page) => {
    for (const effect of ['flyIn', 'zoomIn', 'spin']) {
      const result = await scene(page, {
        steps: [
          step({
            effect,
            direction: effect === 'flyIn' ? 'bottom' : null,
            presetClass: effect === 'spin' ? 'emph' : 'entr',
          }),
        ],
        go: { advance: 1 },
        at: null,
        reduced: true,
        watch: [shape(10)],
      });
      const seen = result[shape(10)];
      assert.equal(seen.animations, 0, effect);
      assertDrawnPlace(result, seen, effect);
      assert.equal(seen.visibility, 'visible', effect);
      assert.ok(!moves(seen.translate), `${effect} translate ${seen.translate}`);
      assert.ok(!turns(seen.rotate), `${effect} rotate ${seen.rotate}`);
      assert.ok(!scales(seen.scale), `${effect} scale ${seen.scale}`);
    }
  });
});
