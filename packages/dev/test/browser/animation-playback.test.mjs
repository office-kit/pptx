// Object animations in the preview's presentation mode.
//
// The decks are built through the public API and saved, so what the player
// reads is what a reload of a real file reports — not an in-memory model the
// writer happened to leave behind.

import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import { compile, Presentation, Slide, Text } from '@office-kit/pptx-dsl';
import {
  getSlideShapes,
  getSlides,
  savePresentation,
  setShapeAnimation,
  setSlideNotes,
  setSlideTransition,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

/** The visibility of every object on the slide being shown, in order. */
const visibilities = (page) =>
  page.evaluate(() =>
    Array.from(slideRoot().querySelectorAll('[data-pptx-shape-id]')).map(
      (el) => el.style.visibility || 'visible',
    ),
  );

const paragraphVisibilities = (page) =>
  page.evaluate(() =>
    Array.from(slideRoot().querySelectorAll('[data-pptx-paragraph]')).map(
      (el) => el.style.visibility || 'visible',
    ),
  );

const cursor = (page) =>
  page.evaluate(() =>
    animationPlayer ? { cursor: animationPlayer.cursor, stops: animationPlayer.stopCount } : null,
  );

/**
 * Waits for the slide to have played out what the last click started — the
 * delayed effects included. Finishing the animations by hand would step over
 * exactly the timer bugs these tests are here to catch.
 */
const idle = (page) =>
  page.waitForFunction(() => !animationPlayer || !animationPlayer.running, null, {
    timeout: 10000,
  });

const deckSource = (dir) =>
  `import {readFile} from 'node:fs/promises';import {Presentation} from '@office-kit/pptx-dsl';` +
  `export default <Presentation source={await readFile(${JSON.stringify(join(dir, 'source.pptx'))})} />;`;

const openPreview = async (dir, deck, slides, prepare) => {
  await writeFile(join(dir, 'source.pptx'), await savePresentation(deck));
  const file = join(dir, 'deck.tsx');
  await writeFile(file, deckSource(dir));
  const preview = await startPreview(file);
  const browser = await chromium.launch({ headless: true });
  const page = await browser.newPage({
    viewport: { width: 1280, height: 800 },
    reducedMotion: 'no-preference',
  });
  const errors = [];
  page.on('pageerror', (error) => errors.push(error.message));
  if (prepare) await prepare(page);
  await page.goto(preview.url);
  await page.getByRole('button', { name: 'Preview', exact: true }).click();
  await page.waitForFunction((n) => state.slides.length === n, slides);
  return { preview, browser, page, errors };
};

test(
  'clicks, keys and buttons drive one cursor through a build and the deck',
  { timeout: 120000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-animations-'));
    let session;
    try {
      const deck = await compile(
        Presentation({
          children: [
            Slide({
              children: [
                Text({ x: 1, y: 1, width: 4, height: 1, children: 'first' }),
                Text({ x: 1, y: 3, width: 4, height: 1, children: 'second' }),
                Text({ x: 1, y: 5, width: 4, height: 1, children: 'third' }),
              ],
            }),
            Slide({
              children: Text({ x: 1, y: 1, width: 8, height: 3, children: 'one\ntwo\nthree' }),
            }),
            Slide({ children: Text({ x: 1, y: 1, width: 8, height: 2, children: 'plain' }) }),
          ],
        }),
      );
      const slides = getSlides(deck);
      const [a, b, c] = getSlideShapes(slides[0]);
      // A click, one that runs with it, and one that follows both — the third
      // may not start before the first two have finished.
      setShapeAnimation(a, { effect: 'fadeIn', durationMs: 600 });
      setShapeAnimation(b, { effect: 'fadeIn', durationMs: 600, start: 'withPrevious' });
      setShapeAnimation(c, { effect: 'fadeIn', durationMs: 300, start: 'afterPrevious' });
      setShapeAnimation(getSlideShapes(slides[1])[0], {
        effect: 'fadeIn',
        durationMs: 200,
        byParagraph: true,
      });
      // The arriving slide is drawn into a layer of its own while the slide
      // being left is still on screen: the build must find the right one.
      setSlideTransition(slides[1], { effect: 'fade', speed: 'fast' });

      session = await openPreview(dir, deck, 3);
      const { page } = session;

      // The still preview shows the slide as it ends: nothing is hidden.
      assert.deepEqual(await visibilities(page), ['visible', 'visible', 'visible']);
      assert.equal(await cursor(page), null);

      await page.getByRole('button', { name: 'Present', exact: true }).click();
      assert.deepEqual(await visibilities(page), ['hidden', 'hidden', 'hidden']);
      assert.deepEqual(await cursor(page), { cursor: 0, stops: 1 });
      // Nothing on this deck is beyond the player, so it says nothing.
      assert.equal(await page.locator('#present-note').isVisible(), false);

      // One click opens the whole stop, but the effect chained onto the end of
      // the first two waits for them rather than starting with them.
      await page.keyboard.press('ArrowRight');
      assert.deepEqual(await cursor(page), { cursor: 1, stops: 1 });
      assert.deepEqual(await visibilities(page), ['visible', 'visible', 'hidden']);
      await idle(page);
      assert.deepEqual(await visibilities(page), ['visible', 'visible', 'visible']);

      // Only once the slide is played out does the next click move on.
      await page.keyboard.press('ArrowRight');
      await page.waitForFunction(() => index === 1);
      assert.deepEqual(await cursor(page), { cursor: 0, stops: 3 });
      assert.deepEqual(await paragraphVisibilities(page), ['hidden', 'hidden', 'hidden']);

      // A paragraph build advances a paragraph per click, by key or by click.
      await page.keyboard.press('Space');
      await idle(page);
      assert.deepEqual(await paragraphVisibilities(page), ['visible', 'hidden', 'hidden']);
      await page.mouse.click(640, 400);
      await idle(page);
      assert.deepEqual(await paragraphVisibilities(page), ['visible', 'visible', 'hidden']);

      // Back walks the build, not the deck.
      await page.keyboard.press('ArrowLeft');
      assert.equal(await page.evaluate(() => index), 1);
      assert.deepEqual(await paragraphVisibilities(page), ['visible', 'hidden', 'hidden']);

      // Stepping back off the top of a slide shows the previous one played out.
      await page.keyboard.press('ArrowLeft');
      await page.keyboard.press('ArrowLeft');
      await page.waitForFunction(() => index === 0);
      assert.deepEqual(await visibilities(page), ['visible', 'visible', 'visible']);
      assert.deepEqual(await cursor(page), { cursor: 1, stops: 1 });

      // Coming back forwards starts the build again.
      await page.keyboard.press('ArrowRight');
      await page.waitForFunction(() => index === 1);
      assert.deepEqual(await cursor(page), { cursor: 0, stops: 3 });
      assert.deepEqual(await paragraphVisibilities(page), ['hidden', 'hidden', 'hidden']);

      // An explicit jump lands with the build at the start too.
      await page.evaluate(() => selectSlide(1));
      assert.deepEqual(await cursor(page), { cursor: 0, stops: 3 });
      await page.keyboard.press('Home');
      await page.waitForFunction(() => index === 0);
      assert.deepEqual(await cursor(page), { cursor: 0, stops: 1 });

      // Fast clicks do not skip a step or leave one half-played.
      await page.keyboard.press('ArrowRight');
      await page.keyboard.press('ArrowRight');
      await page.keyboard.press('ArrowRight');
      await page.waitForFunction(() => index === 1);
      await idle(page);
      assert.deepEqual(await cursor(page), { cursor: 1, stops: 3 });
      assert.deepEqual(await paragraphVisibilities(page), ['visible', 'hidden', 'hidden']);

      // Leaving the show puts the slide back the way the preview draws it.
      await page.keyboard.press('Escape');
      await page.waitForFunction(() => presenting === false);
      assert.equal(await cursor(page), null);
      assert.deepEqual(await paragraphVisibilities(page), ['visible', 'visible', 'visible']);

      assert.deepEqual(session.errors, []);
    } finally {
      await session?.browser.close();
      await session?.preview.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

test(
  'a self-advancing slide waits for its build, and the last slide’s clicks stay reachable',
  { timeout: 120000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-animations-advance-'));
    let session;
    try {
      const deck = await compile(
        Presentation({
          children: [
            Slide({ children: Text({ x: 1, y: 1, width: 8, height: 2, children: 'one' }) }),
            Slide({
              children: [
                Text({ x: 1, y: 1, width: 4, height: 1, children: 'a' }),
                Text({ x: 1, y: 3, width: 4, height: 1, children: 'b' }),
              ],
            }),
          ],
        }),
      );
      const slides = getSlides(deck);
      // Nothing precedes it, so it starts as the slide appears — and runs far
      // longer than the slide's own advance, which may not cut it short.
      setShapeAnimation(getSlideShapes(slides[0])[0], {
        effect: 'fadeIn',
        durationMs: 1500,
        start: 'withPrevious',
      });
      setSlideTransition(slides[0], { effect: 'none', advanceAfterMs: 100 });
      for (const shape of getSlideShapes(slides[1])) {
        setShapeAnimation(shape, { effect: 'fadeIn', durationMs: 100 });
      }

      session = await openPreview(dir, deck, 2);
      const { page } = session;
      await page.getByRole('button', { name: 'Present', exact: true }).click();

      // The leading effect runs without a click…
      assert.deepEqual(await cursor(page), { cursor: 1, stops: 1 });
      assert.equal(await page.evaluate(() => animationPlayer.running), true);
      // …and the 100ms advance does not fire while it is still running.
      await page.waitForTimeout(900);
      assert.equal(await page.evaluate(() => index), 0);
      assert.equal(await page.evaluate(() => animationPlayer.running), true);
      await page.waitForFunction(() => index === 1, null, { timeout: 5000 });

      // On the last slide, with two clicks left, Next is still available.
      const next = page.locator('#present-next');
      const previous = page.locator('#present-prev');
      assert.deepEqual(await cursor(page), { cursor: 0, stops: 2 });
      assert.equal(await next.isDisabled(), false);
      await next.click();
      assert.deepEqual(await cursor(page), { cursor: 1, stops: 2 });
      assert.equal(await next.isDisabled(), false);
      await next.click();
      assert.deepEqual(await cursor(page), { cursor: 2, stops: 2 });
      // Nothing left on the slide and no slide after it.
      assert.equal(await next.isDisabled(), true);
      // Previous walks back into the build rather than being disabled at 0.
      assert.equal(await previous.isDisabled(), false);
      await previous.click();
      assert.deepEqual(await cursor(page), { cursor: 1, stops: 2 });

      assert.deepEqual(session.errors, []);
    } finally {
      await session?.browser.close();
      await session?.preview.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

test(
  'the presenter view keeps pace with what the audience can see',
  { timeout: 120000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-animations-presenter-'));
    let session, presenter;
    try {
      const deck = await compile(
        Presentation({
          children: Slide({
            children: [
              Text({ x: 1, y: 1, width: 4, height: 1, children: 'slow' }),
              Text({ x: 1, y: 3, width: 4, height: 1, children: 'late' }),
            ],
          }),
        }),
      );
      const [slow, late] = getSlideShapes(getSlides(deck)[0]);
      // Both start with the slide; the second arrives long before the first has
      // finished, which is the moment a counter-only mirror would miss.
      setShapeAnimation(slow, { effect: 'fadeIn', durationMs: 2000, start: 'withPrevious' });
      setShapeAnimation(late, {
        effect: 'fadeIn',
        durationMs: 100,
        delayMs: 300,
        start: 'withPrevious',
      });

      session = await openPreview(dir, deck, 1);
      const { page } = session;
      // The presenter button is in the header, which the show hides, so it is
      // pressed from the still preview — it enters the show itself.
      [presenter] = await Promise.all([
        page.waitForEvent('popup'),
        page.getByRole('button', { name: 'Presenter view', exact: true }).click(),
      ]);
      const audienceHidden = () =>
        page.evaluate(
          () =>
            Array.from(slideRoot().querySelectorAll('[data-pptx-shape-id]')).filter(
              (el) => el.style.visibility === 'hidden',
            ).length,
        );
      const presenterHidden = () =>
        presenter.evaluate(
          () =>
            Array.from(
              document
                .getElementById('current')
                .shadowRoot.querySelectorAll('[data-pptx-shape-id]'),
            ).filter((el) => el.style.visibility === 'hidden').length,
        );

      // The slide's own first effect starts without a click; the delayed one
      // has not arrived yet, on either screen.
      await presenter.waitForFunction(
        () =>
          document.getElementById('current').shadowRoot.querySelectorAll('[data-pptx-shape-id]')
            .length > 0,
      );
      assert.equal(await audienceHidden(), 1);
      assert.equal(await presenterHidden(), 1);
      assert.match(await presenter.locator('#count').textContent(), /Click 1 of 1/);

      // Mid-fade the two screens are at the same point in the same effect, not
      // one showing the finished picture while the other is still fading.
      const opacity = (target, where) =>
        target.evaluate(
          (root) =>
            Number(
              getComputedStyle(
                (root === 'presenter'
                  ? document.getElementById('current').shadowRoot
                  : slideRoot()
                ).querySelector('[data-pptx-shape-id]'),
              ).opacity,
            ),
          where,
        );
      const shownInPresenter = await opacity(presenter, 'presenter');
      const shownToAudience = await opacity(page, 'audience');
      assert.ok(
        shownInPresenter > 0 && shownInPresenter < 1,
        `presenter is mid-fade (${shownInPresenter})`,
      );
      assert.ok(
        Math.abs(shownInPresenter - shownToAudience) < 0.2,
        `fades agree (presenter ${shownInPresenter}, audience ${shownToAudience})`,
      );

      // A seek that lands on the cursor it is already on — here, settling the
      // stop in hand — reaches the presenter too, and both stop fading.
      await page.evaluate(() => animationPlayer.jumpTo(animationPlayer.cursor));
      // Well inside the two seconds the fade would otherwise still be running:
      // the presenter has to follow the seek, not wait for its own copy to end.
      await presenter.waitForFunction(
        () =>
          getComputedStyle(
            document.getElementById('current').shadowRoot.querySelector('[data-pptx-shape-id]'),
          ).opacity === '1',
        null,
        { timeout: 700 },
      );
      assert.equal(await opacity(page, 'audience'), 1);
      assert.equal(await presenterHidden(), 0);
      assert.equal(await audienceHidden(), 0);
      assert.equal(await page.evaluate(() => animationPlayer.running), false);

      // Back to a running build: the delayed effect arrives on the presenter's
      // own clock while the first one is still running — no message per effect,
      // and no wait for the slide to settle.
      await page.evaluate(() => animationPlayer.reset());
      await presenter.waitForFunction(() => {
        const shapes = document
          .getElementById('current')
          .shadowRoot.querySelectorAll('[data-pptx-shape-id]');
        return shapes.length > 0 && ![...shapes].some((el) => el.style.visibility === 'hidden');
      });
      assert.equal(await page.evaluate(() => animationPlayer.running), true);
      assert.equal(await audienceHidden(), 0);

      await idle(page);
      assert.equal(await presenterHidden(), 0);
      assert.deepEqual(session.errors, []);
    } finally {
      await presenter?.close();
      await session?.browser.close();
      await session?.preview.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

test(
  'a rebuild keeps the viewer’s place unless the slide’s animation changed',
  { timeout: 120000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-animations-rebuild-'));
    let session;
    try {
      // Three shapes throughout, so the drawing never changes and the only
      // difference between the builds is the timing tree and the notes.
      const build = async (notes, animated) => {
        const deck = await compile(
          Presentation({
            children: Slide({
              children: [1, 2, 3].map((n) =>
                Text({ x: 1, y: n * 1.5, width: 4, height: 1, children: 'shape ' + n }),
              ),
            }),
          }),
        );
        const slide = getSlides(deck)[0];
        setSlideNotes(slide, notes);
        for (const shape of getSlideShapes(slide).slice(0, animated)) {
          setShapeAnimation(shape, { effect: 'fadeIn', durationMs: 100 });
        }
        return savePresentation(deck);
      };
      const source = join(dir, 'source.pptx');
      await writeFile(source, await build('first', 2));
      await writeFile(join(dir, 'deck.tsx'), deckSource(dir));
      const preview = await startPreview(join(dir, 'deck.tsx'));
      const browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({
        viewport: { width: 1280, height: 800 },
        reducedMotion: 'no-preference',
      });
      const errors = [];
      page.on('pageerror', (error) => errors.push(error.message));
      session = { preview, browser, page, errors };
      await page.goto(preview.url);
      await page.getByRole('button', { name: 'Preview', exact: true }).click();
      await page.waitForFunction(() => state.slides.length === 1);
      await page.getByRole('button', { name: 'Present', exact: true }).click();
      await page.keyboard.press('ArrowRight');
      await idle(page);
      assert.deepEqual(await cursor(page), { cursor: 1, stops: 2 });
      const drawn = await page.evaluate(() => state.slides[0]);

      // A rebuild that changed nothing about this slide's animation — new
      // speaker notes — leaves the viewer where they were.
      await writeFile(source, await build('second', 2));
      await page.waitForFunction(() => state.notes?.[0] === 'second', null, { timeout: 20000 });
      assert.equal(await page.evaluate(() => state.slides[0]), drawn);
      assert.deepEqual(await cursor(page), { cursor: 1, stops: 2 });
      // The third shape has no animation of its own in this build.
      assert.deepEqual(await visibilities(page), ['visible', 'hidden', 'visible']);

      // A rebuild that changed the timing does not: the effects the cursor was
      // counting are not the effects there are now, so the build starts again.
      await writeFile(source, await build('second', 3));
      await page.waitForFunction(() => state.animations?.[0]?.length === 3, null, {
        timeout: 20000,
      });
      assert.equal(await page.evaluate(() => state.slides[0]), drawn);
      assert.deepEqual(await cursor(page), { cursor: 0, stops: 3 });
      assert.deepEqual(await visibilities(page), ['hidden', 'hidden', 'hidden']);

      assert.deepEqual(errors, []);
    } finally {
      await session?.browser.close();
      await session?.preview.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

test(
  'a build during a transition belongs to the arriving slide alone',
  { timeout: 120000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-animations-transition-'));
    let session;
    try {
      const deck = await compile(
        Presentation({
          children: [1, 2, 3].map((n) =>
            Slide({ children: Text({ x: 1, y: 1, width: 8, height: 2, children: 'slide ' + n }) }),
          ),
        }),
      );
      for (const slide of getSlides(deck)) {
        setShapeAnimation(getSlideShapes(slide)[0], { effect: 'fadeIn', durationMs: 100 });
        setSlideTransition(slide, { effect: 'fade', speed: 'slow' });
      }

      session = await openPreview(dir, deck, 3);
      const { page } = session;
      await page.getByRole('button', { name: 'Present', exact: true }).click();
      await page.keyboard.press('ArrowRight');
      await idle(page);

      // Both slides are on screen while the transition runs, and they carry
      // shape ids from two different id spaces.
      await page.keyboard.press('ArrowRight');
      assert.equal(await page.locator('.transition-layer').count(), 2);
      assert.deepEqual(
        await page.evaluate(() => ({
          both: canvas.querySelectorAll('[data-pptx-shape-id]').length,
          arriving: slideRoot().querySelectorAll('[data-pptx-shape-id]').length,
          leaving:
            canvas.querySelector('.transition-old [data-pptx-shape-id]')?.style.visibility ||
            'visible',
        })),
        { both: 2, arriving: 1, leaving: 'visible' },
      );
      assert.deepEqual(await visibilities(page), ['hidden']);

      // Clicking through the arriving slide's build while the transition is
      // still running touches that slide only.
      await page.keyboard.press('ArrowRight');
      assert.deepEqual(await cursor(page), { cursor: 1, stops: 1 });
      assert.deepEqual(await visibilities(page), ['visible']);
      assert.equal(
        await page.evaluate(
          () =>
            canvas.querySelector('.transition-old [data-pptx-shape-id]').style.visibility ||
            'visible',
        ),
        'visible',
      );

      // Moving on again mid-transition leaves no layer behind and one copy of
      // the slide, with its own build at the top.
      await page.keyboard.press('ArrowRight');
      await page.waitForFunction(() => index === 2 && !transitionCleanup, null, { timeout: 10000 });
      assert.equal(await page.locator('.transition-layer').count(), 0);
      assert.equal(
        await page.evaluate(() => canvas.querySelectorAll('[data-pptx-shape-id]').length),
        1,
      );
      assert.deepEqual(await cursor(page), { cursor: 0, stops: 1 });
      assert.deepEqual(await visibilities(page), ['hidden']);

      // Back through the transition: a slide arrived at backwards is the one
      // the viewer left, played out.
      await page.keyboard.press('ArrowLeft');
      await page.waitForFunction(() => index === 1 && !transitionCleanup, null, { timeout: 10000 });
      assert.deepEqual(await cursor(page), { cursor: 1, stops: 1 });
      assert.deepEqual(await visibilities(page), ['visible']);

      // A slide arrived at backwards has its build behind it, so the first step
      // back walks that build before the deck.
      await page.keyboard.press('ArrowLeft');
      assert.equal(await page.evaluate(() => index), 1);
      assert.deepEqual(await cursor(page), { cursor: 0, stops: 1 });
      assert.deepEqual(await visibilities(page), ['hidden']);

      // Revisiting it forwards starts it again.
      await page.keyboard.press('ArrowLeft');
      await page.waitForFunction(() => index === 0 && !transitionCleanup, null, { timeout: 10000 });
      await page.keyboard.press('ArrowRight');
      await page.waitForFunction(() => index === 1 && !transitionCleanup, null, { timeout: 10000 });
      assert.deepEqual(await cursor(page), { cursor: 0, stops: 1 });
      assert.deepEqual(await visibilities(page), ['hidden']);

      assert.deepEqual(session.errors, []);
    } finally {
      await session?.browser.close();
      await session?.preview.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

test(
  'reduced motion keeps the click order, and a slide that does not advance on click still builds',
  { timeout: 120000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-animations-reduced-'));
    let session;
    try {
      const deck = await compile(
        Presentation({
          children: [
            Slide({
              children: [
                Text({ x: 1, y: 1, width: 4, height: 1, children: 'a' }),
                Text({ x: 1, y: 3, width: 4, height: 1, children: 'b' }),
              ],
            }),
            Slide({ children: Text({ x: 1, y: 1, width: 8, height: 2, children: 'next' }) }),
          ],
        }),
      );
      const slides = getSlides(deck);
      for (const shape of getSlideShapes(slides[0])) {
        setShapeAnimation(shape, { effect: 'fadeIn', durationMs: 400 });
      }
      // The deck says a click does not move to the next slide.
      setSlideTransition(slides[0], { effect: 'fade', advanceOnClick: false });

      session = await openPreview(dir, deck, 2);
      const { page } = session;
      await page.getByRole('button', { name: 'Present', exact: true }).click();
      await page.emulateMedia({ reducedMotion: 'reduce' });
      assert.deepEqual(await visibilities(page), ['hidden', 'hidden']);

      // A click still plays the build even though the slide does not advance on
      // click — and with reduced motion it arrives at once, in the same order.
      await page.mouse.click(640, 400);
      assert.deepEqual(await cursor(page), { cursor: 1, stops: 2 });
      assert.deepEqual(await visibilities(page), ['visible', 'hidden']);
      assert.equal(
        await page.evaluate(() => slide.shadowRoot.getAnimations({ subtree: true }).length),
        0,
      );
      await page.mouse.click(640, 400);
      assert.deepEqual(await cursor(page), { cursor: 2, stops: 2 });
      assert.deepEqual(await visibilities(page), ['visible', 'visible']);

      // With the build played out, a click is what the deck says it is.
      await page.mouse.click(640, 400);
      await page.waitForTimeout(200);
      assert.equal(await page.evaluate(() => index), 0);
      // The keyboard is not a click, and moves on.
      await page.keyboard.press('ArrowRight');
      await page.waitForFunction(() => index === 1);
      assert.equal(await page.locator('.transition-layer').count(), 0, 'reduced motion, no layers');

      assert.deepEqual(session.errors, []);
    } finally {
      await session?.browser.close();
      await session?.preview.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

test('a slide holds its clicks while the player is on its way', { timeout: 120000 }, async () => {
  const dir = await mkdtemp(join(tmpdir(), 'office-animations-loading-'));
  let session;
  try {
    const deck = await compile(
      Presentation({
        children: [
          Slide({
            children: [
              Text({ x: 1, y: 1, width: 4, height: 1, children: 'a' }),
              Text({ x: 1, y: 3, width: 4, height: 1, children: 'b' }),
            ],
          }),
          Slide({ children: Text({ x: 1, y: 1, width: 8, height: 2, children: 'next' }) }),
        ],
      }),
    );
    const slides = getSlides(deck);
    for (const shape of getSlideShapes(slides[0])) {
      setShapeAnimation(shape, { effect: 'fadeIn', durationMs: 100 });
    }
    // Short enough that a slide which ignored its unplayed effects would have
    // moved on well before the player arrives.
    setSlideTransition(slides[0], { effect: 'none', advanceAfterMs: 100 });

    let held = true;
    session = await openPreview(dir, deck, 2, (page) =>
      page.route('**/animation-player.js', async (route) => {
        while (held) await new Promise((resolve) => setTimeout(resolve, 20));
        await route.continue();
      }),
    );
    const { page } = session;
    await page.getByRole('button', { name: 'Present', exact: true }).click();

    // Clicks, keys and the button all reach a slide whose effects cannot run
    // yet. None of them may spend the click on the deck.
    await page.keyboard.press('ArrowRight');
    await page.keyboard.press('Space');
    await page.mouse.click(640, 400);
    await page.locator('#present-next').click();
    await page.waitForTimeout(400);
    assert.equal(await page.evaluate(() => index), 0, 'the deck waited for the player');
    assert.equal(await cursor(page), null);
    assert.match(await page.locator('#present-note').textContent(), /Preparing animations/);

    // Once it arrives the build is where it should be: at the top, with
    // nothing played and nothing skipped.
    held = false;
    await page.waitForFunction(() => animationPlayer !== undefined, null, { timeout: 15000 });
    assert.deepEqual(await cursor(page), { cursor: 0, stops: 2 });
    assert.deepEqual(await visibilities(page), ['hidden', 'hidden']);
    assert.equal(await page.locator('#present-note').isVisible(), false);
    await page.keyboard.press('ArrowRight');
    assert.deepEqual(await cursor(page), { cursor: 1, stops: 2 });
  } finally {
    await session?.browser.close();
    await session?.preview.close();
    await rm(dir, { recursive: true, force: true });
  }
});

// A deck of its own, with no automatic advance: the slide above moves on by
// itself after 100ms once its effects are known to be unplayable, which is
// what "not held hostage" means — and it would walk out from under every
// assertion about what this slide is showing.
test(
  'a slide says so when the player never comes, and can ask again',
  { timeout: 120000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-animations-failed-'));
    let session;
    try {
      const deck = await compile(
        Presentation({
          children: [
            Slide({
              children: [
                Text({ x: 1, y: 1, width: 4, height: 1, children: 'a' }),
                Text({ x: 1, y: 3, width: 4, height: 1, children: 'b' }),
              ],
            }),
            Slide({ children: Text({ x: 1, y: 1, width: 8, height: 2, children: 'next' }) }),
          ],
        }),
      );
      const slides = getSlides(deck);
      for (const shape of getSlideShapes(slides[0])) {
        setShapeAnimation(shape, { effect: 'fadeIn', durationMs: 100 });
      }

      let refused = true;
      session = await openPreview(dir, deck, 2, (page) =>
        page.route('**/animation-player.js', async (route) => {
          if (refused) await route.abort();
          else await route.continue();
        }),
      );
      const { page } = session;
      await page.getByRole('button', { name: 'Present', exact: true }).click();

      // A player that cannot be fetched is said out loud, in the viewer's
      // language, with a way to ask for it again.
      await page.waitForFunction(() => animationLoad === 'failed', null, { timeout: 15000 });
      assert.match(
        await page.locator('#present-note').textContent(),
        /animations could not be loaded/i,
      );
      const retry = page.locator('#animation-retry');
      assert.equal(await retry.isVisible(), true);
      // The deck is not held hostage by the failure, and the offer to try
      // again is still there on the way back.
      await page.keyboard.press('ArrowRight');
      await page.waitForFunction(() => index === 1);
      assert.equal(await retry.isVisible(), false);
      await page.keyboard.press('ArrowLeft');
      await page.waitForFunction(() => index === 0);
      assert.equal(await retry.isVisible(), true);
      refused = false;
      await retry.click();
      await page.waitForFunction(() => animationPlayer !== undefined, null, { timeout: 15000 });
      // The slide was come back to from the one after it, so it picks up played
      // out — the position it was waiting to be put in.
      assert.deepEqual(await cursor(page), { cursor: 2, stops: 2 });
      assert.deepEqual(await visibilities(page), ['visible', 'visible']);
      assert.equal(await retry.isVisible(), false);
      assert.equal(await page.locator('#present-note').isVisible(), false);
    } finally {
      await session?.browser.close();
      await session?.preview.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

/**
 * What the player does with a timing tree it cannot fully read. The steps are
 * written by hand because the authoring API refuses to produce them — that is
 * the point: they are what a deck from PowerPoint or Google Slides brings.
 */
test('the player runs only what the deck states', { timeout: 120000 }, async () => {
  const dir = await mkdtemp(join(tmpdir(), 'office-animations-rules-'));
  let session;
  try {
    const deck = await compile(
      Presentation({
        children: Slide({ children: Text({ x: 1, y: 1, width: 8, height: 2, children: 'deck' }) }),
      }),
    );
    session = await openPreview(dir, deck, 1);
    const { page } = session;

    // Through the served bundle, which is the file the CLI ships and the
    // editor panel imports — not a copy inlined into the page.
    const plan = (steps) =>
      page.evaluate(async (given) => {
        const { buildAnimationStops } = await import('/animation-player.js');
        const parsed = buildAnimationStops(given);
        return {
          stops: parsed.stops.map((stop) => ({
            auto: stop.auto,
            items: stop.items.map((item) => [item.step.id, item.begin, item.duration]),
          })),
          unsupported: parsed.unsupported.map((entry) => [entry.step.id, entry.reason]),
        };
      }, steps);

    const step = (id, overrides) => ({
      id,
      target: { kind: 'shape', shapeId: 10 },
      targetShapeIds: [10],
      effect: 'fadeIn',
      presetId: 10,
      presetClass: 'entr',
      start: 'click',
      durationMs: 500,
      delayMs: 0,
      valueAfterEnd: 'held',
      buildByParagraph: false,
      buildLevel: null,
      sequence: 'mainSeq',
      playable: true,
      editable: true,
      ...overrides,
    });

    // A step outside the main sequence runs off its own trigger.
    assert.deepEqual(await plan([step(1, { sequence: 'interactiveSeq' }), step(2)]), {
      stops: [{ auto: false, items: [[2, 0, 500]] }],
      unsupported: [[1, 'notInMainSeq']],
    });

    // A leading 'with previous' starts as the slide appears; a leading click
    // does not.
    assert.deepEqual((await plan([step(1, { start: 'withPrevious' })])).stops[0].auto, true);
    assert.deepEqual((await plan([step(1)])).stops[0].auto, false);

    // An effect whose length the tree does not state cannot have anything
    // chained onto its end — and a guessed moment is not an answer. The click
    // group after it is stated outright, so it still plays.
    assert.deepEqual(
      await plan([
        step(1, { durationMs: null, effect: null, playable: false, presetClass: 'entr' }),
        step(2, { start: 'afterPrevious' }),
        step(3, { start: 'withPrevious' }),
        step(4),
        step(5, { start: 'afterPrevious', delayMs: 100 }),
      ]),
      {
        stops: [
          { auto: false, items: [] },
          {
            auto: false,
            items: [
              [4, 0, 500],
              [5, 600, 500],
            ],
          },
        ],
        unsupported: [
          [1, 'notModelled'],
          [2, 'unknownTiming'],
          [3, 'unknownTiming'],
        ],
      },
    );

    // A start condition we do not model says nothing about whether a click
    // stop belongs here, so none is invented for it.
    assert.deepEqual(
      await plan([step(1), step(2, { start: 'unknown' }), step(3, { start: 'withPrevious' })]),
      {
        stops: [{ auto: false, items: [[1, 0, 500]] }],
        unsupported: [
          [2, 'notModelled'],
          [3, 'unknownTiming'],
        ],
      },
    );

    // A delay the tree does not state is not zero.
    assert.deepEqual(await plan([step(1, { delayMs: null })]), {
      stops: [{ auto: false, items: [] }],
      unsupported: [[1, 'notModelled']],
    });

    // 'appear' is instantaneous by definition, so what follows it is placeable.
    assert.deepEqual(
      await plan([
        step(1, { effect: 'appear', durationMs: null }),
        step(2, { start: 'afterPrevious' }),
      ]),
      {
        stops: [
          {
            auto: false,
            items: [
              [1, 0, 0],
              [2, 0, 500],
            ],
          },
        ],
        unsupported: [],
      },
    );

    // Against a real DOM: what the player leaves alone, and which paragraphs
    // one target means.
    const outcome = await page.evaluate(
      async (steps) => {
        const { createAnimationPlayer } = await import('/animation-player.js');
        const root = document.createElement('div');
        root.innerHTML =
          '<div data-pptx-shape-id="10"><span data-pptx-paragraph="0">a</span>' +
          '<span data-pptx-paragraph="1">b</span>' +
          '<div data-pptx-shape-id="11"><span data-pptx-paragraph="0">nested</span></div></div>' +
          '<div data-pptx-shape-id="20">twice</div><div data-pptx-shape-id="20">twice</div>' +
          '<div data-pptx-shape-id="30">composite</div>';
        document.body.append(root);
        const player = createAnimationPlayer({ root, steps, reducedMotion: () => true });
        player.reset();
        const read = (selector) =>
          Array.from(root.querySelectorAll(selector)).map((el) => el.style.visibility || 'visible');
        const before = {
          shapes: read('[data-pptx-shape-id]'),
          paragraphs: read('[data-pptx-paragraph]'),
        };
        player.advance();
        player.advance();
        player.advance();
        const after = {
          shapes: read('[data-pptx-shape-id]'),
          paragraphs: read('[data-pptx-paragraph]'),
        };
        const unsupported = player.unsupported.map((entry) => [
          entry.reason,
          entry.step.targetShapeIds,
        ]);
        player.dispose();
        root.remove();
        return { before, after, unsupported, stops: player.stopCount };
      },
      [
        // A paragraph range wider than the text, on a shape holding another shape.
        step(1, {
          target: { kind: 'paragraphs', shapeId: 10, firstParagraph: 0, lastParagraph: 2147483647 },
          durationMs: 0,
        }),
        // Two shapes are drawn with this id; which one the tree means is unstated.
        step(2, { target: { kind: 'shape', shapeId: 20 }, targetShapeIds: [20], durationMs: 0 }),
        // A composite effect drives 30 and 11 together — we model neither.
        step(3, {
          target: { kind: 'unsupported', shapeId: null },
          targetShapeIds: [30, 11],
          effect: null,
          playable: false,
          durationMs: 0,
        }),
      ],
    );

    // Before the first click: only the paragraphs of shape 10 are held back.
    // Shape 11 sits inside it but is a target of its own, and everything an
    // unplayable step names stays exactly as it was drawn.
    assert.deepEqual(outcome.before.shapes, [
      'visible',
      'visible',
      'visible',
      'visible',
      'visible',
    ]);
    assert.deepEqual(outcome.before.paragraphs, ['hidden', 'hidden', 'visible']);
    assert.deepEqual(outcome.after.paragraphs, ['visible', 'visible', 'visible']);
    assert.deepEqual(outcome.after.shapes, ['visible', 'visible', 'visible', 'visible', 'visible']);
    assert.deepEqual(outcome.unsupported, [
      ['notModelled', [30, 11]],
      ['ambiguousTarget', [20]],
    ]);
    // Every step keeps its place in the click order even when it plays nothing.
    assert.equal(outcome.stops, 3);

    // Two effects of one stop on one object: the deck lists the fade-out last,
    // but the fade-in is the one that ends last. Playing it through, jumping to
    // the same point and resuming into it must all agree — with the deck's
    // order, not with whichever animation happened to finish last.
    const overlap = await page.evaluate(
      async (steps) => {
        const { createAnimationPlayer } = await import('/animation-player.js');
        const root = document.createElement('div');
        root.innerHTML = '<div data-pptx-shape-id="10">both</div>';
        document.body.append(root);
        const target = root.querySelector('[data-pptx-shape-id="10"]');
        const read = () => target.style.visibility || 'visible';
        // Waits on the state itself, never on a stopwatch: a loaded machine
        // must not be able to turn this into a different test.
        const until = async (check) => {
          for (let i = 0; i < 600; i++) {
            if (check()) return true;
            await new Promise((resolve) => setTimeout(resolve, 20));
          }
          return false;
        };
        // An effect that has ended stays in the sandwich, so the count of
        // animations says nothing about where they are: each is asked.
        const at = (ms) => target.getAnimations().find((a) => a.effect.getTiming().duration === ms);
        const player = createAnimationPlayer({ root, steps, reducedMotion: () => false });
        player.reset();
        player.advance();
        // The short fade-out has ended; the long fade-in is still running.
        const midway = await until(
          () => at(50)?.playState === 'finished' && at(3000)?.playState === 'running',
        );
        const early = { visibility: read(), opacity: getComputedStyle(target).opacity };
        const settled = await until(() => !player.running);
        const playedThrough = read();
        player.jumpTo(0);
        player.jumpTo(1);
        const jumped = read();
        player.resume(1, 10000);
        const resumed = read();
        player.dispose();
        root.remove();
        return { midway, early, settled, playedThrough, jumped, resumed };
      },
      [
        step(1, { effect: 'fadeIn', durationMs: 3000 }),
        step(2, {
          effect: 'fadeOut',
          presetClass: 'exit',
          presetId: 10,
          durationMs: 50,
          start: 'withPrevious',
        }),
      ],
    );
    // The fade-out has ended and the fade-in has not. Both began together, so
    // the deck's order decides between them and the fade-out is on top — and a
    // finished effect stays there, frozen on the value it ended on, so the
    // object is out of sight and the fade-in under it does not show through.
    assert.equal(overlap.midway, true, 'the fade-out ended while the fade-in ran');
    assert.equal(overlap.early.visibility, 'visible');
    assert.ok(
      Number(overlap.early.opacity) < 0.05,
      `held out of sight by the fade-out (${overlap.early.opacity})`,
    );
    // Once both have run, the deck's order decides: the fade-out is last.
    assert.equal(overlap.settled, true);
    assert.equal(overlap.playedThrough, 'hidden');
    assert.equal(overlap.jumped, 'hidden');
    assert.equal(overlap.resumed, 'hidden');

    // The same two effects, but the exit runs first because the entrance waits
    // half a second. SMIL composites by the order animations *begin* — "the
    // animation first begun has lowest priority and the most recently begun
    // animation has highest priority" (SMIL 3.0 §12.4.3), which PresentationML
    // timing is built on (ECMA-376 Part 1 §19.5) — so the entrance, written
    // first but begun last, is the one that decides. Reading the tree's order
    // instead would leave the object off the slide it just returned to.
    const delayed = await page.evaluate(
      async (steps) => {
        const { createAnimationPlayer } = await import('/animation-player.js');
        const root = document.createElement('div');
        root.innerHTML = '<div data-pptx-shape-id="10">both</div>';
        document.body.append(root);
        const target = root.querySelector('[data-pptx-shape-id="10"]');
        const read = () => target.style.visibility || 'visible';
        const until = async (check) => {
          for (let i = 0; i < 600; i++) {
            if (check()) return true;
            await new Promise((resolve) => setTimeout(resolve, 20));
          }
          return false;
        };
        const player = createAnimationPlayer({ root, steps, reducedMotion: () => false });
        player.reset();
        const start = read();
        player.advance();
        // Between the two: the exit has run and the entrance has not begun.
        const gone = await until(() => read() === 'hidden' && !target.getAnimations().length);
        const settled = await until(() => !player.running);
        const playedThrough = read();
        player.jumpTo(0);
        player.jumpTo(1);
        const jumped = read();
        player.resume(1, 10000);
        const resumed = read();
        player.dispose();
        root.remove();
        return { start, gone, settled, playedThrough, jumped, resumed };
      },
      [
        step(1, { effect: 'fadeIn', durationMs: 50, delayMs: 500 }),
        step(2, {
          effect: 'fadeOut',
          presetClass: 'exit',
          presetId: 10,
          durationMs: 50,
          start: 'withPrevious',
        }),
      ],
    );
    // The first effect to begin is the exit, so the object starts on the slide.
    assert.equal(delayed.start, 'visible');
    assert.equal(delayed.gone, true, 'the exit ran before the entrance began');
    assert.equal(delayed.settled, true);
    assert.equal(delayed.playedThrough, 'visible');
    assert.equal(delayed.jumped, 'visible');
    assert.equal(delayed.resumed, 'visible');

    // Part-way through, not only at the end — and with the tree's order and the
    // order the two begin in pulling opposite ways. The deck lists the exit
    // first, delayed by 900ms; the entrance it lists second begins at once and
    // runs for four seconds. The exit therefore begins last and outranks the
    // entrance, and a finished effect stays in the sandwich frozen on its last
    // value, so once the exit is over the object must stay out of sight instead
    // of fading back in under it. Joining the same span part-way through has to
    // arrive at the same sandwich, which means building it in the order the
    // effects begin rather than the order the tree lists them.
    const sandwich = await page.evaluate(
      async (steps) => {
        const { createAnimationPlayer } = await import('/animation-player.js');
        const root = document.createElement('div');
        root.innerHTML = '<div data-pptx-shape-id="10">both</div>';
        document.body.append(root);
        const target = root.querySelector('[data-pptx-shape-id="10"]');
        const until = async (check) => {
          for (let i = 0; i < 600; i++) {
            if (check()) return true;
            await new Promise((resolve) => setTimeout(resolve, 20));
          }
          return false;
        };
        // Each effect is asked where it is: a finished one is still listed, so
        // counting them would say nothing about whether the exit has run.
        const at = (ms) => target.getAnimations().find((a) => a.effect.getTiming().duration === ms);
        const state = () => ({
          visibility: target.style.visibility || 'visible',
          opacity: Number(getComputedStyle(target).opacity),
          effects: target.getAnimations().map((a) => [a.effect.getTiming().duration, a.playState]),
        });
        const player = createAnimationPlayer({ root, steps, reducedMotion: () => false });
        player.reset();
        player.advance();
        // The 100ms exit has ended while the 4000ms entrance is still running.
        const overlapped = await until(
          () => at(100)?.playState === 'finished' && at(4000)?.playState === 'running',
        );
        const held = state();
        // None of the entrance's remaining seconds brings it back. Sampled only
        // while it is still running, so a slow machine ends the watch rather
        // than reading a settled slide as if it were a mid-state.
        let peak = 0;
        for (let i = 0; i < 15; i++) {
          await new Promise((resolve) => setTimeout(resolve, 20));
          if (at(4000)?.playState !== 'running') break;
          peak = Math.max(peak, Number(getComputedStyle(target).opacity));
        }
        const settled = await until(() => !player.running);
        const ended = state();

        // Joined 950ms in, with both on their way: the entrance began at 0 and
        // the exit at 900, so the exit is on top and the object is half-way
        // out, not a quarter of the way in.
        player.resume(1, 950);
        const joinedBoth = state();
        // Joined 1200ms in: the exit is over by then and holds the object out
        // of sight, though the entrance under it is still running.
        player.resume(1, 1200);
        const joinedAfter = state();
        player.dispose();
        root.remove();
        return { overlapped, held, peak, settled, ended, joinedBoth, joinedAfter };
      },
      [
        step(1, {
          effect: 'fadeOut',
          presetClass: 'exit',
          presetId: 10,
          durationMs: 100,
          delayMs: 900,
        }),
        step(2, { effect: 'fadeIn', durationMs: 4000, start: 'withPrevious' }),
      ],
    );
    assert.equal(sandwich.overlapped, true, 'the exit ended while the entrance ran');
    assert.ok(sandwich.held.opacity < 0.05, `held out of sight (${sandwich.held.opacity})`);
    assert.ok(sandwich.peak < 0.05, `never shows through while it runs (${sandwich.peak})`);
    assert.equal(sandwich.settled, true);
    assert.equal(sandwich.ended.visibility, 'hidden');
    // Half-way through a 100ms fade-out that started 50ms ago, over a fade-in
    // barely a quarter of the way through: the exit's value is the one shown.
    assert.deepEqual(sandwich.joinedBoth.effects, [
      [4000, 'running'],
      [100, 'running'],
    ]);
    assert.ok(
      sandwich.joinedBoth.opacity > 0.35 && sandwich.joinedBoth.opacity < 0.65,
      `on its way out, not in (${sandwich.joinedBoth.opacity})`,
    );
    // The exit is over, so it is held at its end rather than dropped: it is
    // still above the entrance, which is 1200ms into its four seconds.
    assert.deepEqual(sandwich.joinedAfter.effects, [
      [4000, 'running'],
      [0, 'finished'],
    ]);
    assert.ok(
      sandwich.joinedAfter.opacity < 0.05,
      `joined out of sight (${sandwich.joinedAfter.opacity})`,
    );

    // What the tree says becomes of an effect's value when it ends decides what
    // the slide shows while a second effect over the same object is still
    // running. Both effects here are entrances, so where the object ends up is
    // the same either way and only the moment between them is in question: held,
    // the short one keeps the object fully on; taken away, it gives it back to
    // the long fade underneath, a quarter of the way in. Joining that same
    // moment has to show what playing through to it shows.
    const afterEnd = async (valueAfterEnd) =>
      page.evaluate(
        async (steps) => {
          const { createAnimationPlayer } = await import('/animation-player.js');
          const root = document.createElement('div');
          root.innerHTML = '<div data-pptx-shape-id="10">both</div>';
          document.body.append(root);
          const target = root.querySelector('[data-pptx-shape-id="10"]');
          const until = async (check) => {
            for (let i = 0; i < 600; i++) {
              if (check()) return true;
              await new Promise((resolve) => setTimeout(resolve, 20));
            }
            return false;
          };
          const at = (ms) =>
            target.getAnimations().find((a) => a.effect.getTiming().duration === ms);
          const state = () => ({
            visibility: target.style.visibility || 'visible',
            opacity: Number(getComputedStyle(target).opacity),
          });
          const player = createAnimationPlayer({ root, steps, reducedMotion: () => false });
          player.reset();
          player.advance();
          // The short effect has actually begun. A held one stays listed when it
          // ends and a dropped one does not, so waiting only for its absence
          // would be satisfied before it ever ran.
          const began = await until(() => at(100) !== undefined);
          // …and is over, and the player has done with it: held, it is finished
          // and still in the sandwich; taken away, it has left. Waiting only for
          // 'finished' would read a dropped one in the moment between the two.
          const overlapped = await until(() => {
            const long = at(4000);
            const short = at(100);
            if (long?.playState !== 'running' || Number(long.currentTime) < 1000) return false;
            return steps[1].valueAfterEnd === 'removed'
              ? short === undefined
              : short?.playState === 'finished';
          });
          const midway = state();
          // The same moment, joined rather than played through.
          player.resume(1, 1100);
          const resumed = state();
          const reported = player.unsupported.map((entry) => entry.reason);
          player.dispose();
          root.remove();
          return { began, overlapped, midway, resumed, reported };
        },
        [
          step(1, { effect: 'fadeIn', durationMs: 4000 }),
          step(2, {
            effect: 'fadeIn',
            durationMs: 100,
            delayMs: 900,
            start: 'withPrevious',
            valueAfterEnd,
          }),
        ],
      );

    const held = await afterEnd('held');
    assert.equal(held.began, true);
    assert.equal(held.overlapped, true);
    assert.deepEqual(held.reported, []);
    assert.ok(held.midway.opacity > 0.9, `held fully on (${held.midway.opacity})`);
    assert.ok(held.resumed.opacity > 0.9, `and the same when joined (${held.resumed.opacity})`);

    const dropped = await afterEnd('removed');
    assert.equal(dropped.began, true);
    assert.equal(dropped.overlapped, true);
    assert.deepEqual(dropped.reported, []);
    // A quarter of the way through the long fade, not at either end of it.
    assert.ok(
      dropped.midway.opacity > 0.15 && dropped.midway.opacity < 0.6,
      `given back to the fade underneath (${dropped.midway.opacity})`,
    );
    assert.ok(
      dropped.resumed.opacity > 0.15 && dropped.resumed.opacity < 0.6,
      `and the same when joined (${dropped.resumed.opacity})`,
    );

    // Unstated: the short effect is reported and everything it names is left as
    // the renderer drew it, rather than played on a guess. The click it belongs
    // to is still there.
    const silent = await page.evaluate(
      async (steps) => {
        const { createAnimationPlayer } = await import('/animation-player.js');
        const root = document.createElement('div');
        root.innerHTML =
          '<div data-pptx-shape-id="10">both</div><div data-pptx-shape-id="20">alone</div>';
        document.body.append(root);
        const read = (id) =>
          root.querySelector(`[data-pptx-shape-id="${id}"]`).style.visibility || 'visible';
        const player = createAnimationPlayer({ root, steps, reducedMotion: () => true });
        player.reset();
        const before = [read(10), read(20)];
        player.advance();
        const after = [read(10), read(20)];
        const reported = player.unsupported.map((entry) => [
          entry.reason,
          entry.step.targetShapeIds,
        ]);
        const stops = player.stopCount;
        player.dispose();
        root.remove();
        return { before, after, reported, stops };
      },
      [
        step(1, { effect: 'fadeIn', durationMs: 4000 }),
        step(2, {
          effect: 'fadeOut',
          presetClass: 'exit',
          presetId: 10,
          durationMs: 100,
          delayMs: 900,
          start: 'withPrevious',
          valueAfterEnd: 'unstated',
        }),
        // On its own, an unstated value is no question: nothing is underneath
        // it, so it plays.
        step(3, {
          target: { kind: 'shape', shapeId: 20 },
          targetShapeIds: [20],
          effect: 'fadeIn',
          durationMs: 200,
          start: 'withPrevious',
          valueAfterEnd: 'unstated',
        }),
      ],
    );
    assert.deepEqual(silent.reported, [['notModelled', [10]]]);
    // Which elements a step reaches is a fact about the drawing, so the same
    // reading is made again when the slide is drawn again — a player built
    // against a layer that is still empty must not let the pair through once
    // the shapes arrive.
    const late = await page.evaluate(
      async (steps) => {
        const { createAnimationPlayer } = await import('/animation-player.js');
        const root = document.createElement('div');
        document.body.append(root);
        const player = createAnimationPlayer({
          root: () => root,
          steps,
          reducedMotion: () => true,
        });
        const atFirst = player.unsupported.map((entry) => entry.reason);
        root.innerHTML = '<div data-pptx-shape-id="10">both</div>';
        player.reset();
        player.advance();
        const read = root.querySelector('[data-pptx-shape-id="10"]').style.visibility || 'visible';
        const reported = player.unsupported.map((entry) => entry.reason);
        player.dispose();
        root.remove();
        return { atFirst, reported, read };
      },
      [
        step(1, { effect: 'fadeIn', durationMs: 4000 }),
        step(2, {
          effect: 'fadeOut',
          presetClass: 'exit',
          presetId: 10,
          durationMs: 100,
          delayMs: 900,
          start: 'withPrevious',
          valueAfterEnd: 'unstated',
        }),
      ],
    );
    assert.deepEqual(late.atFirst, [], 'nothing was drawn, so nothing met anything');
    assert.deepEqual(late.reported, ['notModelled']);
    assert.equal(late.read, 'visible', 'and the shape it names is left as drawn');

    // Shape 10 is left exactly as drawn, at both ends; shape 20 builds.
    assert.deepEqual(silent.before, ['visible', 'hidden']);
    assert.deepEqual(silent.after, ['visible', 'visible']);
    assert.equal(silent.stops, 1);

    assert.deepEqual(session.errors, []);
  } finally {
    await session?.browser.close();
    await session?.preview.close();
    await rm(dir, { recursive: true, force: true });
  }
});
