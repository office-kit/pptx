import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  getSlideShapes,
  getSlides,
  savePresentation,
  setCustomShows,
  setShapeClickAction,
  setSlideShowProperties,
  getSlideShowProperties,
} from '@office-kit/pptx';
import { compile, Presentation, Slide, Text } from '@office-kit/pptx-dsl';
import { startPreview } from '../helpers/server.mjs';

async function createPreview() {
  const dir = await mkdtemp(join(tmpdir(), 'office-show-action-order-'));
  const deck = await compile(
    Presentation({
      children: ['Slide Alpha', 'Slide Bravo', 'Slide Charlie'].map((name) =>
        Slide({
          children: [
            Text({ x: 1, y: 1, width: 5, height: 1, children: name }),
            Text({ x: 1, y: 3, width: 2, height: 0.6, children: 'Next' }),
            Text({ x: 3, y: 3, width: 2, height: 0.6, children: 'Previous' }),
            Text({ x: 1, y: 4, width: 2, height: 0.6, children: 'First' }),
            Text({ x: 3, y: 4, width: 2, height: 0.6, children: 'Last' }),
          ],
        }),
      ),
    }),
  );
  const slides = getSlides(deck);
  for (const slide of slides) {
    const shapes = getSlideShapes(slide);
    setShapeClickAction(shapes[1], { kind: 'nextSlide' });
    setShapeClickAction(shapes[2], { kind: 'prevSlide' });
    setShapeClickAction(shapes[3], { kind: 'firstSlide' });
    setShapeClickAction(shapes[4], { kind: 'lastSlide' });
  }
  setCustomShows(deck, [
    { id: 7, name: 'Odd order', slides: [slides[2], slides[0], slides[2], slides[1]] },
  ]);
  setSlideShowProperties(deck, {
    ...getSlideShowProperties(deck),
    slides: { kind: 'customShow', id: 7 },
  });
  const source = join(dir, 'source.pptx');
  await writeFile(source, await savePresentation(deck));
  const file = join(dir, 'deck.tsx');
  await writeFile(
    file,
    `import {readFile} from 'node:fs/promises';import {Presentation} from '@office-kit/pptx-dsl';export default <Presentation source={await readFile(${JSON.stringify(source)})} />;`,
  );
  return { dir, preview: await startPreview(file) };
}

async function assertCustomOrder(page, root) {
  const expectSlide = async (name) => {
    await page.waitForFunction(
      (name) => document.querySelector('#slide')?.shadowRoot?.textContent?.includes(name),
      name,
    );
  };
  const link = (kind) => root.locator(`#slide a[href="#pptx-${kind}-slide"]`);
  assert.equal(await link('next').count(), 1);
  assert.equal(await link('prev').count(), 1);
  assert.equal(await link('first').count(), 1);
  assert.equal(await link('last').count(), 1);
  await expectSlide('Slide Charlie');
  await link('next').click();
  await expectSlide('Slide Alpha');
  await link('next').click();
  await expectSlide('Slide Charlie');
  await link('next').click();
  await expectSlide('Slide Bravo');
  await link('prev').click();
  await expectSlide('Slide Charlie');
  await link('prev').click();
  await expectSlide('Slide Alpha');
  await link('first').click();
  await expectSlide('Slide Charlie');
  await link('last').click();
  await expectSlide('Slide Bravo');
}

test(
  'preset show-action links follow duplicate custom show order during normal playback',
  { timeout: 60000 },
  async () => {
    const { dir, preview } = await createPreview();
    let browser;
    try {
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1400, height: 900 } });
      await page.goto(preview.url);
      await page.getByRole('button', { name: 'Preview', exact: true }).click();
      await page.getByRole('button', { name: 'Present', exact: true }).click();
      await assertCustomOrder(page, page);
      await page.locator('#slide a[href="#pptx-next-slide"]').click();
      await page.waitForFunction(
        () => !document.body.classList.contains('presenting') && !document.fullscreenElement,
      );
    } finally {
      await browser?.close();
      await preview.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

test(
  'preset show-action links follow duplicate custom show order in presenter view',
  { timeout: 60000 },
  async () => {
    const { dir, preview } = await createPreview();
    let browser;
    try {
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1400, height: 900 } });
      await page.goto(preview.url);
      await page.getByRole('button', { name: 'Preview', exact: true }).click();
      await page.getByRole('button', { name: 'Present', exact: true }).click();
      await page.waitForFunction(() => document.body.classList.contains('presenting'));
      await page.keyboard.press('Escape');
      await page.waitForFunction(
        () => !document.body.classList.contains('presenting') && !document.fullscreenElement,
      );
      const popupPromise = page.waitForEvent('popup');
      await page.getByRole('button', { name: 'Presenter view', exact: true }).click();
      const presenter = await popupPromise;
      const current = presenter.locator('#current');
      await presenter.waitForFunction(() => document.querySelector('#current'));
      const expectSlide = async (name) => {
        await page.waitForFunction(
          (name) => document.querySelector('#slide')?.shadowRoot?.textContent?.includes(name),
          name,
        );
      };
      const link = (kind) => current.locator(`a[href="#pptx-${kind}-slide"]`);
      assert.equal(await link('next').count(), 1);
      assert.equal(await link('prev').count(), 1);
      assert.equal(await link('first').count(), 1);
      assert.equal(await link('last').count(), 1);
      await expectSlide('Slide Charlie');
      assert.equal(
        await page.evaluate(() => presenting),
        true,
        'presenter keeps presentation active',
      );
      await link('next').click();
      await expectSlide('Slide Alpha');
      await link('next').click();
      await expectSlide('Slide Charlie');
      assert.equal(
        await page.evaluate(() => presenting),
        true,
        'presenter keeps presentation active',
      );
      await link('next').click();
      await expectSlide('Slide Bravo');
      await link('prev').click();
      await expectSlide('Slide Charlie');
      await link('prev').click();
      await expectSlide('Slide Alpha');
      await link('first').click();
      await expectSlide('Slide Charlie');
      await link('last').click();
      await expectSlide('Slide Bravo');
      await presenter.keyboard.press('Escape');
    } finally {
      await browser?.close();
      await preview.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

test(
  'a fullscreen grant arriving after presenter view starts does not end presenter view',
  { timeout: 60000 },
  async () => {
    const { dir, preview } = await createPreview();
    let browser;
    try {
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1400, height: 900 } });
      await page.goto(preview.url);
      await page.getByRole('button', { name: 'Preview', exact: true }).click();
      const popupPromise = page.waitForEvent('popup');
      // Present, Escape and Presenter view all run in one task, so the
      // fullscreen request from Present is still pending when presenter view
      // starts its own show.
      // Chromium may deliver the grant and its release in one frame, both
      // reporting the released state, so only the count and `presenting` are
      // recorded per event.
      await page.evaluate(() => {
        window.fullscreenTrace = [];
        document.addEventListener('fullscreenchange', () =>
          window.fullscreenTrace.push({ presenting }),
        );
        document.getElementById('present').click();
        document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
        document.getElementById('presenter').click();
      });
      await popupPromise;
      await page.waitForFunction(() => window.fullscreenTrace.length === 2);
      assert.deepEqual(await page.evaluate(() => window.fullscreenTrace), [
        { presenting: true },
        { presenting: true },
      ]);
      assert.equal(await page.evaluate(() => document.fullscreenElement), null);
      assert.equal(await page.evaluate(() => presenting), true);
    } finally {
      await browser?.close();
      await preview.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
