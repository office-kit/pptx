import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import { compile, Presentation, Slide, Text } from '@office-kit/pptx-dsl';
import { getSlideShapes, getSlides, savePresentation, setShapeAnimation } from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

async function makeDeck() {
  const dir = await mkdtemp(join(tmpdir(), 'office-show-playback-'));
  const file = join(dir, 'deck.tsx');
  await writeFile(
    file,
    `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation>{['A','B','C'].map(text => <Slide><Text x={1} y={1} width={7} height={1}>{text}</Text></Slide>)}</Presentation>`,
  );
  return { dir, file };
}

async function makeAnimatedDeck() {
  const dir = await mkdtemp(join(tmpdir(), 'office-show-animation-'));
  const deck = await compile(
    Presentation({
      children: Slide({ children: Text({ x: 1, y: 1, width: 7, height: 1, children: 'A' }) }),
    }),
  );
  setShapeAnimation(getSlideShapes(getSlides(deck)[0])[0], { effect: 'fadeIn', durationMs: 40 });
  await writeFile(join(dir, 'source.pptx'), await savePresentation(deck));
  const file = join(dir, 'deck.tsx');
  await writeFile(
    file,
    `import {readFile} from 'node:fs/promises';import {Presentation} from '@office-kit/pptx-dsl';export default <Presentation source={await readFile(${JSON.stringify(join(dir, 'source.pptx'))})} />;`,
  );
  return { dir, file };
}

test(
  'presentation playback follows custom show order and looping from show properties',
  { timeout: 60000 },
  async () => {
    const { dir, file } = await makeDeck();
    let preview, browser;
    try {
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1400, height: 900 } });
      await page.route('**/state*', async (route) => {
        const response = await route.fetch();
        const body = await response.json();
        body.showProperties = {
          mode: { kind: 'present' },
          slides: { kind: 'customShow', id: 7 },
          loop: true,
          showNarration: false,
          showAnimation: false,
          useTimings: false,
        };
        body.customShows = [{ id: 7, name: 'Odd order', slideIndices: [2, 2, 0] }];
        await route.fulfill({ response, body: JSON.stringify(body) });
      });
      await page.goto(preview.url);
      await page.getByRole('button', { name: 'Preview', exact: true }).click();
      await page.getByRole('button', { name: 'Present', exact: true }).click();
      const slideText = () =>
        page.locator('#slide').evaluate((node) => node.shadowRoot?.textContent ?? '');
      await page.keyboard.press('End');
      await page.waitForFunction(() =>
        document.querySelector('#slide')?.shadowRoot?.textContent?.includes('A'),
      );
      await page.getByRole('button', { name: 'Next slide', exact: true }).click();
      await page.waitForFunction(() =>
        document.querySelector('#slide')?.shadowRoot?.textContent?.includes('C'),
      );
      await page.keyboard.press('Home');
      await page.waitForFunction(() =>
        document.querySelector('#slide')?.shadowRoot?.textContent?.includes('C'),
      );
      await page.getByRole('button', { name: 'Next slide', exact: true }).click();
      await page.waitForFunction(() =>
        document.querySelector('#slide')?.shadowRoot?.textContent?.includes('C'),
      );
      await page.getByRole('button', { name: 'Next slide', exact: true }).click();
      await page.waitForFunction(() =>
        document.querySelector('#slide')?.shadowRoot?.textContent?.includes('A'),
      );
      await page.getByRole('button', { name: 'Next slide', exact: true }).click();
      await page.waitForFunction(() =>
        document.querySelector('#slide')?.shadowRoot?.textContent?.includes('C'),
      );
      assert.match(await slideText(), /C/);
      assert.equal(await page.locator('#present-next').isDisabled(), false);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

test(
  'presentation playback restricts a configured range and does not use timings',
  { timeout: 60000 },
  async () => {
    const { dir, file } = await makeDeck();
    let preview, browser;
    try {
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1400, height: 900 } });
      await page.route('**/state*', async (route) => {
        const response = await route.fetch();
        const body = await response.json();
        body.showProperties = {
          mode: { kind: 'present' },
          slides: { kind: 'range', start: 2, end: 3 },
          loop: false,
          showNarration: false,
          showAnimation: false,
          useTimings: false,
        };
        body.transitions = [{ advanceAfterMs: 100 }, {}, {}];
        body.customShows = [];
        await route.fulfill({ response, body: JSON.stringify(body) });
      });
      await page.goto(preview.url);
      await page.getByRole('button', { name: 'Preview', exact: true }).click();
      await page.getByRole('button', { name: 'Present', exact: true }).click();
      const slideText = () =>
        page.locator('#slide').evaluate((node) => node.shadowRoot?.textContent ?? '');
      await page.waitForFunction(() =>
        document.querySelector('#slide')?.shadowRoot?.textContent?.includes('B'),
      );
      await page.waitForTimeout(250);
      assert.match(await slideText(), /B/);
      await page.getByRole('button', { name: 'Next slide', exact: true }).click();
      await page.waitForFunction(() =>
        document.querySelector('#slide')?.shadowRoot?.textContent?.includes('C'),
      );
      assert.equal(await page.locator('#present-next').isDisabled(), true);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

test(
  'Home and End land on visible ends of a duplicate custom show',
  { timeout: 60000 },
  async () => {
    const { dir, file } = await makeDeck();
    let preview, browser;
    try {
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1400, height: 900 } });
      await page.route('**/state*', async (route) => {
        const response = await route.fetch();
        const body = await response.json();
        body.hiddenSlides = [true, false, false];
        body.showProperties = {
          mode: { kind: 'present' },
          slides: { kind: 'customShow', id: 8 },
          loop: false,
          showNarration: false,
          showAnimation: false,
          useTimings: false,
        };
        body.customShows = [{ id: 8, name: 'Visible duplicates', slideIndices: [0, 2, 1, 2, 0] }];
        await route.fulfill({ response, body: JSON.stringify(body) });
      });
      await page.goto(preview.url);
      await page.getByRole('button', { name: 'Preview', exact: true }).click();
      await page.getByRole('button', { name: 'Present', exact: true }).click();
      const slideText = () =>
        page.locator('#slide').evaluate((node) => node.shadowRoot?.textContent ?? '');
      await page.keyboard.press('Home');
      await page.waitForFunction(() =>
        document.querySelector('#slide')?.shadowRoot?.textContent?.includes('C'),
      );
      await page.getByRole('button', { name: 'Next slide', exact: true }).click();
      await page.waitForFunction(() =>
        document.querySelector('#slide')?.shadowRoot?.textContent?.includes('B'),
      );
      assert.match(await slideText(), /B/);
      await page.keyboard.press('End');
      await page.waitForFunction(() =>
        document.querySelector('#slide')?.shadowRoot?.textContent?.includes('C'),
      );
      await page.getByRole('button', { name: 'Previous slide', exact: true }).click();
      await page.waitForFunction(() =>
        document.querySelector('#slide')?.shadowRoot?.textContent?.includes('B'),
      );
      assert.match(await slideText(), /B/);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

test(
  'presentation advances after a configured timing when timings are enabled',
  { timeout: 60000 },
  async () => {
    const { dir, file } = await makeDeck();
    let preview, browser;
    try {
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1400, height: 900 } });
      await page.route('**/state*', async (route) => {
        const response = await route.fetch();
        const body = await response.json();
        body.showProperties = {
          mode: { kind: 'present' },
          slides: { kind: 'all' },
          loop: false,
          showNarration: false,
          showAnimation: false,
          useTimings: true,
        };
        body.transitions = [{ advanceAfterMs: 100 }, {}, {}];
        body.customShows = [];
        await route.fulfill({ response, body: JSON.stringify(body) });
      });
      await page.goto(preview.url);
      await page.getByRole('button', { name: 'Preview', exact: true }).click();
      await page.getByRole('button', { name: 'Present', exact: true }).click();
      await page.waitForFunction(
        () => document.querySelector('#slide')?.shadowRoot?.textContent?.includes('B'),
        undefined,
        { timeout: 5000 },
      );
      assert.match(
        await page.locator('#slide').evaluate((node) => node.shadowRoot?.textContent ?? ''),
        /B/,
      );
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

test(
  'restarting a consecutively repeated animated slide resets its build',
  { timeout: 60000 },
  async () => {
    const { dir, file } = await makeAnimatedDeck();
    let preview, browser;
    try {
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1400, height: 900 } });
      await page.route('**/state*', async (route) => {
        const response = await route.fetch();
        const body = await response.json();
        body.showProperties = {
          mode: { kind: 'present' },
          slides: { kind: 'customShow', id: 9 },
          loop: false,
          showNarration: false,
          showAnimation: true,
          useTimings: false,
        };
        body.customShows = [{ id: 9, name: 'Repeated animation', slideIndices: [0, 0] }];
        await route.fulfill({ response, body: JSON.stringify(body) });
      });
      await page.goto(preview.url);
      await page.getByRole('button', { name: 'Preview', exact: true }).click();
      await page.getByRole('button', { name: 'Present', exact: true }).click();
      await page.waitForFunction(() => animationPlayer?.cursor === 0);
      await page.getByRole('button', { name: 'Next slide', exact: true }).click();
      await page.waitForFunction(() => animationPlayer?.cursor === 1);
      await page.getByRole('button', { name: 'Next slide', exact: true }).click();
      await page.waitForFunction(() => animationPlayer?.cursor === 0);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

test(
  'presentation is unavailable when every configured slide is hidden',
  { timeout: 60000 },
  async () => {
    const { dir, file } = await makeDeck();
    let preview, browser;
    try {
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1400, height: 900 } });
      await page.route('**/state*', async (route) => {
        const response = await route.fetch();
        const body = await response.json();
        body.hiddenSlides = [true, true, true];
        body.showProperties = {
          mode: { kind: 'present' },
          slides: { kind: 'all' },
          loop: false,
          showNarration: false,
          showAnimation: true,
          useTimings: true,
        };
        await route.fulfill({ response, body: JSON.stringify(body) });
      });
      await page.goto(preview.url);
      await page.getByRole('button', { name: 'Preview', exact: true }).click();
      await page.waitForFunction(() =>
        /3/.test(document.querySelector('#count')?.textContent ?? ''),
      );
      assert.equal(
        await page.getByRole('button', { name: 'Present', exact: true }).isDisabled(),
        true,
      );
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
