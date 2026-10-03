import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  getSlideShapes,
  getSlides,
  getSlideShowProperties,
  savePresentation,
  setCustomShows,
  setShapeClickAction,
  setSlideShowProperties,
} from '@office-kit/pptx';
import { compile, Presentation, Slide, Text } from '@office-kit/pptx-dsl';
import { startPreview } from '../helpers/server.mjs';

for (const [returnToShow, automatic] of [
  [true, false],
  [false, false],
  [true, true],
])
  test(
    `custom show links preserve order and return choice (${returnToShow}, automatic: ${automatic})`,
    { timeout: 30000 },
    async () => {
      const dir = await mkdtemp(join(tmpdir(), 'office-custom-links-'));
      let preview, browser;
      try {
        const deck = await compile(
          Presentation({
            children: ['Opening', 'Caller', 'Part A', 'Part B', 'Closing'].map((name) =>
              Slide({
                children: [name, 'Open custom', 'End show', 'Nested show'].map((children, i) =>
                  Text({ x: 1, y: 1 + i, width: 5, height: 0.8, children }),
                ),
              }),
            ),
          }),
        );
        const slides = getSlides(deck);
        setCustomShows(deck, [
          { id: 7, name: 'Main', slides: [slides[0], slides[1], slides[1], slides[4]] },
          { id: 8, name: 'Detail', slides: [slides[3], slides[2], slides[3]] },
          { id: 9, name: 'Nested', slides: [slides[2]] },
        ]);
        setSlideShowProperties(deck, {
          ...getSlideShowProperties(deck),
          slides: { kind: 'customShow', id: 7 },
        });
        for (const slide of slides) {
          const shapes = getSlideShapes(slide);
          setShapeClickAction(shapes[1], { kind: 'customShow', id: 8, returnToShow });
          setShapeClickAction(shapes[2], { kind: 'endShow' });
          setShapeClickAction(
            shapes[3],
            { kind: 'customShow', id: 9, returnToShow: true },
            { range: { start: 0, end: 11 } },
          );
        }
        const source = join(dir, 'source.pptx');
        await writeFile(source, await savePresentation(deck));
        const file = join(dir, 'deck.tsx');
        await writeFile(
          file,
          `import {readFile} from 'node:fs/promises';import {Presentation} from '@office-kit/pptx-dsl';export default <Presentation source={await readFile(${JSON.stringify(source)})} />;`,
        );
        preview = await startPreview(file);
        browser = await chromium.launch({ headless: true });
        const page = await browser.newPage({ viewport: { width: 1400, height: 900 } });
        if (automatic)
          await page.route('**/state*', async (route) => {
            const response = await route.fetch();
            const body = await response.json();
            body.transitions = [{}, {}, { advanceAfterMs: 250 }, { advanceAfterMs: 250 }, {}];
            await route.fulfill({ response, body: JSON.stringify(body) });
          });
        await page.goto(preview.url);
        await page.getByRole('button', { name: 'Preview', exact: true }).click();
        const expectSlide = async (name) =>
          page.waitForFunction(
            (name) => document.querySelector('#slide')?.shadowRoot?.textContent?.includes(name),
            name,
          );
        const start = async () => {
          await page.getByRole('button', { name: 'Present', exact: true }).click();
          await expectSlide('Opening');
          await page.keyboard.press('ArrowRight');
          await expectSlide('Caller');
          await page.keyboard.press('ArrowRight');
          const link = page.locator(
            `#slide a[href="#pptx-custom-show?id=8&return=${returnToShow}"]`,
          );
          assert.equal(await link.count(), 1);
          await link.click();
          await expectSlide('Part B');
        };
        await start();
        if (automatic) {
          await expectSlide('Caller');
          await page.keyboard.press('ArrowRight');
          await expectSlide('Closing');
          assert.equal(
            await page.locator('body').evaluate((node) => node.classList.contains('presenting')),
            true,
          );
          return;
        }
        const nested = page.locator('#slide a[href="#pptx-custom-show?id=9&return=true"]');
        await nested.locator('span').last().click();
        await expectSlide('Part A');
        await page.locator('#slide a[href="#pptx-end-show"]').click();
        await expectSlide('Part B');
        await page.keyboard.press('ArrowRight');
        await expectSlide('Part A');
        await page.keyboard.press('ArrowRight');
        await expectSlide('Part B');
        await page.getByRole('button', { name: 'Next slide', exact: true }).click();
        if (returnToShow) {
          await expectSlide('Caller');
          assert.equal(
            await page.locator('body').evaluate((node) => node.classList.contains('presenting')),
            true,
          );
          await page.keyboard.press('ArrowRight');
          await expectSlide('Closing');
          await page.locator('#slide a[href="#pptx-end-show"]').click();
        }
        await page.waitForFunction(
          () => !document.body.classList.contains('presenting') && !document.fullscreenElement,
        );
        await start();
        await page.locator('#slide a[href="#pptx-end-show"]').click();
        if (returnToShow) {
          await expectSlide('Caller');
          await page.keyboard.press('Escape');
        }
        await page.waitForFunction(
          () => !document.body.classList.contains('presenting') && !document.fullscreenElement,
        );
        const popup = page.waitForEvent('popup');
        await page.getByRole('button', { name: 'Presenter view', exact: true }).click();
        const presenter = await popup;
        await presenter
          .locator(`#current a[href="#pptx-custom-show?id=8&return=${returnToShow}"]`)
          .click();
        await expectSlide('Part B');
        await presenter.locator('#current a[href="#pptx-end-show"]').click();
        if (returnToShow) await expectSlide('Opening');
        else await page.waitForFunction(() => !document.body.classList.contains('presenting'));
      } finally {
        await browser?.close();
        await preview?.close();
        await rm(dir, { recursive: true, force: true });
      }
    },
  );
