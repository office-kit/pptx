import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  getSlides,
  getSlideShowProperties,
  getSlideTransition,
  loadPresentation,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

const DECK = `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={4} height={1}>One</Text></Slide><Slide><Text x={1} y={1} width={4} height={1}>Two</Text></Slide></Presentation>`;

test(
  'Slide Show tab: Rehearse Timings keeps slide times; show options write presProps',
  { timeout: 180000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-slide-show-tab-'));
    const file = join(dir, 'deck.tsx');
    await writeFile(file, DECK);
    let preview;
    let browser;
    try {
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1512, height: 900 } });
      await page.goto(preview.url);
      const editor = page.frameLocator('#editor-frame');
      const deck = async () =>
        loadPresentation(
          new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
        );
      const changed = async (action) => {
        const before = await (await fetch(preview.url + '/editor/state')).json();
        await action();
        for (let i = 0; i < 200; i += 1) {
          const state = await (await fetch(preview.url + '/editor/state')).json();
          if (!state.building && state.revision !== before.revision) return;
          await new Promise((resolve) => setTimeout(resolve, 50));
        }
        throw new Error('The edit was not saved');
      };
      const panel = editor.locator('#ribbon-panel');
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      await editor.getByRole('tab', { name: 'Slide Show', exact: true }).click();
      assert.equal(
        await panel.getByRole('button', { name: 'Rehearse with Coach' }).isDisabled(),
        true,
      );

      await changed(() => panel.getByRole('checkbox', { name: 'Show Media Controls' }).uncheck());
      assert.equal(getSlideShowProperties(await deck()).showMediaControls, false);

      // Rehearse: a timer runs while presenting; ending asks to keep the times.
      await panel.getByRole('button', { name: 'Rehearse Timings', exact: true }).click();
      await page.waitForFunction(() => document.body.classList.contains('presenting'));
      await page.locator('#rehearsal-timer').waitFor();
      await page.waitForTimeout(1200);
      await page.keyboard.press('ArrowRight');
      await page.waitForTimeout(600);
      await page.keyboard.press('Escape');
      const question = editor.getByRole('dialog', { name: 'Rehearse Timings' });
      await question.waitFor();
      await changed(() => question.getByRole('button', { name: 'Yes', exact: true }).click());
      const slides = getSlides(await deck());
      const first = getSlideTransition(slides[0])?.advanceAfterMs;
      const second = getSlideTransition(slides[1])?.advanceAfterMs;
      assert.ok(first >= 1000 && first < 5000, `slide 1 timed ${first} ms`);
      assert.ok(second >= 400 && second < 5000, `slide 2 timed ${second} ms`);
      assert.equal(getSlideShowProperties(await deck()).useTimings, true);
      assert.equal(await panel.getByRole('checkbox', { name: 'Use Timings' }).isChecked(), true);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
