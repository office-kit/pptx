import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  getShapeParagraphElements,
  getSlideAnimations,
  getSlides,
  getSlideShapes,
  getSlideTransition,
  getSlideTransitionSound,
  isSlideHidden,
  loadPresentation,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

const DECK = `import {Presentation,Slide,Shape} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Shape preset="rect" x={1} y={1} width={3} height={1} text="One" /></Slide><Slide><Shape preset="rect" x={1} y={1} width={3} height={1} text="Two" /></Slide></Presentation>`;

test(
  'Insert, Transitions, Animations and Slide Show tabs edit the deck like PowerPoint',
  { timeout: 180000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-ribbon-tabs-'));
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
      const tab = (name) => editor.getByRole('tab', { name, exact: true }).click();
      const panel = editor.locator('#ribbon-panel');
      await editor.getByText('Saved to this project', { exact: true }).waitFor();

      // Insert ▸ Slide Number puts a live slide-number field in the selected box.
      await editor.locator('.hit').first().click();
      await tab('Insert');
      await changed(() => panel.getByRole('button', { name: 'Slide Number', exact: true }).click());
      const firstShape = getSlideShapes(getSlides(await deck())[0])[0];
      assert.ok(
        getShapeParagraphElements(firstShape, 0).some(
          (element) => element.kind === 'fld' && element.type === 'slidenum',
        ),
      );

      // Transitions: the gallery sets the effect, the timing keeps it, and
      // Apply To All copies it to every slide.
      await page.keyboard.press('Escape');
      await tab('Transitions');
      await changed(() => panel.getByRole('radio', { name: 'Fade', exact: true }).click());
      assert.equal(getSlideTransition(getSlides(await deck())[0])?.effect, 'fade');
      assert.equal(
        await panel.getByRole('radio', { name: 'Fade', exact: true }).getAttribute('aria-checked'),
        'true',
      );
      await changed(() => panel.getByRole('checkbox', { name: 'After:', exact: true }).check());
      await changed(async () => {
        const seconds = panel.getByRole('spinbutton', { name: 'Advance after (seconds)' });
        await seconds.fill('3');
        await seconds.press('Enter');
      });
      let transition = getSlideTransition(getSlides(await deck())[0]);
      assert.equal(transition?.effect, 'fade');
      assert.equal(transition?.advanceAfterMs, 3000);

      // Duration writes p14:dur; Sound stops earlier sounds or embeds a WAV.
      await changed(async () => {
        const duration = panel.getByRole('spinbutton', { name: 'Duration:' });
        await duration.fill('2');
        await duration.press('Enter');
      });
      assert.equal(getSlideTransition(getSlides(await deck())[0])?.durationMs, 2000);
      await panel.getByRole('button', { name: 'Preview', exact: true }).click();
      const sound = panel.getByRole('combobox', { name: 'Sound:' });
      await changed(() => sound.selectOption('stop'));
      assert.deepEqual(getSlideTransitionSound(getSlides(await deck())[0]), { kind: 'stop' });
      const wav = Buffer.alloc(44);
      wav.write('RIFF', 0);
      wav.writeUInt32LE(36, 4);
      wav.write('WAVEfmt ', 8);
      await changed(() =>
        panel
          .locator('input[type="file"]')
          .setInputFiles({ name: 'chime.wav', mimeType: 'audio/wav', buffer: wav }),
      );
      assert.deepEqual(getSlideTransitionSound(getSlides(await deck())[0]), {
        kind: 'play',
        name: 'chime.wav',
        loop: false,
      });
      assert.equal(await sound.inputValue(), 'play');
      await changed(() => panel.getByRole('button', { name: 'Apply To All', exact: true }).click());
      transition = getSlideTransition(getSlides(await deck())[1]);
      assert.equal(transition?.effect, 'fade');
      assert.equal(transition?.advanceAfterMs, 3000);

      // Animations: one gallery effect per shape, replaced in place, None removes it.
      await editor.locator('.thumb').first().click();
      await editor.locator('.hit').first().click();
      await tab('Animations');
      await changed(() => panel.getByRole('radio', { name: 'Fly In', exact: true }).click());
      await changed(() => panel.getByRole('radio', { name: 'Zoom', exact: true }).click());
      let steps = getSlideAnimations(getSlides(await deck())[0]);
      assert.deepEqual(
        steps.map((step) => step.effect),
        ['zoomIn'],
      );
      await changed(() =>
        panel.getByRole('combobox', { name: 'Start', exact: true }).selectOption('afterPrevious'),
      );
      steps = getSlideAnimations(getSlides(await deck())[0]);
      assert.equal(steps[0].start, 'afterPrevious');
      await changed(() => panel.getByRole('radio', { name: 'None', exact: true }).click());
      assert.deepEqual(getSlideAnimations(getSlides(await deck())[0]), []);

      // Slide Show: Hide Slide toggles the selected slide, and Play from
      // Current Slide starts the preview's slide show on that slide.
      await page.keyboard.press('Escape');
      await tab('Slide Show');
      const hide = panel.getByRole('button', { name: 'Hide Slide', exact: true });
      await changed(() => hide.click());
      assert.equal(isSlideHidden(getSlides(await deck())[0]), true);
      assert.equal(await hide.getAttribute('aria-pressed'), 'true');
      await changed(() => hide.click());
      assert.equal(isSlideHidden(getSlides(await deck())[0]), false);
      await editor.locator('.thumb').nth(1).click();
      await panel.getByRole('button', { name: 'Play from Current Slide', exact: true }).click();
      await page.waitForFunction(() => document.body.classList.contains('presenting'));
      await page.waitForFunction(
        () => document.querySelector('#present-count')?.textContent === 'Slide 2 of 2',
      );
      await page.keyboard.press('Escape');
      await page.waitForFunction(() => !document.body.classList.contains('presenting'));

      // The standalone editor has no stage to present on.
      const standalone = await browser.newPage({ viewport: { width: 1512, height: 900 } });
      await standalone.goto(preview.url + '/editor');
      await standalone.getByRole('tab', { name: 'Slide Show', exact: true }).click();
      assert.equal(
        await standalone.getByRole('button', { name: 'Play from Start', exact: true }).isDisabled(),
        true,
      );
      await standalone.close();
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
