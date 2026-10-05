import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  getShapeTextLanguage,
  getSlideComments,
  getSlides,
  getSlideShapes,
  isShapeHidden,
  loadPresentation,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

const DECK = `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={4} height={1}>One</Text></Slide><Slide><Text x={1} y={1} width={4} height={1}>Two</Text></Slide><Slide><Text x={1} y={1} width={4} height={1}>Three</Text></Slide></Presentation>`;

test(
  'Review tab: language, comment navigation and deletion, accessibility and Hide Ink',
  { timeout: 180000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-review-tab-'));
    const file = join(dir, 'deck.tsx');
    await writeFile(file, DECK);
    let preview;
    let browser;
    try {
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1512, height: 900 } });
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      const slides = async () =>
        getSlides(
          await loadPresentation(
            new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
          ),
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
      const button = (name) => panel.getByRole('button', { name, exact: true });
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      await editor.getByRole('tab', { name: 'Review', exact: true }).click();
      assert.equal(await button('Spelling').isDisabled(), true);
      assert.equal(await button('Hide Ink').isDisabled(), true);

      // Language marks the selected box's text.
      await editor.locator('.hit').first().click();
      await button('Language').click();
      await changed(() => panel.getByRole('menuitemradio', { name: 'Japanese (Japan)' }).click());
      assert.equal(getShapeTextLanguage(getSlideShapes((await slides())[0])[0]), 'ja-JP');

      // Comments on slides 1 and 3: Next skips slide 2, Delete clears one slide.
      await page.keyboard.press('Escape');
      await button('New Comment').click();
      const comments = editor.getByRole('dialog', { name: 'Comments', exact: true });
      // The dialog opens with a draft on the current slide; slide 3 needs one added.
      for (const index of ['0', '2']) {
        await comments
          .getByRole('combobox', { name: 'Review slide', exact: true })
          .selectOption(index);
        if ((await comments.getByLabel('Comment text', { exact: true }).count()) === 0)
          await comments.getByRole('button', { name: 'Add comment', exact: true }).click();
        await comments.getByLabel('Author name', { exact: true }).last().fill('Reviewer');
        await comments.getByLabel('Comment text', { exact: true }).last().fill(`Note ${index}`);
      }
      await changed(() => comments.getByRole('button', { name: 'Apply', exact: true }).click());
      await editor.locator('.thumb').nth(0).click();
      await button('Next').click();
      await editor.getByText('Slide 3 of 3', { exact: true }).waitFor();
      assert.equal(await button('Show Comments').getAttribute('aria-pressed'), 'true');
      await page.keyboard.press('Escape');
      await button('Previous').click();
      await editor.getByText('Slide 1 of 3', { exact: true }).waitFor();
      await page.keyboard.press('Escape');
      await button('Delete').click();
      await changed(() =>
        panel.getByRole('menuitem', { name: 'Delete All Comments on This Slide' }).click(),
      );
      const after = await slides();
      assert.equal(getSlideComments(after[0]).length, 0);
      assert.equal(getSlideComments(after[2]).length, 1);

      // Check Accessibility opens the status bar's issue list.
      await button('Check Accessibility').click();
      await editor.getByRole('dialog', { name: 'Accessibility', exact: true }).waitFor();
      await page.keyboard.press('Escape');

      // Hide Ink hides every pen stroke.
      await editor.getByRole('tab', { name: 'Draw', exact: true }).click();
      const stage = await editor.locator('.stage').boundingBox();
      await panel.getByRole('radio', { name: 'Pen: Black, 1 mm', exact: true }).click();
      await changed(async () => {
        await page.mouse.move(stage.x + stage.width * 0.5, stage.y + stage.height * 0.6);
        await page.mouse.down();
        await page.mouse.move(stage.x + stage.width * 0.7, stage.y + stage.height * 0.7, {
          steps: 4,
        });
        await page.mouse.up();
      });
      await page.keyboard.press('Escape');
      await editor.getByRole('tab', { name: 'Review', exact: true }).click();
      await changed(() => button('Hide Ink').click());
      const ink = getSlideShapes((await slides())[0]).find((shape) => isShapeHidden(shape));
      assert.ok(ink, 'the stroke is hidden');
      assert.equal(await button('Hide Ink').getAttribute('aria-pressed'), 'true');
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
