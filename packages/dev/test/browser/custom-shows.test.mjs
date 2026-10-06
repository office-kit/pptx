import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  getCustomShows,
  getSlideShowProperties,
  getSlides,
  loadPresentation,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

test(
  'Custom Shows can be created, edited, reordered, copied, deleted and restored in both languages',
  { timeout: 120000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-custom-shows-'));
    let preview;
    let browser;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={3} height={1}>One</Text></Slide><Slide><Text x={1} y={1} width={3} height={1}>Two</Text></Slide><Slide><Text x={1} y={1} width={3} height={1}>Three</Text></Slide></Presentation>`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      await page.goto(preview.url);
      const editor = page.frameLocator('#editor-frame');
      const saved = () => editor.getByText('Saved to this project', { exact: true }).waitFor();
      const openShows = async (language = 'en') => {
        await editor.locator('.lang select').selectOption(language);
        await editor
          .getByRole('tab', {
            name: language === 'en' ? 'Slide Show' : 'スライド ショー',
            exact: true,
          })
          .click();
        await editor
          .getByRole('button', {
            name: language === 'en' ? 'Custom Show' : 'カスタム スライド ショー',
            exact: true,
          })
          .click();
        return editor.getByRole('dialog', {
          name: language === 'en' ? 'Custom Shows' : 'カスタム ショー',
        });
      };
      const readShows = async () => {
        const pres = await loadPresentation(
          new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
        );
        const slides = getSlides(pres);
        return getCustomShows(pres).map((show) => ({
          name: show.name,
          slides: show.slides.map((slide) => slides.indexOf(slide) + 1),
        }));
      };
      const readShowProperties = async () => {
        const pres = await loadPresentation(
          new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
        );
        return getSlideShowProperties(pres);
      };

      await saved();
      let dialog = await openShows();
      await dialog.getByRole('button', { name: /New/ }).click();
      await dialog.getByLabel('Name', { exact: true }).fill('Morning');
      await dialog.getByRole('button', { name: 'Add slide 3', exact: true }).click();
      await dialog.getByRole('button', { name: 'Add slide 1', exact: true }).click();
      await dialog.getByRole('button', { name: 'Add slide 3', exact: true }).click();
      await dialog
        .getByRole('button', { name: 'Move slide in custom show up', exact: true })
        .nth(1)
        .click();
      await dialog
        .getByRole('button', { name: 'Move slide in custom show down', exact: true })
        .nth(0)
        .click();
      await dialog
        .getByRole('button', { name: 'Move slide in custom show up', exact: true })
        .nth(1)
        .click();
      await dialog
        .getByRole('button', { name: 'Remove slide from show', exact: true })
        .nth(2)
        .click();
      await dialog.getByRole('button', { name: 'Apply', exact: true }).click();
      await saved();

      await dialog.getByRole('button', { name: /Morning/ }).click();
      await dialog.getByLabel('Name', { exact: true }).fill('Edited');
      await dialog.getByRole('button', { name: 'Add slide 3', exact: true }).click();
      await dialog.getByRole('button', { name: 'Apply', exact: true }).click();
      await saved();

      await dialog.getByRole('button', { name: /New/ }).click();
      await dialog.getByLabel('Name', { exact: true }).fill('Evening');
      await dialog.getByRole('button', { name: 'Add slide 2', exact: true }).click();
      await dialog.getByRole('button', { name: 'Apply', exact: true }).click();
      await saved();
      await dialog.getByRole('button', { name: /Evening/ }).click();
      await dialog.getByRole('button', { name: 'Move custom show up', exact: true }).click();
      await saved();
      assert.deepEqual(await readShows(), [
        { name: 'Evening', slides: [2] },
        { name: 'Edited', slides: [1, 3, 3] },
      ]);

      await dialog.getByLabel('Close', { exact: true }).click();
      await editor.getByRole('button', { name: 'Set Up Slide Show', exact: true }).click();
      const setupDialog = editor.getByRole('dialog', { name: 'Set Up Show', exact: true });
      await setupDialog.getByLabel('Custom show:', { exact: true }).check();
      await setupDialog
        .getByLabel('Custom show', { exact: true })
        .selectOption({ label: 'Evening' });
      await setupDialog.getByLabel("Loop continuously until 'Esc'", { exact: true }).check();
      await setupDialog.getByLabel('Show without narration', { exact: true }).uncheck();
      await setupDialog.getByRole('button', { name: 'OK', exact: true }).click();
      await saved();
      const selectedShow = await readShowProperties();
      assert.equal(selectedShow.slides.kind, 'customShow');
      assert.equal(selectedShow.loop, true);
      assert.equal(selectedShow.showNarration, true);

      dialog = await openShows();
      await dialog.getByRole('button', { name: /^Evening 1 slides$/ }).click();
      await dialog.getByRole('button', { name: 'Delete', exact: true }).click();
      await saved();
      await dialog.getByLabel('Close', { exact: true }).click();
      const cleared = await readShowProperties();
      assert.deepEqual(cleared.slides, { kind: 'all' });
      assert.equal(cleared.loop, true);
      assert.equal(cleared.showNarration, true);
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      const restored = await readShowProperties();
      assert.deepEqual(restored.slides, selectedShow.slides);
      assert.equal(restored.loop, true);
      assert.equal(restored.showNarration, true);
      assert.equal(restored.slides.kind, 'customShow');
      assert.equal(restored.slides.id, selectedShow.slides.id);
      assert.equal(
        (await readShows()).some((show) => show.name === 'Evening'),
        true,
      );

      dialog = await openShows();
      await dialog.getByRole('button', { name: /^Edited 3 slides$/ }).click();
      await dialog.getByRole('button', { name: 'Delete', exact: true }).click();
      await saved();
      await dialog.getByLabel('Close', { exact: true }).click();
      const afterUnrelatedDelete = await readShowProperties();
      assert.deepEqual(afterUnrelatedDelete.slides, selectedShow.slides);
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();

      dialog = await openShows();
      await dialog.getByRole('button', { name: /^Evening 1 slides$/ }).click();
      await dialog.getByRole('button', { name: 'Copy', exact: true }).click();
      await saved();
      assert.deepEqual(
        (await readShows()).map((show) => show.name),
        ['Evening', 'Edited', 'Evening Copy'],
      );

      await dialog.getByRole('button', { name: /Evening Copy/ }).click();
      await dialog.getByRole('button', { name: 'Delete', exact: true }).click();
      await saved();
      await dialog.getByLabel('Close', { exact: true }).click();
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual(
        (await readShows()).map((show) => show.name),
        ['Evening', 'Edited', 'Evening Copy'],
      );
      assert.deepEqual(
        (await readShows()).map((show) => show.slides),
        [[2], [1, 3, 3], [2]],
      );

      await page.reload();
      await saved();
      assert.deepEqual(
        (await readShows()).map((show) => show.name),
        ['Evening', 'Edited', 'Evening Copy'],
      );

      const jaDialog = await openShows('ja');
      await jaDialog.getByRole('button', { name: /^Evening 1 slides$/ }).click();
      assert.equal(await jaDialog.getByRole('button', { name: /新規/ }).count(), 1);
      assert.equal(await jaDialog.getByRole('button', { name: 'コピー', exact: true }).count(), 1);
      assert.equal(await jaDialog.getByRole('button', { name: '削除', exact: true }).count(), 1);
      assert.deepEqual(
        (await readShows()).map((show) => show.name),
        ['Evening', 'Edited', 'Evening Copy'],
      );
      await jaDialog.press('Escape');
      assert.equal(await editor.getByRole('dialog', { name: 'カスタム ショー' }).count(), 0);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
