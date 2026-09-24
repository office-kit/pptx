import assert from 'node:assert/strict';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  duplicateSlide,
  getSlides,
  getSlideMasterBackgroundStyles,
  loadPresentation,
  savePresentation,
} from '../../../../dist/index.js';
import { startPreview } from '../helpers/server.mjs';

test(
  'background style gallery applies to the master, supports keyboard and undo, and survives reload',
  { timeout: 120000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-background-graphics-'));
    let preview, browser;
    try {
      const pres = await loadPresentation(
        await readFile(
          new URL('../../../../test/fixtures/minimal/layout-decoration.pptx', import.meta.url),
        ),
      );
      const slide = getSlides(pres)[0];
      duplicateSlide(pres, slide);
      duplicateSlide(pres, slide);
      const source = join(dir, 'source.pptx');
      await writeFile(source, await savePresentation(pres));
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {readFile} from 'node:fs/promises';import {Presentation} from '@office-kit/pptx-dsl';export default <Presentation source={await readFile(${JSON.stringify(source)})} />;`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1500, height: 1100 } });
      const errors = [];
      page.on('pageerror', (error) => errors.push(error.message));
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      const saved = () => editor.getByText('Saved to this project', { exact: true }).waitFor();
      const read = async () => {
        const deck = await loadPresentation(
          new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
        );
        return getSlides(deck).map(
          (slide) => getSlideMasterBackgroundStyles(slide).find((style) => style.selected)?.style,
        );
      };
      const open = async () => {
        await editor.getByRole('tab', { name: 'Design', exact: true }).click();
        await editor.getByRole('button', { name: 'Background Styles', exact: true }).click();
      };
      await saved();
      await editor.locator('.thumb-row').nth(1).click();
      const before = await read();
      await open();
      const gallery = editor.getByRole('menu', { name: 'Background Styles', exact: true });
      assert.equal(await gallery.getByRole('menuitemradio').count(), 12);
      assert.equal(
        await gallery
          .getByRole('menuitem', { name: 'Reset Slide Background', exact: true })
          .isEnabled(),
        false,
      );
      await gallery
        .getByRole('menuitemradio', { name: 'Style 1', exact: true })
        .press('ArrowRight');
      assert.equal(
        await gallery
          .getByRole('menuitemradio', { name: 'Style 2', exact: true })
          .evaluate((el) => el === el.ownerDocument.activeElement),
        true,
      );
      await gallery.getByRole('menuitemradio', { name: 'Style 2', exact: true }).press('Escape');
      assert.equal(await gallery.count(), 0);
      assert.deepEqual(await read(), before);
      await open();
      await gallery.getByRole('menuitemradio', { name: 'Style 7', exact: true }).click();
      await saved();
      assert.deepEqual(await read(), [7, 7, 7]);
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual(await read(), before);
      await open();
      await gallery.getByRole('menuitemradio', { name: 'Style 4', exact: true }).click();
      await saved();
      await page.reload();
      await saved();
      await open();
      assert.equal(
        await gallery
          .getByRole('menuitemradio', { name: 'Style 4', exact: true })
          .getAttribute('aria-checked'),
        'true',
      );
      assert.deepEqual(await read(), [4, 4, 4]);
      await gallery.getByRole('menuitem', { name: 'Format Background...', exact: true }).click();
      await editor.getByRole('region', { name: 'Format Background', exact: true }).waitFor();
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
