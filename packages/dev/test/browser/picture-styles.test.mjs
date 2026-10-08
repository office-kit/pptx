import assert from 'node:assert/strict';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  BUILTIN_PICTURE_STYLES,
  getShapeImageBytes,
  getShapeImageCompressionState,
  getShapePictureStyle,
  getSlides,
  getSlideShapes,
  loadPresentation,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

// The picture the reference desktop app's capture used (test/fixtures/native/compress-pictures/):
// the site's 1200 × 630 social card, a well-compressed PNG.
const OG_PNG = new URL('../../../../site/static/og.png', import.meta.url);

const pixelWidth = (bytes) => Buffer.from(bytes).readUInt32BE(16);

test(
  'Picture Styles gallery applies and detects styles; Compress Pictures labels pictures like the reference desktop app (English and Japanese)',
  { timeout: 180000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-picture-styles-'));
    let preview, browser;
    try {
      const png = await readFile(OG_PNG);
      const file = join(dir, 'deck.tsx');
      // 5 in wide, as in the reference desktop app's capture: 240 ppi effective.
      await writeFile(
        file,
        `import { Presentation, Slide, Image } from '@office-kit/pptx-dsl';
const png = Uint8Array.from(atob('${png.toString('base64')}'), (c) => c.charCodeAt(0));
export default <Presentation><Slide><Image data={png} x={1} y={1} width={5} height={2.625} /></Slide></Presentation>;
`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1512, height: 900 } });
      const errors = [];
      page.on('pageerror', (error) => errors.push(error.message));
      await page.goto(preview.url + '/editor');
      await page.getByText('Saved to this project', { exact: true }).waitFor();
      const panel = page.locator('#ribbon-panel');
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
      const picture = async () =>
        getSlideShapes(
          getSlides(
            await loadPresentation(
              new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
            ),
          )[0],
        )[0];

      await page.locator('.hit').first().click();
      await page.getByRole('tab', { name: 'Picture Format', exact: true }).click();
      const gallery = panel.getByRole('radiogroup', { name: 'Quick Styles', exact: true });
      const tiles = gallery.getByRole('radio');
      assert.equal(await tiles.count(), 3);
      assert.deepEqual(
        await tiles.evaluateAll((items) => items.map((item) => item.getAttribute('title'))),
        BUILTIN_PICTURE_STYLES.slice(0, 3),
      );
      // Every tile is the preview renderer's drawing of the style.
      await tiles.first().locator('img').waitFor();
      assert.match(
        await tiles.first().locator('img').getAttribute('src'),
        /^data:image\/svg\+xml,/,
      );
      assert.equal(await gallery.locator('[aria-checked="true"]').count(), 0);

      // Page to the sixth page and apply "Rotated, White" (index 16).
      const next = panel.getByRole('button', { name: 'Next Quick Styles gallery', exact: true });
      for (let i = 0; i < 5; i++) await next.click();
      const rotated = gallery.getByRole('radio', { name: 'Rotated, White', exact: true });
      await changed(() => rotated.click());
      assert.equal(getShapePictureStyle(await picture()), 'Rotated, White');
      assert.equal(await rotated.getAttribute('aria-checked'), 'true');

      // Picture Quality lists the reference desktop app's two items; Upscale is cloud-only.
      await panel.getByRole('button', { name: 'Picture Quality', exact: true }).click();
      const upscale = panel.getByRole('menuitem', { name: 'Upscale Picture', exact: true });
      assert.equal(await upscale.isDisabled(), true);
      assert.match(await upscale.getAttribute('title'), /cloud AI/);
      await panel.getByRole('menuitem', { name: 'Compress Pictures...', exact: true }).click();
      let dialog = page.getByRole('dialog', { name: 'Compress Pictures' });
      await dialog.getByRole('combobox').selectOption({ label: 'Print (220 ppi)' });
      await changed(() => dialog.getByRole('button', { name: 'OK', exact: true }).click());
      let shape = await picture();
      // Like the reference desktop app: labelled for print, but the pixels are kept (a
      // 1100 px re-encode outgrows the original), and the style is untouched.
      assert.equal(getShapeImageCompressionState(shape), 'print');
      assert.deepEqual(Buffer.from(getShapeImageBytes(shape)), png);
      assert.equal(getShapePictureStyle(shape), 'Rotated, White');

      // Japanese: tooltips, another style, and Email.
      await page.locator('.lang select').selectOption('ja');
      await page.locator('.hit').first().click();
      await page.getByRole('tab', { name: '図の形式', exact: true }).click();
      const jaGallery = panel.getByRole('radiogroup', { name: 'クイック スタイル', exact: true });
      // The gallery shows the page holding the applied style.
      assert.equal(
        await jaGallery
          .getByRole('radio', { name: '回転、白', exact: true })
          .getAttribute('aria-checked'),
        'true',
      );
      const jaNext = panel.getByRole('button', {
        name: '次のクイック スタイル ギャラリー',
        exact: true,
      });
      // Page 6 of 10 (styles 15–17) to the last page, which holds style 27 alone.
      for (let i = 0; i < 4; i++) await jaNext.click();
      assert.equal(await jaGallery.getByRole('radio').count(), 1);
      assert.equal(await jaNext.isDisabled(), true);
      await changed(() =>
        jaGallery.getByRole('radio', { name: '楕円、メタル', exact: true }).click(),
      );
      assert.equal(getShapePictureStyle(await picture()), 'Metal Oval');

      await panel.getByRole('button', { name: '画像の品質', exact: true }).click();
      assert.equal(
        await panel
          .getByRole('menuitem', { name: '画像のアップスケール', exact: true })
          .isDisabled(),
        true,
      );
      await panel.getByRole('menuitem', { name: '図の圧縮...', exact: true }).click();
      dialog = page.getByRole('dialog', { name: '図の圧縮' });
      await dialog.getByRole('combobox').selectOption({ label: 'メール (96 ppi)' });
      await changed(() => dialog.getByRole('button', { name: 'OK', exact: true }).click());
      shape = await picture();
      assert.equal(getShapeImageCompressionState(shape), 'email');
      // 5 in × 96 ppi, as the reference desktop app resampled its capture.
      assert.equal(pixelWidth(getShapeImageBytes(shape)), 480);
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
