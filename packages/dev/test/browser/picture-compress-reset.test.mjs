import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { deflateSync } from 'node:zlib';
import { chromium } from 'playwright';
import {
  getShapeBounds,
  getShapeImageBytes,
  getShapeImageCrop,
  getSlides,
  getSlideShapes,
  inches,
  loadPresentation,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

const WIDTH = 1200;
const HEIGHT = 600;

// A WIDTH × HEIGHT opaque PNG with four colored quadrants, no pHYs (96 dpi).
function quadrantsPng() {
  const crcTable = Array.from({ length: 256 }, (_, n) => {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    return c >>> 0;
  });
  const crc = (bytes) => {
    let c = 0xffffffff;
    for (const byte of bytes) c = crcTable[(c ^ byte) & 0xff] ^ (c >>> 8);
    return (c ^ 0xffffffff) >>> 0;
  };
  const chunk = (type, data) => {
    const body = Buffer.concat([Buffer.from(type, 'ascii'), data]);
    const out = Buffer.alloc(body.length + 8);
    out.writeUInt32BE(data.length, 0);
    body.copy(out, 4);
    out.writeUInt32BE(crc(body), body.length + 4);
    return out;
  };
  const header = Buffer.alloc(13);
  header.writeUInt32BE(WIDTH, 0);
  header.writeUInt32BE(HEIGHT, 4);
  header.set([8, 2, 0, 0, 0], 8);
  const colors = [
    [228, 82, 82],
    [69, 179, 107],
    [56, 124, 220],
    [255, 205, 73],
  ];
  const raw = Buffer.alloc((WIDTH * 3 + 1) * HEIGHT);
  for (let y = 0; y < HEIGHT; y++) {
    const row = y * (WIDTH * 3 + 1);
    for (let x = 0; x < WIDTH; x++)
      raw.set(colors[(y < HEIGHT / 2 ? 0 : 2) + (x < WIDTH / 2 ? 0 : 1)], row + 1 + x * 3);
  }
  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk('IHDR', header),
    chunk('IDAT', deflateSync(raw)),
    chunk('IEND', Buffer.alloc(0)),
  ]);
}

const pixelSize = (bytes) => {
  const view = Buffer.from(bytes);
  return { width: view.readUInt32BE(16), height: view.readUInt32BE(20) };
};

test(
  'Compress Pictures resamples and drops cropped areas; Reset Picture & Size restores the natural size (English and Japanese)',
  { timeout: 120000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-picture-compress-'));
    let preview, browser;
    try {
      const png = quadrantsPng();
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import { Presentation, Slide, Image } from '@office-kit/pptx-dsl';
const png = Uint8Array.from(atob('${png.toString('base64')}'), (c) => c.charCodeAt(0));
export default <Presentation><Slide><Image data={png} x={1} y={1} width={2} height={1} /></Slide></Presentation>;
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

      // Crop half the picture away on the left.
      await panel.getByRole('button', { name: 'Crop', exact: true }).click();
      await panel.getByRole('menuitem', { name: 'Crop', exact: true }).click();
      const crop = page.getByRole('dialog', { name: 'Crop image' });
      await crop.locator('img.source').waitFor();
      for (let i = 0; i < 5; i++)
        await crop
          .getByRole('button', { name: 'Crop left edge', exact: true })
          .press('Shift+ArrowRight');
      await changed(() => crop.getByRole('button', { name: 'Apply', exact: true }).click());
      const cropped = await picture();
      assert.ok(Math.abs(getShapeImageCrop(cropped).left - 0.5) < 0.001);
      const frame = getShapeBounds(cropped);
      // The visible half is 600 × 600 px. A stretched frame keeps the denser
      // axis at the target resolution, so both axes scale by the larger ratio.
      const expected = (ppi) => {
        const scale = Math.max(ppi / (600 / (frame.w / 914400)), ppi / (600 / (frame.h / 914400)));
        return { width: Math.round(600 * scale), height: Math.round(600 * scale) };
      };

      // Compress Pictures ▸ Email (96 ppi), deleting cropped areas.
      await panel.getByRole('button', { name: 'Compress Pictures', exact: true }).click();
      let dialog = page.getByRole('dialog', { name: 'Compress Pictures' });
      assert.equal(
        await dialog
          .getByRole('checkbox', { name: 'Delete cropped areas of pictures' })
          .isChecked(),
        true,
      );
      assert.equal(
        await dialog.getByRole('radio', { name: 'Selected pictures only' }).isChecked(),
        true,
      );
      await dialog.getByRole('combobox').selectOption({ label: 'Email (96 ppi)' });
      await changed(() => dialog.getByRole('button', { name: 'OK', exact: true }).click());
      const compressed = await picture();
      assert.equal(getShapeImageCrop(compressed), null);
      assert.deepEqual(getShapeBounds(compressed), frame);
      assert.deepEqual(pixelSize(getShapeImageBytes(compressed)), expected(96));

      await changed(() => page.keyboard.press('ControlOrMeta+z'));
      const restored = await picture();
      assert.deepEqual(Buffer.from(getShapeImageBytes(restored)), png);
      assert.ok(Math.abs(getShapeImageCrop(restored).left - 0.5) < 0.001);

      // Reset Picture ▸ Reset Picture & Size drops the crop and restores 1200 × 600 px at 96 dpi.
      await page.locator('.hit').first().click();
      await panel.getByRole('button', { name: 'Reset Picture', exact: true }).click();
      await changed(() =>
        panel.getByRole('menuitem', { name: 'Reset Picture & Size', exact: true }).click(),
      );
      const reset = await picture();
      assert.equal(getShapeImageCrop(reset), null);
      const bounds = getShapeBounds(reset);
      assert.deepEqual(
        [bounds.x, bounds.y, bounds.w, bounds.h],
        [frame.x, frame.y, inches(WIDTH / 96), inches(HEIGHT / 96)],
      );

      // Artistic Effects lists the reference desktop app's gallery, the current effect (None) checked.
      await panel.getByRole('button', { name: 'Artistic Effects', exact: true }).click();
      const effects = panel
        .getByRole('group', { name: 'Artistic Effect' })
        .getByRole('menuitemradio');
      assert.equal(await effects.count(), 23);
      assert.equal(await effects.first().getAttribute('aria-checked'), 'true');
      assert.equal(await effects.nth(13).textContent(), 'Mosaic Bubbles');
      await page.keyboard.press('Escape');

      // Japanese.
      await page.locator('.lang select').selectOption('ja');
      await page.locator('.hit').first().click();
      await page.getByRole('tab', { name: '図の形式', exact: true }).click();
      await panel.getByRole('button', { name: 'アート効果', exact: true }).click();
      assert.equal(
        await panel
          .getByRole('menuitemradio', { name: 'モザイク: バブル', exact: true })
          .isDisabled(),
        true,
      );
      await page.keyboard.press('Escape');
      // Undo the reset, then compress every picture to On-screen (150 ppi).
      await changed(() => page.keyboard.press('ControlOrMeta+z'));
      await page.locator('.hit').first().click();
      await page.getByRole('tab', { name: '図の形式', exact: true }).click();
      await panel.getByRole('button', { name: '図の圧縮', exact: true }).click();
      dialog = page.getByRole('dialog', { name: '図の圧縮' });
      await dialog.getByRole('combobox').selectOption({ label: '画面表示 (150 ppi)' });
      await dialog.getByRole('radio', { name: 'このファイル内のすべての画像' }).check();
      await changed(() => dialog.getByRole('button', { name: 'OK', exact: true }).click());
      const screen = await picture();
      assert.equal(getShapeImageCrop(screen), null);
      assert.deepEqual(pixelSize(getShapeImageBytes(screen)), expected(150));
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
