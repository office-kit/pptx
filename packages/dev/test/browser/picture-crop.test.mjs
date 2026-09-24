import { chromium } from 'playwright';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, writeFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { spawn } from 'node:child_process';
import { fileURLToPath } from 'node:url';

test(
  'Picture Crop to Shape gallery supports selection, keyboard, undo and contextual tabs',
  { timeout: 120000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-table-'));
    const file = join(dir, 'deck.tsx');
    await writeFile(
      file,
      `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={5} height={1}>Alpha [a] ALPHA</Text></Slide><Slide><Text x={1} y={1} width={5} height={1}>Beta alpha</Text></Slide></Presentation>`,
    );
    const proc = spawn(process.execPath, [
      fileURLToPath(new URL('../../dist/cli.mjs', import.meta.url)),
      'dev',
      file,
      '--port',
      '0',
    ]);
    let browser;
    try {
      const url = await new Promise((resolve, reject) => {
        let output = '';
        proc.stdout.on('data', (data) => {
          output += data;
          const match = output.match(/Preview: (http:\/\/\S+)/);
          if (match) resolve(match[1]);
        });
        proc.stderr.on('data', (data) => process.stderr.write(data));
        proc.on('error', reject);
        proc.on('exit', (code) => reject(new Error('Server exited: ' + code)));
      });
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1600, height: 1000 } });
      const errors = [];
      page.on('pageerror', (error) => errors.push(error.message));
      await page.context().grantPermissions(['clipboard-read', 'clipboard-write']);
      await page.goto(url);
      await page.locator('.shape-hit').first().waitFor();

      assert.equal(await page.getByRole('tab', { name: 'Picture Format' }).count(), 0);
      await page.locator('[data-edit="picture"]:visible').click();
      const chooserPromise = page.waitForEvent('filechooser');
      await page.getByRole('menuitem', { name: 'Picture from File...' }).click();
      const chooser = await chooserPromise;
      await chooser.setFiles({
        name: 'sample.png',
        mimeType: 'image/png',
        buffer: Buffer.from(
          'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jp1sAAAAASUVORK5CYII=',
          'base64',
        ),
      });
      await page.waitForFunction(
        () =>
          !window.office.getState().building &&
          window.office.getState().editor.slides[0].shapes.some((s) => s.kind === 'picture'),
      );
      await page.getByRole('tab', { name: 'Picture Format' }).click();
      const transparency = page.locator('[data-edit="picture-transparency"]');
      const corrections = page.locator('[data-edit="picture-corrections"]');
      for (const entry of ['ribbon', 'context']) {
        const beforeReplacement = await page.evaluate(() => ({
          revision: window.office.getState().revision,
          shape: window.office.getState().editor.slides[0].shapes.at(-1),
        }));
        if (entry === 'ribbon') {
          await page.locator('[data-edit="picture-replace"]').click();
        } else {
          await page.locator('.shape-hit').first().click({ button: 'right' });
          assert.equal(
            await page.getByRole('menuitem', { name: 'Change Picture', exact: true }).count(),
            0,
          );
          await page.keyboard.press('Escape');
          await page
            .locator(`[data-shape-id="${beforeReplacement.shape.id}"].shape-hit`)
            .click({ button: 'right' });
          await page.getByRole('menuitem', { name: 'Format Picture...', exact: true }).click();
          await page.getByRole('tab', { name: 'Picture', exact: true }).waitFor();
          await page.getByRole('button', { name: 'Close Format Picture' }).click();
          await page
            .locator(`[data-shape-id="${beforeReplacement.shape.id}"].shape-hit`)
            .click({ button: 'right' });
          await page.getByRole('menuitem', { name: 'Change Picture', exact: true }).focus();
          await page.keyboard.press('ArrowRight');
        }
        const replacementChooser = page.waitForEvent('filechooser');
        await page.getByRole('menuitem', { name: 'Picture from File...' }).click();
        await (
          await replacementChooser
        ).setFiles({
          name: 'replacement.gif',
          mimeType: 'image/gif',
          buffer: Buffer.from('R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7', 'base64'),
        });
        await page.waitForFunction(
          (revision) =>
            !window.office.getState().building && window.office.getState().revision > revision,
          beforeReplacement.revision,
        );
        assert.deepEqual(
          await page.evaluate(() => window.office.getState().editor.slides[0].shapes.at(-1)),
          beforeReplacement.shape,
        );
        assert.match(
          await page
            .locator('#slide')
            .evaluate((host) => host.shadowRoot.querySelector('image').getAttribute('href')),
          /^data:image\/gif/,
        );
        await page.locator('[data-edit="undo"]:enabled').click();
        await page.waitForFunction(
          () =>
            !window.office.getState().building &&
            document
              .querySelector('#slide')
              .shadowRoot.querySelector('image')
              .getAttribute('href')
              .startsWith('data:image/png'),
        );
      }
      const pictureId = await page.evaluate(
        () =>
          window.office.getState().editor.slides[0].shapes.find((shape) => shape.kind === 'picture')
            .id,
      );
      await page.locator(`[data-shape-id="${pictureId}"].shape-hit`).click();
      await page.getByRole('tab', { name: 'Picture Format', exact: true }).click();
      const beforeHover = await page.evaluate(() => ({
        revision: window.office.getState().revision,
        editor: window.office.getState().editor,
      }));
      await corrections.click();
      await page
        .getByRole('menuitem', { name: 'Brightness: -40%, Contrast: 0%', exact: true })
        .hover();
      await page.locator('.picture-menu-preview').waitFor();
      assert.deepEqual(
        await page.evaluate(() => ({
          revision: window.office.getState().revision,
          editor: window.office.getState().editor,
        })),
        beforeHover,
      );
      await page.mouse.move(900, 800);
      assert.equal(await page.locator('.picture-menu-preview').count(), 0);
      await page
        .getByRole('menuitem', { name: 'Brightness: -40%, Contrast: 0%', exact: true })
        .hover();
      await page.locator('.picture-menu-preview').waitFor();
      await page.keyboard.press('Escape');
      assert.equal(await page.locator('.picture-menu-preview').count(), 0);
      assert.equal(
        await page
          .locator('#slide')
          .evaluate((host) => host.shadowRoot.querySelector('svg').style.visibility),
        '',
      );
      let releasePreview;
      const previewGate = new Promise((resolve) => {
        releasePreview = resolve;
      });
      let previewFinished;
      const previewDone = new Promise((resolve) => {
        previewFinished = resolve;
      });
      let previewArrived;
      const previewPending = new Promise((resolve) => {
        previewArrived = resolve;
      });
      const delayedPreview = async (route) => {
        if (!route.request().postDataJSON()?.previewRender) return route.continue();
        previewArrived();
        await previewGate;
        await route
          .fulfill({
            status: 200,
            contentType: 'application/json',
            body: JSON.stringify({
              svg: '<svg xmlns="http://www.w3.org/2000/svg"><rect width="10" height="10"/></svg>',
            }),
          })
          .catch(() => {});
        previewFinished();
      };
      await page.route('**/edit', delayedPreview);
      await corrections.click();
      await previewPending;
      await page.keyboard.press('Escape');
      releasePreview();
      await previewDone;
      await page.unroute('**/edit', delayedPreview);
      assert.equal(await page.locator('.picture-menu-preview').count(), 0);
      await corrections.click();
      assert.equal(await page.locator('.correction-gallery svg').count(), 25);
      const firstCorrection = page.getByRole('menuitem', {
        name: 'Brightness: -40%, Contrast: -40%',
        exact: true,
      });
      await firstCorrection.focus();
      await firstCorrection.press('ArrowDown');
      assert.equal(
        await page.evaluate(() => document.activeElement.getAttribute('aria-label')),
        'Brightness: -40%, Contrast: -20%',
      );
      await page.screenshot({ path: '/tmp/pptx-picture-corrections-gallery.png' });
      await page
        .getByRole('menuitem', { name: 'Brightness: 20%, Contrast: -20%', exact: true })
        .click();
      await page.waitForFunction(
        () =>
          !window.office.getState().building &&
          window.office.getState().editor.slides[0].shapes.at(-1).imageBrightness === 0.2 &&
          window.office.getState().editor.slides[0].shapes.at(-1).imageContrast === -0.2,
      );
      await corrections.and(page.locator(':enabled')).waitFor();
      await page.locator('[data-edit="undo"]').click();
      await page.waitForFunction(
        () =>
          !window.office.getState().building &&
          window.office.getState().editor.slides[0].shapes.at(-1).imageBrightness === null &&
          window.office.getState().editor.slides[0].shapes.at(-1).imageContrast === null,
      );
      await corrections.and(page.locator(':enabled')).waitFor();
      await corrections.click();
      await page
        .getByRole('menuitem', { name: 'Picture Corrections Options...', exact: true })
        .click();
      assert.equal(
        await page
          .getByRole('spinbutton', { name: 'Picture brightness', exact: true })
          .inputValue(),
        '0',
      );
      await page.getByRole('button', { name: 'Close Format Picture', exact: true }).click();

      await transparency.click();
      assert.equal(await page.locator('.picture-gallery img').count(), 7);
      const beforeTransparencyPreview = await page.evaluate(() => ({
        revision: window.office.getState().revision,
        editor: window.office.getState().editor,
      }));
      await page.getByRole('menuitem', { name: '50%', exact: true }).hover();
      await page.locator('.picture-menu-preview image[opacity="0.500"]').waitFor();
      assert.deepEqual(
        await page.evaluate(() => ({
          revision: window.office.getState().revision,
          editor: window.office.getState().editor,
        })),
        beforeTransparencyPreview,
      );
      await page.keyboard.press('ArrowLeft');
      await page.locator('.picture-menu-preview image[opacity="0.700"]').waitFor();
      await page.keyboard.press('Escape');
      assert.equal(await page.locator('.picture-menu-preview').count(), 0);
      assert.equal(
        await page
          .locator('#slide')
          .evaluate((host) => host.shadowRoot.querySelector('svg').style.visibility),
        '',
      );
      await transparency.click();
      await page.screenshot({ path: '/tmp/pptx-picture-transparency-gallery.png' });
      await page.getByRole('menuitem', { name: '50%', exact: true }).click();
      await page.waitForFunction(
        () =>
          !window.office.getState().building &&
          window.office.getState().editor.slides[0].shapes.at(-1).imageOpacity === 0.5,
      );
      await transparency.and(page.locator(':enabled')).waitFor();
      assert.equal(await page.locator('.picture-menu-preview').count(), 0);
      await transparency.click();
      await page.getByRole('menuitem', { name: 'Picture Transparency Options...' }).click();
      const opacityInput = page.getByRole('spinbutton', {
        name: 'Picture transparency',
        exact: true,
      });
      assert.equal(await opacityInput.inputValue(), '50');
      await page.screenshot({ path: '/tmp/pptx-picture-transparency-pane.png' });
      await opacityInput.fill('27');
      await opacityInput.press('Tab');
      await page.waitForFunction(
        () =>
          !window.office.getState().building &&
          window.office.getState().editor.slides[0].shapes.at(-1).imageOpacity === 0.73,
      );
      await transparency.and(page.locator(':enabled')).waitFor();
      await page.locator('[data-edit="undo"]').click();
      await page.waitForFunction(
        () =>
          !window.office.getState().building &&
          window.office.getState().editor.slides[0].shapes.at(-1).imageOpacity === 0.5,
      );
      await transparency.and(page.locator(':enabled')).waitFor();
      await page.locator('[data-edit="undo"]').click();
      await page.waitForFunction(
        () =>
          !window.office.getState().building &&
          window.office.getState().editor.slides[0].shapes.at(-1).imageOpacity === null,
      );
      const opacitySlider = page.getByRole('slider', { name: 'Picture transparency slider' });
      await opacitySlider.and(page.locator(':enabled')).waitFor();
      await opacitySlider.focus();
      await opacitySlider.press('ArrowRight');
      await page.waitForFunction(
        () =>
          !window.office.getState().building &&
          window.office.getState().editor.slides[0].shapes.at(-1).imageOpacity === 0.99,
      );
      await transparency.and(page.locator(':enabled')).waitFor();
      assert.equal(await opacityInput.inputValue(), '1');
      await page.locator('[data-edit="undo"]').click();
      await page.waitForFunction(
        () =>
          !window.office.getState().building &&
          window.office.getState().editor.slides[0].shapes.at(-1).imageOpacity === null,
      );
      for (const [name, key, value] of [
        ['brightness', 'imageBrightness', -25],
        ['contrast', 'imageContrast', 40],
      ]) {
        const input = page.getByRole('spinbutton', { name: `Picture ${name}`, exact: true });
        await input.and(page.locator(':enabled')).waitFor();
        await input.fill(String(value));
        await input.press('Tab');
        await page.waitForFunction(
          ({ key, value }) =>
            !window.office.getState().building &&
            window.office.getState().editor.slides[0].shapes.at(-1)[key] === value / 100,
          { key, value },
        );
        await transparency.and(page.locator(':enabled')).waitFor();
        assert.equal(
          await page.getByRole('slider', { name: `Picture ${name} slider` }).inputValue(),
          String(value),
        );
        await page.locator('[data-edit="undo"]').click();
        await page.waitForFunction(
          (key) =>
            !window.office.getState().building &&
            window.office.getState().editor.slides[0].shapes.at(-1)[key] === null,
          key,
        );
      }
      await page.getByRole('button', { name: 'Close Format Picture', exact: true }).click();
      const crop = page.locator('[data-edit="picture-crop-menu"]');
      const cropButton = page.getByRole('button', { name: 'Crop', exact: true });
      await cropButton.and(page.locator(':enabled')).waitFor();
      const beforeMode = await page.evaluate(() =>
        window.office.getState().editor.slides[0].shapes.at(-1),
      );
      await cropButton.click();
      assert.equal(await page.locator('.crop-handle').count(), 8);
      assert.equal(await cropButton.getAttribute('aria-pressed'), 'true');
      assert.equal(await page.getByRole('menu').count(), 0);
      await cropButton.click();
      assert.equal(await page.locator('.crop-handle').count(), 0);
      assert.equal(await cropButton.getAttribute('aria-pressed'), 'false');
      assert.deepEqual(
        await page.evaluate(() => window.office.getState().editor.slides[0].shapes.at(-1)),
        beforeMode,
      );
      await cropButton.focus();
      await page.keyboard.press('Space');
      assert.equal(await page.locator('.crop-handle').count(), 8);
      await page.keyboard.press('Escape');
      assert.equal(await cropButton.getAttribute('aria-pressed'), 'false');
      await crop.and(page.locator(':enabled')).waitFor();
      await crop.click();
      await page.getByRole('menuitem', { name: 'Crop to Shape', exact: true }).click();
      const gallery = page.getByRole('menu', { name: 'Crop to Shape', exact: true });
      await gallery.getByRole('menuitem', { name: 'Oval', exact: true }).click();
      await page.waitForFunction(
        () =>
          !window.office.getState().building &&
          window.office.getState().editor.slides[0].shapes.at(-1).preset === 'ellipse',
      );
      await crop.and(page.locator(':enabled')).waitFor();
      await crop.click();
      await page.keyboard.press('ArrowDown');
      await page.keyboard.press('ArrowRight');
      const first = gallery.getByRole('menuitem').first();
      await first.waitFor();
      const firstName = await page.evaluate(() =>
        document.activeElement.getAttribute('aria-label'),
      );
      await page.keyboard.press('ArrowRight');
      assert.notEqual(
        await page.evaluate(() => document.activeElement.getAttribute('aria-label')),
        firstName,
      );
      await page.keyboard.press('ArrowLeft');
      assert.equal(
        await page.evaluate(() => document.activeElement.getAttribute('aria-label')),
        firstName,
      );
      await page.keyboard.press('Escape');
      assert.equal(await gallery.count(), 0);
      await page.keyboard.press('Escape');
      assert.equal(await crop.evaluate((el) => el === document.activeElement), true);
      await page.locator('[data-edit="undo"]').click();
      await page.waitForFunction(
        () =>
          !window.office.getState().building &&
          window.office.getState().editor.slides[0].shapes.at(-1).preset === 'rect',
      );
      await crop.and(page.locator(':enabled')).waitFor();

      await page.evaluate(async () => {
        const state = window.office.getState(),
          shape = state.editor.slides[0].shapes.at(-1);
        const response = await fetch('/edit', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            revision: state.revision,
            command: {
              type: 'update',
              slide: 0,
              ids: [shape.id],
              changes: { bounds: { ...shape.bounds, w: shape.bounds.h * 2 } },
            },
          }),
        });
        if (!response.ok) throw new Error(await response.text());
      });
      await page.waitForFunction(() => {
        const state = window.office.getState(),
          bounds = state.editor.slides[0].shapes.at(-1).bounds;
        return !state.building && bounds.w === bounds.h * 2;
      });
      for (const mode of ['Fill', 'Fit']) {
        await crop.and(page.locator(':enabled')).waitFor();
        await crop.click();
        await page.getByRole('menuitem', { name: mode, exact: true }).click();
        await page.waitForFunction((mode) => {
          const state = window.office.getState(),
            crop = state.editor.slides[0].shapes.at(-1).imageCrop;
          return (
            !state.building && crop && (mode === 'Fill' ? crop.top === 0.25 : crop.left === -0.5)
          );
        }, mode);
        await page.locator('.crop-handle[data-handle="se"]').waitFor();
        assert.equal(await cropButton.getAttribute('aria-pressed'), 'true');
        assert.equal(await page.locator('.picture-source-handle').count(), 4);
        await page.keyboard.press('Escape');
        assert.equal(await page.locator('.crop-handle').count(), 0);
        assert.equal(await cropButton.getAttribute('aria-pressed'), 'false');
      }
      await crop.and(page.locator(':enabled')).waitFor();
      await crop.click();
      await page.getByRole('menuitem', { name: 'Aspect Ratio', exact: true }).click();
      const ratios = page.getByRole('menu', { name: 'Aspect Ratio', exact: true });
      assert.equal(await ratios.evaluate((el) => el.classList.contains('shape-gallery')), false);
      await ratios.getByRole('menuitem', { name: '1:1', exact: true }).click();
      await page.waitForFunction(() => {
        const state = window.office.getState(),
          b = state.editor.slides[0].shapes.at(-1).bounds;
        return !state.building && b.w === b.h;
      });
      await page.locator('.crop-handle[data-handle="se"]').waitFor();
      await crop.and(page.locator(':enabled')).waitFor();
      const ratioBefore = await page.evaluate(
        () => window.office.getState().editor.slides[0].shapes.at(-1).bounds,
      );
      const corner = await page.locator('.crop-handle[data-handle="se"]').boundingBox();
      await page.mouse.move(corner.x + corner.width / 2, corner.y + corner.height / 2);
      await page.mouse.down();
      await page.mouse.move(corner.x + corner.width / 2 - 30, corner.y + corner.height / 2 - 5, {
        steps: 4,
      });
      await page.mouse.up();
      await page.waitForFunction((w) => {
        const state = window.office.getState(),
          b = state.editor.slides[0].shapes.at(-1).bounds;
        return !state.building && b.w < w && b.w === b.h;
      }, ratioBefore.w);
      await crop.and(page.locator(':enabled')).waitFor();
      await page.locator('[data-edit="undo"]').click();
      await page.waitForFunction(
        (w) => window.office.getState().editor.slides[0].shapes.at(-1).bounds.w === w,
        ratioBefore.w,
      );
      await crop.and(page.locator(':enabled')).waitFor();
      await crop.click();
      await page.getByRole('menuitem', { name: 'Crop', exact: true }).click();
      const east = page.locator('.crop-handle[data-handle="e"]');
      await east.waitFor();
      const beforeCrop = await page.evaluate(() =>
        window.office.getState().editor.slides[0].shapes.at(-1),
      );
      const frame = await page.locator('.picture-cropping').boundingBox();
      const edge = await east.boundingBox();
      await page.mouse.move(edge.x + edge.width / 2, edge.y + edge.height / 2);
      await page.mouse.down();
      await page.mouse.move(edge.x + edge.width / 2 - frame.width / 4, edge.y + edge.height / 2, {
        steps: 5,
      });
      await page.mouse.up();
      await page.waitForFunction((w) => {
        const state = window.office.getState(),
          shape = state.editor.slides[0].shapes.at(-1);
        return !state.building && shape.bounds.w < w && shape.imageCrop.right > 0.2;
      }, beforeCrop.bounds.w);
      await crop.and(page.locator(':enabled')).waitFor();
      const committed = await page.evaluate(() =>
        window.office.getState().editor.slides[0].shapes.at(-1),
      );
      const sourceOverlay = page.locator('.picture-crop-source');
      assert.equal(await sourceOverlay.count(), 1);
      const overlayGeometry = await sourceOverlay.evaluate((svg) => ({
        frame: svg.viewBox.baseVal.width,
        source: Number(svg.querySelector('image').getAttribute('width')),
        mask: svg.querySelector('mask rect[fill="black"]').getAttribute('width'),
      }));
      assert.ok(overlayGeometry.source > overlayGeometry.frame);
      assert.equal(Number(overlayGeometry.mask), overlayGeometry.frame);
      const cropScreenshot = await page.screenshot({ path: '/tmp/pptx-crop-source-overlay.png' });
      const cropBox = await page.locator('.picture-cropping').boundingBox();
      const pixels = await page.evaluate(
        async ({ png, box }) => {
          const image = new Image();
          image.src = png;
          await image.decode();
          const canvas = document.createElement('canvas');
          canvas.width = image.width;
          canvas.height = image.height;
          const context = canvas.getContext('2d');
          context.drawImage(image, 0, 0);
          const pixel = (x) =>
            Array.from(
              context.getImageData(Math.round(x), Math.round(box.y + box.height / 2), 1, 1).data,
            );
          return { inside: pixel(box.x + box.width / 2), outside: pixel(box.x + box.width + 20) };
        },
        { png: 'data:image/png;base64,' + cropScreenshot.toString('base64'), box: cropBox },
      );
      assert.deepEqual(pixels.inside, [255, 255, 255, 255]);
      assert.ok(pixels.outside.slice(0, 3).every((channel) => channel >= 110 && channel <= 120));
      assert.equal(pixels.outside[3], 255);
      assert.equal(await page.locator('.picture-source-handle').count(), 4);
      const sourceCorner = await page.locator('[data-handle="source-se"]').boundingBox();
      await page.mouse.move(
        sourceCorner.x + sourceCorner.width / 2,
        sourceCorner.y + sourceCorner.height / 2,
      );
      await page.mouse.down();
      await page.mouse.move(
        sourceCorner.x + sourceCorner.width / 2 + 25,
        sourceCorner.y + sourceCorner.height / 2 + 25,
        { steps: 3 },
      );
      await page.mouse.up();
      await page.waitForFunction((right) => {
        const state = window.office.getState();
        return !state.building && state.editor.slides[0].shapes.at(-1).imageCrop.right > right;
      }, committed.imageCrop.right);
      assert.deepEqual(
        await page.evaluate(() => window.office.getState().editor.slides[0].shapes.at(-1).bounds),
        committed.bounds,
      );
      await page.locator('[data-edit="undo"]').click();
      await page.waitForFunction((right) => {
        const state = window.office.getState();
        return !state.building && state.editor.slides[0].shapes.at(-1).imageCrop.right === right;
      }, committed.imageCrop.right);
      await crop.and(page.locator(':enabled')).waitFor();
      const panFrame = await page.locator('.picture-cropping').boundingBox();
      await page.mouse.move(panFrame.x + panFrame.width / 2, panFrame.y + panFrame.height / 2);
      await page.mouse.down();
      await page.mouse.move(
        panFrame.x + panFrame.width / 2 + 15,
        panFrame.y + panFrame.height / 2,
        { steps: 3 },
      );
      await page.mouse.up();
      await page.waitForFunction((left) => {
        const state = window.office.getState(),
          pic = state.editor.slides[0].shapes.at(-1);
        return !state.building && pic.imageCrop.left < left;
      }, committed.imageCrop.left);
      await crop.and(page.locator(':enabled')).waitFor();
      assert.deepEqual(
        await page.evaluate(() => window.office.getState().editor.slides[0].shapes.at(-1).bounds),
        committed.bounds,
      );
      await page.locator('[data-edit="undo"]').click();
      await page.waitForFunction(
        (left) => window.office.getState().editor.slides[0].shapes.at(-1).imageCrop.left === left,
        committed.imageCrop.left,
      );
      await crop.and(page.locator(':enabled')).waitFor();
      const nextEdge = await east.boundingBox();
      await page.mouse.move(nextEdge.x + nextEdge.width / 2, nextEdge.y + nextEdge.height / 2);
      await page.mouse.down();
      await page.mouse.move(nextEdge.x - 20, nextEdge.y, { steps: 3 });
      await page.keyboard.press('Escape');
      await page.mouse.up();
      assert.deepEqual(
        await page.evaluate(() => window.office.getState().editor.slides[0].shapes.at(-1)),
        committed,
      );
      await page.keyboard.press('Escape');
      assert.equal(await page.locator('.crop-handle').count(), 0);
      assert.equal(await page.locator('.picture-crop-source').count(), 0);
      await page.locator('[data-edit="undo"]').click();
      await page.waitForFunction(
        (w) => window.office.getState().editor.slides[0].shapes.at(-1).bounds.w === w,
        beforeCrop.bounds.w,
      );
      await crop.and(page.locator(':enabled')).waitFor();
      await page
        .locator('.shape-hit')
        .first()
        .click({ position: { x: 5, y: 5 } });
      await page.locator('#tab-picture-format').waitFor({ state: 'hidden' });
      assert.equal(await page.getByRole('tab', { name: 'Picture Format' }).count(), 0);
      assert.equal(
        await page.getByRole('tab', { name: 'Home', exact: true }).getAttribute('aria-selected'),
        'true',
      );
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      proc.kill('SIGTERM');
      await new Promise((resolve) => {
        if (proc.exitCode !== null) resolve();
        else proc.once('exit', resolve);
      });
      await rm(dir, { recursive: true, force: true });
    }
  },
);
