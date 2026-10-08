import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  addBlankSlide,
  addSlideMedia,
  createPresentation,
  getShapeBoundsResolved,
  isShapeAspectRatioLocked,
  getSlides,
  getSlideShapes,
  inches,
  loadPresentation,
  savePresentation,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

const makeVideo = (page) =>
  page.evaluate(async () => {
    const canvas = document.createElement('canvas');
    canvas.width = 160;
    canvas.height = 90;
    const context = canvas.getContext('2d');
    const stream = canvas.captureStream(20);
    const recorder = new MediaRecorder(stream, { mimeType: 'video/webm;codecs=vp8' });
    const chunks = [];
    const stopped = new Promise((resolve) => (recorder.onstop = resolve));
    recorder.ondataavailable = (event) => chunks.push(event.data);
    recorder.start();
    let frame = 0;
    const timer = setInterval(() => {
      context.fillStyle = frame++ % 2 ? '#2d9cdb' : '#eb5757';
      context.fillRect(0, 0, canvas.width, canvas.height);
    }, 50);
    await new Promise((resolve) => setTimeout(resolve, 700));
    recorder.stop();
    await stopped;
    clearInterval(timer);
    stream.getTracks().forEach((track) => track.stop());
    return Array.from(new Uint8Array(await new Blob(chunks).arrayBuffer()));
  });

test(
  'Video Format ribbon plays the selected video and exposes Arrange and Format Pane',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-video-format-ribbon-'));
    let browser;
    let preview;
    try {
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1500, height: 900 } });
      const media = Uint8Array.from(await makeVideo(page));
      const deck = createPresentation();
      const slide = addBlankSlide(deck);
      addSlideMedia(slide, {
        kind: 'video',
        data: media,
        x: inches(1),
        y: inches(1),
        w: inches(6),
        h: inches(3.375),
      });
      const source = join(dir, 'source.pptx');
      await writeFile(source, await savePresentation(deck));
      const entry = join(dir, 'deck.tsx');
      await writeFile(
        entry,
        `import {readFile} from 'node:fs/promises';import {Presentation} from '@office-kit/pptx-dsl';export default <Presentation source={await readFile(${JSON.stringify(source)})} />;`,
      );
      preview = await startPreview(entry);
      await page.goto(preview.url);
      const editor = page.frameLocator('#editor-frame');
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      await editor.locator('.hit').first().click();
      await editor.getByRole('tab', { name: 'Video Format', exact: true }).click();
      const panel = editor.locator('#ribbon-panel');
      const video = editor.locator('.media-preview video');
      await video.waitFor({ state: 'attached' });
      await video.evaluate(
        (element) =>
          new Promise((resolve, reject) => {
            if (element.readyState >= 2) return resolve();
            const timer = setTimeout(() => reject(new Error('video did not decode')), 5000);
            element.addEventListener(
              'loadeddata',
              () => {
                clearTimeout(timer);
                resolve();
              },
              { once: true },
            );
          }),
      );

      // The reference desktop app keeps media transport on Video Format. Verify the command
      // controls the actual selected <video>, rather than a visual-only state.
      const play = panel.getByRole('button', { name: 'Play', exact: true });
      assert.equal(await play.count(), 1);
      await play.click();
      await video.evaluate(
        (element) =>
          new Promise((resolve, reject) => {
            const deadline = Date.now() + 3000;
            const check = () => {
              if (!element.paused && element.currentTime > 0.1) resolve();
              else if (Date.now() >= deadline)
                reject(new Error('Video Format Play did not advance'));
              else setTimeout(check, 20);
            };
            check();
          }),
      );
      const pause = panel.getByRole('button', { name: 'Pause', exact: true });
      assert.equal(await pause.count(), 1);
      await pause.click();
      assert.equal(await video.evaluate((element) => element.paused), true);

      // This intermediate ribbon groups the existing arrange commands under
      // one menu while keeping the same Selection Pane action available.
      await panel.getByRole('button', { name: 'Arrange', exact: true }).click();
      const arrange = editor.getByRole('menu', { name: 'Arrange', exact: true });
      const selectionPane = arrange.getByRole('menuitemcheckbox', { name: /Selection Pane/ });
      await selectionPane.click();
      await editor.getByRole('region', { name: 'Selection Pane', exact: true }).waitFor();

      // Video Effects is a fixed-position menu. Its keyboard loop must remain
      // usable even when the trigger is near the right edge of the ribbon.
      await editor.getByRole('button', { name: 'Close Selection Pane', exact: true }).click();
      const effectsTrigger = panel.getByRole('button', { name: /Video Effects/ });
      await effectsTrigger.click();
      const effects = editor.getByRole('menu', { name: 'Video Effects', exact: true });
      await effects.waitFor();
      const effectsBox = await effects.boundingBox();
      assert.ok(effectsBox);
      assert.ok(effectsBox.x >= 0 && effectsBox.y >= 0);
      assert.ok(effectsBox.x + effectsBox.width <= 1500);
      assert.ok(effectsBox.y + effectsBox.height <= 900);
      const firstEffect = effects.getByRole('menuitem').first();
      await firstEffect.focus();
      await firstEffect.press('ArrowDown');
      assert.equal(
        await effects
          .getByRole('menuitem')
          .nth(1)
          .evaluate((node) => node === node.ownerDocument.activeElement),
        true,
      );
      await effects.getByRole('menuitem').nth(1).press('Escape');
      assert.equal(await effects.count(), 0);

      // The aspect-ratio lock is persisted in DrawingML's noChangeAspect
      // attribute. An unlocked resize changes width only, and undoing that
      // resize must leave the unlocked state intact.
      const readBounds = async () => {
        const saved = await loadPresentation(
          new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
        );
        return getShapeBoundsResolved(saved, getSlideShapes(getSlides(saved)[0])[0]);
      };
      const readAspectLock = async () => {
        const saved = await loadPresentation(
          new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
        );
        return isShapeAspectRatioLocked(getSlideShapes(getSlides(saved)[0])[0]);
      };
      const initialBounds = await readBounds();
      assert.ok(initialBounds);
      const lock = panel.getByLabel('Lock aspect ratio', { exact: true });
      assert.equal(await lock.isChecked(), true);
      assert.equal(await readAspectLock(), true);
      const initialWidthCm = Number(
        await panel.locator('input[type="number"]').nth(1).inputValue(),
      );
      const initialHeightCm = Number(
        await panel.locator('input[type="number"]').nth(0).inputValue(),
      );
      await lock.uncheck();
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      assert.equal(await readAspectLock(), false);
      // The reference desktop app does not add a history entry for changing this checkbox.
      assert.equal(await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).isDisabled(), true);

      // Reloading must hydrate the checkbox from the saved OOXML value.
      await page.reload();
      const reloadedEditor = page.frameLocator('#editor-frame');
      await reloadedEditor.getByText('Saved to this project', { exact: true }).waitFor();
      await reloadedEditor.locator('.hit').first().click();
      await reloadedEditor.getByRole('tab', { name: 'Video Format', exact: true }).click();
      const reloadedPanel = reloadedEditor.locator('#ribbon-panel');
      const reloadedLock = reloadedPanel.getByLabel('Lock aspect ratio', { exact: true });
      assert.equal(await reloadedLock.isChecked(), false);
      assert.equal(await readAspectLock(), false);

      const unlockedInitialBounds = await readBounds();
      assert.deepEqual(unlockedInitialBounds, initialBounds);
      const reloadedHeight = reloadedPanel.locator('input[type="number"]').nth(0);
      const reloadedWidth = reloadedPanel.locator('input[type="number"]').nth(1);
      await reloadedWidth.fill(String(initialWidthCm + 1));
      await reloadedWidth.press('Tab');
      await reloadedEditor.getByText('Saved to this project', { exact: true }).waitFor();
      const resizedBounds = await readBounds();
      assert.ok(resizedBounds);
      assert.ok(resizedBounds.w > initialBounds.w);
      assert.equal(resizedBounds.h, initialBounds.h);
      assert.ok(Math.abs(Number(await reloadedHeight.inputValue()) - initialHeightCm) < 0.01);
      assert.ok(Math.abs(Number(await reloadedWidth.inputValue()) - (initialWidthCm + 1)) < 0.01);
      await reloadedEditor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await reloadedEditor.getByText('Saved to this project', { exact: true }).waitFor();
      const undoneBounds = await readBounds();
      assert.deepEqual(undoneBounds, initialBounds);
      assert.equal(await reloadedLock.isChecked(), false);
      assert.equal(await readAspectLock(), false);

      // Re-enabling the lock restores proportional resizing coverage. The
      // subsequent Undo is for the resize only, so the lock remains enabled.
      await reloadedLock.check();
      await reloadedEditor.getByText('Saved to this project', { exact: true }).waitFor();
      assert.equal(await readAspectLock(), true);
      const lockedInitialBounds = await readBounds();
      const lockedInitialWidthCm = Number(await reloadedWidth.inputValue());
      await reloadedWidth.fill(String(lockedInitialWidthCm + 1));
      await reloadedWidth.press('Tab');
      await reloadedEditor.getByText('Saved to this project', { exact: true }).waitFor();
      const lockedResizedBounds = await readBounds();
      assert.ok(lockedResizedBounds.w > lockedInitialBounds.w);
      assert.ok(lockedResizedBounds.h > lockedInitialBounds.h);
      await reloadedEditor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await reloadedEditor.getByText('Saved to this project', { exact: true }).waitFor();
      assert.deepEqual(await readBounds(), lockedInitialBounds);
      assert.equal(await reloadedLock.isChecked(), true);
      assert.equal(await readAspectLock(), true);

      // Format Pane opens the properties pane for the selected media shape.
      await reloadedPanel.getByRole('button', { name: 'Format Pane', exact: true }).click();
      await reloadedEditor.locator('#format-panel').waitFor({ state: 'visible' });
      const sizePaneLock = reloadedEditor
        .locator('#format-panel')
        .getByLabel('Lock aspect ratio', { exact: true });
      if (await sizePaneLock.count()) assert.equal(await sizePaneLock.isChecked(), true);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
