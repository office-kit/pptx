import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  getShapeGradientFill,
  getSlideShapes,
  getSlides,
  loadPresentation,
} from '../../../../dist/index.js';
import { startPreview } from '../helpers/server.mjs';

test(
  'gradient stops edit brightness, opacity, position and count with undo and persistence',
  { timeout: 180000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-gradient-'));
    let preview, browser;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={3} height={2} fill={{stops:[{offset:0,color:'accent1',brightness:0.95},{offset:1,color:'accent1',brightness:0.7}],angleDeg:90,scaled:false}}>Gradient</Text><Text x={6} y={1} width={3} height={2} fill="#00FF00">Solid</Text></Slide></Presentation>`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      const errors = [];
      page.on('pageerror', (error) => errors.push(error.message));
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      const saved = () => editor.getByText('Saved to this project', { exact: true }).waitFor();
      const gradient = async (index = 0) => {
        const deck = await loadPresentation(
          new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
        );
        return getShapeGradientFill(getSlideShapes(getSlides(deck)[0])[index]);
      };
      await saved();
      await editor.locator('.hit').nth(0).click();
      let initialGradient = await gradient();
      const direction = editor.getByRole('button', { name: 'Gradient direction', exact: true });
      await direction.click();
      const directionMenu = editor.getByRole('menu', { name: 'Gradient direction', exact: true });
      const directions = [
        ['Linear Diagonal - Top Left to Bottom Right', 45],
        ['Linear Down', 90],
        ['Linear Diagonal - Top Right to Bottom Left', 135],
        ['Linear Right', 0],
        ['Linear Left', 180],
        ['Linear Diagonal - Bottom Left to Top Right', 315],
        ['Linear Up', 270],
        ['Linear Diagonal - Bottom Right to Top Left', 225],
      ];
      assert.deepEqual(
        await directionMenu
          .getByRole('menuitemradio')
          .evaluateAll((items) => items.map((item) => item.getAttribute('aria-label'))),
        directions.map(([label]) => label),
      );
      await directionMenu.press('Escape');
      assert.deepEqual(await gradient(), initialGradient);
      for (const [label, angle] of directions) {
        await direction.click();
        await directionMenu.getByRole('menuitemradio', { name: label, exact: true }).click();
        await saved();
        const changed = await gradient();
        assert.equal(changed.angleDeg, angle);
        assert.equal(changed.scaled, true);
        assert.deepEqual(changed.stops, initialGradient.stops);
        if (angle === 45) {
          await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
          await saved();
          assert.deepEqual(await gradient(), initialGradient);
          await editor.getByTitle('Redo (Ctrl+Y)', { exact: true }).click();
          await saved();
          assert.equal((await gradient()).scaled, true);
        }
      }
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.equal((await gradient()).angleDeg, 270);
      await editor.getByTitle('Redo (Ctrl+Y)', { exact: true }).click();
      await saved();
      assert.equal((await gradient()).angleDeg, 225);
      await direction.click();
      await page.keyboard.press('Home');
      await page.keyboard.press('ArrowRight');
      await page.keyboard.press('Enter');
      await saved();
      initialGradient = { ...initialGradient, scaled: true };
      assert.deepEqual(await gradient(), initialGradient);
      const noFill = editor.getByRole('radio', { name: 'No fill', exact: true });
      const solidFill = editor.getByRole('radio', { name: 'Solid fill', exact: true });
      const gradientFill = editor.getByRole('radio', { name: 'Gradient fill', exact: true });
      assert.equal(await gradientFill.isChecked(), true);
      await noFill.check();
      await saved();
      assert.equal(await gradient(), null);
      assert.equal(
        await editor.getByRole('spinbutton', { name: 'Gradient angle', exact: true }).count(),
        0,
      );
      await solidFill.check();
      await saved();
      assert.equal(
        await editor
          .getByRole('spinbutton', { name: 'Fill transparency', exact: true })
          .isVisible(),
        true,
      );
      await editor.getByRole('button', { name: 'Close Format Shape', exact: true }).click();
      await editor.locator('.hit').nth(0).click({ button: 'right' });
      await editor.getByRole('menuitem', { name: 'Format Shape...', exact: true }).click();
      await gradientFill.check();
      await saved();
      assert.deepEqual(await gradient(), initialGradient);
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.equal(await solidFill.isChecked(), true);
      await editor.getByTitle('Redo (Ctrl+Y)', { exact: true }).click();
      await saved();
      assert.deepEqual(await gradient(), initialGradient);
      await editor.locator('.hit').nth(1).click();
      assert.equal(await solidFill.isChecked(), true);
      await gradientFill.check();
      await saved();
      const createdGradient = await gradient(1);
      assert.deepEqual(
        createdGradient.stops.map((stop) => stop.offset),
        [0, 0.74, 0.83, 1],
      );
      assert.deepEqual(
        createdGradient.stops.map((stop) => stop.brightness),
        [0.95, 0.55, 0.55, 0.7],
      );
      assert.equal(createdGradient.angleDeg, 90);
      assert.equal(createdGradient.scaled, true);
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.equal(await solidFill.isChecked(), true);
      await editor.locator('.hit').nth(0).click();
      const brightness = editor.getByRole('spinbutton', {
        name: 'Gradient stop brightness',
        exact: true,
      });
      const opacity = editor.getByRole('spinbutton', {
        name: 'Gradient stop transparency',
        exact: true,
      });
      const position = editor.getByRole('spinbutton', {
        name: 'Gradient stop position',
        exact: true,
      });
      assert.equal(await brightness.inputValue(), '95');
      const paintedStop = editor.locator('.paint linearGradient stop').first();
      const initialColor = await paintedStop.getAttribute('stop-color');
      assert.equal(
        await editor
          .getByRole('button', { name: 'Remove gradient stop', exact: true })
          .isDisabled(),
        true,
      );
      await brightness.fill('-25');
      await brightness.press('Tab');
      await saved();
      assert.equal((await gradient()).stops[0].brightness, -0.25);
      assert.notEqual(await paintedStop.getAttribute('stop-color'), initialColor);
      await opacity.fill('60');
      await opacity.press('Tab');
      await saved();
      assert.equal((await gradient()).stops[0].opacity, 0.4);
      assert.equal(await paintedStop.getAttribute('stop-opacity'), '0.4');
      await editor.getByRole('button', { name: 'Gradient stop 2', exact: true }).click();
      assert.equal(await brightness.inputValue(), '70');
      await position.fill('80');
      await position.press('Tab');
      await saved();
      assert.equal((await gradient()).stops[1].offset, 0.8);
      assert.equal(await position.inputValue(), '80');
      await editor.getByRole('button', { name: 'Add gradient stop', exact: true }).click();
      await saved();
      assert.equal((await gradient()).stops.length, 3);
      assert.equal(await position.inputValue(), '90');
      assert.equal(await brightness.inputValue(), '0');
      assert.match((await gradient()).stops[2].color, /^#[0-9A-F]{6}$/i);
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.equal((await gradient()).stops.length, 2);
      await position.fill('100');
      await position.press('Tab');
      await saved();
      assert.equal((await gradient()).stops[1].offset, 1);
      await editor.getByRole('button', { name: 'Add gradient stop', exact: true }).click();
      await saved();
      assert.equal(await position.inputValue(), '50');
      assert.equal(await brightness.inputValue(), '0');
      assert.equal(await opacity.inputValue(), '30');
      await editor.getByRole('button', { name: 'Remove gradient stop', exact: true }).click();
      await saved();
      assert.equal((await gradient()).stops.length, 2);
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.equal((await gradient()).stops.length, 3);
      await page.reload();
      await saved();
      await editor.locator('.hit').nth(0).click();
      assert.equal(await brightness.inputValue(), '-25');
      assert.equal(await opacity.inputValue(), '60');
      await brightness.fill('101');
      await brightness.press('Tab');
      assert.equal(await brightness.inputValue(), '-25');
      const handle = editor.getByRole('button', { name: 'Gradient stop 1', exact: true });
      const bounds = await handle.boundingBox();
      const track = await editor
        .getByRole('group', { name: 'Gradient stops', exact: true })
        .boundingBox();
      assert.ok(bounds && track);
      await page.mouse.move(bounds.x + bounds.width / 2, bounds.y + bounds.height / 2);
      await page.mouse.down();
      await page.mouse.move(track.x + track.width / 4, bounds.y + bounds.height / 2, { steps: 5 });
      assert.equal(await position.inputValue(), '25');
      await page.mouse.up();
      await saved();
      assert.equal((await gradient()).stops[0].offset, 0.25);
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.equal((await gradient()).stops[0].offset, 0);
      await page.mouse.move(bounds.x + bounds.width / 2, bounds.y + bounds.height / 2);
      await page.mouse.down();
      await page.mouse.move(track.x + track.width / 3, bounds.y + bounds.height / 2, { steps: 5 });
      await page.keyboard.press('Escape');
      await page.mouse.up();
      assert.equal(await position.inputValue(), '0');
      assert.equal((await gradient()).stops[0].offset, 0);
      const type = editor.getByRole('combobox', { name: 'Gradient type', exact: true });
      await type.selectOption('circle');
      await saved();
      assert.equal((await gradient()).path, 'circle');
      assert.equal(
        await editor.getByRole('spinbutton', { name: 'Gradient angle', exact: true }).isDisabled(),
        true,
      );
      const radial = [
        ['From Bottom Right Corner', [1, 1, 0, 0], [0, 0, -1, -1]],
        ['From Bottom Left Corner', [0, 1, 1, 0], [-1, 0, 0, -1]],
        ['From Center', [0.5, 0.5, 0.5, 0.5], [0, 0, 0, 0]],
        ['From Top Right Corner', [1, 0, 0, 1], [0, -1, -1, 0]],
        ['From Top Left Corner', [0, 0, 1, 1], [-1, -1, 0, 0]],
      ];
      const radialInitial = await gradient();
      for (const path of ['circle', 'rect']) {
        await type.selectOption(path);
        await saved();
        for (const [label, focus, tile] of radial) {
          const before = await gradient();
          await direction.click();
          assert.deepEqual(
            await directionMenu
              .getByRole('menuitemradio')
              .evaluateAll((items) => items.map((item) => item.getAttribute('aria-label'))),
            radial.map(([name]) => name),
          );
          await directionMenu.getByRole('menuitemradio', { name: label, exact: true }).click();
          await saved();
          const changed = await gradient();
          const rect = ([left, top, right, bottom]) => ({ left, top, right, bottom });
          assert.equal(changed.path, path);
          assert.deepEqual(changed.focus, rect(focus));
          assert.deepEqual(changed.tileRect, rect(tile));
          assert.deepEqual(changed.stops, radialInitial.stops);
          assert.equal(changed.rotateWithShape, radialInitial.rotateWithShape);
          await direction.click();
          assert.equal(
            await directionMenu
              .getByRole('menuitemradio', { name: label, exact: true })
              .getAttribute('aria-checked'),
            'true',
          );
          await directionMenu.press('Escape');
          await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
          await saved();
          assert.deepEqual(await gradient(), before);
          await editor.getByTitle('Redo (Ctrl+Y)', { exact: true }).click();
          await saved();
          assert.deepEqual(await gradient(), changed);
        }
      }
      const radialSaved = await gradient();
      await page.reload();
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      await saved();
      await editor.locator('.hit').nth(0).click();
      assert.deepEqual(await gradient(), radialSaved);
      await direction.click();
      assert.equal(
        await directionMenu
          .getByRole('menuitemradio', { name: 'From Top Left Corner', exact: true })
          .getAttribute('aria-checked'),
        'true',
      );
      await directionMenu.press('Escape');
      await type.selectOption('shape');
      await saved();
      assert.equal(await direction.isDisabled(), true);
      assert.equal((await gradient()).path, 'shape');
      assert.deepEqual((await gradient()).focus, radialSaved.focus);
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.equal(await direction.isEnabled(), true);
      assert.deepEqual(await gradient(), radialSaved);
      await page.screenshot({ path: '/tmp/pptx-pr287-gradient-panel.png', fullPage: true });
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
