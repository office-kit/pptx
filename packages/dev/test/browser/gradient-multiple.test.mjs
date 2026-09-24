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
  'matching gradient stops edit together while retaining per-shape placement',
  { timeout: 120000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-gradient-multiple-'));
    let preview, browser;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';const stops=[{offset:0,color:'accent1'},{offset:1,color:'accent2'}];export default <Presentation><Slide><Text x={1} y={1} width={3} height={2} fill={{stops,angleDeg:0,scaled:false,rotateWithShape:true}}>First</Text><Text x={6} y={1} width={3} height={2} fill={{stops,angleDeg:90,scaled:true,rotateWithShape:false}}>Second</Text></Slide></Presentation>`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1500, height: 1100 } });
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      const saved = () => editor.getByText('Saved to this project', { exact: true }).waitFor();
      const read = async () => {
        const pres = await loadPresentation(
          new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
        );
        return getSlideShapes(getSlides(pres)[0]).map(getShapeGradientFill);
      };
      await saved();
      const initial = await read();
      await editor.locator('.hit').nth(0).click();
      await editor
        .locator('.hit')
        .nth(1)
        .click({ modifiers: ['Shift'] });
      assert.equal(
        await editor.getByRole('spinbutton', { name: 'Gradient angle', exact: true }).inputValue(),
        '',
      );
      const rotation = editor.getByRole('checkbox', { name: 'Rotate with shape', exact: true });
      assert.equal(await rotation.evaluate((input) => input.indeterminate), true);
      const transparency = editor.getByRole('spinbutton', {
        name: 'Gradient stop transparency',
        exact: true,
      });
      await transparency.fill('35');
      await transparency.press('Tab');
      await saved();
      const changed = initial.map((value) => ({
        ...value,
        stops: value.stops.map((stop, index) => (index === 0 ? { ...stop, opacity: 0.65 } : stop)),
      }));
      assert.deepEqual(await read(), changed);
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual(await read(), initial);
      await editor.getByTitle('Redo (Ctrl+Y)', { exact: true }).click();
      await saved();
      assert.deepEqual(await read(), changed);
      await page.reload();
      await saved();
      assert.deepEqual(await read(), changed);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
