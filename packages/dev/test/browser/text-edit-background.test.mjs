import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import { startPreview } from '../helpers/server.mjs';

test(
  'editing text preserves the slide background without duplicate painted text',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-text-background-'));
    let preview, browser;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide,Text,Table} from '@office-kit/pptx-dsl';export default <Presentation><Slide background="#15171C"><Text x={1} y={1} width={8} height={1} color="#B4B9C4" size={22}>Revenue and margin</Text><Table x={1} y={3} width={8} height={1} rows={[['First cell','Second cell']]} cellStyle={{fill:'#336699',format:{color:'#FFFFFF'}}}/></Slide></Presentation>`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage();
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      await editor.locator('.hit').first().dblclick();
      const input = editor.locator('.inline-edit');
      assert.equal(
        await input.evaluate((el) => getComputedStyle(el).backgroundColor),
        'rgba(0, 0, 0, 0)',
      );
      const painted = editor.locator('.paint [data-pptx-paragraph]').first();
      assert.equal(await painted.evaluate((el) => getComputedStyle(el).visibility), 'hidden');
      assert.equal(
        await input
          .locator('[data-text-paragraph] span')
          .first()
          .evaluate((el) => getComputedStyle(el).color),
        'rgb(180, 185, 196)',
      );
      await input.press('ControlOrMeta+a');
      await page.keyboard.insertText('Updated revenue');
      await input.press('Escape');
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      assert.equal(await painted.evaluate((el) => getComputedStyle(el).visibility), 'visible');
      assert.match(await painted.textContent(), /Updated revenue/);
      const cell = editor.locator('.paint [data-pptx-cell="0,0"]');
      const cellText = cell.locator('foreignObject');
      const rect = await cellText.boundingBox();
      await page.mouse.dblclick(rect.x + rect.width / 2, rect.y + rect.height / 2);
      await input.waitFor();
      assert.equal(
        await input.evaluate((el) => getComputedStyle(el).backgroundColor),
        'rgba(0, 0, 0, 0)',
      );
      assert.equal(await cellText.evaluate((el) => getComputedStyle(el).visibility), 'hidden');
      assert.equal(
        await cell
          .locator('rect')
          .first()
          .evaluate((el) => getComputedStyle(el).visibility),
        'visible',
      );
      assert.equal(await cell.locator('rect').first().getAttribute('fill'), '#336699');
      assert.equal(
        await editor
          .locator('.paint [data-pptx-cell="0,1"] foreignObject')
          .evaluate((el) => getComputedStyle(el).visibility),
        'visible',
      );
      await input.press('Escape');
      assert.equal(await cellText.evaluate((el) => getComputedStyle(el).visibility), 'visible');
      await page.reload();
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      assert.match(await painted.textContent(), /Updated revenue/);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
