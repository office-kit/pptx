import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import { startPreview } from '../helpers/server.mjs';

const DECK = `import {Presentation,Slide,Chart,Table} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Chart x={0.5} y={0.5} width={5} height={3} spec={{kind:'column',categories:['Q1','Q2'],series:[{name:'Revenue',values:[120,180]}]}} /><Table x={6.5} y={0.5} width={5} height={2} rows={[["A","B"],["C","D"]]} /></Slide></Presentation>`;

test(
  'charts get Chart Design and Format tabs, tables get Table Design and Table Layout, like PowerPoint',
  { timeout: 180000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-contextual-tabs-'));
    const file = join(dir, 'deck.tsx');
    await writeFile(file, DECK);
    let preview;
    let browser;
    try {
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1512, height: 900 } });
      await page.goto(preview.url + '/editor');
      await page.getByText('Saved to this project', { exact: true }).waitFor();
      const tabs = () =>
        page.getByRole('tablist', { name: 'Ribbon' }).getByRole('tab').allTextContents();
      const contextual = async () => (await tabs()).map((name) => name.trim()).slice(10);

      await page.locator('.hit').nth(0).click();
      assert.deepEqual(await contextual(), ['Chart Design', 'Format']);
      await page.getByRole('tab', { name: 'Chart Design', exact: true }).click();
      const panel = page.locator('#ribbon-panel');
      for (const name of ['Quick Layout', 'Chart Styles', 'Switch Row/Column'])
        assert.equal(await panel.getByRole('button', { name, exact: true }).isDisabled(), true);
      await panel.getByRole('button', { name: 'Edit Data', exact: true }).click();
      await page.getByRole('table', { name: 'Chart data' }).waitFor();
      await page.keyboard.press('Escape');

      await page.locator('.hit').nth(1).click();
      assert.deepEqual(await contextual(), ['Table Design', 'Table Layout']);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
