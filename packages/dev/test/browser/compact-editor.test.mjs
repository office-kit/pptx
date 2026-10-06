import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import { startPreview } from '../helpers/server.mjs';

test(
  'embedded editor keeps its compact command chrome usable at desktop widths',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-compact-editor-'));
    const file = join(dir, 'deck.tsx');
    await writeFile(
      file,
      `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={6} height={1}>Compact header</Text></Slide></Presentation>`,
    );
    let preview;
    let browser;
    try {
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      for (const width of [1500, 900]) {
        const page = await browser.newPage({ viewport: { width, height: 900 } });
        try {
          await page.goto(preview.url);
          // The shell has no header of its own: the editor fills the window.
          assert.equal(await page.locator('body > header').count(), 0);
          assert.deepEqual(await page.locator('#editor-frame').boundingBox(), {
            x: 0,
            y: 0,
            width,
            height: 900,
          });
          const editor = page.frameLocator('#editor-frame');
          await editor.getByText('Saved to this project', { exact: true }).waitFor();

          const header = editor.locator('.topbar');
          const saveStatus = editor.locator('.save-status');
          const [headerBox, statusBox] = await Promise.all([
            header.boundingBox(),
            saveStatus.boundingBox(),
          ]);
          assert.ok(
            headerBox && headerBox.height <= 45,
            `compact header should stay short at ${width}px`,
          );
          assert.ok(
            statusBox && statusBox.y >= headerBox.y && statusBox.y < headerBox.y + headerBox.height,
          );
          await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).waitFor();
          assert.equal(
            await page
              .locator('.slide-edit-tools')
              .evaluate((node) => getComputedStyle(node).display),
            'none',
          );

          const ribbon = editor.locator('.ribbon:has(> .tab-row)');
          const expandedHeight = (await ribbon.boundingBox()).height;
          await editor.getByRole('button', { name: 'Collapse ribbon', exact: true }).click();
          await editor.locator('#ribbon-panel').waitFor({ state: 'hidden' });
          assert.ok(expandedHeight - (await ribbon.boundingBox()).height >= 60);
          await editor.getByRole('tab', { name: 'Insert', exact: true }).click();
          await editor.locator('#ribbon-panel').waitFor({ state: 'visible' });
          await editor.getByRole('button', { name: 'Collapse ribbon', exact: true }).click();
          await editor.getByRole('button', { name: 'Expand ribbon', exact: true }).press('Enter');
          await editor.locator('#ribbon-panel').waitFor({ state: 'visible' });

          await editor.locator('.lang select').selectOption('ja');
          await editor.getByText('このプロジェクトに保存済み', { exact: true }).waitFor();
          const [jaHeader, jaStatus] = await Promise.all([
            header.boundingBox(),
            saveStatus.boundingBox(),
          ]);
          assert.ok(jaHeader && jaStatus && jaStatus.y < jaHeader.y + jaHeader.height);

          await editor
            .locator('.statusbar')
            .getByRole('button', { name: '閲覧表示', exact: true })
            .click();
          await page.locator('.slide-edit-tools').waitFor({ state: 'visible' });
          assert.ok((await page.locator('.slide-edit-tools').boundingBox())?.height > 0);
        } finally {
          await page.close();
        }
      }
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
