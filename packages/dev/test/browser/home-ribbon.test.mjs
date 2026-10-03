import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import { startPreview } from '../helpers/server.mjs';

test(
  'home ribbon groups stay usable without horizontal scrolling in both locales',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-home-ribbon-'));
    const file = join(dir, 'deck.tsx');
    await writeFile(
      file,
      `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={6} height={1}>Ribbon test</Text></Slide></Presentation>`,
    );
    let preview;
    let browser;
    try {
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      for (const width of [756, 900, 1500, 1601, 1900, 2100]) {
        const page = await browser.newPage({ viewport: { width, height: 900 } });
        try {
          await page.goto(preview.url);
          const toggle = page.locator('#toggle-editor');
          if ((await toggle.textContent())?.trim() === 'Edit') await toggle.click();
          await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
          const editor = page.frameLocator('#editor-frame');
          await editor.getByText('Saved to this project', { exact: true }).waitFor();
          for (const locale of ['en', 'ja']) {
            await editor.locator('.lang select').selectOption(locale);
            const groups = editor.locator('#ribbon-panel');
            await groups.waitFor();
            const groupMetrics = await groups.evaluate((node) => ({
              fits: node.scrollWidth <= node.clientWidth + 1,
              scrollWidth: node.scrollWidth,
              clientWidth: node.clientWidth,
              overflow: getComputedStyle(node).overflow,
              children: [...node.children].map((child) => ({
                cls: child.className,
                width: child.getBoundingClientRect().width,
                display: getComputedStyle(child).display,
              })),
            }));
            assert.equal(
              groupMetrics.fits,
              true,
              `${width}px ${locale} ribbon scrolls: ${JSON.stringify(groupMetrics)}`,
            );
            await editor
              .getByRole('button', { name: locale === 'ja' ? '配置' : 'Arrange', exact: true })
              .click();
            const selectionPane = editor.getByRole('menuitemcheckbox', {
              name: locale === 'ja' ? '選択ウィンドウ...' : 'Selection Pane...',
              exact: true,
            });
            await selectionPane.waitFor();
            await page.keyboard.press('Escape');
            await selectionPane.waitFor({ state: 'hidden' });
            const compact = editor.locator('.compact-groups');
            if (width > 2000) {
              assert.equal(await compact.isVisible(), false);
              continue;
            }
            await compact.waitFor();
            assert.equal(
              await compact.evaluate((node) => node.scrollWidth <= node.clientWidth + 1),
              true,
              `${width}px ${locale} group selector scrolls`,
            );

            const fontTrigger = compact.locator('.font-trigger');
            if (width < 800) {
              assert.equal(await editor.locator('.font-group').isVisible(), false);
              await fontTrigger.click();
              const fontMenu = editor.locator('.group-menu');
              await fontMenu.waitFor();
              assert.ok((await fontMenu.locator('input').count()) > 0);
              assert.equal(
                await fontMenu.evaluate((node) => node.scrollWidth <= node.clientWidth + 1),
                true,
                `${width}px ${locale} Font menu scrolls`,
              );
              await page.keyboard.press('Escape');
              await fontMenu.waitFor({ state: 'hidden' });
            } else {
              assert.equal(await fontTrigger.isVisible(), false);
              assert.equal(await editor.locator('.font-group').isVisible(), true);
            }

            const paragraph = compact
              .locator('.group-menu-trigger')
              .filter({ hasText: locale === 'ja' ? '段落' : 'Paragraph' });
            if (width >= 1100) {
              assert.equal(await paragraph.isVisible(), false);
              assert.equal(await editor.locator('.paragraph-group').isVisible(), true);
            } else {
              await paragraph.click();
              await editor.locator('.group-menu').waitFor();
              const paragraphMenu = editor.locator('.group-menu');
              assert.ok(
                (await paragraphMenu.getByRole('button').count()) > 0,
                `${width}px ${locale} Paragraph menu is empty`,
              );
              assert.equal(
                await paragraphMenu
                  .getByRole('button', { name: locale === 'ja' ? '行間' : 'Line spacing' })
                  .count(),
                1,
                `${width}px ${locale} line spacing control is missing`,
              );
              await page.keyboard.press('Escape');
              await editor.locator('.group-menu').waitFor({ state: 'hidden' });
            }
            await page.screenshot({
              path: `/tmp/home-ribbon-${width}-${locale}.png`,
              fullPage: true,
            });
          }
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
