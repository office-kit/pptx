import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import { startPreview, waitForState } from '../helpers/server.mjs';

test(
  'rulers follow slide coordinates through zoom, scrolling and viewport resizing',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-rulers-'));
    let preview, browser;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={3} height={1}>Ruler</Text></Slide></Presentation>`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      const revision = (await waitForState(preview.url, () => true)).revision;
      await editor.getByRole('tab', { name: 'View', exact: true }).click();
      const panel = editor.getByRole('tabpanel', { name: 'View', exact: true });
      await panel.getByRole('checkbox', { name: 'Ruler', exact: true }).check();
      const verifyAlignment = async () => {
        await editor.locator('.rulers').evaluate(async (node) => {
          await new Promise((resolve) =>
            requestAnimationFrame(() => requestAnimationFrame(resolve)),
          );
          const stage = node.parentElement.querySelector('.stage').getBoundingClientRect();
          for (const [axis, start, extent] of [
            ['horizontal', 'left', 'width'],
            ['vertical', 'top', 'height'],
          ]) {
            const svg = node.querySelector(`.${axis}`);
            const labels = [...svg.querySelectorAll('text')].map((label) =>
              label.getBoundingClientRect(),
            );
            for (let i = 1; i < labels.length; i++) {
              assertLabelGap(labels[i][start] - labels[i - 1][start] - labels[i - 1][extent]);
            }
            function assertLabelGap(gap) {
              if (gap < 2) throw new Error('Ruler labels overlap');
            }
            const origin = svg.getBoundingClientRect()[start];
            const band = svg.querySelector('rect').getBoundingClientRect();
            if (
              Math.abs(band[start] - stage[start]) > 1 ||
              Math.abs(band[extent] - stage[extent]) > 1
            )
              throw new Error(`${axis} slide band is displaced`);
            for (const line of svg.querySelectorAll('line')) {
              const pxPerInch = stage.width / (13 + 1 / 3);
              const expected =
                stage[start] + stage[extent] / 2 + (Number(line.dataset.value) * pxPerInch) / 2.54;
              const actual =
                origin + Number(line.getAttribute(axis === 'horizontal' ? 'x1' : 'y1'));
              if (Math.abs(expected - actual) > 1) throw new Error(`${axis} tick is displaced`);
            }
          }
        });
      };
      await verifyAlignment();
      await editor.getByTitle('Zoom...', { exact: true }).click();
      const dialog = editor.getByRole('dialog', { name: 'Zoom', exact: true });
      await dialog.getByRole('radio', { name: '200%', exact: true }).check();
      await dialog.getByRole('button', { name: 'OK', exact: true }).click();
      await verifyAlignment();
      await editor.locator('.canvas-area').evaluate((node) => node.scrollTo(300, 200));
      await verifyAlignment();
      await page.setViewportSize({ width: 1100, height: 800 });
      await verifyAlignment();
      await panel.getByRole('checkbox', { name: 'Gridlines', exact: true }).check();
      await editor.getByRole('button', { name: 'View', exact: true }).click();
      const menuRuler = editor.getByRole('menuitemcheckbox', { name: 'Ruler', exact: true });
      assert.equal(await menuRuler.getAttribute('aria-checked'), 'true');
      await menuRuler.click();
      await editor.locator('.rulers').waitFor({ state: 'detached' });
      await panel.getByRole('checkbox', { name: 'Ruler', exact: true }).check();
      await panel.getByRole('button', { name: 'Slide Sorter', exact: true }).click();
      assert.equal(
        await panel.getByRole('checkbox', { name: 'Ruler', exact: true }).isEnabled(),
        false,
      );
      await editor.locator('.rulers').waitFor({ state: 'detached' });
      await panel.getByRole('button', { name: 'Normal', exact: true }).click();
      await editor.locator('.rulers').waitFor();
      await verifyAlignment();
      await page.reload();
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      await editor.locator('.rulers').waitFor();
      await verifyAlignment();
      assert.equal((await waitForState(preview.url, () => true)).revision, revision);
      await page.screenshot({ path: '/tmp/pptx-rulers.png' });
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
