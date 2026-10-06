import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import { startPreview } from '../helpers/server.mjs';

// ANSI red (SGR 31) from the terminal's light and dark palettes.
const RED = { light: 'rgb(201, 42, 42)', dark: 'rgb(241, 112, 122)' };

test(
  'the Agents task pane docks beside the editor and follows its theme',
  { timeout: 120000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-agents-pane-'));
    await writeFile(
      join(dir, 'claude'),
      `#!${process.execPath}
console.log('\\x1b[31mRed output\\x1b[0m ready');
process.stdin.on('data', () => {});
`,
      { mode: 0o755 },
    );
    const file = join(dir, 'deck.tsx');
    await writeFile(
      file,
      `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={4} height={1}>One</Text></Slide><Slide><Text x={1} y={1} width={4} height={1}>Two</Text></Slide></Presentation>`,
    );
    let preview;
    let browser;
    try {
      preview = await startPreview(file, {
        env: { ...process.env, PATH: dir + ':' + process.env.PATH },
      });
      browser = await chromium.launch({ headless: true });
      const context = await browser.newContext({
        viewport: { width: 1400, height: 900 },
        colorScheme: 'light',
      });
      const page = await context.newPage();
      const errors = [];
      page.on('pageerror', (error) => errors.push(error.message));
      await page.goto(preview.url);
      const editor = page.frameLocator('#editor-frame');
      await editor.getByText('Saved to this project', { exact: true }).waitFor();

      // No shell header: the editor fills the window and the pane starts closed.
      assert.equal(await page.locator('header').count(), 0);
      assert.equal(await page.locator('#chat').isVisible(), false);
      assert.equal((await page.locator('#editor-frame').boundingBox()).width, 1400);

      const agents = editor
        .locator('.tab-row')
        .getByRole('button', { name: 'Agents', exact: true });
      await agents.waitFor();
      await editor.locator('.tab-row .agents[aria-pressed="false"]').waitFor();
      await agents.click();
      await page.locator('#chat').waitFor({ state: 'visible' });
      await editor.locator('.tab-row .agents[aria-pressed="true"]').waitFor();
      assert.equal(await page.locator('#agents-title').textContent(), 'Agents');
      assert.ok((await page.locator('#editor-frame').boundingBox()).width < 1400);

      // The pane is painted with the editor's own tokens in both schemes.
      const editorPanel = () =>
        page
          .frames()
          .find((frame) => frame.url().endsWith('/editor'))
          .evaluate(() => {
            const probe = document.createElement('div');
            probe.style.background = 'var(--ok-panel)';
            document.querySelector('.ok-editor').append(probe);
            const color = getComputedStyle(probe).backgroundColor;
            probe.remove();
            return color;
          });
      const paneColor = () =>
        page.locator('#chat').evaluate((node) => getComputedStyle(node).backgroundColor);
      const agentFrame = page.frames().find((frame) => /\/agents\//.test(frame.url()));
      const agentColor = () =>
        agentFrame.evaluate(() => getComputedStyle(document.body).backgroundColor);
      const terminalRed = () =>
        agentFrame.evaluate(() => {
          const span = [...document.querySelectorAll('.xterm-rows span')].find(
            (node) => node.textContent === 'Red output',
          );
          return span && getComputedStyle(span).color;
        });

      await agentFrame.locator('#terminal-start').click();
      await agentFrame.waitForFunction(() =>
        document.querySelector('#terminal').textContent.includes('Red output ready'),
      );
      assert.equal(await paneColor(), 'rgb(255, 255, 255)');
      assert.equal(await paneColor(), await editorPanel());
      assert.equal(await agentColor(), await editorPanel());
      assert.equal(await terminalRed(), RED.light);
      assert.match(
        await agentFrame
          .locator('.xterm-rows')
          .evaluate((node) => getComputedStyle(node).fontFamily),
        /^ui-monospace, "SF Mono"/,
      );

      await page.emulateMedia({ colorScheme: 'dark' });
      assert.equal(await paneColor(), 'rgb(43, 43, 43)');
      assert.equal(await paneColor(), await editorPanel());
      assert.equal(await agentColor(), await editorPanel());
      await agentFrame.waitForFunction(
        (red) =>
          [...document.querySelectorAll('.xterm-rows span')].some(
            (node) => node.textContent === 'Red output' && getComputedStyle(node).color === red,
          ),
        RED.dark,
      );

      // × closes the pane and hands focus back to the title-row button.
      await page.getByRole('button', { name: 'Close Agents', exact: true }).click();
      await page.locator('#chat').waitFor({ state: 'hidden' });
      await editor.locator('.tab-row .agents[aria-pressed="false"]').waitFor();
      await page.waitForFunction(() => document.activeElement?.id === 'editor-frame');
      await editor.locator('.tab-row .agents:focus').waitFor();
      assert.equal((await page.locator('#editor-frame').boundingBox()).width, 1400);

      // View ▸ Reading View shows the rendered viewer; Normal returns.
      await editor.getByRole('tab', { name: 'View', exact: true }).click();
      await editor
        .locator('#ribbon-panel')
        .getByRole('button', { name: 'Reading View', exact: true })
        .click();
      await page.locator('#stage').waitFor({ state: 'visible' });
      assert.equal(await page.locator('#editor-frame').isVisible(), false);
      await page.locator('#slide svg').waitFor({ state: 'attached' });
      await page.getByRole('button', { name: 'Normal', exact: true }).click();
      await page.locator('#editor-frame').waitFor({ state: 'visible' });
      assert.equal(await page.locator('#stage').isVisible(), false);

      // The pane state survives a reload; it is the tab's, like the view.
      await agents.click();
      await page.locator('#chat').waitFor({ state: 'visible' });
      await page.reload();
      await page.locator('#chat').waitFor({ state: 'visible' });
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      await editor.locator('.tab-row .agents[aria-pressed="true"]').waitFor();
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
