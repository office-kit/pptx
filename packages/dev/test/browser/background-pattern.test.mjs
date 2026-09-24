import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  getSlideBackgroundPatternFill,
  getSlides,
  loadPresentation,
} from '../../../../dist/index.js';
import { startPreview } from '../helpers/server.mjs';

test(
  'background patterns edit selected slides, preserve theme colors, undo and reload',
  { timeout: 120000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-background-pattern-'));
    let preview, browser;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation>{['A','B','C'].map(text => <Slide><Text x={1} y={1} width={7} height={1}>{text}</Text></Slide>)}</Presentation>`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1500, height: 1100 } });
      const errors = [];
      page.on('pageerror', (error) => errors.push(error.message));
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      const saved = () => editor.getByText('Saved to this project', { exact: true }).waitFor();
      const read = async () => {
        const pres = await loadPresentation(
          new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
        );
        return getSlides(pres).map((slide) =>
          getSlideBackgroundPatternFill(pres, slide, { preserveTheme: true }),
        );
      };
      await saved();
      const thumbs = editor.locator('.thumb-row');
      await thumbs.nth(0).click();
      await thumbs.nth(1).click({ modifiers: ['Shift'] });
      const pane = editor.getByRole('region', { name: 'Slide options', exact: true });
      await pane.getByRole('radio', { name: 'Pattern fill', exact: true }).check();
      await saved();
      const initial = { preset: 'pct5', foreground: 'accent1', background: 'bg1' };
      assert.deepEqual(await read(), [initial, initial, null]);
      await pane.getByRole('button', { name: 'Wave', exact: true }).click();
      await saved();
      await pane.getByRole('button', { name: 'Foreground', exact: true }).click();
      await editor.getByRole('menuitemradio', { name: 'Accent 2', exact: true }).click();
      await saved();
      const changed = { ...initial, preset: 'wave', foreground: 'accent2' };
      assert.deepEqual(await read(), [changed, changed, null]);
      await page.screenshot({ path: '/tmp/pptx-background-pattern-panel.png' });
      await page.reload();
      await saved();
      assert.deepEqual(await read(), [changed, changed, null]);
      await thumbs.nth(0).click();
      await pane.getByRole('button', { name: 'Background', exact: true }).click();
      await editor.getByRole('menuitemradio', { name: 'Red', exact: true }).click();
      await saved();
      assert.deepEqual(await read(), [{ ...changed, background: '#FF0000' }, changed, null]);
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual(await read(), [changed, changed, null]);
      await pane.getByRole('button', { name: 'Reset background', exact: true }).click();
      await saved();
      assert.deepEqual(await read(), [null, changed, null]);
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual(await read(), [changed, changed, null]);
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
