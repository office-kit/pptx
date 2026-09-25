import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  getSlideBackgroundPatternFill,
  getSlideLayout,
  getSlideMasterBackgroundPatternFill,
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
      const read = async (effective = false) => {
        const pres = await loadPresentation(
          new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
        );
        return getSlides(pres).map(
          (slide) =>
            getSlideBackgroundPatternFill(pres, slide, { preserveTheme: true }) ??
            (effective
              ? getSlideMasterBackgroundPatternFill(pres, getSlideLayout(slide), {
                  preserveTheme: true,
                })
              : null),
        );
      };
      await saved();
      const thumbs = editor.locator('.thumb-row');
      await thumbs.nth(0).click();
      await thumbs.nth(1).click({ modifiers: ['Shift'] });
      await editor.getByRole('tab', { name: 'Design', exact: true }).click();
      await editor
        .getByRole('tabpanel', { name: 'Design', exact: true })
        .getByRole('button', { name: 'Background Styles', exact: true })
        .click();
      await editor.getByRole('menuitem', { name: 'Format Background...', exact: true }).click();
      const pane = editor.getByRole('region', { name: 'Format Background', exact: true });
      assert.equal(
        await editor.getByRole('region', { name: 'Slide options', exact: true }).count(),
        0,
      );
      await pane.locator('summary').click();
      assert.equal(
        await pane.getByRole('radio', { name: 'Pattern fill', exact: true }).isVisible(),
        false,
      );
      assert.equal(
        await pane.getByRole('button', { name: 'Apply to All', exact: true }).isVisible(),
        true,
      );
      await pane.locator('summary').click();
      await editor.getByRole('button', { name: 'Close Format Background', exact: true }).click();
      assert.equal(await pane.isVisible(), false);
      await editor.getByRole('tab', { name: 'Design', exact: true }).click();
      await editor
        .getByRole('tabpanel', { name: 'Design', exact: true })
        .getByRole('button', { name: 'Background Styles', exact: true })
        .click();
      await editor.getByRole('menuitem', { name: 'Format Background...', exact: true }).click();
      await pane.getByRole('radio', { name: 'Pattern fill', exact: true }).check();
      await saved();
      const initial = { preset: 'pct5', foreground: 'accent1', background: 'bg1' };
      assert.deepEqual(await read(), [initial, initial, null]);
      await page.setViewportSize({ width: 1280, height: 650 });
      const actions = pane.locator('.actions');
      const footerBefore = await actions.boundingBox();
      const paneBounds = await pane.boundingBox();
      assert.ok(footerBefore && paneBounds);
      assert.ok(
        Math.abs(footerBefore.y + footerBefore.height - paneBounds.y - paneBounds.height) < 2,
      );
      const settings = pane.locator('details');
      assert.equal(await settings.evaluate((el) => el.scrollHeight > el.clientHeight), true);
      await settings.hover();
      await page.mouse.wheel(0, 1000);
      await pane.getByRole('button', { name: 'Background', exact: true }).scrollIntoViewIfNeeded();
      const footerAfter = await actions.boundingBox();
      assert.ok(footerAfter && Math.abs(footerAfter.y - footerBefore.y) < 2);
      assert.ok(await settings.evaluate((el) => el.scrollTop > 0));
      await page.screenshot({ path: '/tmp/pptx-background-footer.png' });
      await page.setViewportSize({ width: 1500, height: 1100 });
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
      await editor.getByRole('tab', { name: 'Design', exact: true }).click();
      await editor
        .getByRole('tabpanel', { name: 'Design', exact: true })
        .getByRole('button', { name: 'Background Styles', exact: true })
        .click();
      await editor.getByRole('menuitem', { name: 'Format Background...', exact: true }).click();

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
      await pane.getByRole('button', { name: 'Apply to All', exact: true }).click();
      await saved();
      assert.deepEqual(await read(), [null, null, null]);
      assert.deepEqual(await read(true), [changed, changed, changed]);
      await pane
        .getByRole('button', { name: 'Reset background', exact: true })
        .isDisabled()
        .then((value) => assert.equal(value, true));
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual(await read(), [changed, changed, null]);
      await pane.getByRole('button', { name: 'Apply to All', exact: true }).click();
      await saved();
      await page.reload();
      await saved();
      await editor.getByRole('tab', { name: 'Design', exact: true }).click();
      await editor
        .getByRole('tabpanel', { name: 'Design', exact: true })
        .getByRole('button', { name: 'Background Styles', exact: true })
        .click();
      await editor.getByRole('menuitem', { name: 'Format Background...', exact: true }).click();

      assert.deepEqual(await read(true), [changed, changed, changed]);
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
