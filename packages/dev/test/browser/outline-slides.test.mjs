import assert from 'node:assert/strict';
import { copyFile, mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import { getSlides, getSlideText, loadPresentation } from '@office-kit/pptx';
import { startPreview, waitForState } from '../helpers/server.mjs';

for (const locale of ['en', 'ja']) {
  test(
    `outline slide moves preserve selection, history and saved order (${locale})`,
    { timeout: 60000 },
    async () => {
      const dir = await mkdtemp(join(tmpdir(), 'office-outline-slides-'));
      let preview, browser;
      try {
        await copyFile(
          new URL('../../../../test/fixtures/minimal/one-text-slide.pptx', import.meta.url),
          join(dir, 'template.pptx'),
        );
        const file = join(dir, 'deck.tsx');
        await writeFile(
          file,
          `import {readFileSync} from 'node:fs'; import {Presentation,Slide,Fill} from '@office-kit/pptx-dsl'; const source = new Uint8Array(readFileSync(new URL('./template.pptx', import.meta.url))); export default <Presentation source={source} mode="compose">{['A 日本語','B','C','D'].map(title => <Slide layout={{name:'Title and Content'}}><Fill target={{placeholder:{type:'title'}}}>{title}</Fill><Fill target={{placeholder:{idx:1}}}>{title + ' body'}</Fill></Slide>)}</Presentation>;`,
        );
        preview = await startPreview(file);
        browser = await chromium.launch({ headless: true });
        const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
        const errors = [];
        page.on('pageerror', (error) => errors.push(error.message));
        await page.addInitScript(
          (language) => localStorage.setItem('ok-editor-locale', language),
          locale,
        );
        await page.goto(preview.url);
        const editor = page.frameLocator('#editor-frame');
        const labels =
          locale === 'en'
            ? { view: 'View', outline: 'Outline View', up: 'Move Up', down: 'Move Down' }
            : { view: '表示', outline: 'アウトライン表示', up: '上へ移動', down: '下へ移動' };
        await editor.getByRole('tab', { name: labels.view, exact: true }).click();
        await editor
          .getByRole('tabpanel', { name: labels.view, exact: true })
          .getByRole('button', { name: labels.outline, exact: true })
          .click();
        const outline = editor.getByRole('navigation', { name: labels.outline, exact: true });
        const icon = (index) => outline.locator(`[data-outline-slide="${index}"] > button`);
        const menu = editor.getByRole('menu');
        const change = async (action) => {
          const before = (await waitForState(preview.url, () => true)).revision;
          await action();
          await waitForState(preview.url, (state) => state.revision !== before);
        };
        const order = async () =>
          getSlides(
            await loadPresentation(
              new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
            ),
          ).map((slide) => getSlideText(slide));
        const expected = (titles) => titles.map((title) => `${title}\n${title} body`);
        await icon(0).click({ button: 'right' });
        assert.equal(
          await menu.getByRole('menuitem', { name: labels.up, exact: true }).isEnabled(),
          false,
        );
        await change(() => menu.getByRole('menuitem', { name: labels.down, exact: true }).click());
        assert.deepEqual(await order(), expected(['B', 'A 日本語', 'C', 'D']));
        await change(() => icon(1).press('Control+z'));
        assert.deepEqual(await order(), expected(['A 日本語', 'B', 'C', 'D']));
        await change(() => icon(0).press('Control+Shift+z'));
        assert.deepEqual(await order(), expected(['B', 'A 日本語', 'C', 'D']));
        await icon(1).click();
        await icon(2).click({ modifiers: ['Shift'] });
        await icon(2).click({ button: 'right' });
        await change(() => menu.getByRole('menuitem', { name: labels.down, exact: true }).click());
        assert.deepEqual(await order(), expected(['B', 'D', 'A 日本語', 'C']));
        await icon(3).click({ button: 'right' });
        assert.equal(
          await menu.getByRole('menuitem', { name: labels.down, exact: true }).isEnabled(),
          false,
        );
        await change(() => menu.getByRole('menuitem', { name: labels.up, exact: true }).click());
        assert.deepEqual(await order(), expected(['B', 'A 日本語', 'C', 'D']));
        await icon(1).click();
        await icon(2).click({ modifiers: ['Shift'] });
        const lastRow = outline.locator('[data-outline-slide="3"]');
        const rowBounds = await lastRow.boundingBox();
        await change(() =>
          icon(1).dragTo(lastRow, { targetPosition: { x: 20, y: rowBounds.height - 2 } }),
        );
        assert.deepEqual(await order(), expected(['B', 'D', 'A 日本語', 'C']));
        assert.equal(await icon(2).getAttribute('aria-pressed'), 'true');
        assert.equal(await icon(3).getAttribute('aria-pressed'), 'true');
        await change(() => icon(2).press('Control+z'));
        assert.deepEqual(await order(), expected(['B', 'A 日本語', 'C', 'D']));
        // Disjoint selections move together without disturbing the order of unselected slides.
        await icon(0).click();
        await icon(2).click({ modifiers: ['Meta'] });
        assert.equal(await icon(0).getAttribute('aria-pressed'), 'true');
        assert.equal(await icon(2).getAttribute('aria-pressed'), 'true');
        const nextBounds = await lastRow.boundingBox();
        await change(() =>
          icon(0).dragTo(lastRow, { targetPosition: { x: 20, y: nextBounds.height - 2 } }),
        );
        assert.deepEqual(await order(), expected(['A 日本語', 'D', 'B', 'C']));
        await change(() => icon(2).press('Control+z'));
        assert.deepEqual(await order(), expected(['B', 'A 日本語', 'C', 'D']));
        // Text/files dragged from outside this pane cannot reorder slides.
        await icon(3).dispatchEvent('drop');
        assert.deepEqual(await order(), expected(['B', 'A 日本語', 'C', 'D']));
        await page.reload();
        assert.deepEqual(await order(), expected(['B', 'A 日本語', 'C', 'D']));
        assert.deepEqual(errors, []);
      } finally {
        await browser?.close();
        await preview?.close();
        await rm(dir, { recursive: true, force: true });
      }
    },
  );
}
