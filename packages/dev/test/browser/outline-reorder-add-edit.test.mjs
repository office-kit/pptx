import assert from 'node:assert/strict';
import { copyFile, mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import { getSlideText, getSlides, loadPresentation } from '@office-kit/pptx';
import { startPreview, waitForState } from '../helpers/server.mjs';

test(
  'outline keeps editing identity after reorder and adding a slide',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-outline-reorder-add-'));
    let preview, browser;
    try {
      await copyFile(
        new URL('../../../../test/fixtures/minimal/one-text-slide.pptx', import.meta.url),
        join(dir, 'template.pptx'),
      );
      await writeFile(
        join(dir, 'deck.tsx'),
        `import {readFileSync} from 'node:fs'; import {Presentation,Slide,Fill} from '@office-kit/pptx-dsl'; const source = new Uint8Array(readFileSync(new URL('./template.pptx', import.meta.url))); export default <Presentation source={source} mode="compose">{['A','B','C'].map(title => <Slide layout={{name:'Title and Content'}}><Fill target={{placeholder:{type:'title'}}}>{title}</Fill><Fill target={{placeholder:{idx:1}}}>{title} body</Fill></Slide>)}</Presentation>;`,
      );
      preview = await startPreview(join(dir, 'deck.tsx'));
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1400, height: 900 } });
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      await editor.getByRole('tab', { name: 'View', exact: true }).click();
      await editor
        .getByRole('tabpanel', { name: 'View', exact: true })
        .getByRole('button', { name: 'Outline View', exact: true })
        .click();
      const outline = editor.getByRole('navigation', { name: 'Outline View', exact: true });
      const icon = (index) => outline.locator(`[data-outline-slide="${index}"] > button`);
      const menu = editor.getByRole('menu');
      const read = async () =>
        loadPresentation(
          new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
        );
      const change = async (action) => {
        const before = (await waitForState(preview.url, () => true)).revision;
        await action();
        await waitForState(preview.url, (state) => state.revision !== before);
      };
      await icon(0).click({ button: 'right' });
      await change(() => menu.getByRole('menuitem', { name: 'New Slide', exact: true }).click());
      assert.equal(getSlides(await read()).length, 4);
      const last = outline.locator('[data-outline-slide="3"]');
      const bounds = await last.boundingBox();
      await change(() => icon(0).dragTo(last, { targetPosition: { x: 20, y: bounds.height - 2 } }));
      const firstTitle = outline.getByRole('textbox').first();
      const beforeEdit = (await waitForState(preview.url, () => true)).revision;
      await firstTitle.fill('Added after reorder');
      await firstTitle.blur();
      await waitForState(preview.url, (state) => state.revision !== beforeEdit);
      const text = getSlides(await read())
        .map(getSlideText)
        .join('\n');
      assert.match(text, /Added after reorder/);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
