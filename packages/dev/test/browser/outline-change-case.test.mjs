import assert from 'node:assert/strict';
import { copyFile, mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import { getShapeText, getSlideShapes, getSlides, loadPresentation } from '@office-kit/pptx';
import { installRichTextSelection } from '../helpers/rich-text.mjs';
import { startPreview, waitForState } from '../helpers/server.mjs';

test(
  'outline Change Case preserves selection history and expands a caret to its word',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-outline-case-'));
    let preview;
    let browser;
    try {
      await copyFile(
        new URL('../../../../test/fixtures/minimal/one-text-slide.pptx', import.meta.url),
        join(dir, 'template.pptx'),
      );
      await writeFile(
        join(dir, 'deck.tsx'),
        `import {readFileSync} from 'node:fs'; import {Presentation,Slide,Fill} from '@office-kit/pptx-dsl'; const source = new Uint8Array(readFileSync(new URL('./template.pptx', import.meta.url))); export default <Presentation source={source} mode="compose"><Slide layout={{name:"Title and Content"}}><Fill target={{placeholder:{type:'title'}}}>hello world</Fill><Fill target={{placeholder:{idx:1}}}>Body</Fill></Slide></Presentation>;`,
      );
      preview = await startPreview(join(dir, 'deck.tsx'));
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1400, height: 900 } });
      page.on('console', (message) => console.log('PAGE', message.type(), message.text()));
      await installRichTextSelection(page);
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      await editor.getByRole('tab', { name: 'View', exact: true }).click();
      await editor
        .getByRole('tabpanel', { name: 'View', exact: true })
        .getByRole('button', { name: 'Outline View', exact: true })
        .click();
      await editor.getByRole('tab', { name: 'Home', exact: true }).click();
      const title = editor
        .getByRole('navigation', { name: 'Outline View', exact: true })
        .getByRole('textbox')
        .first();
      const caseButton = editor
        .locator('.ribbon .font-ribbon')
        .getByRole('button', { name: 'Change Case', exact: true });
      const readTitle = async () => {
        const pres = await loadPresentation(
          new Uint8Array(await (await fetch(`${preview.url}/deck.pptx`)).arrayBuffer()),
        );
        return getSlideShapes(getSlides(pres)[0])
          .map(getShapeText)
          .find(
            (text) => text.includes('world') || text.includes('WORLD') || text.includes('HELLO'),
          );
      };

      await title.focus();
      await title.evaluate((input) => window.selectEditorText(input, 0, input.textContent.length));
      const saved = editor.getByText('Saved to this project', { exact: true });
      let revision = (await waitForState(preview.url, () => true)).revision;
      await caseButton.click();
      await editor.getByRole('menuitem', { name: 'UPPERCASE', exact: true }).click();
      const upper = await waitForState(preview.url, (state) => state.revision !== revision);
      await saved.waitFor();
      assert.equal(await readTitle(), 'HELLO WORLD');
      revision = upper.revision;
      await page.keyboard.press('Control+z');
      const undone = await waitForState(preview.url, (state) => state.revision !== revision);
      await saved.waitFor();
      assert.equal(await readTitle(), 'hello world');
      revision = undone.revision;
      await page.keyboard.press('Control+y');
      const redone = await waitForState(preview.url, (state) => state.revision !== revision);
      await saved.waitFor();
      assert.equal(await readTitle(), 'HELLO WORLD');

      revision = redone.revision;
      await page.keyboard.press('Control+z');
      await waitForState(preview.url, (state) => state.revision !== revision);
      await saved.waitFor();
      await title.focus();
      await title.evaluate((input) => window.selectEditorText(input, 1, 1));
      revision = (await waitForState(preview.url, () => true)).revision;
      await caseButton.click();
      await editor.getByRole('menuitem', { name: 'UPPERCASE', exact: true }).click();
      await waitForState(preview.url, (state) => state.revision !== revision);
      await saved.waitFor();
      assert.equal(await readTitle(), 'HELLO world');
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
