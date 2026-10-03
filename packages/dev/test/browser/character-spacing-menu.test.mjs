import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  getShapeParagraphElements,
  getSlideShapes,
  getSlides,
  loadPresentation,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';
import { installRichTextSelection } from '../helpers/rich-text.mjs';

test(
  'character spacing menu applies presets and custom spacing to selected runs',
  { timeout: 90000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-character-spacing-'));
    let preview;
    let browser;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={8} height={2} paragraphs={[{runs:[{text:'Before '},{text:'Tight',format:{kern:1200}},{text:' After'}]}]} /></Slide></Presentation>`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      await installRichTextSelection(page);
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      const saved = () => editor.getByText('Saved to this project', { exact: true }).waitFor();
      const readRuns = async () => {
        const pres = await loadPresentation(
          new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
        );
        const shape = getSlideShapes(getSlides(pres)[0])[0];
        return getShapeParagraphElements(shape, 0)
          .filter((element) => element.kind === 'r')
          .map((element) => ({
            text: element.text,
            spc: element.format?.spc ?? 0,
            kern: element.format?.kern ?? 0,
          }));
      };
      const selectRange = async (start, end) => {
        const input = editor.locator('.inline-edit');
        await input.evaluate(
          (node, range) => {
            node.focus({ preventScroll: true });
            window.selectEditorText(node, range[0], range[1]);
            node.dispatchEvent(new Event('select', { bubbles: true }));
          },
          [start, end],
        );
      };
      const selectTight = () => selectRange(7, 12);
      const trigger = editor.getByRole('button', { name: 'Character Spacing', exact: true });
      const menu = editor.getByRole('menu', { name: 'Character Spacing', exact: true });
      await saved();
      await editor.locator('.hit').first().dblclick();
      await editor.locator('.inline-edit').waitFor();
      await selectTight();

      await trigger.click();
      assert.deepEqual(await menu.getByRole('menuitemradio').allTextContents(), [
        'Very Tight',
        'Tight',
        'Normal',
        'Loose',
        'Very Loose',
      ]);
      assert.equal(
        await menu.getByRole('menuitem', { name: 'More Spacing...', exact: true }).count(),
        1,
      );
      await menu.getByRole('menuitemradio', { name: 'Loose', exact: true }).click();
      await saved();
      assert.deepEqual(await readRuns(), [
        { text: 'Before ', spc: 0, kern: 0 },
        { text: 'Tight', spc: 300, kern: 1200 },
        { text: ' After', spc: 0, kern: 0 },
      ]);

      await selectTight();
      await trigger.click();
      await menu.getByRole('menuitemradio', { name: 'Normal', exact: true }).click();
      await saved();
      assert.deepEqual(await readRuns(), [
        { text: 'Before ', spc: 0, kern: 0 },
        { text: 'Tight', spc: 0, kern: 1200 },
        { text: ' After', spc: 0, kern: 0 },
      ]);

      await selectTight();
      await trigger.click();
      await menu.getByRole('menuitemradio', { name: 'Tight', exact: true }).click();
      await saved();
      assert.deepEqual(await readRuns(), [
        { text: 'Before ', spc: 0, kern: 0 },
        { text: 'Tight', spc: -150, kern: 1200 },
        { text: ' After', spc: 0, kern: 0 },
      ]);
      await trigger.click();
      assert.equal(
        await menu
          .getByRole('menuitemradio', { name: 'Tight', exact: true })
          .getAttribute('aria-checked'),
        'true',
      );
      await menu.press('Escape');

      // A mixed selection has no checked preset; applying Normal resolves every
      // selected run while preserving the pre-existing kerning on Tight.
      await selectRange(0, 18);
      await trigger.click();
      assert.equal(
        await menu
          .getByRole('menuitemradio', { name: 'Normal', exact: true })
          .getAttribute('aria-checked'),
        'false',
      );
      const mixedRuns = await readRuns();
      await menu.getByRole('menuitem', { name: 'More Spacing...', exact: true }).click();
      const mixedDialog = editor.getByRole('dialog', { name: 'Font', exact: true });
      assert.equal(
        await mixedDialog
          .getByRole('tab', { name: 'Character Spacing', exact: true })
          .getAttribute('aria-selected'),
        'true',
      );
      assert.equal(await mixedDialog.getByLabel('Spacing', { exact: true }).inputValue(), '');
      assert.equal(await mixedDialog.getByLabel('By (pt)', { exact: true }).inputValue(), '');
      assert.equal(
        await mixedDialog.getByLabel('Use kerning for fonts', { exact: true }).isChecked(),
        true,
      );
      assert.equal(
        await mixedDialog
          .getByRole('spinbutton', { name: 'For fonts points and above', exact: true })
          .inputValue(),
        '12',
      );
      await mixedDialog.getByLabel('Spacing', { exact: true }).selectOption('expanded');
      assert.equal(await mixedDialog.getByLabel('By (pt)', { exact: true }).inputValue(), '1');
      await mixedDialog.getByRole('button', { name: 'Cancel', exact: true }).click();
      assert.deepEqual(await readRuns(), mixedRuns);
      await trigger.click();
      await menu.getByRole('menuitem', { name: 'More Spacing...', exact: true }).click();
      await mixedDialog.getByRole('button', { name: 'OK', exact: true }).click();
      assert.deepEqual(await readRuns(), mixedRuns);
      await trigger.click();
      await menu.getByRole('menuitem', { name: 'More Spacing...', exact: true }).click();
      await mixedDialog.getByLabel('Spacing', { exact: true }).selectOption('normal');
      await mixedDialog.getByRole('button', { name: 'OK', exact: true }).click();
      await saved();
      assert.deepEqual(await readRuns(), [
        { text: 'Before ', spc: 0, kern: 0 },
        { text: 'Tight', spc: 0, kern: 1200 },
        { text: ' After', spc: 0, kern: 0 },
      ]);
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();

      const beforeCancel = await readRuns();
      await selectTight();
      await trigger.click();
      await menu.getByRole('menuitem', { name: 'More Spacing...', exact: true }).click();
      const custom = editor.getByRole('dialog', { name: 'Font', exact: true });
      await page.screenshot({ path: '/tmp/pptx-character-spacing-dialog.png' });
      assert.equal(await custom.getByLabel('By (pt)', { exact: true }).inputValue(), '1.5');
      await custom.getByLabel('Spacing', { exact: true }).selectOption('normal');
      assert.equal(await custom.getByLabel('By (pt)', { exact: true }).isDisabled(), false);
      await custom.getByLabel('By (pt)', { exact: true }).fill('2');
      await custom.getByLabel('By (pt)', { exact: true }).press('Tab');
      assert.equal(await custom.getByLabel('Spacing', { exact: true }).inputValue(), 'expanded');
      assert.equal(await custom.getByLabel('By (pt)', { exact: true }).inputValue(), '2');
      await custom.getByLabel('Spacing', { exact: true }).selectOption('expanded');
      assert.equal(await custom.getByLabel('By (pt)', { exact: true }).inputValue(), '2');
      await custom.getByLabel('By (pt)', { exact: true }).fill('2.5');
      await custom.getByRole('button', { name: 'Cancel', exact: true }).click();
      assert.equal(
        await trigger.evaluate((node) => node === node.ownerDocument.activeElement),
        true,
      );
      assert.deepEqual(await readRuns(), beforeCancel);

      await trigger.click();
      await menu.getByRole('menuitem', { name: 'More Spacing...', exact: true }).click();
      assert.equal(await custom.getByLabel('By (pt)', { exact: true }).inputValue(), '1.5');
      await custom.getByLabel('Spacing', { exact: true }).selectOption('expanded');
      await custom.getByLabel('By (pt)', { exact: true }).fill('2.5');
      await custom.getByRole('button', { name: 'OK', exact: true }).click();
      await saved();
      const customRuns = await readRuns();
      assert.deepEqual(customRuns, [
        { text: 'Before ', spc: 0, kern: 0 },
        { text: 'Tight', spc: 250, kern: 1200 },
        { text: ' After', spc: 0, kern: 0 },
      ]);

      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual(await readRuns(), beforeCancel);
      await editor.getByTitle('Redo (Ctrl+Y)', { exact: true }).click();
      await saved();
      assert.deepEqual(await readRuns(), customRuns);
      await page.reload();
      await saved();
      assert.deepEqual(await readRuns(), customRuns);

      await editor.locator('.hit').first().dblclick();
      await editor.locator('.inline-edit').waitFor();
      await selectTight();
      await trigger.click();
      await menu.getByRole('menuitem', { name: 'More Spacing...', exact: true }).click();
      const kerningDialog = editor.getByRole('dialog', { name: 'Font', exact: true });
      await kerningDialog.getByLabel('Use kerning for fonts', { exact: true }).uncheck();
      await kerningDialog.getByRole('button', { name: 'OK', exact: true }).click();
      await saved();
      assert.equal((await readRuns())[1].kern, 0);
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual(await readRuns(), customRuns);

      await trigger.press('Enter');
      assert.equal(await trigger.getAttribute('aria-expanded'), 'true');
      const firstPreset = menu.getByRole('menuitemradio').first();
      await firstPreset.focus();
      await firstPreset.press('ArrowDown');
      assert.equal(
        await menu
          .getByRole('menuitemradio')
          .nth(1)
          .evaluate((node) => node === node.ownerDocument.activeElement),
        true,
      );
      await menu.getByRole('menuitemradio').nth(1).press('ArrowUp');
      assert.equal(
        await menu
          .getByRole('menuitemradio')
          .first()
          .evaluate((node) => node === node.ownerDocument.activeElement),
        true,
      );
      await menu.press('Escape');
      assert.equal(await trigger.getAttribute('aria-expanded'), 'false');
      await trigger.press('Enter');
      assert.equal(await trigger.getAttribute('aria-expanded'), 'true');
      const viewport = await editor.locator('body').evaluate((body) => ({
        scrollWidth: document.documentElement.scrollWidth,
        clientWidth: document.documentElement.clientWidth,
        bodyScrollWidth: body.scrollWidth,
        bodyClientWidth: body.clientWidth,
      }));
      assert.ok(viewport.scrollWidth <= viewport.clientWidth + 1, JSON.stringify(viewport));
      assert.ok(viewport.bodyScrollWidth <= viewport.bodyClientWidth + 1, JSON.stringify(viewport));
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
