import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import { getSlideShowProperties, loadPresentation } from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

test(
  'Set Up Show preserves Cancel, saves one transaction, supports Undo and bilingual validation',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-show-properties-'));
    let preview, browser, page;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation>{['A','B'].map(text => <Slide><Text x={1} y={1} width={7} height={1}>{text}</Text></Slide>)}</Presentation>`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      const errors = [];
      page.on('pageerror', (error) => errors.push(error.message));
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      const saved = () => editor.getByText('Saved to this project', { exact: true }).waitFor();
      const deck = async () =>
        loadPresentation(
          new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
        );
      const settings = async () => getSlideShowProperties(await deck());
      const state = async () => (await fetch(preview.url + '/state')).json();
      const waitPreviewState = async (predicate) => {
        const deadline = Date.now() + 10000;
        while (Date.now() < deadline) {
          const snapshot = await state();
          if (predicate(snapshot)) return snapshot;
          await new Promise((resolve) => setTimeout(resolve, 50));
        }
        throw new Error('Preview state did not settle');
      };
      const open = async (locale = 'en') => {
        const japanese = locale === 'ja';
        const show = japanese ? 'スライド ショー' : 'Slide Show';
        const setup = japanese ? 'スライド ショーの設定' : 'Set Up Show';
        await editor.getByRole('tab', { name: show, exact: true }).click();
        await editor
          .getByRole('button', { name: japanese ? setup : 'Set Up Slide Show', exact: true })
          .click();
        return editor.getByRole('dialog', { name: setup, exact: true });
      };
      await saved();
      const initial = await settings();
      assert.deepEqual((await state()).showProperties, initial);

      await editor.locator('.lang select').selectOption('ja');
      const japaneseDialog = await open('ja');
      assert.equal(
        await japaneseDialog.getByLabel('アニメーションを表示しない', { exact: true }).isChecked(),
        !initial.showAnimation,
      );
      assert.equal(
        await japaneseDialog.getByLabel('ナレーションを表示しない', { exact: true }).isChecked(),
        !initial.showNarration,
      );
      await japaneseDialog.getByRole('button', { name: 'キャンセル', exact: true }).click();
      await editor.locator('.lang select').selectOption('en');

      // Cancel leaves both the file and the server snapshot untouched.
      let dialog = await open();
      await dialog.getByLabel("Loop continuously until 'Esc'", { exact: true }).check();
      await dialog.getByLabel('Show without narration', { exact: true }).uncheck();
      await dialog.getByRole('button', { name: 'Cancel', exact: true }).click();
      await saved();
      assert.deepEqual(await settings(), initial);
      assert.deepEqual((await state()).showProperties, initial);

      // An invalid range is surfaced before submission and cannot be committed.
      dialog = await open();
      await dialog.getByLabel('From:', { exact: true }).check();
      await dialog.getByLabel('From slide', { exact: true }).fill('2');
      await dialog.getByLabel('To slide', { exact: true }).fill('1');
      await dialog.getByRole('alert').waitFor();
      assert.equal(
        await dialog.getByRole('button', { name: 'OK', exact: true }).isDisabled(),
        true,
      );
      await dialog.getByRole('button', { name: 'Cancel', exact: true }).click();
      assert.deepEqual(await settings(), initial);

      const beforeRevision = (await state()).revision;
      dialog = await open();
      await dialog.getByLabel('From:', { exact: true }).check();
      await dialog.getByLabel('From slide', { exact: true }).fill('1');
      await dialog.getByLabel('To slide', { exact: true }).fill('2');
      await dialog.getByLabel("Loop continuously until 'Esc'", { exact: true }).check();
      await dialog.getByLabel('Show without narration', { exact: true }).uncheck();
      await dialog.getByLabel('Using timings, if present', { exact: true }).check();
      await dialog.getByRole('button', { name: 'OK', exact: true }).click();
      await saved();
      const changed = await waitPreviewState((snapshot) => snapshot.revision > beforeRevision);
      assert.deepEqual(changed.showProperties.slides, { kind: 'range', start: 1, end: 2 });
      const edited = await settings();
      assert.deepEqual(edited.slides, { kind: 'range', start: 1, end: 2 });
      assert.equal(edited.loop, true);
      assert.equal(edited.showNarration, true);
      assert.equal(edited.useTimings, true);

      // The whole dialog edit is one history entry.
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual(await settings(), initial);
      assert.deepEqual((await state()).showProperties, initial);

      // Japanese suppression controls persist the inverse OOXML flags.
      await editor.locator('.lang select').selectOption('ja');
      dialog = await open('ja');
      await dialog.getByLabel('アニメーションを表示しない', { exact: true }).check();
      await dialog.getByLabel('ナレーションを表示しない', { exact: true }).uncheck();
      await dialog.getByRole('button', { name: 'OK', exact: true }).click();
      await editor.getByText('このプロジェクトに保存済み', { exact: true }).waitFor();
      assert.equal((await settings()).showAnimation, false);
      assert.equal((await settings()).showNarration, true);
      dialog = await open('ja');
      assert.equal(
        await dialog.getByLabel('アニメーションを表示しない', { exact: true }).isChecked(),
        true,
      );
      assert.equal(
        await dialog.getByLabel('ナレーションを表示しない', { exact: true }).isChecked(),
        false,
      );
      await dialog.getByLabel('アニメーションを表示しない', { exact: true }).uncheck();
      await dialog.getByLabel('ナレーションを表示しない', { exact: true }).check();
      await dialog.getByRole('button', { name: 'OK', exact: true }).click();
      await editor.getByText('このプロジェクトに保存済み', { exact: true }).waitFor();
      assert.deepEqual(await settings(), initial);
      assert.deepEqual(errors, []);
    } catch (error) {
      await page?.screenshot({ path: '/tmp/pptx-show-properties-failure.png', fullPage: true });
      throw error;
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
