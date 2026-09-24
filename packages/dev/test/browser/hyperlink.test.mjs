import { chromium } from 'playwright';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, writeFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { spawn } from 'node:child_process';
import { fileURLToPath } from 'node:url';

test(
  'ribbon, shortcut and context hyperlink edits persist and undo',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-link-'));
    const file = join(dir, 'deck.tsx');
    await writeFile(
      file,
      `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={8} height={1}>Link target</Text></Slide><Slide><Text x={1} y={1} width={8} height={1}>Return</Text></Slide></Presentation>`,
    );
    const proc = spawn(process.execPath, [
      fileURLToPath(new URL('../../dist/cli.mjs', import.meta.url)),
      'dev',
      file,
      '--port',
      '0',
    ]);
    let browser;
    try {
      const url = await new Promise((resolve, reject) => {
        let output = '';
        proc.stdout.on('data', (data) => {
          output += data;
          const match = output.match(/Preview: (http:\/\/\S+)/);
          if (match) resolve(match[1]);
        });
        proc.on('error', reject);
        proc.on('exit', (code) => reject(new Error('Server exited ' + code)));
      });
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1600, height: 1000 } });
      const errors = [];
      page.on('pageerror', (error) => errors.push(error.message));
      await page.goto(url);
      await page.locator('.shape-hit').first().click();
      await page.locator('#tab-insert').click();
      await page.locator('[data-edit="link"]').click();
      await page.getByLabel('Address:', { exact: true }).fill('https://example.com/');
      await page.getByRole('button', { name: 'OK', exact: true }).click();
      await page.waitForFunction(
        () =>
          window.office.getState().editor.slides[0].shapes[0].link?.action.url ===
          'https://example.com/',
      );
      await page.waitForFunction(
        () =>
          !document.querySelector('[data-edit="link"]').disabled &&
          !document.querySelector('.link-dialog'),
      );
      await page.keyboard.press('Meta+k');
      await page.getByRole('dialog', { name: 'Edit Hyperlink' }).waitFor();
      await page.getByRole('tab', { name: 'This Document' }).click();
      await page.getByLabel('Select a place in this document:').selectOption('lastSlide');
      await page.getByRole('button', { name: 'OK', exact: true }).click();
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].shapes[0].link?.action.kind === 'lastSlide',
      );
      await page.waitForFunction(
        () =>
          !document.querySelector('[data-edit="link"]').disabled &&
          !document.querySelector('.link-dialog'),
      );
      await page.locator('.shape-hit').first().click({ button: 'right' });
      await page.getByRole('menuitem', { name: 'Open Hyperlink', exact: true }).click();
      await page.waitForFunction(() => window.office.getState().index === 1);
      assert.equal(await page.evaluate(() => window.office.getState().presenting), false);
      await page.evaluate(() => window.office.selectSlide(0, true));
      await page.locator('.shape-hit').first().click({ button: 'right' });
      await page.getByRole('menuitem', { name: 'Remove Hyperlink', exact: true }).click();
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].shapes[0].link === null,
      );
      await page.locator('[data-edit="undo"]').click();
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].shapes[0].link?.action.kind === 'lastSlide',
      );
      await page.reload();
      await page.waitForFunction(
        () =>
          window.office.getState().editor?.slides[0].shapes[0].link?.action.kind === 'lastSlide',
      );
      await page.locator('.shape-hit').first().dblclick();
      const field = page.locator('.direct-text');
      await field.waitFor();
      await field.evaluate((f) => {
        f.focus();
        f.setSelectionRange(0, 4);
      });
      await page.keyboard.press('Meta+k');
      await page.getByRole('dialog', { name: 'Insert Hyperlink' }).waitFor();
      assert.equal(await page.getByLabel('Text to Display:', { exact: true }).inputValue(), 'Link');
      await page.getByLabel('Address:', { exact: true }).fill('https://example.com/selected');
      await page.getByRole('button', { name: 'OK', exact: true }).click();
      await field.waitFor();
      assert.deepEqual(await field.evaluate((f) => [f.selectionStart, f.selectionEnd]), [0, 4]);
      assert.deepEqual(
        await page.evaluate(() => window.office.getState().editor.slides[0].shapes[0].textLinks),
        [
          {
            start: 0,
            end: 4,
            link: { action: { kind: 'url', url: 'https://example.com/selected' }, tooltip: null },
          },
        ],
      );
      await page.keyboard.press('Meta+k');
      await page.getByRole('dialog', { name: 'Edit Hyperlink' }).waitFor();
      await page.getByRole('button', { name: 'Cancel', exact: true }).click();
      await field.waitFor();
      assert.deepEqual(await field.evaluate((f) => [f.selectionStart, f.selectionEnd]), [0, 4]);
      await page.keyboard.press('Meta+k');
      await page.getByRole('button', { name: 'Remove Link', exact: true }).click();
      await field.waitFor();
      assert.deepEqual(
        await page.evaluate(() => window.office.getState().editor.slides[0].shapes[0].textLinks),
        [],
      );
      await page.locator('[data-edit="undo"]').click();
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].shapes[0].textLinks.length === 1,
      );
      await page.reload();
      await page.waitForFunction(
        () => window.office.getState().editor?.slides[0].shapes[0].textLinks.length === 1,
      );
      assert.equal(
        await page.evaluate(
          () => window.office.getState().editor.slides[0].shapes[0].link.action.kind,
        ),
        'lastSlide',
      );
      await page.locator('.shape-hit').first().dblclick();
      await field.waitFor();
      await field.evaluate((f) => {
        f.focus();
        f.setSelectionRange(0, 4);
      });
      await page.keyboard.press('Meta+k');
      await page.getByRole('dialog', { name: 'Edit Hyperlink' }).waitFor();
      await page.getByLabel('Text to Display:', { exact: true }).fill('参照資料😀');
      await page.getByRole('button', { name: 'OK', exact: true }).click();
      await field.waitFor();
      assert.equal(await field.evaluate((f) => f.value), '参照資料😀 target');
      assert.deepEqual(await field.evaluate((f) => [f.selectionStart, f.selectionEnd]), [0, 6]);
      await page.locator('[data-edit="undo"]').click();
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].shapes[0].text === 'Link target',
      );
      assert.equal(
        await page.evaluate(
          () => window.office.getState().editor.slides[0].shapes[0].textLinks[0].link.action.url,
        ),
        'https://example.com/selected',
      );
      await field.waitFor();
      await field.evaluate((f) => {
        f.focus();
        f.setSelectionRange(2, 2);
      });
      await page.keyboard.press('Meta+k');
      await page.getByRole('dialog', { name: 'Edit Hyperlink' }).waitFor();
      assert.equal(await page.getByLabel('Text to Display:', { exact: true }).inputValue(), 'Link');
      await page.getByRole('button', { name: 'Cancel', exact: true }).click();
      await field.waitFor();
      assert.deepEqual(await field.evaluate((f) => [f.selectionStart, f.selectionEnd]), [2, 2]);
      await field.evaluate((f) => {
        f.focus();
        f.setSelectionRange(11, 11);
      });
      await page.keyboard.press('Meta+k');
      await page.getByRole('dialog', { name: 'Insert Hyperlink' }).waitFor();
      assert.equal(await page.getByLabel('Text to Display:', { exact: true }).inputValue(), '');
      await page.getByLabel('Address:', { exact: true }).fill('https://example.com/new');
      await page.getByRole('button', { name: 'OK', exact: true }).click();
      await field.waitFor();
      assert.equal(await field.evaluate((f) => f.value), 'Link targethttps://example.com/new');
      assert.deepEqual(await field.evaluate((f) => [f.selectionStart, f.selectionEnd]), [11, 34]);
      await page.locator('[data-edit="undo"]').click();
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].shapes[0].text === 'Link target',
      );
      assert.equal(
        await page.evaluate(
          () => window.office.getState().editor.slides[0].shapes[0].textLinks.length,
        ),
        1,
      );
      await page.keyboard.press('Escape');
      await page.getByRole('tab', { name: 'Insert', exact: true }).click();
      await page.locator('[data-edit=table]').click();
      await page.getByRole('gridcell', { name: '2 columns, 2 rows', exact: true }).click();
      await page.getByRole('tab', { name: 'Table Layout', exact: true }).waitFor();
      const box = await page.locator('#stage [data-pptx-cell="0,0"]').first().boundingBox();
      await page.mouse.dblclick(box.x + box.width / 2, box.y + box.height / 2);
      const cellField = page.getByRole('textbox', { name: 'Edit cell 1, 1', exact: true });
      await cellField.fill('Cell target');
      await cellField.evaluate((f) => f.setSelectionRange(0, 4));
      await page.keyboard.press('Meta+k');
      await page.getByRole('dialog', { name: 'Insert Hyperlink' }).waitFor();
      assert.equal(await page.getByLabel('Text to Display:', { exact: true }).inputValue(), 'Cell');
      await page.getByLabel('Address:', { exact: true }).fill('https://example.com/cell');
      await page.getByRole('button', { name: 'OK', exact: true }).click();
      await cellField.waitFor();
      assert.deepEqual(await cellField.evaluate((f) => [f.selectionStart, f.selectionEnd]), [0, 4]);
      assert.equal(
        await page.evaluate(
          () =>
            window.office.getState().editor.slides[0].shapes.find((s) => s.table).table.cells[0][0]
              .textLinks[0].link.action.url,
        ),
        'https://example.com/cell',
      );
      await page
        .context()
        .route('https://example.com/cell', (route) =>
          route.fulfill({ status: 200, body: 'Cell destination' }),
        );
      await cellField.evaluate((f) => f.setSelectionRange(f.value.length, f.value.length));
      await page.keyboard.insertText('!');
      await cellField.evaluate((f) => f.setSelectionRange(0, 4));
      await cellField.dispatchEvent('contextmenu', { clientX: 400, clientY: 400 });
      await page.waitForFunction(
        () =>
          window.office.getState().editor.slides[0].shapes.find((s) => s.table).table.cells[0][0]
            .text === 'Cell target!',
      );
      const editPopupPromise = page.waitForEvent('popup');
      await page.getByRole('menuitem', { name: 'Open Hyperlink', exact: true }).click();
      const editPopup = await editPopupPromise;
      await editPopup.waitForLoadState();
      assert.equal(editPopup.url(), 'https://example.com/cell');
      assert.equal(await page.evaluate(() => window.office.getState().presenting), false);
      await editPopup.close();
      await cellField.focus();
      await page.keyboard.press('Meta+k');
      await page.getByRole('dialog', { name: 'Edit Hyperlink' }).waitFor();
      await page.getByRole('button', { name: 'Remove Link', exact: true }).click();
      await cellField.waitFor();
      assert.deepEqual(
        await page.evaluate(
          () =>
            window.office.getState().editor.slides[0].shapes.find((s) => s.table).table.cells[0][0]
              .textLinks,
        ),
        [],
      );
      await page.locator('[data-edit="undo"]').click();
      await page.waitForFunction(
        () =>
          window.office.getState().editor.slides[0].shapes.find((s) => s.table).table.cells[0][0]
            .textLinks.length === 1,
      );
      await page.reload();
      await page.waitForFunction(
        () =>
          window.office.getState().editor?.slides[0].shapes.find((s) => s.table)?.table.cells[0][0]
            .textLinks.length === 1,
      );
      await page.locator('#present').click();
      await page.waitForFunction(() => window.office.getState().presenting);
      await page
        .context()
        .route('https://example.com/cell', (route) =>
          route.fulfill({ status: 200, body: 'Cell destination' }),
        );
      const popupPromise = page.waitForEvent('popup');
      await page.locator('#slide a[href="https://example.com/cell"]').click();
      const popup = await popupPromise;
      await popup.waitForLoadState();
      assert.equal(popup.url(), 'https://example.com/cell');
      assert.equal(await page.evaluate(() => window.office.getState().index), 0);
      await popup.close();
      const lastLink = page.locator('#slide a[href="#slide-lastSlide"]').first();
      await lastLink.waitFor();
      await lastLink.dispatchEvent('click');
      await page.waitForFunction(() => window.office.getState().index === 1);
      assert.equal(await page.evaluate(() => window.office.getState().presenting), true);
      assert.equal(await page.locator('#show-screen').isVisible(), false);
      await page.keyboard.press('Escape');
      await page.waitForFunction(() => !window.office.getState().presenting);
      // Persist real actions through the editor API, then exercise slideshow history.
      const setAction = async (slide, kind, hover = false) => {
        const revision = await page.evaluate(
          async ({ slide, kind, hover }) => {
            const state = window.office.getState();
            const response = await fetch('/edit', {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({
                revision: state.revision,
                command: {
                  type: hover ? 'object-actions' : 'object-link',
                  ...(hover
                    ? { hoverLink: kind ? { action: { kind }, tooltip: null } : null }
                    : {}),
                  slide,
                  ids: [state.editor.slides[slide].shapes[0].id],
                  link: hover ? null : { action: { kind }, tooltip: null },
                },
              }),
            });
            if (!response.ok) throw new Error(await response.text());
            return state.revision;
          },
          { slide, kind, hover },
        );
        await page.waitForFunction(
          (revision) =>
            window.office.getState().revision > revision && !window.office.getState().building,
          revision,
        );
      };
      await setAction(0, 'lastSlideViewed');
      await setAction(1, 'lastSlideViewed');
      await page.evaluate(() => window.office.selectSlide(0, true));
      await page.locator('#present').click();
      await page.waitForFunction(() => window.office.getState().presenting);
      const historyLink = page.locator('#slide a[href="#slide-lastSlideViewed"]').first();
      await historyLink.dispatchEvent('click');
      assert.equal(await page.evaluate(() => window.office.getState().index), 0);
      await page.keyboard.press('ArrowRight');
      await page.waitForFunction(() => window.office.getState().index === 1);
      await historyLink.dispatchEvent('click');
      await page.waitForFunction(() => window.office.getState().index === 0);
      await historyLink.dispatchEvent('click');
      await page.waitForFunction(() => window.office.getState().index === 1);
      await page.keyboard.press('Escape');
      await page.waitForFunction(() => !window.office.getState().presenting);
      await setAction(0, 'endShow');
      await page.evaluate(() => window.office.selectSlide(0, true));
      await page.locator('#present').click();
      await page.waitForFunction(() => window.office.getState().presenting);
      await page.locator('#slide a[href="#slide-endShow"]').first().dispatchEvent('click');
      await page.waitForFunction(() => !window.office.getState().presenting);
      assert.equal(await page.evaluate(() => window.office.getState().index), 0);
      await setAction(0, 'nextSlide', true);
      await setAction(1, 'endShow', true);
      await page.reload();
      await page.waitForFunction(
        () =>
          window.office.getState().editor?.slides[0].shapes[0].hoverLink?.action.kind ===
          'nextSlide',
      );
      const hoverNext = page.locator('#slide [data-hover-href="#slide-nextSlide"]').first();
      await hoverNext.dispatchEvent('pointerover', { pointerType: 'mouse' });
      assert.equal(await page.evaluate(() => window.office.getState().index), 0);
      await page.locator('#present').click();
      await page.waitForFunction(() => window.office.getState().presenting);
      await hoverNext.dispatchEvent('pointerover', { pointerType: 'touch' });
      assert.equal(await page.evaluate(() => window.office.getState().index), 0);
      await page.mouse.move(0, 0);
      await hoverNext.hover();
      await page.waitForFunction(() => window.office.getState().index === 1);
      assert.equal(await page.evaluate(() => window.office.getState().presenting), true);
      await page.mouse.move(0, 0);
      const endHoverBox = await page
        .locator('#slide [data-hover-href="#slide-endShow"]')
        .first()
        .boundingBox();
      assert.ok(endHoverBox);
      await page.mouse.move(
        endHoverBox.x + endHoverBox.width / 2,
        endHoverBox.y + endHoverBox.height / 2,
      );
      await page.waitForFunction(() => !window.office.getState().presenting);
      await setAction(0, null, true);
      await setAction(1, null, true);
      await page.evaluate(() => window.office.selectSlide(0, true));
      await page.locator('#tab-insert').click();
      await page.locator('[data-edit="shapes"]:visible').first().click();
      await page
        .getByRole('menuitem', { name: 'Action Button: Forward or Next', exact: true })
        .click();
      const canvas = await page.locator('#slide').boundingBox();
      await page.mouse.move(canvas.x + 50, canvas.y + 200);
      await page.mouse.down();
      await page.mouse.move(canvas.x + 130, canvas.y + 280, { steps: 5 });
      await page.mouse.up();
      await page.waitForFunction(() =>
        window.office
          .getState()
          .editor.slides[0].shapes.some(
            (s) => s.preset === 'actionButtonForwardNext' && s.link?.action.kind === 'nextSlide',
          ),
      );
      const actions = page.getByRole('dialog', { name: 'Action Settings', exact: true });
      await actions.waitFor();
      assert.equal(
        await actions.getByLabel('Hyperlink to', { exact: true }).inputValue(),
        'nextSlide',
      );
      await actions.getByLabel('Hyperlink to', { exact: true }).selectOption('endShow');
      await actions.getByRole('button', { name: 'Cancel', exact: true }).click();
      await page.locator('[data-edit="action-settings"]').click();
      await actions.waitFor();
      assert.equal(
        await actions.getByLabel('Hyperlink to', { exact: true }).inputValue(),
        'nextSlide',
      );
      await actions.getByLabel('Hyperlink to', { exact: true }).selectOption('url');
      const urlDialog = page.getByRole('dialog', { name: 'Hyperlink to URL', exact: true });
      await urlDialog.getByLabel('URL:', { exact: true }).fill('https://example.com/button');
      await urlDialog.getByRole('button', { name: 'OK', exact: true }).click();
      await actions.getByRole('button', { name: 'OK', exact: true }).click();
      await page.waitForFunction(
        () =>
          window.office.getState().editor.slides[0].shapes.at(-1).link?.action.url ===
          'https://example.com/button',
      );
      await page.waitForFunction(
        () =>
          !document.querySelector('[data-edit="action-settings"]').disabled &&
          !document.querySelector('.action-dialog'),
      );
      await page.locator('.shape-hit').last().click({ button: 'right' });
      await page.getByRole('menuitem', { name: 'Action Settings...', exact: true }).click();
      await actions.getByRole('radio', { name: 'None', exact: true }).check();
      await actions.getByLabel('Play sound:', { exact: true }).check();
      const wav = Buffer.alloc(46);
      wav.write('RIFF', 0);
      wav.writeUInt32LE(38, 4);
      wav.write('WAVEfmt ', 8);
      wav.writeUInt32LE(16, 16);
      wav.writeUInt16LE(1, 20);
      wav.writeUInt16LE(1, 22);
      wav.writeUInt32LE(8000, 24);
      wav.writeUInt32LE(16000, 28);
      wav.writeUInt16LE(2, 32);
      wav.writeUInt16LE(16, 34);
      wav.write('data', 36);
      wav.writeUInt32LE(2, 40);
      await actions.locator('[data-sound-file]').setInputFiles({
        name: 'Chime.wav',
        mimeType: 'audio/wav',
        buffer: wav,
      });
      await page.waitForFunction(() => document.querySelector('[data-sound]').value === 'embedded');
      await actions.getByLabel('Play sound:', { exact: true }).uncheck();
      await actions.getByRole('tab', { name: 'Mouse Over', exact: true }).click();
      await actions.getByRole('tab', { name: 'Mouse Click', exact: true }).click();
      assert.equal(await actions.getByLabel('Play sound:', { exact: true }).isChecked(), false);
      assert.equal(await actions.getByLabel('Sound', { exact: true }).inputValue(), 'embedded');
      assert.equal(await actions.getByLabel('Sound', { exact: true }).isDisabled(), true);
      await actions.getByLabel('Play sound:', { exact: true }).check();
      await actions.getByLabel('Sound', { exact: true }).selectOption('stop');
      await actions.getByLabel('Sound', { exact: true }).selectOption('embedded');
      await actions.getByRole('tab', { name: 'Mouse Over', exact: true }).click();
      await actions.getByLabel('Play sound:', { exact: true }).check();
      await actions.getByLabel('Sound', { exact: true }).selectOption('stop');
      await actions.getByRole('radio', { name: 'Hyperlink to:', exact: true }).check();
      await actions.getByLabel('Hyperlink to', { exact: true }).selectOption('endShow');
      await actions.getByRole('tab', { name: 'Mouse Click', exact: true }).click();
      assert.equal(
        await actions.getByRole('radio', { name: 'None', exact: true }).isChecked(),
        true,
      );
      assert.equal(await actions.getByLabel('Sound', { exact: true }).inputValue(), 'embedded');
      await actions.getByRole('tab', { name: 'Mouse Over', exact: true }).click();
      assert.equal(await actions.getByLabel('Sound', { exact: true }).inputValue(), 'stop');
      assert.equal(
        await actions.getByLabel('Hyperlink to', { exact: true }).inputValue(),
        'endShow',
      );
      await actions.getByRole('button', { name: 'OK', exact: true }).click();
      await page.waitForFunction(
        () =>
          window.office.getState().editor.slides[0].shapes.at(-1).link === null &&
          window.office.getState().editor.slides[0].shapes.at(-1).hoverLink?.action.kind ===
            'endShow',
      );
      assert.deepEqual(
        await page.evaluate(
          () => window.office.getState().editor.slides[0].shapes.at(-1).actionSounds,
        ),
        {
          click: {
            sound: { name: 'Chime.wav', base64: wav.toString('base64') },
            stopPrevious: false,
          },
          hover: { sound: null, stopPrevious: true },
        },
      );
      await page.waitForFunction(
        () => !document.querySelector('[data-edit="action-settings"]').disabled,
      );
      const soundRevision = await page.evaluate(() => window.office.getState().revision);
      await page.locator('.shape-hit').last().click({ button: 'right' });
      await page.getByRole('menuitem', { name: 'Action Settings...', exact: true }).click();
      assert.equal(await actions.getByLabel('Sound', { exact: true }).inputValue(), 'embedded');
      await actions.locator('[data-sound-file]').setInputFiles({
        name: 'invalid.wav',
        mimeType: 'audio/wav',
        buffer: Buffer.from('invalid audio'),
      });
      await actions.getByRole('alert').filter({ hasText: 'Choose a WAV sound.' }).waitFor();
      assert.equal(await actions.getByLabel('Sound', { exact: true }).inputValue(), 'embedded');
      await actions.getByLabel('Play sound:', { exact: true }).uncheck();
      await actions.getByRole('button', { name: 'Cancel', exact: true }).click();
      assert.equal(await page.evaluate(() => window.office.getState().revision), soundRevision);
      assert.equal(
        await page.evaluate(
          () =>
            window.office.getState().editor.slides[0].shapes.at(-1).actionSounds.click.sound.name,
        ),
        'Chime.wav',
      );
      // Observe media lifecycle without depending on the machine's audio output.
      await page.evaluate(() => {
        window.soundEvents = [];
        window.realAudio = window.Audio;
        window.Audio = class extends EventTarget {
          constructor(src) {
            super();
            this.src = src;
          }
          play() {
            window.soundEvents.push(['play', this.src]);
            return Promise.resolve();
          }
          pause() {
            window.soundEvents.push(['pause']);
          }
          removeAttribute() {}
          load() {}
        };
      });
      const soundShape = page.locator('#slide [data-click-sound]').last();
      await soundShape.dispatchEvent('click');
      assert.deepEqual(await page.evaluate(() => window.soundEvents), []);
      await page.locator('#status-present').click();
      const actionCount = await page
        .locator('#slide a[href],#slide [data-click-sound],#slide [data-click-stop-sound]')
        .count();
      let soundFocused = false;
      for (let i = 0; i < actionCount; i++) {
        await page.keyboard.press('Tab');
        soundFocused = await page.evaluate(() =>
          document
            .querySelector('#slide')
            .shadowRoot.activeElement?.hasAttribute('data-click-sound'),
        );
        if (soundFocused) break;
      }
      assert.equal(soundFocused, true);
      await page.keyboard.press('Enter');
      await page.keyboard.press('Enter');
      assert.equal(await page.evaluate(() => window.office.getState().index), 0);
      assert.equal(await page.evaluate(() => window.soundEvents.length), 2);
      assert.equal(
        await page.evaluate(() => window.soundEvents[0][1]),
        'data:audio/wav;base64,' + wav.toString('base64'),
      );
      await soundShape.dispatchEvent('pointerover', {
        pointerType: 'mouse',
        clientX: 41,
        clientY: 51,
      });
      await page.waitForFunction(() => !window.office.getState().presenting);
      assert.deepEqual(await page.evaluate(() => window.soundEvents.map((event) => event[0])), [
        'play',
        'play',
        'pause',
        'pause',
      ]);
      await page.locator('#status-present').click();
      await soundShape.dispatchEvent('click');
      await page.keyboard.press('Escape');
      await page.waitForFunction(() => !window.office.getState().presenting);
      assert.equal(await page.evaluate(() => window.soundEvents.at(-1)[0]), 'pause');
      await page.evaluate(() => {
        window.Audio = window.realAudio;
      });
      await page.evaluate(async (base64) => {
        const audio = new Audio('data:audio/wav;base64,' + base64);
        await new Promise((resolve, reject) => {
          const timeout = setTimeout(() => reject(new Error('WAV playback did not end')), 5000);
          audio.onended = () => {
            clearTimeout(timeout);
            resolve();
          };
          audio.onerror = () => {
            clearTimeout(timeout);
            reject(new Error('WAV decode failed'));
          };
          audio.play().catch((error) => {
            clearTimeout(timeout);
            reject(error);
          });
        });
      }, wav.toString('base64'));

      await page.locator('[data-edit="undo"]').click();
      await page.waitForFunction(
        () =>
          window.office.getState().editor.slides[0].shapes.at(-1).link?.action.url ===
          'https://example.com/button',
      );
      assert.equal(
        await page.evaluate(
          () => window.office.getState().editor.slides[0].shapes.at(-1).hoverLink,
        ),
        null,
      );
      await page.locator('[data-edit="undo"]').click();
      await page.waitForFunction(
        () =>
          window.office.getState().editor.slides[0].shapes.at(-1).link?.action.kind === 'nextSlide',
      );
      await page.reload();
      await page.waitForFunction(() =>
        window.office
          .getState()
          .editor?.slides[0].shapes.some(
            (s) => s.preset === 'actionButtonForwardNext' && s.link?.action.kind === 'nextSlide',
          ),
      );
      await page.locator('#present').click();
      await page.waitForFunction(() => window.office.getState().presenting);
      const hotspotHrefs = await page
        .locator('#slide a[href]')
        .evaluateAll((links) => links.map((link) => link.getAttribute('href')));
      assert.ok(hotspotHrefs.length >= 2);
      const focusedHref = () =>
        page.evaluate(() =>
          document.querySelector('#slide').shadowRoot.activeElement?.getAttribute('href'),
        );
      for (const href of hotspotHrefs) {
        await page.keyboard.press('Tab');
        assert.equal(await focusedHref(), href);
        assert.equal(await page.evaluate(() => window.office.getState().index), 0);
      }
      await page.keyboard.press('Tab');
      assert.equal(await focusedHref(), hotspotHrefs[0]);
      await page.keyboard.press('Shift+Tab');
      assert.equal(await focusedHref(), hotspotHrefs.at(-1));
      assert.equal(hotspotHrefs.at(-1), '#slide-nextSlide');
      await page.keyboard.press('Enter');
      await page.waitForFunction(() => window.office.getState().index === 1);
      // A slide without hotspots retains slide focus instead of tabbing into browser chrome.
      await page.keyboard.press('Tab');
      assert.equal(await page.evaluate(() => window.office.getState().index), 1);
      assert.equal(await page.evaluate(() => window.office.getState().presenting), true);
      await page.keyboard.press('Escape');
      await page.evaluate(() => window.office.selectSlide(0, true));
      await page.locator('#tab-home').click();
      await page.locator('[data-edit="selection"]:visible').first().click();
      const actionName = await page.evaluate(
        () => window.office.getState().editor.slides[0].shapes.at(-1).name,
      );
      await page.getByRole('button', { name: `Hide ${actionName}`, exact: true }).click();
      await page.waitForFunction(
        () =>
          window.office.getState().editor.slides[0].shapes.at(-1).hidden &&
          !window.office.getState().building,
      );
      await page.locator('#present').click();
      await page.waitForFunction(() => window.office.getState().presenting);
      const visibleHrefs = await page
        .locator('#slide a[href]')
        .evaluateAll((links) => links.map((link) => link.getAttribute('href')));
      assert.deepEqual(visibleHrefs, hotspotHrefs.slice(0, -1));
      for (const href of [...visibleHrefs, visibleHrefs[0]]) {
        await page.keyboard.press('Tab');
        assert.equal(await focusedHref(), href);
      }
      assert.equal(await page.evaluate(() => window.office.getState().index), 0);
      await page.keyboard.press('Escape');
      await page.getByRole('button', { name: `Show ${actionName}`, exact: true }).click();
      await page.waitForFunction(
        () =>
          !window.office.getState().editor.slides[0].shapes.at(-1).hidden &&
          !window.office.getState().building,
      );
      await page.locator('#present').click();
      await page.waitForFunction(() => window.office.getState().presenting);
      assert.deepEqual(
        await page
          .locator('#slide a[href]')
          .evaluateAll((links) => links.map((link) => link.getAttribute('href'))),
        hotspotHrefs,
      );
      await page.keyboard.press('Escape');
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      proc.kill('SIGTERM');
      await new Promise((resolve) =>
        proc.exitCode !== null ? resolve() : proc.once('exit', resolve),
      );
      await rm(dir, { recursive: true, force: true });
    }
  },
);
