import { chromium } from 'playwright';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { mkdtemp, writeFile, rm } from 'node:fs/promises';
import { spawn } from 'node:child_process';
import assert from 'node:assert/strict';
test(
  'live edits preserve the viewport and patch only changed slides',
  { timeout: 30000 },
  async (t) => {
    const dir = await mkdtemp(join(tmpdir(), 'office-hmr-live-'));
    const file = dir + '/deck.tsx';
    const source = (changed = '', other = '', count = 50) =>
      `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation>{Array.from({length:${count}},(_,i)=><Slide><Text x={1} y={1} width={8} height={1}>Slide {i} {i===2?'${changed}':i===40?'${other}':''}</Text></Slide>)}</Presentation>`;
    await writeFile(file, source());
    const proc = spawn(process.execPath, [
      fileURLToPath(new URL('../../dist/cli.mjs', import.meta.url)),
      'dev',
      file,
      '--port',
      '0',
    ]);
    let browser;
    try {
      const url = await new Promise((res, rej) => {
        let out = '';
        proc.stdout.on('data', (d) => {
          out += d;
          const m = out.match(/Preview: (http:\/\/\S+)/);
          if (m) res(m[1]);
        });
        proc.stderr.on('data', (d) => process.stderr.write(d));
        proc.once('error', rej);
        proc.on('exit', (c) => rej(new Error('exit ' + c)));
      });
      browser = await chromium.launch({
        channel: process.env.PLAYWRIGHT_CHANNEL || undefined,
        headless: true,
      });
      const page = await browser.newPage({ viewport: { width: 1280, height: 800 } });
      const errors = [];
      page.on('pageerror', (e) => errors.push(e.message));
      await page.goto(url);
      await page.getByRole('button', { name: 'Slide 3', exact: true }).click();
      const slideSvg = (label) => {
        const svg = document.querySelector('#slide').shadowRoot.querySelector('svg');
        return svg !== null && (label === undefined || svg.textContent.includes(label));
      };
      await page.waitForFunction(slideSvg);
      await page.selectOption('#zoom', '2');
      await page.evaluate(() => {
        window.oldSvg = document.querySelector('#slide').shadowRoot.querySelector('svg');
        window.oldThumb = document.querySelectorAll('.thumbnail')[2];
        document.querySelector('#stage').scrollTop = 100;
        document.querySelector('.filmstrip').scrollTop = 300;
      });
      await writeFile(file, source('', 'Offscreen'));
      await page.waitForFunction(() => state.slides[40].includes('Offscreen'));
      assert.deepEqual(
        await page.evaluate(() => ({
          sameSvg: oldSvg === document.querySelector('#slide').shadowRoot.querySelector('svg'),
          sameThumb: oldThumb === document.querySelectorAll('.thumbnail')[2],
          zoom: document.querySelector('#zoom').value,
          stage: document.querySelector('#stage').scrollTop,
          film: document.querySelector('.filmstrip').scrollTop,
        })),
        { sameSvg: true, sameThumb: true, zoom: '2', stage: 100, film: 300 },
      );
      const times = [];
      for (let i = 0; i < 5; i++) {
        const label = 'Revision' + i;
        const start = performance.now();
        await writeFile(file, source(label, 'Offscreen'));
        await page.waitForFunction(slideSvg, label, { polling: 'raf' });
        times.push(Math.round(performance.now() - start));
      }
      await writeFile(file, 'export default <');
      await page.locator('#error').waitFor({ state: 'visible' });
      assert.ok(await page.evaluate(slideSvg, 'Revision4'));
      await writeFile(file, source('Recovered'));
      await page.locator('#error').waitFor({ state: 'hidden' });
      await page.waitForFunction(slideSvg, 'Recovered');
      assert.equal(await page.locator('#count').textContent(), 'Slide 3 of 50');
      // Slide text is real DOM inside the shadow root: select it and read the selection.
      await page.locator('#slide p').first().click({ clickCount: 3 });
      assert.ok((await page.evaluate(() => getSelection().toString())).includes('Slide 2'));
      await page.keyboard.press('ArrowRight');
      assert.equal(await page.locator('#count').textContent(), 'Slide 4 of 50');
      await page.keyboard.press('ArrowLeft');
      await page.getByRole('button', { name: 'Present', exact: true }).click();
      await writeFile(file, source('Presenting'));
      await page.waitForFunction(slideSvg, 'Presenting');
      assert.ok(await page.locator('body').evaluate((el) => el.classList.contains('presenting')));
      await page.keyboard.press('Escape');
      await page.context().setOffline(true);
      await writeFile(file, source('Reconnect'));
      await page.context().setOffline(false);
      await page.waitForFunction(slideSvg, 'Reconnect');
      await writeFile(file, source('', '', 2));
      await page.waitForFunction(
        () => document.querySelector('#count').textContent === 'Slide 2 of 2',
      );
      assert.equal(await page.locator('.thumbnail').count(), 2);
      await writeFile(file, source('', '', 0));
      await page.waitForFunction(
        () => document.querySelector('#count').textContent === 'No slides',
      );
      assert.equal(await page.evaluate(slideSvg), false);
      await writeFile(file, source('Restored'));
      await page.waitForFunction(
        () => document.querySelector('#count').textContent === 'Slide 1 of 50',
      );
      assert.equal(await page.locator('.thumbnail').count(), 50);
      assert.deepEqual(errors, []);
      t.diagnostic(
        JSON.stringify({
          saveToVisibleMs: times,
          offscreenIdentityAndScroll: 'passed',
          errorRecovery: 'passed',
        }),
      );
    } finally {
      await browser?.close();
      proc.kill('SIGTERM');
      await rm(dir, { recursive: true, force: true });
    }
  },
);
