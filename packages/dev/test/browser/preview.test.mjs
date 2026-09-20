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
    await writeFile(
      dir + '/codex',
      `#!${process.execPath}
import { readFileSync, writeFileSync } from 'node:fs';
let input='';for await(const chunk of process.stdin)input+=chunk;
writeFileSync('chat-prompt.txt',input);
writeFileSync('deck.tsx',readFileSync('deck.tsx','utf8').replace("i===2?'':","i===2?'ChatEdited':"));
console.log(JSON.stringify({type:'item.completed',item:{type:'agent_message',text:'Updated the focused slide. **Verified**'}}));
console.log(JSON.stringify({type:'turn.completed'}));
`,
      { mode: 0o755 },
    );
    await writeFile(
      dir + '/claude',
      `#!${process.execPath}
const {writeFileSync}=require('node:fs');
console.log('Claude Code terminal ready');
process.stdin.setEncoding('utf8');
process.stdin.on('data',async data=>{
 const hook=JSON.parse(process.argv[3]).hooks.UserPromptSubmit[0].hooks[0];
 const response=await fetch(hook.url,{method:'POST',headers:hook.headers,body:'{}'});
 writeFileSync('terminal-context.json',await response.text());
 console.log('Model menu: '+data.trim());
});
`,
      { mode: 0o755 },
    );
    const source = (changed = '', other = '', count = 50) =>
      `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation>{Array.from({length:${count}},(_,i)=><Slide><Text x={1} y={1} width={8} height={1}>Slide {i} {i===2?'${changed}':i===40?'${other}':''}</Text></Slide>)}</Presentation>`;
    await writeFile(file, source());
    const proc = spawn(
      process.execPath,
      [fileURLToPath(new URL('../../dist/cli.mjs', import.meta.url)), 'dev', file, '--port', '0'],
      { env: { ...process.env, PATH: dir + ':' + process.env.PATH } },
    );
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
      assert.equal(await page.getByText('What would you like to change?').count(), 0);
      await page.locator('#terminal-start').click();
      await page.waitForFunction(() =>
        document.querySelector('#terminal').textContent.includes('Claude Code terminal ready'),
      );
      await page.locator('.xterm-helper-textarea').pressSequentially('/model');
      await page.locator('.xterm-helper-textarea').press('Enter');
      await page.waitForFunction(() =>
        document.querySelector('#terminal').textContent.includes('Model menu: /model'),
      );
      await page.locator('.xterm-helper-textarea').press('ArrowDown');
      assert.equal(await page.locator('#count').textContent(), 'Slide 3 of 50');
      await page.reload();
      await page.waitForFunction(() =>
        document.querySelector('#terminal').textContent.includes('Model menu: /model'),
      );
      assert.equal(await page.locator('#terminal-start').isVisible(), false);
      const blocked = await page.request.post(url + '/chat', {
        headers: { origin: url },
        data: { provider: 'codex', message: 'conflicting edit', slide: 2, revision: 0 },
      });
      assert.equal(blocked.status(), 409);
      const other = await browser.newPage();
      await other.goto(url);
      await other.getByText('Open in another tab · view only').waitFor();
      assert.equal(await other.locator('#terminal-stop').isVisible(), false);
      await other.close();
      await page.locator('#terminal-stop').click();
      await page.locator('#terminal-start').waitFor({ state: 'visible' });
      await page.getByRole('button', { name: 'Slide 3', exact: true }).click();
      await page.selectOption('#chat-provider', 'codex');
      await page.locator('#chat-input').fill('Make this slide clearer 日本語');
      await page.keyboard.press('ArrowLeft');
      assert.equal(await page.locator('#count').textContent(), 'Slide 3 of 50');
      await page.locator('#chat-send').click();
      await page.waitForFunction(
        () => document.querySelector('#chat-status').textContent === 'Done',
      );
      await page.waitForFunction(slideSvg, 'ChatEdited');
      const chat = await (await page.request.get(url + '/chat')).json();
      assert.equal(chat.messages[0].context.slide, 3);
      assert.equal(chat.messages[0].context.count, 50);
      assert.match(chat.messages[0].context.text, /Slide 2/);
      assert.match(chat.messages[1].text, /Updated the focused slide/);
      assert.equal(await page.locator('#messages strong').textContent(), 'Verified');
      await page.reload();
      await page.selectOption('#chat-provider', 'codex');
      await page.waitForFunction(() =>
        document.querySelector('#messages').textContent.includes('Updated the focused slide'),
      );
      await page.getByRole('button', { name: 'Slide 3', exact: true }).click();
      await page.locator('#toggle-chat').click();
      assert.equal(await page.locator('#chat').isVisible(), false);
      await page.locator('#toggle-chat').click();
      assert.equal(await page.locator('#chat').isVisible(), true);
      await page.setViewportSize({ width: 640, height: 800 });
      assert.equal(await page.locator('#chat-input').isVisible(), true);
      assert.equal(await page.evaluate(() => document.body.scrollWidth <= innerWidth), true);
      await page.setViewportSize({ width: 1280, height: 800 });
      await writeFile(file, source());
      await page.waitForFunction(() => !state.slides[2].includes('ChatEdited'));
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
      await page.locator('#toggle-chat').click();
      await page.getByRole('button', { name: 'Present', exact: true }).click();
      await writeFile(file, source('Presenting'));
      await page.waitForFunction(slideSvg, 'Presenting');
      assert.ok(await page.locator('body').evaluate((el) => el.classList.contains('presenting')));
      assert.equal(await page.locator('#chat').isVisible(), false);
      assert.equal(
        await page.locator('main').evaluate((el) => el.clientWidth),
        await page.evaluate(() => innerWidth),
      );
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
