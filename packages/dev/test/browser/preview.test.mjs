import { chromium } from 'playwright';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { mkdtemp, writeFile, readFile, rm } from 'node:fs/promises';
import { spawn } from 'node:child_process';
import assert from 'node:assert/strict';
test(
  'live edits preserve the viewport and patch only changed slides',
  { timeout: 60000 },
  async (t) => {
    const dir = await mkdtemp(join(tmpdir(), 'office-hmr-live-'));
    const file = dir + '/deck.tsx';
    await writeFile(
      dir + '/codex',
      `#!${process.execPath}
import { readFileSync, writeFileSync, existsSync } from 'node:fs';
let input='';for await(const chunk of process.stdin)input+=chunk;
writeFileSync('chat-prompt.txt',input);
let source=readFileSync('deck.tsx','utf8');
if(!existsSync('repair-started')) {
 writeFileSync('repair-started','1');
 source=source.replace('</Slide>', '<Bullets /></Slide>');
} else {
 source=source.replace('<Bullets />','').replace("i===2?'':","i===2?'ChatEdited':");
 if(input.includes('Bullets is not defined'))writeFileSync('repair-observed','1');
}
writeFileSync('deck.tsx',source);
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
 writeFileSync('terminal-context-'+hook.url.split('/')[4]+'.json',await response.text());
 console.log('Model menu: '+data.replaceAll('\\x1b[13;2u','').trim());
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
      const chatWidth = () =>
        page.locator('#chat').evaluate((el) => el.getBoundingClientRect().width);
      const initialWidth = await chatWidth();
      const resizer = page.getByRole('separator', { name: 'Chat width' });
      const handle = await resizer.boundingBox();
      await page.mouse.move(handle.x + handle.width / 2, handle.y + 100);
      await page.mouse.down();
      await page.mouse.move(handle.x + handle.width / 2 - 100, handle.y + 100);
      await page.mouse.up();
      assert.ok(Math.abs((await chatWidth()) - initialWidth - 100) < 2);
      await resizer.press('ArrowLeft');
      assert.ok(Math.abs((await chatWidth()) - initialWidth - 110) < 2);
      assert.equal(await page.locator('#count').textContent(), 'Slide 3 of 50');
      await page.reload();
      assert.ok(Math.abs((await chatWidth()) - initialWidth - 110) < 2);
      await resizer.press('End');
      assert.ok(await page.locator('main').evaluate((el) => el.clientWidth >= 240));
      await page.setViewportSize({ width: 800, height: 800 });
      assert.ok(await page.locator('main').evaluate((el) => el.clientWidth >= 240));
      await page.setViewportSize({ width: 600, height: 800 });
      assert.equal(await resizer.isVisible(), false);
      assert.equal(Math.round(await chatWidth()), 600);
      await page.setViewportSize({ width: 1280, height: 800 });
      await resizer.dblclick();
      assert.ok(Math.abs((await chatWidth()) - initialWidth) < 2);
      await page.getByRole('button', { name: 'Slide 3', exact: true }).click();

      let agent = page.frames().find((frame) => /\/agents\//.test(frame.url()));
      const agentPath = new URL(agent.url()).pathname;
      await agent.locator('#terminal-start').click();
      await agent.waitForFunction(() =>
        document.querySelector('#terminal').textContent.includes('Claude Code terminal ready'),
      );
      const assertTerminalFits = async (frame) => {
        await frame.waitForFunction(() => {
          const host = document.querySelector('#terminal').getBoundingClientRect();
          const screen = document.querySelector('.xterm-screen').getBoundingClientRect();
          return screen.width > 0 && screen.right <= host.right && screen.bottom <= host.bottom;
        });
      };
      await assertTerminalFits(agent);
      await resizer.press('ArrowLeft');
      await assertTerminalFits(agent);
      await resizer.press('ArrowRight');
      await assertTerminalFits(agent);
      await page.screenshot({ path: '/tmp/office-kit-studio.png' });
      const inputs = [];
      page.on('request', (request) => {
        if (request.url().endsWith('/terminal/input')) inputs.push(request.postDataJSON().data);
      });
      const newlineRequest = page.waitForRequest((request) =>
        request.url().endsWith('/terminal/input'),
      );
      await agent.locator('.xterm-helper-textarea').press('Shift+Enter');
      await newlineRequest;
      assert.deepEqual(inputs, ['\x1b[13;2u']);
      await agent.locator('.xterm-helper-textarea').pressSequentially('/model');
      await agent.locator('.xterm-helper-textarea').press('Enter');
      await agent.waitForFunction(() =>
        document.querySelector('#terminal').textContent.includes('Model menu: /model'),
      );
      await agent.locator('.xterm-helper-textarea').press('ArrowDown');
      assert.equal(await page.locator('#count').textContent(), 'Slide 3 of 50');
      await page.reload();
      agent = page.frames().find((frame) => frame.url().endsWith(agentPath));
      await agent.waitForFunction(() =>
        document.querySelector('#terminal').textContent.includes('Model menu: /model'),
      );
      assert.equal(await agent.locator('#terminal-start').isVisible(), false);
      const blocked = await page.request.post(url + agentPath + '/chat', {
        headers: { origin: url },
        data: { provider: 'codex', message: 'conflicting edit', slide: 2, revision: 0 },
      });
      assert.equal(blocked.status(), 409);
      // Split without interrupting the running first terminal; start an independent second.
      await page.getByRole('button', { name: 'Split down', exact: true }).first().click();
      await page.locator('iframe').nth(1).waitFor();
      const secondPath = await page.locator('iframe').nth(1).getAttribute('src');
      await page.waitForFunction(() =>
        document.querySelectorAll('iframe')[1].contentDocument?.querySelector('#terminal-start'),
      );
      const second = page.frames().find((frame) => frame.url().endsWith(secondPath));
      await second.locator('#terminal-start').click();
      await second.waitForFunction(() =>
        document.querySelector('#terminal').textContent.includes('Claude Code terminal ready'),
      );
      assert.equal(await agent.locator('#terminal-stop').isVisible(), true);
      await assertTerminalFits(agent);
      await assertTerminalFits(second);
      await page.getByRole('button', { name: 'Slide 3', exact: true }).click();
      await second.waitForFunction(
        () => document.querySelector('#chat-context').textContent === 'Slide 3 of 50',
      );
      await second.locator('.xterm-helper-textarea').pressSequentially('second-agent');
      await second.locator('.xterm-helper-textarea').press('Enter');
      await second.waitForFunction(() =>
        document.querySelector('#terminal').textContent.includes('Model menu: second-agent'),
      );
      assert.ok(!(await agent.locator('#terminal').textContent()).includes('second-agent'));
      const secondContext = await readFile(
        dir + '/terminal-context-' + secondPath.split('/').at(-1) + '.json',
        'utf8',
      );
      assert.match(JSON.parse(secondContext).hookSpecificOutput.additionalContext, /"slide":3/);
      // A third Codex pane can work while both Claude sessions are running.
      await page.getByRole('button', { name: 'Split right', exact: true }).nth(1).click();
      await page.waitForFunction(() =>
        document.querySelectorAll('iframe')[2]?.contentDocument?.querySelector('#chat-provider'),
      );
      const thirdPath = await page.locator('iframe').nth(2).getAttribute('src');
      const third = page.frames().find((frame) => frame.url().endsWith(thirdPath));
      await assertTerminalFits(second);
      await third.selectOption('#chat-provider', 'codex');
      await third.locator('#chat-input').fill('Parallel edit');
      await third.locator('#chat-send').click();
      await third.waitForFunction(
        () => document.querySelector('#chat-status').textContent === 'Done',
      );
      assert.equal(await agent.locator('#terminal-stop').isVisible(), true);
      assert.equal(await second.locator('#terminal-stop').isVisible(), true);
      assert.equal(
        (await (await page.request.get(url + agentPath + '/chat')).json()).messages.length,
        0,
      );
      assert.equal(await readFile(dir + '/repair-observed', 'utf8'), '1');
      await third.locator('#chat-input').fill('independent draft');
      await page.getByRole('button', { name: 'Split down', exact: true }).nth(2).click();
      assert.equal(await third.locator('#chat-input').inputValue(), 'independent draft');
      assert.equal(await page.locator('.agent-pane').count(), 4);
      assert.equal(
        await page.getByRole('button', { name: 'Split right', exact: true }).first().isDisabled(),
        true,
      );
      await page.screenshot({ path: '/tmp/office-kit-splits.png' });
      const divider = page.getByRole('separator', { name: 'Agent split', exact: true }).first();
      await divider.press('ArrowUp');
      assert.equal(await divider.getAttribute('aria-valuenow'), '45');
      await page.reload();
      await page.waitForFunction(() => document.querySelectorAll('iframe').length === 4);
      agent = page.frames().find((frame) => frame.url().endsWith(agentPath));
      await agent.waitForFunction(() =>
        document.querySelector('#terminal').textContent.includes('Model menu: /model'),
      );
      assert.equal(
        await page
          .getByRole('separator', { name: 'Agent split', exact: true })
          .first()
          .getAttribute('aria-valuenow'),
        '45',
      );
      for (let count = 4; count > 1; count--) {
        await page.getByRole('button', { name: 'Close', exact: true }).last().click();
        await page.waitForFunction(
          (count) => document.querySelectorAll('.agent-pane').length === count - 1,
          count,
        );
      }
      assert.equal((await page.request.get(url + secondPath + '/chat')).status(), 404);
      assert.equal(await agent.locator('#terminal-stop').isVisible(), true);
      await agent.locator('#terminal-stop').click();
      await agent.locator('#terminal-start').waitFor({ state: 'visible' });
      await page.getByRole('button', { name: 'Slide 3', exact: true }).click();
      await agent.selectOption('#chat-provider', 'codex');
      await agent.locator('#chat-input').fill('Make this slide clearer 日本語');
      await page.keyboard.press('ArrowLeft');
      assert.equal(await page.locator('#count').textContent(), 'Slide 3 of 50');
      await agent.locator('#chat-send').click();
      await agent.waitForFunction(
        () => document.querySelector('#chat-status').textContent === 'Done',
      );
      await page.waitForFunction(slideSvg, 'ChatEdited');
      const chat = await (await page.request.get(url + agentPath + '/chat')).json();
      assert.equal(chat.messages[0].context.slide, 3);
      assert.equal(chat.messages[0].context.count, 50);
      assert.match(chat.messages[0].context.text, /Slide 2/);
      assert.match(chat.messages[1].text, /Updated the focused slide/);
      assert.equal(await agent.locator('#messages strong').textContent(), 'Verified');
      await page.reload();
      agent = page.frames().find((frame) => frame.url().endsWith(agentPath));
      await agent.selectOption('#chat-provider', 'codex');
      await agent.waitForFunction(() =>
        document.querySelector('#messages').textContent.includes('Updated the focused slide'),
      );
      await page.getByRole('button', { name: 'Slide 3', exact: true }).click();
      await page.locator('#toggle-chat').click();
      assert.equal(await page.locator('#chat').isVisible(), false);
      await page.locator('#toggle-chat').click();
      assert.equal(await page.locator('#chat').isVisible(), true);
      await page.setViewportSize({ width: 640, height: 800 });
      assert.equal(await agent.locator('#chat-input').isVisible(), true);
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
      // Text editing owns arrow keys until the user exits the in-place field.
      await page.locator('#slide p').first().dblclick();
      assert.ok(
        (
          await page.getByRole('textbox', { name: 'Edit slide text', exact: true }).inputValue()
        ).includes('Slide 2'),
      );
      await page.keyboard.press('ArrowRight');
      assert.equal(await page.locator('#count').textContent(), 'Slide 3 of 50');
      await page.keyboard.press('Escape');
      await page.locator('#slide').focus();
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
