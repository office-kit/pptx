import { chromium } from 'playwright';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, writeFile, readFile, rm, readdir } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { spawn } from 'node:child_process';
import { fileURLToPath } from 'node:url';

test(
  'region instructions, exact text saves, screenshots, undo and stale selections',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-inline-'));
    const source = `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={8} height={1}>Hello world</Text></Slide><Slide><Text x={1} y={1} width={8} height={1}>Keep this</Text></Slide></Presentation>`;
    await writeFile(join(dir, 'deck.tsx'), source);
    await writeFile(
      join(dir, 'codex'),
      `#!${process.execPath}
const fs=require('node:fs');let input='';process.stdin.on('data',d=>input+=d);process.stdin.on('end',()=>{fs.appendFileSync('requests.jsonl',JSON.stringify({input,args:process.argv})+'\\n');if(input.includes('Move down')&&!fs.readFileSync('deck.tsx','utf8').includes('y={1.2}'))fs.writeFileSync('deck.tsx',fs.readFileSync('deck.tsx','utf8').replace('y={1}','y={1.2}'));console.log(JSON.stringify({type:'item.completed',item:{type:'agent_message',text:'Reviewed'}}));console.log(JSON.stringify({type:'turn.completed'}));});`,
      { mode: 0o755 },
    );
    const proc = spawn(
      process.execPath,
      [
        fileURLToPath(new URL('../../dist/cli.mjs', import.meta.url)),
        'dev',
        join(dir, 'deck.tsx'),
        '--port',
        '0',
      ],
      { env: { ...process.env, PATH: dir + ':' + process.env.PATH } },
    );
    let browser;
    try {
      const url = await new Promise((resolve, reject) => {
        let out = '';
        proc.stdout.on('data', (data) => {
          out += data;
          const m = out.match(/Preview: (http:\/\/\S+)/);
          if (m) resolve(m[1]);
        });
        proc.once('error', reject);
        proc.once('exit', (code) => reject(new Error('exit ' + code)));
      });
      browser = await chromium.launch({
        channel: process.env.PLAYWRIGHT_CHANNEL || undefined,
        headless: true,
      });
      const page = await browser.newPage({ viewport: { width: 1400, height: 900 } });
      const errors = [];
      page.on('pageerror', (e) => errors.push(e.message));
      await page.goto(url);
      await page.locator('#slide p').first().waitFor();
      const agent = page.frames().find((f) => /\/agents\//.test(f.url()));
      await agent.selectOption('#chat-provider', 'codex');
      assert.equal(await page.getByRole('button', { name: 'Select area' }).count(), 0);
      const bounds = await page.locator('#slide').boundingBox();
      await page.mouse.move(bounds.x + bounds.width * 0.04, bounds.y + bounds.height * 0.05);
      await page.mouse.down();
      await page.mouse.move(bounds.x + bounds.width * 0.9, bounds.y + bounds.height * 0.5);
      await page.mouse.up();
      await page.getByRole('textbox', { name: 'Selection edit' }).fill('Move down');
      await page.getByRole('textbox', { name: 'Selection edit' }).press('Shift+Enter');
      assert.equal(
        await page.getByRole('textbox', { name: 'Selection edit' }).inputValue(),
        'Move down\n',
      );
      await page.getByRole('button', { name: 'Apply with AI', exact: true }).click();
      await agent.waitForFunction(
        () => document.querySelector('#chat-status').textContent === 'Done',
      );
      let requests = (await readFile(join(dir, 'requests.jsonl'), 'utf8'))
        .trim()
        .split('\n')
        .map(JSON.parse);
      assert.match(requests[0].input, /fractions of the full slide/);
      assert.match(requests[0].input, /Hello world/);
      assert.equal(requests.length, 2);
      const image = requests[1].args[requests[1].args.indexOf('--image') + 1];
      assert.match(image, /slide-1\.png$/);
      assert.deepEqual([...(await readFile(image))].slice(0, 8), [137, 80, 78, 71, 13, 10, 26, 10]);
      await Promise.all([
        page.waitForResponse((r) => r.url().endsWith('/history/undo')),
        page.getByRole('button', { name: '↶ Undo', exact: true }).click(),
      ]);
      assert.equal(await readFile(join(dir, 'deck.tsx'), 'utf8'), source);
      await Promise.all([
        page.waitForResponse((r) => r.url().endsWith('/history/redo')),
        page.getByRole('button', { name: '↷ Redo', exact: true }).click(),
      ]);
      assert.match(await readFile(join(dir, 'deck.tsx'), 'utf8'), /y=\{1.2\}/);
      await page.locator('#slide p').first().click();
      assert.equal(await page.getByRole('form', { name: 'Edit selection' }).isVisible(), true);
      await page.locator('#slide p').first().dblclick();
      await page
        .getByRole('textbox', { name: 'Edit slide text', exact: true })
        .fill('Hello 日本語');
      await page.screenshot({ path: '/tmp/office-kit-inline-edit.png' });
      await page.getByRole('button', { name: 'Save text', exact: true }).click();
      await page.waitForFunction(() =>
        document.querySelector('.slide-edit-panel small').textContent.includes('Text saved ·'),
      );
      await agent.waitForFunction(
        () => document.querySelector('#chat-status').textContent === 'Done',
      );
      assert.match(await readFile(join(dir, 'deck.tsx'), 'utf8'), /\{"Hello 日本語"\}/);
      assert.equal(await page.locator('#slide p').first().textContent(), 'Hello 日本語');
      requests = (await readFile(join(dir, 'requests.jsonl'), 'utf8'))
        .trim()
        .split('\n')
        .map(JSON.parse);
      assert.ok(requests.at(-1).args.includes('--image'));
      assert.ok((await readdir(join(dir, '.office-kit', 'reviews'))).length >= 2);
      await Promise.all([
        page.waitForResponse((r) => r.url().endsWith('/history/undo')),
        page.getByRole('button', { name: '↶ Undo', exact: true }).click(),
      ]);
      await page.waitForFunction(
        () =>
          document.querySelector('#slide').shadowRoot.querySelector('p').textContent ===
          'Hello world',
      );
      await page.locator('#slide').focus();
      await page.keyboard.press('Control+Shift+z');
      await page.waitForFunction(
        () =>
          document.querySelector('#slide').shadowRoot.querySelector('p').textContent ===
          'Hello 日本語',
      );
      await page.waitForFunction(() => !document.querySelector('[data-undo]').disabled);
      await page.keyboard.press('Control+z');
      await page.waitForFunction(
        () =>
          document.querySelector('#slide').shadowRoot.querySelector('p').textContent ===
          'Hello world',
      );
      await page.locator('#slide p').first().click();
      await page.getByRole('button', { name: 'Slide 2', exact: true }).click();
      assert.equal(await page.getByRole('form', { name: 'Edit selection' }).isVisible(), false);
      const bad = await page.request.post(url + '/text-edit', {
        headers: { origin: 'https://evil.example' },
        data: {},
      });
      assert.equal(bad.status(), 403);
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      proc.kill('SIGTERM');
      await rm(dir, { recursive: true, force: true });
    }
  },
);
