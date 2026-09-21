import { test } from 'node:test';
import assert from 'node:assert/strict';
import { createServer } from 'node:http';
import { mkdtemp, writeFile, readFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { createTerminal } from '../src/terminal.ts';

const pause = () => new Promise((resolve) => setTimeout(resolve, 20));
test(
  'interactive Claude preserves CLI settings, receives current focus through a hook and reconnects',
  { timeout: 15000 },
  async (t) => {
    const directory = await mkdtemp(join(tmpdir(), 'pptx-terminal-'));
    const previousPath = process.env.PATH;
    process.env.PATH = directory + ':' + previousPath;
    await writeFile(
      join(directory, 'claude'),
      `#!${process.execPath}
const {writeFileSync}=require('node:fs');
writeFileSync('args.json',JSON.stringify(process.argv.slice(2)));
console.log('Interactive Claude ready');
process.stdin.setEncoding('utf8');
process.stdin.on('data',async data=>{
 const settings=JSON.parse(process.argv[3]);
 const hook=settings.hooks.UserPromptSubmit[0].hooks[0];
 const response=await fetch(hook.url,{method:'POST',headers:hook.headers,body:'{}'});
 writeFileSync('context.json',await response.text());
 console.log('REPLY '+data.trim());
});
`,
      { mode: 0o755 },
    );
    let busy = true;
    let buildError = 'ReferenceError: Bullets is not defined';
    const terminal = createTerminal(
      join(directory, 'deck.tsx'),
      () => busy,
      '',
      async () => buildError,
    );
    const server = createServer(
      (req, res) =>
        void terminal.handle(req, res, (slide, revision) => {
          if (revision !== 7) throw new Error('Preview changed');
          return {
            slide: slide === null ? null : slide + 1,
            count: 10,
            revision,
            text: 'Revenue 日本語',
            entry: 'deck.tsx',
            files: ['slides/revenue.tsx'],
            buildError: null,
          };
        }),
    );
    await new Promise((done) => server.listen(0, '127.0.0.1', done));
    const url = `http://127.0.0.1:${server.address().port}`;
    t.after(async () => {
      await terminal.close();
      server.closeAllConnections();
      await new Promise((done) => server.close(done));
      process.env.PATH = previousPath;
      await rm(directory, { recursive: true, force: true });
    });
    const post = (path, body, origin = url) =>
      fetch(url + '/terminal/' + path, {
        method: 'POST',
        headers: { origin, 'Content-Type': 'application/json' },
        body: JSON.stringify({ client: 'tab-one', ...body }),
      });
    const start = { cols: 80, rows: 24, slide: 2, revision: 7 };
    assert.equal((await post('start', start, 'https://example.com')).status, 403);
    assert.equal((await post('start', { ...start, cols: 0 })).status, 400);
    assert.equal((await post('start', { ...start, revision: 6 })).status, 400);
    assert.match(await (await post('start', start)).text(), /Codex edit/);
    busy = false;
    const response = await post('start', start);
    assert.equal(response.status, 200, await response.text());
    const waitFile = async (name, predicate = () => true) => {
      for (let i = 0; i < 150; i++) {
        try {
          const value = JSON.parse(await readFile(join(directory, name), 'utf8'));
          if (predicate(value)) return value;
        } catch (error) {
          if (error.code !== 'ENOENT') throw error;
        }
        await pause();
      }
      throw new Error('File not written: ' + name);
    };
    const args = await waitFile('args.json');
    assert.deepEqual(
      args.filter((arg) => arg.startsWith('--')),
      ['--settings'],
    );
    assert.equal(
      (await post('input', { data: 'blocked\r', slide: 2, revision: 7, client: 'another-tab' }))
        .status,
      409,
    );
    assert.equal((await post('input', { data: 'stale\r', slide: 2, revision: 6 })).status, 400);
    assert.equal(
      (await post('input', { data: 'このスライド\r', slide: 2, revision: 7 })).status,
      200,
    );
    let context = await waitFile('context.json');
    assert.match(context.hookSpecificOutput.additionalContext, /"slide":3/);
    assert.match(context.hookSpecificOutput.additionalContext, /Revenue 日本語/);
    assert.equal(
      (await post('input', { data: '次のスライド\r', slide: 5, revision: 7 })).status,
      200,
    );
    context = await waitFile('context.json', (value) =>
      value.hookSpecificOutput.additionalContext.includes('"slide":6'),
    );
    assert.match(context.hookSpecificOutput.additionalContext, /NOT a restriction/);
    const hook = JSON.parse(args[1]).hooks.Stop[0].hooks[0];
    const verify = async () =>
      (
        await fetch(hook.url, {
          method: 'POST',
          headers: hook.headers,
          body: '{}',
        })
      ).json();
    assert.equal((await post('verify', {})).status, 403);
    for (let attempt = 0; attempt < 3; attempt++) {
      const feedback = await verify();
      assert.equal(feedback.decision, 'block');
      assert.match(feedback.reason, /Bullets is not defined/);
    }
    assert.equal((await verify()).decision, undefined);
    buildError = null;
    assert.equal((await verify()).decision, undefined);
    assert.equal((await post('resize', { cols: 12, rows: 3 })).status, 200);
    assert.equal((await post('resize', { cols: 1, rows: 0 })).status, 400);
    assert.equal((await post('resize', { cols: 100, rows: 30 })).status, 200);
    assert.equal((await post('context', {})).status, 403);
    const controller = new AbortController();
    const stream = await fetch(url + '/terminal/events', { signal: controller.signal });
    const reader = stream.body.getReader();
    const replay = new TextDecoder().decode((await reader.read()).value);
    assert.match(replay, /Interactive Claude ready/);
    assert.match(replay, /event: replay/);
    controller.abort();
    assert.equal((await post('stop', {})).status, 200);
    assert.equal((await post('input', { data: 'x', slide: 2, revision: 7 })).status, 400);
    assert.equal((await post('start', { ...start, client: 'another-tab' })).status, 200);
  },
);
