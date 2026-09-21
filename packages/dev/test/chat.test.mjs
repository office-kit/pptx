import { test } from 'node:test';
import assert from 'node:assert/strict';
import { createServer, request as httpRequest } from 'node:http';
import { mkdtemp, writeFile, readFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { createChat } from '../src/chat.ts';

const pause = () => new Promise((done) => setTimeout(done, 20));
test('CLI chat includes focus/history, renders Markdown, serializes edits and stops', async (t) => {
  const directory = await mkdtemp(join(tmpdir(), 'pptx-chat-'));
  const previousPath = process.env.PATH;
  process.env.PATH = directory + ':' + previousPath;
  const script = `#!${process.execPath}
import {writeFileSync} from 'node:fs';
let input='';for await(const chunk of process.stdin)input+=chunk;
writeFileSync('prompt.txt',input);writeFileSync('args.json',JSON.stringify(process.argv));
if(input.includes('WAIT_FOREVER'))setInterval(()=>{},1000);
else {
 console.log(JSON.stringify({type:'item.completed',item:{type:'agent_message',text:input.includes('MARKDOWN')?'**Bold** and '+String.fromCharCode(96)+'code'+String.fromCharCode(96)+'\\n\\n- item\\n\\n<script>alert(1)</script>\\n\\n[unsafe](javascript:alert(1))':'Codex edit complete'}}));
 console.log(JSON.stringify(input.includes('FAIL_RESULT')?{type:'turn.failed',error:{message:'Result'}}:{type:'turn.completed'}));
}`;
  for (const provider of ['claude', 'codex'])
    await writeFile(join(directory, provider), script, { mode: 0o755 });
  const chat = createChat(join(directory, 'deck.tsx'), () => {});
  const focus = {
    slide: 3,
    count: 5,
    revision: 7,
    text: 'Revenue 日本語',
    entry: 'deck.tsx',
    files: ['slides/revenue.tsx'],
    buildError: null,
  };
  const server = createServer(
    (req, res) =>
      void chat.handle(req, res, (index, revision) => {
        assert.equal(index, 2);
        if (revision !== 7) throw new Error('Preview changed');
        return focus;
      }),
  );
  await new Promise((done) => server.listen(0, '127.0.0.1', done));
  const url = `http://127.0.0.1:${server.address().port}`;
  t.after(async () => {
    await chat.close();
    await new Promise((done) => server.close(done));
    process.env.PATH = previousPath;
    await rm(directory, { recursive: true, force: true });
  });
  const post = (body, path = '/chat', origin = url) =>
    fetch(url + path, {
      method: 'POST',
      headers: { origin, 'Content-Type': 'application/json' },
      body: JSON.stringify(body),
    });
  const request = (message, provider = 'codex') => ({ message, provider, slide: 2, revision: 7 });
  const state = async () => (await fetch(url + '/chat')).json();
  const finished = async () => {
    for (let i = 0; i < 150; i++) {
      const value = await state();
      if (!value.running) return value;
      await pause();
    }
    throw new Error('Chat did not finish');
  };
  assert.equal((await post(request('cross origin'), '/chat', 'https://example.com')).status, 403);
  assert.equal((await post({ ...request('stale'), revision: 6 })).status, 400);
  // Split inside a Japanese character to exercise HTTP's incremental UTF-8 decoder.
  const encoded = Buffer.from(JSON.stringify(request('この見出しを変更')));
  const split = encoded.indexOf(Buffer.from('こ')) + 1;
  const splitStatus = await new Promise((resolve, reject) => {
    const req = httpRequest(
      url + '/chat',
      {
        method: 'POST',
        headers: { origin: url, 'Content-Type': 'application/json' },
      },
      (res) => {
        res.resume();
        res.on('end', () => resolve(res.statusCode));
      },
    );
    req.on('error', reject);
    req.write(encoded.subarray(0, split));
    setTimeout(() => req.end(encoded.subarray(split)), 30);
  });
  assert.equal(splitStatus, 202);
  let result = await finished();
  assert.equal(result.status, 'Done');
  assert.equal(result.messages[1].text, 'Codex edit complete');
  let prompt = await readFile(join(directory, 'prompt.txt'), 'utf8');
  assert.match(prompt, /Revenue 日本語/);
  assert.match(prompt, /"slide":3/);
  assert.match(prompt, /slides\/revenue.tsx/);
  assert.match(prompt, /NOT a restriction/);
  assert.equal((await post(request('Now edit the entire deck', 'codex'))).status, 202);
  result = await finished();
  assert.equal(result.messages[3].text, 'Codex edit complete');
  prompt = await readFile(join(directory, 'prompt.txt'), 'utf8');
  assert.match(prompt, /この見出しを変更/);
  const args = JSON.parse(await readFile(join(directory, 'args.json'), 'utf8'));
  assert.ok(args.includes('workspace-write'));
  assert.ok(!args.includes('--dangerously-bypass-approvals-and-sandbox'));
  await post(request('WAIT_FOREVER'));
  assert.equal((await post(request('concurrent'))).status, 409);
  assert.equal((await post({}, '/chat/reset')).status, 409);
  await post({}, '/chat/stop');
  assert.match((await finished()).status, /Stopped/);
  await post({}, '/chat/reset');
  assert.deepEqual((await state()).messages, []);
  await post(request('FAIL_RESULT'));
  assert.equal((await finished()).status, 'Result');
  await post(request('MARKDOWN'));
  const rendered = (await finished()).messages.at(-1).html;
  assert.match(rendered, /<strong>Bold<\/strong>/);
  assert.match(rendered, /<code>code<\/code>/);
  assert.ok(!rendered.includes('<script>'));
  assert.ok(!rendered.includes('href="javascript:'));
  assert.equal((await post(request('legacy', 'claude'))).status, 400);
  await rm(join(directory, 'codex'));
  // Keep the real installed CLI out of PATH for this missing-command case.
  process.env.PATH = directory;
  await post(request('missing executable'));
  assert.match((await finished()).status, /Install it and log in/);
});
