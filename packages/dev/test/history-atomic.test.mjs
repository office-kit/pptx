import { test } from 'node:test';
import assert from 'node:assert/strict';
import { chmod, mkdir, mkdtemp, readFile, readdir, rm, stat, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { createHistory } from '../src/history.ts';

async function project(t) {
  const dir = await mkdtemp(join(tmpdir(), 'office-history-atomic-'));
  t.after(() => rm(dir, { recursive: true, force: true }));
  return dir;
}

test('an undo that cannot write every file leaves all files at their newer contents', async (t) => {
  // Directory permissions do not stop root, so the failure cannot be provoked.
  if (process.getuid?.() === 0) return t.skip('running as root');
  const dir = await project(t);
  const deck = join(dir, 'deck.tsx');
  const locked = join(dir, 'locked');
  const data = join(locked, 'data.json');
  await mkdir(locked);
  await writeFile(deck, 'v1');
  await writeFile(data, 'v1');
  const history = await createHistory(dir);
  await history.begin('agent', 'AI edit');
  await writeFile(deck, 'v2');
  await writeFile(data, 'v2');
  await history.end('agent');

  await chmod(locked, 0o555);
  try {
    await assert.rejects(history.move('undo'), { code: 'EACCES' });
    assert.equal(await readFile(deck, 'utf8'), 'v2');
    assert.equal(await readFile(data, 'utf8'), 'v2');
    for (const folder of [dir, locked])
      assert.deepEqual(
        (await readdir(folder)).filter((name) => name.endsWith('.tmp')),
        [],
      );
    assert.equal(history.state().undo, 'AI edit');
  } finally {
    // Restore before the directory cleanup registered in `project` runs.
    await chmod(locked, 0o755);
  }

  await history.move('undo');
  assert.equal(await readFile(deck, 'utf8'), 'v1');
  assert.equal(await readFile(data, 'utf8'), 'v1');
});

test('undo and redo keep the permissions of the files they replace', async (t) => {
  const dir = await project(t);
  const script = join(dir, 'tool.mjs');
  await writeFile(script, 'v1');
  await chmod(script, 0o755);
  const history = await createHistory(dir);
  await history.begin('agent', 'AI edit');
  await writeFile(script, 'v2');
  await history.end('agent');
  await history.move('undo');
  assert.equal(await readFile(script, 'utf8'), 'v1');
  assert.equal((await stat(script)).mode & 0o777, 0o755);
  await history.move('redo');
  assert.equal((await stat(script)).mode & 0o777, 0o755);
});
