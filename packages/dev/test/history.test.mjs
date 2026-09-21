import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, writeFile, readFile, rm, mkdir, symlink, truncate } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { createHistory } from '../src/history.ts';

async function project(t) {
  const dir = await mkdtemp(join(tmpdir(), 'office-history-'));
  t.after(() => rm(dir, { recursive: true, force: true }));
  await writeFile(join(dir, 'deck.tsx'), 'original');
  return { dir, file: join(dir, 'deck.tsx'), history: await createHistory(dir) };
}

test('AI turns group repairs and overlapping agents; restores created/deleted binary assets and source', async (t) => {
  const { dir, file, history } = await project(t);
  await history.begin('agent1', 'AI edit');
  await writeFile(file, 'broken source');
  await history.capture();
  await history.begin('agent2', 'AI edit');
  await assert.rejects(history.move('undo'), /Wait for all agents/);
  await writeFile(file, 'repaired source');
  const image = join(dir, 'new.png');
  await writeFile(image, Buffer.from([0, 1, 255, 13]));
  await history.end('agent1');
  assert.equal(history.state().busy, true);
  await history.end('agent2');
  assert.equal(history.state().undo, 'Concurrent edits');
  await history.move('undo');
  assert.equal(await readFile(file, 'utf8'), 'original');
  await assert.rejects(readFile(image), { code: 'ENOENT' });
  await history.move('redo');
  assert.equal(await readFile(file, 'utf8'), 'repaired source');
  assert.deepEqual(await readFile(image), Buffer.from([0, 1, 255, 13]));
  await history.begin('agent1', 'AI edit');
  await rm(image);
  await history.end('agent1');
  await history.move('undo');
  assert.deepEqual(await readFile(image), Buffer.from([0, 1, 255, 13]));
});

test('external saves become separate undo entries and invalidate redo, unchanged turns do not', async (t) => {
  const { file, history } = await project(t);
  await history.begin('agent', 'AI edit');
  await writeFile(file, 'AI');
  await history.end('agent');
  await writeFile(file, 'new user save');
  await history.move('undo');
  assert.equal(await readFile(file, 'utf8'), 'AI');
  await history.move('undo');
  assert.equal(await readFile(file, 'utf8'), 'original');
  await history.begin('agent', 'AI edit');
  await history.end('agent');
  await history.move('redo');
  assert.equal(await readFile(file, 'utf8'), 'AI');
  await writeFile(file, 'new branch');
  await history.capture();
  assert.equal(history.state().redo, null);
});

test('history excludes dependencies, generated reviews and symlinks, refuses symlink replacement', async (t) => {
  const { dir, file, history } = await project(t);
  await mkdir(join(dir, '.office-kit'));
  await writeFile(join(dir, '.office-kit', 'review.json'), 'private');
  await mkdir(join(dir, 'node_modules'));
  await writeFile(join(dir, 'node_modules', 'package.json'), 'unchanged');
  await history.capture();
  assert.equal(history.state().undo, null);
  await history.begin('agent', 'AI edit');
  await writeFile(file, 'AI');
  await history.end('agent');
  const outside = await mkdtemp(join(tmpdir(), 'office-outside-'));
  t.after(() => rm(outside, { recursive: true, force: true }));
  const other = join(outside, 'deck.tsx');
  await writeFile(other, 'untouched');
  await rm(file);
  await symlink(other, file);
  await assert.rejects(history.move('undo'), /changed type/);
  assert.equal(await readFile(other, 'utf8'), 'untouched');
});

test('oversized projects keep the preview available and expose a recoverable history error', async (t) => {
  const { dir } = await project(t);
  const asset = join(dir, 'large.png');
  await writeFile(asset, '');
  await truncate(asset, 65 * 1024 * 1024);
  const history = await createHistory(dir);
  assert.match(history.state().error, /64 MiB/);
  await assert.rejects(history.begin('agent', 'AI edit'), /64 MiB/);
  assert.equal(history.state().busy, false);
  await rm(asset);
  await history.capture();
  assert.equal(history.state().error, null);
  assert.equal(history.state().undo, null);
});
