import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, readFile, writeFile, rm, symlink, readdir } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { spawn, execFile } from 'node:child_process';
import { promisify } from 'node:util';
import { setTimeout } from 'node:timers/promises';
import { buildDeck, initProject, inspectTemplate } from '../dist/index.mjs';

const execute = promisify(execFile);
const cli = new URL('../dist/cli.mjs', import.meta.url).pathname;
async function fixture(t) {
  const directory = await mkdtemp(join(tmpdir(), 'pptx-authoring-test-'));
  t.after(() => rm(directory, { recursive: true, force: true }));
  return directory;
}

test('initialized TSX typechecks, builds, and uses source-relative assets in imported modules', async (t) => {
  const directory = await fixture(t);
  const project = await initProject(join(directory, 'slides'));
  await symlink(resolve('node_modules'), join(project, 'node_modules'), 'dir');
  await execute(process.execPath, [
    resolve('node_modules/typescript/bin/tsc'),
    '--project',
    project,
  ]);
  const deck = join(project, 'deck.tsx');
  const result = await buildDeck(deck);
  assert.equal(result.slides.length, 1);
  assert.match(result.slides[0], /Your next presentation/);
  const entry = await readFile(deck, 'utf8');
  const coverPath = join(project, 'slides/cover.tsx');
  const cover = await readFile(coverPath, 'utf8');
  await writeFile(coverPath, cover.replace('Your next presentation', 'Quarterly review'));
  assert.match((await buildDeck(deck)).slides[0], /Quarterly review/);
  const themePath = join(project, 'theme.ts');
  const theme = await readFile(themePath, 'utf8');
  await writeFile(themePath, theme.replace('#E5481F', '#123456'));
  assert.match((await buildDeck(deck)).slides[0], /#123456/i);
  assert.equal(await readFile(deck, 'utf8'), entry);
  await writeFile(join(project, 'template.pptx'), result.bytes);
  await writeFile(
    join(project, 'template.ts'),
    `import { readFile } from 'node:fs/promises';
export default await readFile(new URL('./template.pptx', import.meta.url));`,
  );
  await writeFile(
    deck,
    `import { Presentation, Slide, Text } from '@office-kit/pptx-dsl';
import source from './template.ts';
export default <Presentation source={source} mode="edit"><Slide target={{index:0}}>
<Text x={1} y={5} width={8} height={1}>Template loaded</Text></Slide></Presentation>;`,
  );
  assert.match((await buildDeck(deck)).slides[0], /Template loaded/);
  const inspection = await inspectTemplate(join(project, 'template.pptx'));
  assert.equal(inspection.slides.length, 1);
  assert.ok(inspection.layouts.length > 0);
  await assert.rejects(initProject(project), /EEXIST/);
});

test('builds are isolated and recover from compile and runtime errors', async (t) => {
  const directory = await fixture(t);
  const deck = join(directory, 'deck.tsx');
  const source = `import { Presentation, Slide, Text } from '@office-kit/pptx-dsl';
const counter = globalThis.__deckBuildCounter = (globalThis.__deckBuildCounter ?? 0) + 1;
export default <Presentation><Slide><Text x={1} y={1} width={4} height={1}>Build {counter}</Text></Slide></Presentation>;`;
  await writeFile(deck, source);
  assert.match((await buildDeck(deck)).slides[0], /Build 1/);
  assert.match((await buildDeck(deck)).slides[0], /Build 1/);
  await writeFile(deck, 'export default <Presentation');
  await assert.rejects(buildDeck(deck));
  await writeFile(deck, `throw new Error('intentional source error');`);
  await assert.rejects(buildDeck(deck), /intentional source error/);
  await writeFile(
    deck,
    `import { Presentation, Text } from '@office-kit/pptx-dsl';
export default <Presentation>
  <Text x={1} y={1} width={4} height={1}>Wrong parent</Text>
</Presentation>;`,
  );
  await assert.rejects(buildDeck(deck), /deck\.tsx:3:\d+ <Text>/);
  await writeFile(deck, source);
  assert.match((await buildDeck(deck)).slides[0], /Build 1/);
  await assert.rejects(
    execute(process.execPath, [cli, 'build', deck, '--unknown', 'x']),
    /Unexpected arguments/,
  );
});

test('watch updates and retains the last successful deck through an error', async (t) => {
  const directory = await fixture(t);
  const deck = join(directory, 'deck.tsx');
  const source = (text) => `import { Presentation, Slide, Text } from '@office-kit/pptx-dsl';
export default <Presentation><Slide><Text x={1} y={1} width={4} height={1}>${text}</Text></Slide></Presentation>;`;
  await writeFile(deck, source('Before edit'));
  const workerDirectory = await fixture(t);
  const child = spawn(process.execPath, [cli, 'dev', deck, '--port', '0'], {
    env: { ...process.env, TMPDIR: workerDirectory },
    stdio: ['ignore', 'pipe', 'pipe'],
  });
  t.after(() => {
    child.kill('SIGTERM');
  });
  const url = await new Promise((resolveUrl, reject) => {
    let output = '';
    const timeout = setTimeout(15_000, undefined, { ref: false }).then(() =>
      reject(new Error(`Server startup timed out: ${output}`)),
    );
    void timeout;
    child.stdout.on('data', (data) => {
      output += data;
      const match = output.match(/Preview: (http:\/\/[^\s]+)/);
      if (match) resolveUrl(match[1]);
    });
    child.stderr.on('data', (data) => {
      output += data;
    });
    child.once('error', reject);
    child.once('exit', (code) => reject(new Error(`Server exited: ${code}\n${output}`)));
  });
  async function until(predicate) {
    const end = Date.now() + 10_000;
    while (Date.now() < end) {
      const state = await fetch(`${url}/state`).then((response) => response.json());
      if (predicate(state)) return state;
      await setTimeout(50);
    }
    throw new Error('Preview did not update');
  }
  const initial = await until((state) => state.slides[0]?.includes('Before edit'));
  await writeFile(deck, source('After edit'));
  const edited = await until((state) => state.slides[0]?.includes('After edit'));
  const delta = await fetch(`${url}/state?since=${initial.revision}`).then((r) => r.json());
  assert.equal(delta.slides, undefined);
  assert.deepEqual(delta.changes, { 0: edited.slides[0] });
  assert.equal(delta.count, 1);
  const unchanged = await fetch(`${url}/state?since=${edited.revision}`).then((r) => r.json());
  assert.deepEqual(unchanged.changes, {});
  const stale = await fetch(`${url}/state?since=-10`).then((r) => r.json());
  assert.deepEqual(stale.slides, edited.slides);
  const imported = join(directory, 'content.ts');
  await writeFile(imported, "export const title = 'Imported title';");
  const importedEntry = `import { title } from './content.ts';\n${source('{title}')}`;
  await writeFile(deck, importedEntry);
  await until((state) => state.slides[0]?.includes('Imported title'));
  await writeFile(imported, "export const title = 'Patched title';");
  await until((state) => state.slides[0]?.includes('Patched title'));
  assert.equal(await readFile(deck, 'utf8'), importedEntry);
  await writeFile(deck, source('After edit'));
  await until((state) => state.slides[0]?.includes('After edit'));
  const beforeError = new Uint8Array(
    await fetch(`${url}/deck.pptx`).then((response) => response.arrayBuffer()),
  );
  await writeFile(deck, 'export default <');
  await until((state) => state.error && state.slides[0]?.includes('After edit'));
  assert.deepEqual(
    new Uint8Array(await fetch(`${url}/deck.pptx`).then((response) => response.arrayBuffer())),
    beforeError,
  );
  await writeFile(deck, source('Recovered'));
  await until((state) => !state.error && state.slides[0]?.includes('Recovered'));
  assert.match(await readFile(deck, 'utf8'), /Recovered/);
  await writeFile(deck, `await new Promise(r => setTimeout(r, 500));\n${source('Superseded')}`);
  await until((state) => state.building);
  await writeFile(deck, source('Newest'));
  const newest = await until((state) => {
    assert.ok(!state.slides[0]?.includes('Superseded'));
    return state.slides[0]?.includes('Newest');
  });
  assert.equal(newest.error, null);
  const isolated = (suffix) => `import { Presentation, Slide, Text } from '@office-kit/pptx-dsl';
const counter = globalThis.__counter = (globalThis.__counter ?? 0) + 1;
export default <Presentation><Slide><Text x={1} y={1} width={4} height={1}>${suffix} {counter}</Text></Slide></Presentation>`;
  await writeFile(deck, isolated('First'));
  await until((state) => state.slides[0]?.includes('First 1'));
  await writeFile(deck, isolated('Second'));
  await until((state) => state.slides[0]?.includes('Second 1'));

  await writeFile(deck, `while (true) {}\n${source('Never finishes')}`);
  await until((state) => state.building);
  const recoveryStarted = Date.now();
  await writeFile(deck, source('Cancelled obsolete evaluation'));
  await until(
    (state) => !state.error && state.slides[0]?.includes('Cancelled obsolete evaluation'),
  );
  assert.equal(
    (await readdir(workerDirectory)).length,
    1,
    'only the next preloaded worker retains a temporary directory',
  );
  assert.ok(
    Date.now() - recoveryStarted < 3000,
    'an obsolete infinite loop must not delay the next edit',
  );
});
