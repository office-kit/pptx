import { spawn } from 'node:child_process';
import { once } from 'node:events';
import { fileURLToPath } from 'node:url';

export async function startPreview(file) {
  const process = spawn(globalThis.process.execPath, [
    fileURLToPath(new URL('../../dist/cli.mjs', import.meta.url)),
    'dev',
    file,
    '--port',
    '0',
  ]);
  let output = '';
  process.stderr.on('data', (chunk) => {
    output += chunk;
  });
  const url = await new Promise((resolve, reject) => {
    const timer = setTimeout(() => {
      process.kill();
      reject(new Error(`Preview startup timed out: ${output}`));
    }, 15000);
    process.stdout.on('data', (chunk) => {
      output += chunk;
      const match = output.match(/Preview: (http:\/\/\S+)/);
      if (match) {
        clearTimeout(timer);
        resolve(match[1]);
      }
    });
    process.once('error', (error) => {
      clearTimeout(timer);
      reject(error);
    });
    process.once('exit', (code) => {
      clearTimeout(timer);
      reject(new Error(`Preview exited (${code}): ${output}`));
    });
  });
  return {
    url,
    async close() {
      if (process.exitCode !== null || process.signalCode !== null) return;
      const closed = once(process, 'exit');
      process.kill('SIGTERM');
      await closed;
    },
  };
}

export async function waitForState(url, predicate) {
  const deadline = Date.now() + 10000;
  while (Date.now() < deadline) {
    const response = await fetch(url + '/editor/state');
    const state = await response.json();
    if (!state.building && predicate(state)) return state;
    await new Promise((resolve) => setTimeout(resolve, 50));
  }
  throw new Error('Editor state did not settle');
}
