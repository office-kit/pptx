import { Worker } from 'node:worker_threads';
import { resolve, join } from 'node:path';
import { mkdtempSync } from 'node:fs';
import { rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import type { BuildResult } from './build.ts';

const BUILD_TIMEOUT_MS = 60_000;

function prepareWorker(entry: string) {
  const directory = mkdtempSync(join(tmpdir(), 'office-pptx-'));
  const worker = new Worker(new URL('./worker.mjs', import.meta.url), {
    workerData: { entry: resolve(entry), directory },
    execArgv: ['--enable-source-maps'],
  });
  const result = new Promise<BuildResult>((resolveBuild, reject) => {
    worker.once('message', resolveBuild);
    worker.once('error', reject);
    worker.once('exit', (code) =>
      reject(new Error(`Build worker exited without a result (code ${code}).`)),
    );
  });
  // Startup may fail before the next edit asks to use this worker.
  void result.catch(() => {});
  return { worker, result, directory };
}

/** Internal watch session. Preload libraries, but never reuse a user module cache. */
export function createDeckBuilder(entry: string, keepWarm = true) {
  let prepared: ReturnType<typeof prepareWorker> | undefined = prepareWorker(entry);
  let closed = false;
  let active = false;
  return {
    async build(): Promise<BuildResult> {
      if (closed || active) throw new Error('Deck builder is closed or already building.');
      active = true;
      const current = prepared ?? prepareWorker(entry);
      prepared = current;
      let timer: ReturnType<typeof setTimeout> | undefined;
      try {
        current.worker.postMessage('build');
        return await Promise.race([
          current.result,
          new Promise<never>((_, reject) => {
            timer = setTimeout(
              () =>
                reject(
                  new Error(
                    'Build exceeded 60 seconds. Check the deck for loops or unresolved promises.',
                  ),
                ),
              BUILD_TIMEOUT_MS,
            );
          }),
        ]);
      } finally {
        clearTimeout(timer);
        await current.worker.terminate();
        await rm(current.directory, { recursive: true, force: true });
        active = false;
        prepared = !closed && keepWarm ? prepareWorker(entry) : undefined;
      }
    },
    async cancel() {
      if (active) await prepared?.worker.terminate();
    },
    async close() {
      closed = true;
      const current = prepared;
      await current?.worker.terminate();
      if (current) await rm(current.directory, { recursive: true, force: true });
    },
  };
}
