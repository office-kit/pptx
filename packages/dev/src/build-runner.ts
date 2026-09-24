import { Worker } from 'node:worker_threads';
import { join } from 'node:path';
import { mkdtempSync } from 'node:fs';
import { rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import type { BuildResult } from './build.ts';
import type { PreviewCache } from './preview-cache.ts';
import { createDeckCompiler } from './bundle.ts';

const BUILD_TIMEOUT_MS = 60_000;

function prepareWorker() {
  const directory = mkdtempSync(join(tmpdir(), 'office-pptx-'));
  const worker = new Worker(new URL('./worker.mjs', import.meta.url), {
    workerData: { directory },
    execArgv: ['--enable-source-maps'],
  });
  const result = new Promise<{ result: BuildResult; cache: PreviewCache }>(
    (resolveBuild, reject) => {
      worker.once('message', resolveBuild);
      worker.once('error', reject);
      worker.once('exit', (code) =>
        reject(new Error(`Build worker exited without a result (code ${code}).`)),
      );
    },
  );
  // Startup may fail before the next edit asks to use this worker.
  void result.catch(() => {});
  return { worker, result, directory };
}

/** Internal watch session. Preload libraries, but never reuse a user module cache. */
export function createDeckBuilder(entry: string, keepWarm = true) {
  const compiler = createDeckCompiler(entry);
  void compiler.catch(() => {});
  let prepared: ReturnType<typeof prepareWorker> | undefined = prepareWorker();
  let cache: PreviewCache | undefined;
  let closed = false;
  let active = false;
  return {
    async build(): Promise<BuildResult> {
      if (closed || active) throw new Error('Deck builder is closed or already building.');
      active = true;
      const current = prepared ?? prepareWorker();
      prepared = current;
      let timer: ReturnType<typeof setTimeout> | undefined;
      try {
        const built = await Promise.race([
          (async () => {
            const dependencies = await (await compiler).build(join(current.directory, 'deck.mjs'));
            current.worker.postMessage({ dependencies, cache, entry });
            return current.result;
          })(),
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
        cache = built.cache;
        return built.result;
      } finally {
        clearTimeout(timer);
        await compiler.then(
          (compiler) => compiler.cancel(),
          () => {},
        );
        await current.worker.terminate();
        await rm(current.directory, { recursive: true, force: true });
        active = false;
        prepared = !closed && keepWarm ? prepareWorker() : undefined;
      }
    },
    async cancel() {
      if (!active) return;
      const current = prepared;
      await compiler.then(
        (compiler) => compiler.cancel(),
        () => {},
      );
      await current?.worker.terminate();
    },
    async close() {
      closed = true;
      const current = prepared;
      await compiler.then(
        (compiler) => compiler.cancel(),
        () => {},
      );
      await current?.worker.terminate();
      await compiler.then(
        (compiler) => compiler.close(),
        () => {},
      );
      if (current) await rm(current.directory, { recursive: true, force: true });
    },
  };
}
