import { Worker } from 'node:worker_threads';
import { writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import type { BuildResult } from './build.ts';
export type { BuildResult } from './build.ts';
export { initProject } from './init.ts';
export { inspectTemplate } from './inspect.ts';

/** Each evaluation gets a fresh module cache and releases it when finished. */
export async function buildDeck(entry: string): Promise<BuildResult> {
  const worker = new Worker(new URL('./worker.mjs', import.meta.url), {
    workerData: { entry: resolve(entry) },
    execArgv: ['--enable-source-maps'],
  });
  let timer: ReturnType<typeof setTimeout> | undefined;
  try {
    return await new Promise<BuildResult>((resolveBuild, reject) => {
      timer = setTimeout(
        () =>
          reject(
            new Error(
              'Build exceeded 60 seconds. Check the deck for loops or unresolved promises.',
            ),
          ),
        60_000,
      );
      worker.once('message', (result: BuildResult) => resolveBuild(result));
      worker.once('error', reject);
      worker.once('exit', (code) =>
        reject(new Error(`Build worker exited without a result (code ${code}).`)),
      );
    });
  } finally {
    clearTimeout(timer);
    await worker.terminate();
  }
}
export async function exportDeck(entry: string, output: string): Promise<BuildResult> {
  const result = await buildDeck(entry);
  await writeFile(output, result.bytes);
  return result;
}
