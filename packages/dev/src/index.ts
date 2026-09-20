import { writeFile } from 'node:fs/promises';
import { createDeckBuilder } from './build-runner.ts';
import type { BuildResult } from './build.ts';
export type { BuildResult } from './build.ts';
export { initProject } from './init.ts';
export { inspectTemplate } from './inspect.ts';

/** Each evaluation gets a fresh module cache and releases it when finished. */
export async function buildDeck(entry: string): Promise<BuildResult> {
  const builder = createDeckBuilder(entry, false);
  try {
    return await builder.build();
  } finally {
    await builder.close();
  }
}
export async function exportDeck(entry: string, output: string): Promise<BuildResult> {
  const result = await buildDeck(entry);
  await writeFile(output, result.bytes);
  return result;
}
