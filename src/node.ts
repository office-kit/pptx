// Node-specific entry point. Adds fs-backed convenience helpers on top of
// the platform-neutral public API. The browser bundle does not include
// this file, so `node:fs/promises` stays out of the browser tree.

import { readFile, writeFile } from 'node:fs/promises';
import {
  type PresentationData,
  loadPresentation,
  savePresentation,
} from './api/index.ts';

export * from './api/index.ts';

/**
 * Reads a `.pptx` from disk and returns a `PresentationData`. Convenience
 * over `loadPresentation(await fs.readFile(path))`.
 */
export const loadPresentationFile = async (path: string): Promise<PresentationData> => {
  const bytes = await readFile(path);
  return loadPresentation(bytes);
};

/**
 * Serializes a `PresentationData` and writes the bytes to disk.
 */
export const savePresentationToFile = async (
  pres: PresentationData,
  path: string,
): Promise<void> => {
  await writeFile(path, await savePresentation(pres));
};
