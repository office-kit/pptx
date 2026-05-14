// Node-specific entry point. Adds fs-backed convenience helpers on top of
// the platform-neutral public API. The browser bundle does not include
// this file, so `node:fs/promises` stays out of the browser tree.

import { readFile, writeFile } from 'node:fs/promises';
import {
  Presentation as BasePresentation,
  type PresentationInput,
  _internalPackageOf,
} from './api/index.ts';

export * from './api/index.ts';

/**
 * Node-only subclass of `Presentation` that adds `loadFile` and `saveTo`
 * helpers. Inherits everything else from the base class.
 *
 * Browser code should import the same name from `pptx-kit` (not
 * `pptx-kit/node`). The two paths are intentionally distinct to keep the
 * browser bundle from pulling in `node:fs`.
 */
export class Presentation extends BasePresentation {
  /**
   * Reads a `.pptx` from disk and returns a `Presentation`. Convenience
   * over `Presentation.load(await fs.readFile(path))`.
   */
  static async loadFile(path: string): Promise<Presentation> {
    const bytes = await readFile(path);
    return Presentation.load(bytes);
  }

  /**
   * Equivalent to `Presentation.load(input)` but returns the Node
   * subclass so callers always get `saveTo`.
   */
  static override async load(input: PresentationInput): Promise<Presentation> {
    const base = await BasePresentation.load(input);
    // Re-wrap the OpcPackage into the Node subclass so the caller's static
    // type matches the instance type and `saveTo` is available.
    // biome-ignore lint/complexity/useLiteralKeys: bracket access of an internal symbol property
    return Presentation['_fromPackage'](_internalPackageOf(base)) as Presentation;
  }

  /** Writes the serialized PPTX to disk. */
  async saveTo(path: string): Promise<void> {
    await writeFile(path, await this.save());
  }
}
