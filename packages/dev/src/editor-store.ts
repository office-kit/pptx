import { randomUUID } from 'node:crypto';
import { mkdir, readFile, rename, rm, writeFile } from 'node:fs/promises';
import { basename, dirname, join, resolve } from 'node:path';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { renderDeck, type BuildResult } from './build.ts';
import { describeConflict, mergeDecks } from '@office-kit/pptx-editor/merge';
import { sourceFingerprint } from './fingerprint.ts';

export interface SavedEdits {
  sourceHash: string;
  bytes: Uint8Array;
  /**
   * The source build the edits were made against, which lets a later source
   * change merge with them. Version 1 sidecars did not store it.
   */
  base: Uint8Array | null;
}

const STORE_VERSION = 2;

/** One atomic sidecar keeps the saved deck and its source basis in agreement. */
export function editorStore(entry: string) {
  const path = join(dirname(resolve(entry)), '.office-kit', `${basename(entry)}.editor.zip`);
  return {
    path,
    async read(): Promise<SavedEdits | undefined> {
      let bytes: Uint8Array;
      try {
        bytes = await readFile(path);
      } catch (cause) {
        if (cause instanceof Error && 'code' in cause && cause.code === 'ENOENT') return;
        throw cause;
      }
      const parts = unzipSync(bytes);
      if (!parts['state.json'] || !parts['document.pptx'])
        throw new Error(`Invalid editor file: ${path}`);
      const state: unknown = JSON.parse(strFromU8(parts['state.json']));
      if (
        !state ||
        typeof state !== 'object' ||
        !('version' in state) ||
        (state.version !== 1 && state.version !== STORE_VERSION) ||
        !('sourceHash' in state) ||
        typeof state.sourceHash !== 'string' ||
        !/^[a-f0-9]{64}$/.test(state.sourceHash)
      ) {
        throw new Error(`Invalid editor metadata: ${path}`);
      }
      const base = parts['base.pptx'];
      if (state.version === STORE_VERSION && !base) throw new Error(`Invalid editor file: ${path}`);
      return { sourceHash: state.sourceHash, bytes: parts['document.pptx'], base: base ?? null };
    },
    async write(edits: SavedEdits & { base: Uint8Array }): Promise<void> {
      await mkdir(dirname(path), { recursive: true });
      const temporary = `${path}.${randomUUID()}.tmp`;
      try {
        await writeFile(
          temporary,
          zipSync({
            'state.json': strToU8(
              JSON.stringify({ version: STORE_VERSION, sourceHash: edits.sourceHash }),
            ),
            'document.pptx': edits.bytes,
            'base.pptx': edits.base,
          }),
          { flag: 'wx' },
        );
        await rename(temporary, path);
      } finally {
        await rm(temporary, { force: true });
      }
    },
    async clear(): Promise<void> {
      await rm(path, { force: true });
    },
  };
}

export async function applySavedEdits(entry: string, source: BuildResult): Promise<BuildResult> {
  const edits = await editorStore(entry).read();
  if (!edits) return source;
  let bytes = edits.bytes;
  if (edits.sourceHash !== sourceFingerprint(source.bytes)) {
    const merged = edits.base ? mergeDecks(edits.base, edits.bytes, source.bytes) : null;
    if (!merged?.ok) {
      throw new Error(
        [
          'Source changed since the last editor save. Open office-pptx dev and resolve the editor conflict before exporting.',
          ...(merged?.conflicts.map(describeConflict) ?? []),
        ].join('\n'),
      );
    }
    bytes = merged.bytes;
  }
  return (await renderDeck(bytes, source.dependencies)).result;
}
