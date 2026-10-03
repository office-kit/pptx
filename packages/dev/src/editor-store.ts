import { createHash, randomUUID } from 'node:crypto';
import { mkdir, readFile, rename, rm, writeFile } from 'node:fs/promises';
import { basename, dirname, join, resolve } from 'node:path';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { renderDeck, type BuildResult } from './build.ts';

/** ZIP container dates and generated core timestamps are not source edits. */
export function sourceFingerprint(bytes: Uint8Array): string {
  const parts = unzipSync(bytes);
  const hash = createHash('sha256');
  for (const name of Object.keys(parts).sort()) {
    let content = parts[name]!;
    if (name === 'docProps/core.xml') {
      content = strToU8(
        strFromU8(content).replace(/<(dcterms:(?:created|modified))\b[^>]*>[\s\S]*?<\/\1>/g, ''),
      );
    }
    hash.update(JSON.stringify([name, createHash('sha256').update(content).digest('hex')]));
  }
  return hash.digest('hex');
}

export interface SavedEdits {
  sourceHash: string;
  bytes: Uint8Array;
}

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
        state.version !== 1 ||
        !('sourceHash' in state) ||
        typeof state.sourceHash !== 'string' ||
        !/^[a-f0-9]{64}$/.test(state.sourceHash)
      ) {
        throw new Error(`Invalid editor metadata: ${path}`);
      }
      return { sourceHash: state.sourceHash, bytes: parts['document.pptx'] };
    },
    async write(edits: SavedEdits): Promise<void> {
      await mkdir(dirname(path), { recursive: true });
      const temporary = `${path}.${randomUUID()}.tmp`;
      try {
        await writeFile(
          temporary,
          zipSync({
            'state.json': strToU8(JSON.stringify({ version: 1, sourceHash: edits.sourceHash })),
            'document.pptx': edits.bytes,
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
  if (edits.sourceHash !== sourceFingerprint(source.bytes)) {
    throw new Error(
      'Source changed since the last editor save. Open office-pptx dev and resolve the editor conflict before exporting.',
    );
  }
  return (await renderDeck(edits.bytes, source.dependencies)).result;
}
