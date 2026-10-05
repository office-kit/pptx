import { loadPresentation } from '@office-kit/pptx';
import type { JsxSource } from '@office-kit/pptx-dsl';
import { planPropEdit } from '@office-kit/pptx-dsl/source-edit';
import { readFile, realpath, writeFile } from 'node:fs/promises';
import { dirname, isAbsolute, relative } from 'node:path';
import type { BuildResult } from './build.ts';
import { diffDecks, type DeckChange } from './deck-diff.ts';
import { sourceFingerprint } from './editor-store.ts';

export interface SyncResult {
  /** Changes now stated in the TSX source. */
  written: DeckChange[];
  /** Changes the source does not state yet; the caller keeps them in the sidecar. */
  pending: (DeckChange & { reason?: string })[];
  /** The rebuilt source produces exactly the edited deck, so no sidecar is needed. */
  complete: boolean;
}

interface Target {
  change: Extract<DeckChange, { kind: 'geometry' }>;
  anchor: JsxSource;
}

const key = (change: DeckChange) =>
  `${change.kind}:${change.slide}:${change.shapeId}:${change.kind === 'unsupported' ? change.description : ''}`;

/**
 * Writes the editor's changes back into the TSX source as typed prop edits.
 * Every write is verified by rebuilding: the changes it claims must disappear
 * from the diff and no other change may appear, or every file is restored.
 */
export function createSourceSync(
  entry: string,
  source: () => BuildResult | undefined,
  rebuild: () => Promise<string | null>,
) {
  let busy = false;

  async function localFile(fileName: string): Promise<string | null> {
    const root = await realpath(dirname(entry));
    let path: string;
    try {
      path = await realpath(fileName);
    } catch {
      return null;
    }
    const local = relative(root, path);
    if (
      local.startsWith('..') ||
      isAbsolute(local) ||
      local.split(/[\\/]/).includes('node_modules')
    )
      return null;
    return path;
  }

  return {
    async sync(edited: Uint8Array): Promise<SyncResult> {
      if (busy) throw new Error('A source write-back is already running.');
      busy = true;
      try {
        const built = source();
        if (!built?.shapeSources) throw new Error('The source deck is not built yet.');
        const shapeSources = built.shapeSources;
        const editedDeck = await loadPresentation(edited);
        const changes = diffDecks(await loadPresentation(built.bytes), editedDeck);
        // An element evaluated more than once (a `.map` callback, a reused
        // component) made several shapes; its attributes are not one shape's.
        const uses = new Map<string, number>();
        for (const bySlide of Object.values(shapeSources))
          for (const chain of Object.values(bySlide)) {
            const at = chain.at(-1);
            if (at) {
              const id = `${at.fileName}:${at.lineNumber}:${at.columnNumber}`;
              uses.set(id, (uses.get(id) ?? 0) + 1);
            }
          }

        const pending: SyncResult['pending'] = [];
        const targets: Target[] = [];
        for (const change of changes) {
          if (change.kind !== 'geometry') {
            pending.push(change);
            continue;
          }
          const anchor = shapeSources[change.slide]?.[change.shapeId]?.at(-1);
          if (!anchor) {
            pending.push({ ...change, reason: 'no source element' });
            continue;
          }
          if (
            (uses.get(`${anchor.fileName}:${anchor.lineNumber}:${anchor.columnNumber}`) ?? 0) > 1
          ) {
            pending.push({ ...change, reason: 'element makes several shapes' });
            continue;
          }
          targets.push({ change, anchor });
        }

        // Plan per file, later elements first, so an edit never shifts the
        // position of an anchor still to be planned.
        const byFile = new Map<string, Target[]>();
        for (const target of targets) {
          const path = await localFile(target.anchor.fileName);
          if (!path) {
            pending.push({ ...target.change, reason: 'source outside the project' });
            continue;
          }
          byFile.set(path, [...(byFile.get(path) ?? []), target]);
        }
        const files: { path: string; before: string; after: string }[] = [];
        const written: DeckChange[] = [];
        for (const [path, fileTargets] of byFile) {
          const before = await readFile(path, 'utf8');
          let text = before;
          fileTargets.sort(
            (a, b) =>
              b.anchor.lineNumber - a.anchor.lineNumber ||
              b.anchor.columnNumber - a.anchor.columnNumber,
          );
          for (const { change, anchor } of fileTargets) {
            const plan = planPropEdit({
              file: { path, source: text },
              anchor: { path, lineNumber: anchor.lineNumber, columnNumber: anchor.columnNumber },
              edits: change.edits,
            });
            if (!plan.ok) {
              pending.push({ ...change, reason: plan.reason });
              continue;
            }
            text = plan.change.after;
            written.push(change);
          }
          if (text !== before) files.push({ path, before, after: text });
        }

        if (files.length === 0) return { written: [], pending, complete: false };

        const restore = async () => {
          for (const file of files) {
            if ((await readFile(file.path, 'utf8')) !== file.after)
              throw new Error(
                'Source changed during write-back. Your newer changes were preserved.',
              );
          }
          for (const file of files) await writeFile(file.path, file.before);
          await rebuild();
        };
        for (const file of files) {
          if ((await readFile(file.path, 'utf8')) !== file.before)
            throw new Error('Source changed. Save again.');
        }
        for (const file of files) await writeFile(file.path, file.after);
        try {
          const error = await rebuild();
          if (error) throw new Error('The written source did not build: ' + error);
          const rebuilt = source();
          if (!rebuilt) throw new Error('The written source did not build.');
          const remaining = diffDecks(await loadPresentation(rebuilt.bytes), editedDeck);
          const expected = new Set(pending.map(key));
          if (remaining.some((change) => !expected.has(key(change))))
            throw new Error('The written source changed more than the edited shapes.');
          return {
            written,
            pending,
            complete:
              remaining.length === 0 &&
              sourceFingerprint(rebuilt.bytes) === sourceFingerprint(edited),
          };
        } catch (error) {
          await restore();
          throw error;
        }
      } finally {
        busy = false;
      }
    },
  };
}
