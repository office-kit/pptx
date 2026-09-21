import { lstat, mkdir, readFile, readdir, realpath, unlink, writeFile } from 'node:fs/promises';
import { dirname, join, relative, sep } from 'node:path';

const SNAPSHOT_LIMIT = 64 * 1024 * 1024;
const HISTORY_LIMIT = 128 * 1024 * 1024;
const ENTRY_LIMIT = 50;
const ignored = new Set(['node_modules', '.git', 'dist', '.office-kit']);
export const historyFile = (name: string) =>
  !name.split(sep).some((part) => ignored.has(part)) &&
  /\.([cm]?[jt]sx?|json|css|md|ya?ml|pptx|png|jpe?g|gif|bmp|tiff?|emf|wmf|svg)$/i.test(name);
type Snapshot = Map<string, Buffer>;
type Change = { file: string; before: Buffer | undefined; after: Buffer | undefined };
type Entry = { label: string; changes: Change[] };
const equal = (a?: Buffer, b?: Buffer) => a === b || (!!a && !!b && a.equals(b));

/** Session-local project history. Overlapping agent turns form one transaction. */
export async function createHistory(directory: string, notify: () => void = () => {}) {
  const root = await realpath(directory);
  let baseline: Snapshot = new Map();
  let initialized = false;
  const past: Entry[] = [],
    future: Entry[] = [];
  const active = new Set<string>();
  let label = '';
  let error: string | null = null;
  let queue = Promise.resolve();
  let restoring = false;
  function serial<T>(action: () => Promise<T>): Promise<T> {
    const result = queue.then(action);
    queue = result.then(
      () => {},
      () => {},
    );
    return result
      .catch((cause) => {
        error = cause instanceof Error ? cause.message : String(cause);
        throw cause;
      })
      .finally(notify);
  }
  async function snapshot(): Promise<Snapshot> {
    const files = new Map<string, Buffer>();
    let bytes = 0;
    async function visit(dir: string): Promise<void> {
      const entries = await readdir(dir, { withFileTypes: true });
      await Promise.all(
        entries.map(async (item) => {
          if (ignored.has(item.name) || item.isSymbolicLink()) return;
          const file = join(dir, item.name);
          if (item.isDirectory()) return visit(file);
          if (!item.isFile() || !historyFile(file)) return;
          const info = await lstat(file);
          bytes += info.size;
          if (bytes > SNAPSHOT_LIMIT)
            throw new Error(
              'Undo history exceeds 64 MiB of project source/assets. Move unused assets outside the deck folder.',
            );
          if (info.isSymbolicLink() || (await realpath(file)) !== file)
            throw new Error('Project paths changed while recording history. Try again.');
          const content = await readFile(file);
          files.set(relative(root, file), content);
        }),
      );
    }
    await visit(root);
    return files;
  }
  function record(next: Snapshot, description: string) {
    if (!initialized) {
      baseline = next;
      initialized = true;
      error = null;
      return;
    }
    const changes: Change[] = [];
    for (const file of new Set([...baseline.keys(), ...next.keys()])) {
      const before = baseline.get(file),
        after = next.get(file);
      if (!equal(before, after)) changes.push({ file, before, after });
    }
    if (changes.length) {
      past.push({ label: description, changes });
      future.length = 0;
      const size = () =>
        past.reduce(
          (sum, item) =>
            sum +
            item.changes.reduce(
              (n, change) => n + (change.before?.length ?? 0) + (change.after?.length ?? 0),
              0,
            ),
          0,
        );
      while (past.length > ENTRY_LIMIT || (past.length > 1 && size() > HISTORY_LIMIT)) past.shift();
    }
    baseline = next;
    error = null;
  }
  async function contents(file: string): Promise<Buffer | undefined> {
    const path = join(root, file);
    let parent = dirname(path);
    while (true) {
      try {
        if ((await realpath(parent)) !== parent)
          throw new Error('Undo refused: a project directory is now a symbolic link.');
        break;
      } catch (cause) {
        if ((cause as NodeJS.ErrnoException).code !== 'ENOENT') throw cause;
        parent = dirname(parent);
      }
    }
    try {
      const info = await lstat(path);
      if (!info.isFile() || info.isSymbolicLink())
        throw new Error('Undo refused: a project file changed type.');
      return await readFile(path);
    } catch (cause) {
      if ((cause as NodeJS.ErrnoException).code === 'ENOENT') return undefined;
      throw cause;
    }
  }
  async function replace(change: Change, expected?: Buffer, next?: Buffer) {
    if (!equal(await contents(change.file), expected))
      throw new Error(`History conflict in ${change.file}. Newer changes were preserved.`);
    const file = join(root, change.file);
    if (next) {
      await mkdir(dirname(file), { recursive: true });
      await writeFile(file, next);
    } else await unlink(file);
  }
  try {
    record(await snapshot(), 'Initial state');
  } catch (cause) {
    error = cause instanceof Error ? cause.message : String(cause);
  }
  return {
    state: () => ({
      undo: past.at(-1)?.label ?? null,
      redo: future.at(-1)?.label ?? null,
      busy: active.size > 0 || restoring,
      error,
    }),
    capture: () =>
      serial(async () => {
        if (!active.size && !restoring) record(await snapshot(), 'Source edit');
      }),
    begin: (id: string, description: string) =>
      serial(async () => {
        if (active.has(id)) return;
        if (!active.size) {
          record(await snapshot(), 'Source edit');
          label = description;
        } else label = 'Concurrent edits';
        active.add(id);
      }),
    end: (id: string) =>
      serial(async () => {
        if (!active.delete(id) || active.size) return;
        record(await snapshot(), label);
      }),
    move: (direction: 'undo' | 'redo') =>
      serial(async () => {
        if (active.size)
          throw new Error(
            'Wait for all agents and text edits to finish before using Undo or Redo.',
          );
        // A save made outside the preview is also an edit; never silently overwrite it.
        record(await snapshot(), 'Source edit');
        const from = direction === 'undo' ? past : future;
        const to = direction === 'undo' ? future : past;
        const item = from.at(-1);
        if (!item) throw new Error(`Nothing to ${direction}.`);
        const before = (c: Change) => (direction === 'undo' ? c.after : c.before);
        const after = (c: Change) => (direction === 'undo' ? c.before : c.after);
        const applied: Change[] = [];
        restoring = true;
        try {
          await Promise.all(
            item.changes.map(async (change) => {
              if (!equal(await contents(change.file), before(change)))
                throw new Error(
                  `History conflict in ${change.file}. Newer changes were preserved.`,
                );
            }),
          );
          // Ordered writes let a failed transaction restore only the files it wrote.
          for (const change of item.changes) {
            await replace(change, before(change), after(change));
            applied.push(change);
          }
          for (const change of item.changes) {
            const content = after(change);
            if (content) baseline.set(change.file, content);
            else baseline.delete(change.file);
          }
          from.pop();
          to.push(item);
          error = null;
        } catch (cause) {
          const failures = [];
          for (const change of applied.reverse()) {
            try {
              await replace(change, after(change), before(change));
            } catch (rollback) {
              failures.push(String(rollback));
            }
          }
          if (failures.length)
            throw new Error(`${String(cause)}; rollback incomplete: ${failures.join('; ')}`);
          throw cause;
        } finally {
          restoring = false;
        }
      }),
  };
}
