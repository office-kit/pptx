import type { PresentationData } from '@office-kit/pptx';
import EditorApp from './EditorApp.svelte';
import { EditorController } from './core/controller.svelte.ts';
import type { ChangeSource } from './core/change-source.ts';
import type { EditorTool, JsonValue } from './core/agent-tools.ts';
import { reconcileSelection, selectionShapeRefs, type ShapeRef } from './core/shape-ref.ts';
import { watchSelection } from './core/watch-selection.svelte.ts';
import { setLocale, t } from './i18n/i18n.svelte.ts';
import { mountInShadowRoot } from './mount.ts';

export { resolveShape, type ShapeRef } from './core/shape-ref.ts';
export type { ChangeSource } from './core/change-source.ts';
export type { EditorTool, JsonSchema, JsonValue } from './core/agent-tools.ts';

/** A language the editor's interface is available in. */
export type EditorLocale = 'en' | 'ja';

export interface EditorOptions {
  /**
   * The `.pptx` file to open. Without it the editor starts a new presentation
   * with one title slide, as File ▸ New does.
   */
  source?: Uint8Array;
  /**
   * The interface language. Without it the editor uses the language last chosen
   * in its header, else the browser's. All editors on a page share one language.
   */
  locale?: EditorLocale;
  /**
   * Called with the presentation's `.pptx` bytes when the user saves (the Save
   * button, File ▸ Save or Ctrl/Cmd+S). A rejected promise is shown to the user
   * as a failed save. Without `onSave`, saving downloads the file instead.
   */
  onSave?: (pptx: Uint8Array) => Promise<void> | void;
}

/** The events `EditorHandle.on` reports, by name. */
export interface EditorEventMap {
  /**
   * The presentation changed: an edit was made (by the user, or by `apply`),
   * undone or redone, or File ▸ New / Open replaced it.
   */
  change: { readonly source: ChangeSource };
  /** The shapes the user has selected changed. */
  selectionchange: { readonly selection: readonly ShapeRef[] };
}

export interface EditorHandle {
  /**
   * Settles once `source` is open and the editor is shown. Rejects, showing no
   * editor, when `source` is not a presentation the library can read.
   */
  readonly ready: Promise<void>;
  /**
   * The presentation as it is now, as `.pptx` bytes. Does not call `onSave`,
   * and does not mark the presentation saved.
   */
  snapshot(): Promise<Uint8Array>;
  /**
   * The shapes the user has selected, in selection order. For a selected
   * table cell, the table. Empty when slides or nothing are selected.
   */
  selection(): readonly ShapeRef[];
  /**
   * Edits the open presentation with the `@office-kit/pptx` API as one undo
   * step, named "Agent: `label`" in the interface language. The edit is shown
   * at once and marks the presentation as changed, as an edit by the user
   * does; `onSave` still runs only when the user saves.
   *
   * `edit` must make its changes synchronously. If it throws, the presentation
   * is restored to how it was and the returned promise rejects with that error
   * once it has been. Rejects without editing before `ready` resolves, after
   * `destroy()`, and while the user is in the middle of dragging a shape.
   */
  apply(label: string, edit: (presentation: PresentationData) => void): Promise<void>;
  /**
   * Every editing capability as a tool definition for a language model: one
   * per mutating `@office-kit/pptx` export, named after it, plus `listSlides`,
   * `listShapes`, `listLayouts` and `listComments` to find what the others
   * take. Sorted by name. The definitions load on first use.
   */
  tools(): Promise<readonly EditorTool[]>;
  /**
   * Runs the tool `name` with `input` as a model gave it, and resolves with the
   * tool's result as JSON (refs for the slides and shapes it made). An editing
   * tool is one undo step, as `apply` makes it, named "Agent: `name`".
   *
   * Rejects without editing when there is no such tool or `input` does not
   * match its schema, with a message that lists each problem for the model to
   * correct; and for the reasons `apply` does.
   */
  run(name: string, input: unknown): Promise<JsonValue>;
  /** Calls `listener` for each `type` event until the returned function is called. */
  on<K extends keyof EditorEventMap>(
    type: K,
    listener: (event: EditorEventMap[K]) => void,
  ): () => void;
  /** Removes the editor and its event listeners, leaving `target` as it was. */
  destroy(): void;
}

/**
 * Mounts the presentation editor in `target`, which it fills; give `target` a
 * height. The editor renders in a shadow root, so the page's styles do not
 * reach it and its styles do not reach the page.
 */
export function mountEditor(target: HTMLElement, options: EditorOptions = {}): EditorHandle {
  const { source, locale, onSave } = options;
  if (locale) setLocale(locale, false);
  const editor = new EditorController();
  const doc = editor.doc;
  const listeners: { [K in keyof EditorEventMap]: Set<(event: EditorEventMap[K]) => void> } = {
    change: new Set(),
    selectionchange: new Set(),
  };
  let unmount: (() => void) | undefined;
  let opened = false;
  let destroyed = false;

  function emit<K extends keyof EditorEventMap>(type: K, event: EditorEventMap[K]): void {
    for (const listener of [...listeners[type]]) {
      // As with DOM events: a failing listener is reported on the page and
      // does not stop the others, nor undo the change it was told about.
      try {
        listener(event);
      } catch (error) {
        reportError(error);
      }
    }
  }

  async function saveToHost(save: (pptx: Uint8Array) => Promise<void> | void): Promise<void> {
    const version = doc.version;
    try {
      await save(await doc.toBytes());
      doc.markSaved(version);
    } catch (error) {
      editor.toast(
        'error',
        `${t('Save failed')}: ${error instanceof Error ? error.message : String(error)}`,
      );
    }
  }

  const ready = (async () => {
    if (source) await doc.loadBytes(source, doc.fileName);
    if (destroyed) return;
    unmount = mountInShadowRoot(target, EditorApp, {
      editor,
      embedded: true,
      ...(onSave ? { onsave: () => saveToHost(onSave) } : {}),
    });
    doc.onChange = (source) => emit('change', { source });
    const stopWatching = watchSelection(doc, (selection) => emit('selectionchange', { selection }));
    const unmountApp = unmount;
    unmount = () => {
      stopWatching();
      doc.onChange = undefined;
      unmountApp();
    };
    opened = true;
  })();

  function assertOpen(): void {
    if (destroyed) throw new Error('The editor has been destroyed.');
    if (!opened) throw new Error('The editor is not ready yet; await `ready` first.');
  }

  // The tool schemas are large; embeddings that never call them never load them.
  const agentTools = () => import('./core/agent-tools.ts');

  const handle: EditorHandle = {
    ready,
    async snapshot() {
      await ready;
      return doc.toBytes();
    },
    selection() {
      return selectionShapeRefs(doc.pres, doc.selection);
    },
    async apply(label, edit) {
      assertOpen();
      if (doc.liveEditing)
        throw new Error('The user is dragging or resizing a shape; try again when they finish.');
      try {
        doc.transact(
          t('Agent: {label}').replace('{label}', () => label),
          () => {
            const result: unknown = edit(doc.pres);
            if (result instanceof Promise)
              throw new TypeError('apply() edits must be synchronous; `edit` returned a promise.');
            doc.select(reconcileSelection(doc.pres, doc.selection));
          },
          'agent',
        );
      } catch (error) {
        await doc.rollback;
        throw error;
      }
    },
    async tools() {
      return (await agentTools()).editorTools;
    },
    async run(name, input) {
      const { runTool } = await agentTools();
      const host = {
        presentation() {
          assertOpen();
          return doc.pres;
        },
        apply: handle.apply,
      };
      return runTool(host, name, input);
    },
    on(type, listener) {
      const set: Set<typeof listener> = listeners[type];
      set.add(listener);
      return () => set.delete(listener);
    },
    destroy() {
      destroyed = true;
      listeners.change.clear();
      listeners.selectionchange.clear();
      unmount?.();
      unmount = undefined;
    },
  };
  return handle;
}
