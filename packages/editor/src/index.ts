import type { PresentationData } from '@office-kit/pptx';
import EditorApp from './EditorApp.svelte';
import { EditorController } from './core/controller.svelte.ts';
import type { ChangeSource } from './core/change-source.ts';
import type { EditorTool, JsonValue } from './core/agent-tools.ts';
import { downloadPptx } from './core/download.ts';
import { previewFrameOptions } from './core/preview-frame.ts';
import { propose } from './core/propose.svelte.ts';
import type { ProposeOptions, ProposeResult } from './core/proposal.ts';
import { reconcileSelection, selectionShapeRefs, type ShapeRef } from './core/shape-ref.ts';
import { watchSelection } from './core/watch-selection.svelte.ts';
import { watchValue } from './core/watch-value.svelte.ts';
import { getLocale, setLocale, t } from './i18n/i18n.svelte.ts';
import { mountInPage, mountInShadowRoot } from './mount.ts';

export { resolveShape, type ShapeRef } from './core/shape-ref.ts';
export type { ChangeSource } from './core/change-source.ts';
export type { EditorTool, JsonSchema, JsonValue } from './core/agent-tools.ts';
export type { ProposalOrigin, ProposeOptions, ProposeResult } from './core/proposal.ts';
export type { DeckConflict } from './merge/deck-merge.ts';

/** A language the editor's interface is available in. */
export type EditorLocale = 'en' | 'ja';

export interface EditorOptions {
  /**
   * The `.pptx` file to open. Without it the editor starts a new presentation
   * with one title slide, as File ▸ New does.
   */
  source?: Uint8Array;
  /**
   * The file name shown in the title bar and given to a downloaded copy.
   * Defaults to `Untitled.pptx`.
   */
  fileName?: string;
  /**
   * The interface language. Without it the editor uses the language last chosen
   * in its header, else the browser's. All editors on a page share one language.
   */
  locale?: EditorLocale;
  /**
   * Called with the presentation's `.pptx` bytes when the user saves (the Save
   * button, File ▸ Save or Ctrl/Cmd+S), on `save()`, and for AutoSave. A
   * rejected promise is shown to the user as a failed save. Resolve `false`
   * when the deck was not saved and you have told the user why yourself: it
   * stays marked as changed, the editor shows nothing and does not retry, so
   * call `save()` once you can save again. A call never starts before the
   * previous one has settled. Without `onSave`, saving downloads the file.
   */
  onSave?: (pptx: Uint8Array) => Promise<boolean | void> | boolean | void;
  /**
   * Shows the title bar's AutoSave switch (on at first). While it is on, the
   * editor calls `onSave` shortly after each edit. Needs `onSave`.
   */
  autoSave?: boolean;
  /** A slimmer title bar without the product name, for a host with little room. */
  compact?: boolean;
  /**
   * An element of yours shown in the title bar, such as your own save status.
   * It stays in your page's DOM (slotted into the editor's shadow root), so
   * your styles apply to it.
   */
  status?: HTMLElement;
  /**
   * `false` renders the editor into the page instead of a shadow root. Only for
   * a page that is the editor's alone, such as its own `<iframe>`: the page's
   * styles then reach the editor, and the editor's reach the page.
   */
  isolate?: boolean;
}

/** The events `EditorHandle.on` reports, by name. */
export interface EditorEventMap {
  /**
   * The presentation changed: an edit was made (by the user, or by `apply` or
   * `propose`), undone or redone, or File ▸ New / Open (or `open`) replaced it.
   */
  change: { readonly source: ChangeSource };
  /** The shapes the user has selected changed. */
  selectionchange: { readonly selection: readonly ShapeRef[] };
  /** The presentation got changes `onSave` has not saved, or they were saved. */
  dirtychange: { readonly dirty: boolean };
  /** The interface language changed (the user picked another in the header). */
  localechange: { readonly locale: EditorLocale };
}

export interface EditorHandle {
  /**
   * Settles once `source` is open and the editor is shown. Rejects, showing no
   * editor, when `source` is not a presentation the library can read.
   */
  readonly ready: Promise<void>;
  /** Whether the presentation has changes `onSave` has not saved. */
  readonly dirty: boolean;
  /** The interface language in use. */
  readonly locale: EditorLocale;
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
   * Hands the editor a version of the presentation made elsewhere: `edited`,
   * made from `base` (for an agent, the `snapshot()` it worked on; for the
   * source, the previous build). It is merged with the open presentation,
   * including edits the user has made since and not saved, and applied as one
   * undo step, named "Agent: `label`" (or `label` for `from: 'source'`).
   * Changes to different slides, shapes and relationships combine; a
   * proposal that changes nothing the user has not already got is a no-op.
   *
   * When the same item changed on both sides, nothing is applied until the
   * user chooses in the title bar: keep their edits (`kept-mine`), or replace
   * the presentation with `edited` as one undo step (`took-theirs`; Undo
   * brings their version back). With `base: null` there is nothing to merge
   * against, so any difference asks the user.
   *
   * With `from: 'source'`, `edited` is what the host has saved: a result equal
   * to it does not count as an unsaved change. Rejects before `ready`, after
   * `destroy()` (also while the user is choosing), and while an earlier
   * proposal has not finished.
   */
  propose(
    base: Uint8Array | null,
    edited: Uint8Array,
    options: ProposeOptions,
  ): Promise<ProposeResult>;
  /**
   * Opens another presentation in place of this one, as File ▸ Open does, and
   * Undo starts over. It counts as saved unless `unsaved` is set (for a
   * recovered copy, say). Rejects while a proposal awaits the user's choice.
   */
  open(pptx: Uint8Array, options?: { fileName?: string; unsaved?: boolean }): Promise<void>;
  /** Saves as the Save button does: through `onSave`, or as a download without it. */
  save(): Promise<void>;
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
  const { source, fileName, locale, onSave, autoSave = false, compact = false, status } = options;
  if (locale) setLocale(locale, false);
  const editor = new EditorController({ hostFrame: previewFrameOptions.has(options) });
  const doc = editor.doc;
  const listeners: { [K in keyof EditorEventMap]: Set<(event: EditorEventMap[K]) => void> } = {
    change: new Set(),
    selectionchange: new Set(),
    dirtychange: new Set(),
    localechange: new Set(),
  };
  let unmount: (() => void) | undefined;
  let opened = false;
  let destroyed = false;
  let proposing = false;

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

  let saving = Promise.resolve();
  function saveToHost(save: NonNullable<EditorOptions['onSave']>): Promise<void> {
    saving = saving.then(async () => {
      const version = doc.version;
      try {
        if ((await save(await doc.toBytes())) !== false) doc.markSaved(version);
      } catch (error) {
        editor.toast(
          'error',
          `${t('Save failed')}: ${error instanceof Error ? error.message : String(error)}`,
        );
      }
    });
    return saving;
  }

  const ready = (async () => {
    if (source) await doc.loadBytes(source, fileName ?? doc.fileName);
    else if (fileName) doc.fileName = fileName;
    if (destroyed) return;
    const props = {
      editor,
      embedded: true,
      compactHost: compact,
      autoSave,
      ...(status ? { status } : {}),
      ...(onSave ? { onsave: () => saveToHost(onSave) } : {}),
    };
    unmount =
      options.isolate === false
        ? mountInPage(target, EditorApp, props)
        : mountInShadowRoot(target, EditorApp, props);
    doc.onChange = (source) => emit('change', { source });
    const stopWatching = [
      watchSelection(doc, (selection) => emit('selectionchange', { selection })),
      watchValue(
        () => doc.dirty,
        (dirty) => emit('dirtychange', { dirty }),
      ),
      watchValue(getLocale, (locale) => emit('localechange', { locale })),
    ];
    const unmountApp = unmount;
    unmount = () => {
      for (const stop of stopWatching) stop();
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
    get dirty() {
      return doc.dirty;
    },
    get locale() {
      return getLocale();
    },
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
    async propose(base, edited, proposeOptions) {
      assertOpen();
      if (proposing) throw new Error('An earlier proposal has not finished yet.');
      proposing = true;
      try {
        return await propose(editor, base, edited, proposeOptions);
      } finally {
        proposing = false;
      }
    },
    async open(pptx, { fileName: name, unsaved = false } = {}) {
      assertOpen();
      if (editor.proposal) throw new Error('A proposal is waiting for the user to choose.');
      await doc.loadBytes(pptx, name ?? doc.fileName, unsaved);
    },
    async save() {
      assertOpen();
      if (onSave) await saveToHost(onSave);
      else doc.markSaved(await downloadPptx(doc));
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
      editor.proposal?.abandon(new Error('The editor has been destroyed.'));
      for (const set of Object.values(listeners)) set.clear();
      unmount?.();
      unmount = undefined;
    },
  };
  return handle;
}
