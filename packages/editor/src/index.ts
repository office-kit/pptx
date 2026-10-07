import EditorApp from './EditorApp.svelte';
import { EditorController } from './core/controller.svelte.ts';
import { setLocale, t } from './i18n/i18n.svelte.ts';
import { mountInShadowRoot } from './mount.ts';

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

export interface EditorHandle {
  /**
   * Settles once `source` is open and the editor is shown. Rejects, showing no
   * editor, when `source` is not a presentation the library can read.
   */
  readonly ready: Promise<void>;
  /** The presentation as it is now, as `.pptx` bytes. Does not call `onSave`. */
  save(): Promise<Uint8Array>;
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
  let unmount: (() => void) | undefined;
  let destroyed = false;

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
  })();

  return {
    ready,
    async save() {
      await ready;
      return doc.toBytes();
    },
    destroy() {
      destroyed = true;
      unmount?.();
      unmount = undefined;
    },
  };
}
