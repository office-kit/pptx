// Svelte context plumbing so any editor component can reach the controller
// without prop-drilling. The root <EditorApp> sets it; everything else gets it.

import { getContext, setContext } from 'svelte';
import type { EditorController } from './controller.svelte.ts';

const KEY = Symbol('ok-editor');

export function setEditor(controller: EditorController): void {
  setContext(KEY, controller);
}

export function getEditor(): EditorController {
  const c = getContext<EditorController>(KEY);
  if (!c) throw new Error('getEditor() called outside <EditorApp>');
  return c;
}
