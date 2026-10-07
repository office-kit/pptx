import { untrack } from 'svelte';
import type { EditorDocument } from './document.svelte.ts';
import { selectionShapeRefs, type ShapeRef } from './shape-ref.ts';

/**
 * Calls `listener` whenever the shapes `doc` has selected change, including a
 * selected shape being renamed. Returns a function that stops watching.
 */
export function watchSelection(
  doc: EditorDocument,
  listener: (selection: readonly ShapeRef[]) => void,
): () => void {
  // Many paths assign `doc.selection` (often to an equal value), so compare the
  // refs rather than reporting every assignment.
  let last = JSON.stringify(selectionShapeRefs(doc.pres, doc.selection));
  return $effect.root(() => {
    $effect(() => {
      doc.version;
      const selection = selectionShapeRefs(doc.pres, doc.selection);
      const key = JSON.stringify(selection);
      if (key === last) return;
      last = key;
      untrack(() => listener(selection));
    });
  });
}
