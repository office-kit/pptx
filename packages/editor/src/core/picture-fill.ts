import { getShapeId, getShapeKind, getSlidePartName } from '@office-kit/pptx';
import type { EditorController } from './controller.svelte.ts';
import { insertRememberedPictureFill, insertRememberedTextureFill } from './remembered-fill.ts';
import { t } from '../i18n/i18n.svelte.ts';

/** Whether every selected object can take a picture or texture fill. */
export function canFillWithPicture(editor: EditorController): boolean {
  const shapes = editor.selectedShapes();
  return (
    editor.doc.selection.kind === 'shape' &&
    shapes.length > 0 &&
    !editor.selectionLocked() &&
    shapes.every((shape) => getShapeKind(shape) === 'shape')
  );
}

/**
 * Fills the selected shapes with a picture (stretched) or a gallery texture (tiled)
 * in one undo step. `read` may wait on a file chooser or the clipboard, so a
 * selection that changed meanwhile is refused rather than filled.
 */
export async function fillSelectionWithPicture(
  editor: EditorController,
  read: () => Promise<Uint8Array>,
  kind: 'picture' | 'texture',
): Promise<void> {
  const { doc } = editor;
  const selection = doc.selection;
  if (selection.kind !== 'shape' || !canFillWithPicture(editor)) return;
  const slide = doc.slideAt(selection.slideIndex);
  if (!slide) return;
  const targets = editor.selectedShapes(),
    presentation = doc.pres,
    version = doc.version;
  const slideKey = getSlidePartName(slide);
  const bytes = await read();
  if (
    doc.pres !== presentation ||
    doc.version !== version ||
    doc.selection !== selection ||
    editor.selectionLocked()
  ) {
    throw new Error(t('The selection changed. Choose the picture again.'));
  }
  const insert = kind === 'texture' ? insertRememberedTextureFill : insertRememberedPictureFill;
  doc.transact(t('Picture or texture fill'), () => {
    for (const target of targets) {
      const key = `${slideKey}:${getShapeId(target)}`;
      const remembered = doc.rememberedFills.get(key) ?? {};
      insert(doc.pres, target, bytes, remembered);
      doc.rememberedFills.set(key, remembered);
    }
  });
}
