// Arrange ▸ Rotate's quarter turns and flips, shared by the collapsed Arrange
// menu and the expanded Arrange group of the contextual tabs.
import {
  getShapeFlip,
  getShapeId,
  getShapeRotation,
  getSlideShapes,
  setShapeFlip,
  setShapeRotation,
} from '@office-kit/pptx';
import type { EditorController } from '../core/controller.svelte.ts';
import { selectedShapeIds } from '../core/selection.ts';
import { t } from '../i18n/i18n.svelte.ts';

export type RotateAction = 'right' | 'left' | 'horizontal' | 'vertical';

export const ROTATE_ITEMS: readonly { readonly action: RotateAction; readonly label: string }[] = [
  { action: 'right', label: 'Rotate Right 90°' },
  { action: 'left', label: 'Rotate Left 90°' },
  { action: 'vertical', label: 'Flip Vertical' },
  { action: 'horizontal', label: 'Flip Horizontal' },
];

export const ALIGN_ITEMS = [
  { value: 'left', label: 'Align Left' },
  { value: 'center', label: 'Align Center' },
  { value: 'right', label: 'Align Right' },
  { value: 'top', label: 'Align Top' },
  { value: 'middle', label: 'Align Middle' },
  { value: 'bottom', label: 'Align Bottom' },
] as const;

export function rotateSelection(editor: EditorController, action: RotateAction): void {
  const doc = editor.doc;
  if (!doc.currentSlide || editor.selectionLocked()) return;
  const ids = new Set(selectedShapeIds(doc.selection));
  const shapes = getSlideShapes(doc.currentSlide).filter((shape) => ids.has(getShapeId(shape)));
  if (!shapes.length) return;
  doc.transact(t('Rotate'), () => {
    for (const shape of shapes) {
      if (action === 'right' || action === 'left')
        setShapeRotation(shape, getShapeRotation(shape) + (action === 'right' ? 90 : -90));
      else setShapeFlip(shape, { [action]: !getShapeFlip(shape)?.[action] });
    }
  });
}
