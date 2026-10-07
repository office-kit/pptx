// Paragraph alignment of the text being edited, or of every paragraph of the
// selected text shapes: Home ▸ Paragraph's buttons and Format ▸ Alignment.
import {
  getParagraphPropertiesEffective,
  getShapeKind,
  getShapeParagraphCount,
} from '@office-kit/pptx';
import type { EditorController } from './controller.svelte.ts';
import { shapeTextDefaults } from './text-layout-defaults.ts';

export const PARAGRAPH_ALIGNMENTS = [
  { value: 'left', label: 'Align Left' },
  { value: 'center', label: 'Center' },
  { value: 'right', label: 'Align Right' },
  { value: 'justify', label: 'Justify' },
  { value: 'distribute', label: 'Distributed' },
] as const;

export function canAlignParagraphs(editor: EditorController): boolean {
  const shapes = editor.selectedShapes();
  return (
    !!editor.inlineTextFormat ||
    (shapes.length > 0 && shapes.every((shape) => getShapeKind(shape) === 'shape'))
  );
}

/** The shared alignment, or '' when the paragraphs differ or none can align. */
export function paragraphAlignment(editor: EditorController): string {
  editor.doc.version;
  if (editor.inlineTextFormat) return editor.inlineTextFormat.alignment;
  if (!canAlignParagraphs(editor)) return '';
  const values = new Set<string>();
  for (const shape of editor.selectedShapes()) {
    const defaultAlign = shapeTextDefaults(shape).align;
    const count = getShapeParagraphCount(shape);
    for (let index = 0; index < count; index++)
      values.add(
        getParagraphPropertiesEffective(editor.doc.pres, shape, index).align ?? defaultAlign,
      );
    if (!count) values.add(defaultAlign);
  }
  return values.size === 1 ? [...values][0]! : '';
}

export function alignParagraphs(editor: EditorController, value: string): void {
  if (editor.inlineTextFormat) editor.inlineTextFormat.align(value);
  else editor.invoke('setShapeAlignment', { align: value });
}
