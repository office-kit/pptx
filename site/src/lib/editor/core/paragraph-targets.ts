import {
  getParagraphPropertiesEffective,
  getShapeKind,
  getShapeParagraphCount,
  getTableCellParagraphs,
  getTableCells,
  getTableCellSpan,
  type ParagraphAlignment,
  type ParagraphProperties,
  type setParagraphAlignment,
} from '@office-kit/pptx';
import type { EditorController } from './controller.svelte.ts';
import { tableCellsInRange, tableSelectionBlock } from './table-selection.ts';
import { shapeTextDefaults } from './text-layout-defaults.ts';
import { t } from '../i18n/i18n.svelte.ts';

export type ParagraphEdit = (
  shape: Parameters<typeof setParagraphAlignment>[0],
  index: number,
) => void;

interface ParagraphTarget {
  readonly shape: Parameters<ParagraphEdit>[0];
  readonly index: number;
  readonly defaultAlign: ParagraphAlignment;
}

/**
 * The paragraphs a Home ribbon paragraph command acts on when no text is being
 * edited: every paragraph of the selected text shapes, or of the selected
 * table cells (merged-away cells excluded).
 */
function selectionParagraphs(editor: EditorController): ParagraphTarget[] {
  const doc = editor.doc;
  const selection = doc.selection;
  if (selection.kind === 'cell') {
    const table = doc.shapeById(selection.slideIndex, selection.shapeId);
    if (!table) return [];
    return [...tableCellsInRange(getTableCells(table), tableSelectionBlock(selection))]
      .filter((cell) => {
        const span = getTableCellSpan(cell);
        return !span.hMerge && !span.vMerge;
      })
      .flatMap((shape) =>
        getTableCellParagraphs(shape).map((_, index) => ({
          shape,
          index,
          defaultAlign: 'left' as const,
        })),
      );
  }
  const shapes = editor.selectedShapes();
  if (!shapes.every((shape) => getShapeKind(shape) === 'shape')) return [];
  return shapes.flatMap((shape) =>
    Array.from({ length: getShapeParagraphCount(shape) }, (_, index) => ({
      shape,
      index,
      defaultAlign: shapeTextDefaults(shape).align,
    })),
  );
}

/** Effective properties of the targeted paragraphs, inline editing first. */
export function targetParagraphProperties(editor: EditorController): ParagraphProperties[] {
  if (editor.inlineTextFormat) return editor.inlineTextFormat.paragraphs;
  return selectionParagraphs(editor).map(({ shape, index, defaultAlign }) => {
    const props = getParagraphPropertiesEffective(editor.doc.pres, shape, index);
    return { ...props, align: props.align ?? defaultAlign };
  });
}

/** Applies `edit` to every targeted paragraph as one undo step. */
export function editTargetParagraphs(editor: EditorController, edit: ParagraphEdit): void {
  if (editor.inlineTextFormat) {
    editor.inlineTextFormat.editParagraphs(edit);
    return;
  }
  const targets = selectionParagraphs(editor);
  editor.doc.transact(t('Format paragraphs'), () => {
    for (const { shape, index } of targets) edit(shape, index);
  });
}
