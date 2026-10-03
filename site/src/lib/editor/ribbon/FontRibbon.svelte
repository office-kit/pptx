<script lang="ts">
  import { getPresentationFonts, getShapeKind, getShapeText, getTableCells, getTableCellText, isTableShape, setShapeText, setTableCellText, setTableCellTextFormat, type TextCase, type TextFormat } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { textFormatsInRange } from '../core/text-format-selection.ts';
  import { toggleTextFormat, type TextFormatToggle } from '../core/text-format-toggle.ts';
  import { stepShapeFontSize, stepTableCellFontSize } from '../core/font-size.ts';
  import { tableCellsInRange, tableSelectionBlock } from '../core/table-selection.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import TextFormatBar from '../ui/TextFormatBar.svelte';

  const editor = getEditor();
  const objectFormats = $derived.by(() => {
    editor.doc.version;
    const shapes = editor.selectedShapes();
    if (!shapes.length || !shapes.every(shape => getShapeKind(shape) === 'shape')) return null;
    return shapes.flatMap(shape => {
      const formats = textFormatsInRange(shape, { start: 0, end: getShapeText(shape).length }, undefined, { pres: editor.doc.pres });
      return formats.length ? formats : [{}];
    });
  });
  const cellFormats = $derived.by(() => {
    const selection = editor.doc.selection;
    editor.doc.version;
    if (selection.kind !== 'cell') return null;
    const table = editor.doc.shapeById(selection.slideIndex, selection.shapeId);
    if (!table) return null;
    const tableCells = getTableCells(table);
    const block = tableSelectionBlock(selection);
    const result: TextFormat[] = [];
    for (let row = block.row; row < block.row + block.rowSpan; row++) {
      for (let col = block.col; col < block.col + block.colSpan; col++) {
        const cell = tableCells[row]?.[col];
        if (cell) result.push(...textFormatsInRange(table, { start: 0, end: getTableCellText(cell).length }, { row, col }, { pres: editor.doc.pres }));
      }
    }
    return result;
  });
  const activeTextFormat = $derived(editor.inlineTextFormat ?? editor.notesInlineTextFormat);
  const formats = $derived(activeTextFormat?.formats ?? cellFormats ?? objectFormats ?? []);
  const themeFonts = $derived.by(() => {
    editor.doc.version;
    const fonts = getPresentationFonts(editor.doc.pres);
    return fonts ? [fonts.majorLatin, fonts.majorEastAsian, fonts.majorComplexScript, fonts.minorLatin, fonts.minorEastAsian, fonts.minorComplexScript].filter((font): font is string => !!font) : [];
  });
  function apply(format: TextFormat, reset = false) {
    if (activeTextFormat) activeTextFormat.apply(format, reset);
    else if (editor.doc.selection.kind === 'cell') {
      const selection = editor.doc.selection;
      const table = editor.doc.shapeById(selection.slideIndex, selection.shapeId);
      if (!table) return;
      const cells = tableCellsInRange(getTableCells(table), tableSelectionBlock(selection));
      editor.doc.transact(t(reset ? 'Clear text formatting' : 'Format selected cells'), () => {
        for (const cell of cells) setTableCellTextFormat(cell, format, { reset });
      });
    } else editor.invoke('setShapeTextFormat', { format, options: { reset } });
  }
  function toggle(property: TextFormatToggle) {
    if (activeTextFormat) activeTextFormat.toggle(property);
    else apply(toggleTextFormat(formats, property));
  }
  function changeCase(value: TextCase) {
    if (activeTextFormat) { activeTextFormat.changeCase?.(value); return; }
    if (editor.doc.selection.kind === 'cell') {
      const selection = editor.doc.selection;
      const table = editor.doc.shapeById(selection.slideIndex, selection.shapeId);
      if (!table) return;
      const cells = tableCellsInRange(getTableCells(table), tableSelectionBlock(selection));
      editor.doc.transact(t('Change Case'), () => { for (const cell of cells) setTableCellText(cell, { case: value }); });
      return;
    }
    editor.doc.transact(t('Change Case'), () => {
      for (const shape of editor.selectedShapes()) if (getShapeKind(shape) === 'shape') setShapeText(shape, { case: value });
    });
  }
  function stepFontSize(direction: 1 | -1) {
    if (activeTextFormat?.fontSize) {
      activeTextFormat.fontSize(direction);
      return;
    }
    if (editor.doc.selection.kind === 'cell') {
      const selection = editor.doc.selection;
      const table = editor.doc.shapeById(selection.slideIndex, selection.shapeId);
      if (!table) return;
      const cells = tableCellsInRange(getTableCells(table), tableSelectionBlock(selection));
      editor.doc.transact(t(direction > 0 ? 'Increase Font Size' : 'Decrease Font Size'), () => {
        for (const cell of cells) stepTableCellFontSize(cell, direction, { start: 0, end: getTableCellText(cell).length });
      });
      return;
    }
    const shapes = editor.selectedShapes();
    if (!shapes.length) return;
    editor.doc.transact(t(direction > 0 ? 'Increase Font Size' : 'Decrease Font Size'), () => {
      for (const shape of shapes) if (getShapeKind(shape) === 'shape') stepShapeFontSize(editor.doc.pres, shape, direction);
    });
  }
</script>

<TextFormatBar ribbon {formats} fontFamilies={themeFonts} selected={!!activeTextFormat || editor.doc.selection.kind === 'cell' || editor.selectedShapes().some(isTableShape) || !!objectFormats} onformat={apply} oncase={changeCase} onfontsize={stepFontSize} ontoggle={toggle} />
