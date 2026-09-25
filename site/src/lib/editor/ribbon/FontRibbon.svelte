<script lang="ts">
  import { getShapeKind, getShapeText, type TextFormat } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { textFormatsInRange } from '../core/text-format-selection.ts';
  import { toggleTextFormat, type TextFormatToggle } from '../core/text-format-toggle.ts';
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
  const formats = $derived(editor.inlineTextFormat?.formats ?? objectFormats ?? []);
  function apply(format: TextFormat, reset = false) {
    if (editor.inlineTextFormat) editor.inlineTextFormat.apply(format, reset);
    else editor.invoke('setShapeTextFormat', { format, options: { reset } });
  }
  function toggle(property: TextFormatToggle) {
    if (editor.inlineTextFormat) editor.inlineTextFormat.toggle(property);
    else apply(toggleTextFormat(formats, property));
  }
</script>

<TextFormatBar ribbon {formats} selected={!!editor.inlineTextFormat || !!objectFormats} onformat={apply} ontoggle={toggle} />
