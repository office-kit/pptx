<script lang="ts">
  import { getShapeKind, getShapeParagraphCount, getParagraphPropertiesEffective } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { shapeTextDefaults } from '../core/text-layout-defaults.ts';
  import { t } from '../i18n/i18n.svelte.ts';

  const editor = getEditor();
  const shapes = $derived(editor.selectedShapes());
  const enabled = $derived(!!editor.inlineTextFormat || (shapes.length > 0 && shapes.every(shape => getShapeKind(shape) === 'shape')));
  const alignment = $derived.by(() => {
    editor.doc.version;
    if (editor.inlineTextFormat) return editor.inlineTextFormat.alignment;
    if (!enabled) return '';
    const values = new Set<string>();
    for (const shape of shapes) {
      const defaultAlign = shapeTextDefaults(shape).align;
      const count = getShapeParagraphCount(shape);
      for (let index = 0; index < count; index++) values.add(getParagraphPropertiesEffective(editor.doc.pres, shape, index).align ?? defaultAlign);
      if (!count) values.add(defaultAlign);
    }
    return values.size === 1 ? [...values][0] : '';
  });
  function align(value: string) {
    if (editor.inlineTextFormat) editor.inlineTextFormat.align(value);
    else editor.invoke('setShapeAlignment', { align: value });
  }
  const options = [
    { value: 'left', label: 'Align Left' },
    { value: 'center', label: 'Center' },
    { value: 'right', label: 'Align Right' },
    { value: 'justify', label: 'Justify' },
    { value: 'distribute', label: 'Distributed' },
  ];
</script>

<div class="paragraph-alignment" role="group" aria-label={t('Paragraph alignment')}>
  {#each options as option}
    <button class="ok-btn" aria-label={t(option.label)} title={t(option.label)} aria-pressed={alignment === option.value} disabled={!enabled} onmousedown={event => event.preventDefault()} onclick={() => align(option.value)}>
      <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" aria-hidden="true">
        <path d="M3 5h18 M3 13h18" />
        {#if option.value === 'left'}<path d="M3 9h12 M3 17h12" />
        {:else if option.value === 'center'}<path d="M6 9h12 M6 17h12" />
        {:else if option.value === 'right'}<path d="M9 9h12 M9 17h12" />
        {:else}<path d="M3 9h18 M3 17h18" />{/if}
        {#if option.value === 'distribute'}<path d="M3 21h18 M5 19l-2 2 2 2 M19 19l2 2-2 2" />{/if}
      </svg>
    </button>
  {/each}
</div>

<style>
  .paragraph-alignment { display: flex; align-self: flex-end; gap: 2px; }
  button { display: grid; place-items: center; padding: 3px; min-width: 25px; }
  button[aria-pressed='true'] { background: var(--ok-hover); border-color: var(--ok-accent); }
</style>
