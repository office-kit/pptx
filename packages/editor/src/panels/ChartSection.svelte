<script lang="ts">
  import { getShapeChartSpec } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { selectedShapeIds } from '../core/selection.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  const editor = getEditor();
  const chart = $derived.by(() => {
    editor.doc.version;
    const ids = selectedShapeIds(editor.doc.selection);
    const shape = ids.length === 1 ? editor.doc.shapeById(editor.doc.selection.slideIndex, ids[0]!) : null;
    return shape ? getShapeChartSpec(shape) : null;
  });
</script>

{#if chart}
  <section aria-label={t('Chart options')}>
    <strong>{t('Chart options')}</strong>
    <button class="ok-btn" onclick={() => editor.runOrPrompt('setChartSpec')}>{t('Edit chart')}</button>
  </section>
{/if}

<style>
  section { display: grid; gap: 10px; padding: 12px; border-bottom: 1px solid var(--ok-border); }
  strong { font-size: 12px; }
</style>
