<script lang="ts">
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';

  const editor = getEditor();
  const doc = editor.doc;
  const count = $derived(doc.selection.kind === 'shape' ? doc.selection.shapeIds.length : 0);
  const alignments = [
    ['left', 'Align left'], ['center', 'Align center'], ['right', 'Align right'],
    ['top', 'Align top'], ['middle', 'Align middle'], ['bottom', 'Align bottom'],
  ] as const;
</script>

{#if count > 0}
  <section class="arrange" aria-label={t('Arrange')}>
    <strong>{t('Arrange')}</strong>
    {#if count === 1}
      <p>{t('Align to slide')}</p>
    {:else}
      <select aria-label={t('Alignment reference')} bind:value={editor.alignmentReference}>
        <option value="selection">{t('Align to selection')}</option>
        <option value="slide">{t('Align to slide')}</option>
      </select>
    {/if}
    <div class="buttons">
      {#each alignments as [alignment, label]}
        <button onclick={() => editor.alignSelection(alignment, editor.alignmentReference)}>{t(label)}</button>
      {/each}
      <button disabled={count === 2 && editor.alignmentReference !== 'slide'} onclick={() => editor.distributeSelection('horizontal')}>{t('Distribute horizontally')}</button>
      <button disabled={count === 2 && editor.alignmentReference !== 'slide'} onclick={() => editor.distributeSelection('vertical')}>{t('Distribute vertically')}</button>
      <button disabled={!editor.canRun('groupShapes')} onclick={() => editor.invoke('groupShapes')}>{t('Group')}</button>
      <button disabled={!editor.canRun('ungroupShapes')} onclick={() => editor.invoke('ungroupShapes')}>{t('Ungroup')}</button>
    </div>
  </section>
{/if}

<style>
  .arrange { padding: 12px; border-bottom: 1px solid var(--ok-border); }
  p { font-size: 11px; color: var(--ok-text-2); margin: 6px 0; }
  select { width: 100%; margin: 6px 0; padding: 5px; font: inherit; font-size: 11px; background: var(--ok-panel); color: var(--ok-text); border: 1px solid var(--ok-border); border-radius: var(--ok-radius); }
  .buttons { display: grid; grid-template-columns: 1fr 1fr; gap: 4px; }
  button { font: inherit; font-size: 11px; padding: 6px 4px; border: 1px solid var(--ok-border); border-radius: var(--ok-radius); background: var(--ok-panel); color: var(--ok-text); cursor: pointer; }
  button:hover:not(:disabled) { background: var(--ok-hover); }
  button:disabled { opacity: 0.4; cursor: default; }
</style>
