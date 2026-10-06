<script lang="ts">
  // View ▸ Notes Page: the printed notes page — the slide image above its
  // speaker notes on a portrait page — with the notes editable in place.
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import NotesPane from './NotesPane.svelte';

  const editor = getEditor();
  const doc = editor.doc;
</script>

<div class="notes-page-view" role="region" aria-label={t('Notes Page')}>
  <div class="page">
    <button class="slide" aria-label={t('Edit slide')} ondblclick={() => editor.setViewMode('normal')}>{@html doc.currentSvg}</button>
    {#if doc.currentSlide}{#key doc.currentSlide}<div class="notes"><NotesPane /></div>{/key}{/if}
  </div>
</div>

<style>
  .notes-page-view { display: flex; justify-content: center; min-width: 0; min-height: 0; overflow: auto; padding: 24px; background: var(--ok-canvas-bg, #d9d9d9); }
  .page { display: flex; flex-direction: column; gap: 24px; width: min(640px, 100%); aspect-ratio: 7.5 / 10; padding: 6% 8%; background: #fff; color: #000; box-shadow: 0 1px 4px #0003; }
  .slide { flex-shrink: 0; padding: 0; border: 1px solid #0003; background: #fff; cursor: default; }
  .slide :global(svg) { display: block; width: 100%; height: auto; }
  .notes { flex: 1; min-height: 0; display: flex; }
  .notes :global(> *) { flex: 1; }
</style>
