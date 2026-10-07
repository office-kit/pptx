<script lang="ts">
  import { getEditor } from '../core/context.ts';
  import { alignParagraphs, canAlignParagraphs, paragraphAlignment, PARAGRAPH_ALIGNMENTS } from '../core/paragraph-alignment.ts';
  import { t } from '../i18n/i18n.svelte.ts';

  const editor = getEditor();
  const enabled = $derived(canAlignParagraphs(editor));
  const alignment = $derived(paragraphAlignment(editor));
  const align = (value: string) => alignParagraphs(editor, value);
  const options = PARAGRAPH_ALIGNMENTS;
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
  .paragraph-alignment { display: flex; align-self: flex-end; gap: 0; }
  button { display: grid; place-items: center; width: 26px; height: 26px; padding: 0; }
  button[aria-pressed='true'] { background: var(--ok-hover); border-color: var(--ok-accent); }
</style>
