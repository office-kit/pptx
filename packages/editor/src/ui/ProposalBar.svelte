<script lang="ts">
  import type { DeckConflict } from '../merge/deck-merge.ts';
  import type { PendingProposal } from '../core/proposal.ts';
  import { t } from '../i18n/i18n.svelte.ts';

  let { proposal }: { proposal: PendingProposal } = $props();

  const source = $derived(proposal.from === 'source');
  // The proposed version, downloadable before choosing.
  let download = $state('');
  $effect(() => {
    const url = URL.createObjectURL(
      new Blob([new Uint8Array(proposal.edited)], {
        type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      }),
    );
    download = url;
    return () => URL.revokeObjectURL(url);
  });

  function message(item: DeckConflict): string {
    if (item.part === 'ppt/presentation.xml' && !item.shape)
      return t(
        source
          ? 'The slide list or presentation settings were changed both here and in the source.'
          : 'The slide list or presentation settings were changed both here and by the agent.',
      );
    const deleted = item.reason === 'deleted-and-changed';
    const where =
      item.slide === null ? item.part : t('Slide {n}').replace('{n}', String(item.slide));
    if (item.shape) {
      const shape = item.shape.name || `#${item.shape.id}`;
      return t(
        deleted
          ? '{where}: {shape} was deleted on one side and changed on the other.'
          : source
            ? '{where}: {shape} was changed both here and in the source.'
            : '{where}: {shape} was changed both here and by the agent.',
      )
        .replace('{where}', () => where)
        .replace('{shape}', () => shape);
    }
    return t(
      deleted
        ? '{where} was deleted on one side and changed on the other.'
        : source
          ? '{where} was changed both here and in the source.'
          : '{where} was changed both here and by the agent.',
    ).replace('{where}', () => where);
  }
  const messages = $derived([...new Set(proposal.conflicts.map(message))]);
</script>

<div class="conflict" role="alert">
  {#if messages.length}
    <span>{t(source ? 'These edits could not be merged with the source. Your edits are still here.' : "These edits could not be merged with the agent's changes. Your edits are still here.")}</span>
    <ul>{#each messages as item (item)}<li>{item}</li>{/each}</ul>
  {:else}
    <span>{t(source ? 'The source or saved deck changed. Your edits are still here.' : 'The agent changed this presentation as well. Your edits are still here.')}</span>
  {/if}
  <button onclick={() => proposal.choose('mine')}>{t('Keep my edits')}</button>
  <button onclick={() => proposal.choose('theirs')}>{t(source ? 'Use source' : "Use the agent's version")}</button>
  <a href={download} download={source ? 'source.pptx' : 'agent.pptx'}>{t(source ? 'Download source' : "Download the agent's version")}</a>
</div>

<style>
  .conflict {
    display: flex;
    flex-basis: 100%;
    align-items: center;
    flex-wrap: wrap;
    gap: 10px;
    padding: 5px 12px;
    font: 12px system-ui;
    color: #583b00;
    background: #fff0cd;
  }
  .conflict ul { flex-basis: 100%; order: 1; margin: 0; padding-left: 18px; }
  button, a { font: inherit; color: inherit; cursor: pointer; }
  @media (prefers-color-scheme: dark) {
    .conflict { background: #4a3a12; color: #ffe2a3; }
  }
</style>
