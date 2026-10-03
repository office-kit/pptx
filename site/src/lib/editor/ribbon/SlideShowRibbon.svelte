<script lang="ts">
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';

  const editor = getEditor();
</script>

<div class="group">
  <div class="items">
    <button
      aria-label={t('Set Up Show')}
      title={t('Set Up Show')}
      onclick={() => (editor.activeDialog = 'showProperties')}
    >
      <svg viewBox="0 0 24 24" aria-hidden="true">
        <rect x="3" y="4" width="18" height="16" rx="1" />
        <path d="M7 9h10M7 13h6M7 17h8" />
      </svg>
      {t('Set Up Show')}
    </button>
    <button
      aria-label={t('Custom Shows')}
      title={t('Custom Shows')}
      onclick={() => editor.openCustomShows()}
    >
      <svg viewBox="0 0 24 24" aria-hidden="true">
        <rect x="3" y="4" width="18" height="16" rx="1" />
        <path d="M7 8h10M7 12h7M7 16h4" />
      </svg>
      {t('Custom Shows')}
    </button>
  </div>
  <span>{t('Set Up')}</span>
</div>

<style>
  .group { display: flex; flex-direction: column; justify-content: space-between; flex-shrink: 0; padding: 0 8px; border-right: 1px solid var(--ok-border); }
  .items { display: flex; align-items: center; flex: 1; gap: 2px; }
  button { display: flex; flex-direction: column; align-items: center; justify-content: center; gap: 4px; width: 82px; min-height: 55px; padding: 4px; font: inherit; font-size: 11px; color: var(--ok-text); background: transparent; border: 1px solid transparent; border-radius: var(--ok-radius); cursor: pointer; }
  button:hover { background: var(--ok-hover); border-color: var(--ok-border); }
  svg { width: 24px; height: 24px; stroke: currentColor; fill: none; stroke-width: 1.2; }
  .group > span { text-align: center; font-size: 10px; color: var(--ok-text-2); padding: 4px 0 1px; }
</style>
