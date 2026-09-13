<script lang="ts">
  // Ctrl/Cmd+K command palette. Lists every capability (all 147), searchable,
  // grouped by category, with a live enabled/disabled state from the current
  // selection. This is the guaranteed-reachable path: whatever the ribbon has
  // not surfaced, the user can still run from here.
  import { getEditor } from '../core/context.ts';
  import { capabilities, CATEGORY_LABELS } from '../manifest/index.ts';
  import { t, capLabel, catLabel } from '../i18n/i18n.svelte.ts';

  const editor = getEditor();
  let query = $state('');
  let activeIndex = $state(0);

  const filtered = $derived.by(() => {
    const q = query.trim().toLowerCase();
    const list = capabilities.filter((c) => {
      if (!q) return true;
      return (
        c.id.toLowerCase().includes(q) ||
        c.labelEn.toLowerCase().includes(q) ||
        c.labelJa.includes(query.trim()) ||
        c.category.includes(q)
      );
    });
    return list;
  });

  function pick(id: string) {
    editor.togglePalette(false);
    editor.runOrPrompt(id);
  }

  function onKey(e: KeyboardEvent) {
    if (e.key === 'ArrowDown') {
      e.preventDefault();
      activeIndex = Math.min(activeIndex + 1, filtered.length - 1);
    } else if (e.key === 'ArrowUp') {
      e.preventDefault();
      activeIndex = Math.max(activeIndex - 1, 0);
    } else if (e.key === 'Enter') {
      e.preventDefault();
      const item = filtered[activeIndex];
      if (item) pick(item.id);
    }
  }

  $effect(() => {
    query;
    activeIndex = 0;
  });
</script>

<div
  class="backdrop"
  role="presentation"
  onclick={(e) => {
    if (e.target === e.currentTarget) editor.togglePalette(false);
  }}
>
  <div class="palette ok-editor" role="dialog" aria-modal="true" aria-label="Command palette">
    <!-- svelte-ignore a11y_autofocus -- palette is a transient modal; focusing the search on open is the expected UX -->
    <input
      class="search"
      type="text"
      placeholder={t('Search capabilities…')}
      bind:value={query}
      onkeydown={onKey}
      autofocus
    />
    <div class="ok-scroll list" role="listbox">
      {#each filtered as cap, i (cap.id)}
        {@const enabled = editor.canRun(cap.id)}
        <button
          class="item"
          class:active={i === activeIndex}
          class:disabled={!enabled}
          role="option"
          aria-selected={i === activeIndex}
          onmouseenter={() => (activeIndex = i)}
          onclick={() => pick(cap.id)}
        >
          <span class="label">{capLabel(cap)}</span>
          <span class="cat">{catLabel(CATEGORY_LABELS[cap.category])}</span>
          <code class="id">{cap.id}</code>
        </button>
      {:else}
        <p class="empty">{t('No capability matches')} “{query}”.</p>
      {/each}
    </div>
    <footer>
      <span>↑↓ {t('navigate')} · Enter {t('run')} · Esc {t('close')}</span>
      <span>{filtered.length} / {capabilities.length}</span>
    </footer>
  </div>
</div>

<style>
  .backdrop {
    position: fixed;
    inset: 0;
    background: rgba(0, 0, 0, 0.28);
    display: grid;
    place-items: start center;
    padding-top: 10vh;
    z-index: 210;
  }
  .palette {
    width: min(620px, 94vw);
    max-height: 70vh;
    background: var(--ok-panel);
    border-radius: var(--ok-radius-lg);
    box-shadow: var(--ok-shadow-lg);
    display: flex;
    flex-direction: column;
    overflow: hidden;
  }
  .search {
    border: none;
    border-bottom: 1px solid var(--ok-border);
    padding: 14px 16px;
    font-size: 15px;
    font-family: var(--ok-font);
    outline: none;
  }
  .list {
    overflow: auto;
    padding: 4px;
  }
  .item {
    display: grid;
    grid-template-columns: 1fr auto auto;
    align-items: center;
    gap: 10px;
    width: 100%;
    text-align: left;
    padding: 7px 10px;
    border: none;
    background: none;
    border-radius: var(--ok-radius);
    cursor: pointer;
    font: inherit;
  }
  .item.active {
    background: var(--ok-selected);
  }
  .item.disabled {
    opacity: 0.5;
  }
  .label {
    font-size: 13px;
  }
  .cat {
    font-size: 10px;
    color: var(--ok-text-2);
    background: var(--ok-bg);
    padding: 1px 7px;
    border-radius: 10px;
  }
  .id {
    font-family: var(--ok-mono);
    font-size: 10px;
    color: var(--ok-text-3);
  }
  .empty {
    padding: 20px;
    text-align: center;
    color: var(--ok-text-2);
  }
  footer {
    display: flex;
    justify-content: space-between;
    padding: 8px 14px;
    border-top: 1px solid var(--ok-border);
    background: var(--ok-panel-2);
    color: var(--ok-text-3);
    font-size: 11px;
  }
</style>
