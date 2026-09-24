<script lang="ts">
  // The ribbon. Renders the tab/group/command layout from config.ts. Contextual
  // tabs (Shape Format, Table) only appear when the matching selection is
  // active, mirroring PowerPoint. Buttons dispatch through runOrPrompt, so a
  // command needing arguments opens its (auto-generated or bespoke) dialog.
  import { getEditor } from '../core/context.ts';
  import { RIBBON, type RibbonTab } from './config.ts';
  import { capabilityById } from '../manifest/index.ts';
  import Icon from '../ui/Icon.svelte';
  import ViewRibbon from './ViewRibbon.svelte';
  import ArrangeMenu from './ArrangeMenu.svelte';
  import { t, capLabel } from '../i18n/i18n.svelte.ts';

  const editor = getEditor();
  const doc = editor.doc;

  let activeTab = $state('home');

  const visibleTabs = $derived.by<RibbonTab[]>(() => {
    const sel = doc.selection;
    return RIBBON.filter((t) => {
      if (!t.contextual) return true;
      if (t.contextual === 'shape') return sel.kind === 'shape';
      if (t.contextual === 'cell' || t.contextual === 'table') return sel.kind === 'cell';
      return false;
    });
  });

  // If the active tab disappears (selection changed), fall back to Home.
  $effect(() => {
    if (!visibleTabs.some((t) => t.id === activeTab)) activeTab = 'home';
  });

  const current = $derived(visibleTabs.find((t) => t.id === activeTab) ?? visibleTabs[0]);

  function tabKeys(event: KeyboardEvent) {
    if (!['ArrowLeft', 'ArrowRight', 'Home', 'End'].includes(event.key)) return;
    event.preventDefault(); event.stopPropagation();
    const index = visibleTabs.findIndex(tab => tab.id === activeTab);
    const next = event.key === 'Home' ? 0 : event.key === 'End' ? visibleTabs.length - 1 : (index + (event.key === 'ArrowLeft' ? -1 : 1) + visibleTabs.length) % visibleTabs.length;
    activeTab = visibleTabs[next]!.id;
    (event.currentTarget as HTMLElement).querySelectorAll<HTMLButtonElement>('[role="tab"]')[next]?.focus();
  }

  function tip(id: string): string {
    const cap = capabilityById.get(id);
    return cap ? `${capLabel(cap)} — ${cap.id}` : id;
  }
</script>

<div class="ribbon">
  <div class="tabs" role="tablist" tabindex="-1" aria-label={t('Ribbon')} onkeydown={tabKeys}>
    {#each visibleTabs as tab (tab.id)}
      <button
        class="tab"
        role="tab"
        id="ribbon-tab-{tab.id}"
        aria-selected={activeTab === tab.id}
        aria-controls="ribbon-panel"
        tabindex={activeTab === tab.id ? 0 : -1}
        class:active={activeTab === tab.id}
        class:contextual={tab.contextual}
        onclick={() => (activeTab = tab.id)}
      >
        {t(tab.title)}
      </button>
    {/each}
  </div>

  <div class="groups ok-scroll" id="ribbon-panel" role="tabpanel" aria-labelledby="ribbon-tab-{current?.id}">
    {#if current?.id === 'view'}<ViewRibbon />{/if}
    {#each current?.groups ?? [] as group (group.title)}
      <div class="group">
        <div class="group-items">
          {#if current?.id === 'home' && group.title === 'Arrange'}<ArrangeMenu />
          {:else if current?.id === 'design' && group.title === 'Background'}
            <button class="cmd format-background-trigger" disabled={!doc.currentSlide} aria-label={t('Format Background')} onclick={() => editor.showBackgroundFormat()}>
              <span class="icon"><Icon name="background" /></span>
              <span class="cmd-label">{t('Format Background')}</span>
            </button>
          {:else}
          {#each group.items as item (item.id + (item.label ?? ''))}
            {@const cap = capabilityById.get(item.id)}
            <button
              class="cmd"
              disabled={!editor.canRun(item.id)}
              title={tip(item.id)}
              aria-label={item.label ? t(item.label) : cap ? capLabel(cap) : item.id}
              onclick={() => editor.runOrPrompt(item.id, item.preset ?? {})}
            >
              <span class="icon"><Icon name={item.icon ?? 'dot'} /></span>
              <span class="cmd-label">{item.compactLabel ? t(item.compactLabel) : item.label ? t(item.label) : cap ? capLabel(cap) : item.id}</span>
            </button>
          {/each}
          {/if}
        </div>
        <div class="group-title">{t(group.title)}</div>
      </div>
    {/each}
  </div>
</div>

<style>
  .ribbon {
    min-width: 0;
    background: var(--ok-ribbon);
    border-bottom: 1px solid var(--ok-border);
    display: flex;
    flex-direction: column;
  }
  .tabs {
    display: flex;
    gap: 2px;
    padding: 0 8px;
    height: 30px;
    align-items: flex-end;
  }
  .tab {
    border: none;
    background: none;
    padding: 6px 14px;
    font: inherit;
    font-size: 12px;
    color: var(--ok-text-2);
    cursor: pointer;
    border-radius: var(--ok-radius) var(--ok-radius) 0 0;
  }
  .tab:hover {
    background: var(--ok-hover);
  }
  .tab.active {
    background: var(--ok-ribbon-active);
    color: var(--ok-accent);
    font-weight: 600;
    box-shadow: 0 -2px 0 var(--ok-accent) inset;
  }
  .tab.contextual {
    color: var(--ok-accent);
  }
  .tab.contextual.active {
    color: var(--ok-accent);
  }
  .groups {
    display: flex;
    gap: 0;
    background: var(--ok-ribbon-active);
    min-height: calc(var(--ok-ribbon-h) - 30px);
    padding: 4px 6px 2px;
    overflow-x: auto;
  }
  .group {
    flex-shrink: 0;
    display: flex;
    flex-direction: column;
    justify-content: space-between;
    padding: 0 8px;
    border-right: 1px solid var(--ok-border);
    min-width: 0;
  }
  .group-items {
    display: flex;
    gap: 2px;
    flex: 1;
    align-items: center;
  }
  .cmd {
    display: flex;
    flex-direction: column;
    align-items: center;
    gap: 3px;
    width: 62px;
    padding: 4px 2px;
    border: 1px solid transparent;
    background: none;
    border-radius: var(--ok-radius);
    cursor: pointer;
    font: inherit;
  }
  .cmd:hover:not(:disabled) {
    background: var(--ok-hover);
    border-color: var(--ok-border);
  }
  .cmd:disabled {
    opacity: 0.4;
    cursor: default;
  }
  .icon {
    display: flex;
    align-items: center;
    justify-content: center;
    height: 20px;
    color: var(--ok-text);
  }
  .cmd:disabled .icon {
    color: var(--ok-text-3);
  }
  .cmd-label {
    font-size: 10px;
    color: var(--ok-text-2);
    text-align: center;
    line-height: 1.15;
    max-width: 100%;
    overflow: hidden;
    text-overflow: ellipsis;
    display: -webkit-box;
    -webkit-line-clamp: 2;
    line-clamp: 2;
    -webkit-box-orient: vertical;
  }
  .group-title {
    text-align: center;
    font-size: 10px;
    color: var(--ok-text-3);
    padding-top: 2px;
  }
</style>
