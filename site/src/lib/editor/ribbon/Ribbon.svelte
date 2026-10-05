<script lang="ts">
  // The ribbon. Renders the tab/group/command layout from config.ts. Contextual
  // tabs (Shape Format, Table) only appear when the matching selection is
  // active, mirroring PowerPoint. Buttons dispatch through runOrPrompt, so a
  // command needing arguments opens its (auto-generated or bespoke) dialog.
  import { getShapeMedia, getShapeMediaPlayback } from '@office-kit/pptx';
  import MediaPlaybackRibbon from './MediaPlaybackRibbon.svelte';
  import { getEditor } from '../core/context.ts';
  import { RIBBON, type RibbonTab } from './config.ts';
  import { capabilityById } from '../manifest/index.ts';
  import Icon from '../ui/Icon.svelte';
  import ViewRibbon from './ViewRibbon.svelte';
  import SlideShowRibbon from './SlideShowRibbon.svelte';
  import TransitionsRibbon from './TransitionsRibbon.svelte';
  import DrawRibbon from './DrawRibbon.svelte';
  import RecordRibbon from './RecordRibbon.svelte';
  import AnimationsRibbon from './AnimationsRibbon.svelte';
  import BackgroundStyles from './BackgroundStyles.svelte';
  import ShapeQuickStyles from './ShapeQuickStyles.svelte';
  import HomeRibbon from './HomeRibbon.svelte';
  import VideoFormatRibbon from './VideoFormatRibbon.svelte';
  import { t, capLabel } from '../i18n/i18n.svelte.ts';

  const editor = getEditor();
  const doc = editor.doc;

  let activeTab = $state('home');
  let collapsed = $state(false);

  const visibleTabs = $derived.by<RibbonTab[]>(() => {
    const sel = doc.selection;
    doc.version;
    const shapes = editor.selectedShapes();
    const media = shapes.length === 1 && getShapeMediaPlayback(shapes[0]!) !== null;
    const kind = shapes.length === 1 ? getShapeMedia(shapes[0]!)?.kind : undefined;
    return RIBBON.filter((t) => {
      if (!t.contextual) return true;
      if (t.contextual === 'media') return media;
      if (t.contextual === 'shape') return sel.kind === 'shape';
      if (t.contextual === 'cell' || t.contextual === 'table') return sel.kind === 'cell';
      return false;
    }).map(tab => tab.id === 'shape' && kind === 'video' ? { ...tab, title: 'Video Format' } : tab);
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
    collapsed = false;
    (event.currentTarget as HTMLElement).querySelectorAll<HTMLButtonElement>('[role="tab"]')[next]?.focus();
  }

  function tip(id: string): string {
    const cap = capabilityById.get(id);
    return cap ? `${capLabel(cap)} — ${cap.id}` : id;
  }

</script>

<div class="ribbon">
  <div class="tab-row">
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
          onclick={() => { activeTab = tab.id; collapsed = false; }}
        >
          {t(tab.title)}
        </button>
      {/each}
    </div>

    <button class="ribbon-toggle" aria-label={t(collapsed ? 'Expand ribbon' : 'Collapse ribbon')} title={t(collapsed ? 'Expand ribbon' : 'Collapse ribbon')} aria-expanded={!collapsed} aria-controls="ribbon-panel" onclick={() => (collapsed = !collapsed)}>{collapsed ? '⌄' : '⌃'}</button>
  </div>

  <div class="groups ok-scroll" hidden={collapsed} id="ribbon-panel" role="tabpanel" aria-labelledby="ribbon-tab-{current?.id}">
    {#if current?.title === 'Video Format'}
      <VideoFormatRibbon />
    {:else}
    {#if current?.id === 'playback'}<MediaPlaybackRibbon />{/if}
    {#if current?.id === 'view'}<ViewRibbon />{/if}
    {#if current?.id === 'slideShow'}<SlideShowRibbon />{/if}
    {#if current?.id === 'draw'}<DrawRibbon />{/if}
    {#if current?.id === 'record'}<RecordRibbon />{/if}
    {#if current?.id === 'transitions'}<TransitionsRibbon />{/if}
    {#if current?.id === 'animations'}<AnimationsRibbon />{/if}
    {#if current?.id === 'home'}<HomeRibbon />{/if}
    {#if current?.id === 'shape'}
      <div class="group shape-style-group">
        <div class="group-items"><ShapeQuickStyles inline /></div>
      </div>
    {/if}
    {#each current?.groups ?? [] as group (group.title)}
      <div class="group" role="group" aria-label={t(group.title)}>
        <div class="group-items">
          {#if current?.id === 'design' && group.title === 'Background'}
            <BackgroundStyles />
          {:else}
          {#each group.items as item (item.id + (item.label ?? ''))}
            {@const cap = capabilityById.get(item.id)}
            <button
              class="cmd"
              disabled={!editor.canRun(item.id)}
              title={tip(item.id)}
              aria-label={item.label ? t(item.label) : cap ? capLabel(cap) : item.id}
              onclick={() => (item.run ? item.run(editor) : editor.runOrPrompt(item.id, item.preset ?? {}))}
            >
              <span class="icon"><Icon name={item.icon ?? 'dot'} size={32} /></span>
              <span class="cmd-label">{item.compactLabel ? t(item.compactLabel) : item.label ? t(item.label) : cap ? capLabel(cap) : item.id}</span>
            </button>
          {/each}
          {/if}
        </div>
      </div>
    {/each}
    {/if}
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
  .tab-row { display: flex; min-width: 0; align-items: center; }
  .ribbon-toggle { flex: none; width: 30px; height: 28px; margin: 0 4px; border: none; background: none; color: var(--ok-text-2); cursor: pointer; font-size: 18px; }
  .ribbon-toggle:hover { background: var(--ok-hover); }
  .groups[hidden] { display: none; }
  .tabs {
    flex: 1;
    min-width: 0;
    overflow-x: auto;
    display: flex;
    gap: 2px;
    padding: 0 8px;
    height: 30px;
    align-items: flex-end;
  }
  .tab {
    flex-shrink: 0;
    border: none;
    background: none;
    padding: 6px 14px;
    font: inherit;
    font-size: 13px;
    color: var(--ok-text-2);
    cursor: pointer;
    border-radius: var(--ok-radius) var(--ok-radius) 0 0;
  }
  .tab:hover {
    background: var(--ok-hover);
  }
  /* Mac PowerPoint marks the active tab with bold text and an underline bar,
     on the same background as the commands below it. */
  .tab.active {
    color: var(--ok-text);
    font-weight: 600;
    box-shadow: 0 -3px 0 var(--ok-accent) inset;
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
    background: var(--ok-ribbon);
    min-height: calc(var(--ok-ribbon-h) - 30px);
    padding: 4px 6px 2px;
    overflow-x: auto;
  }
  .group {
    flex-shrink: 0;
    display: flex;
    align-items: center;
    padding: 0 8px;
    border-right: 1px solid var(--ok-border);
    min-width: 0;
  }
  .group-items {
    display: flex;
    gap: 2px;
    align-items: center;
  }
  /* Mac PowerPoint's large ribbon button: a 32px icon over a one- or two-line
     caption, with no group captions under the commands. */
  .cmd {
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: flex-start;
    gap: 2px;
    min-width: 52px;
    min-height: 66px;
    padding: 3px 4px;
    border: 1px solid transparent;
    background: none;
    border-radius: var(--ok-radius);
    color: var(--ok-text);
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
    height: 32px;
  }
  .cmd-label {
    font-size: 11px;
    text-align: center;
    line-height: 1.15;
    max-width: 84px;
  }
</style>
