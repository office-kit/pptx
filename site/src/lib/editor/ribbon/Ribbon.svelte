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
  import BackgroundStyles from './BackgroundStyles.svelte';
  import ArrangeMenu from './ArrangeMenu.svelte';
  import FontRibbon from './FontRibbon.svelte';
  import LineSpacingMenu from './LineSpacingMenu.svelte';
  import ParagraphAlignment from './ParagraphAlignment.svelte';
  import VideoFormatRibbon from './VideoFormatRibbon.svelte';
  import { t, capLabel } from '../i18n/i18n.svelte.ts';

  const editor = getEditor();
  const doc = editor.doc;

  let activeTab = $state('home');
  let collapsed = $state(false);
  let openGroup = $state<string | null>(null);

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
    openGroup = null;
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

  function dismissGroup(event: PointerEvent) {
    const target = event.target;
    if (!(target instanceof Element) || !target.closest('.group-menu, .group-menu-trigger')) openGroup = null;
  }

  function groupKeydown(event: KeyboardEvent) {
    if (event.key === 'Escape' && openGroup) { event.preventDefault(); openGroup = null; }
  }
</script>

<svelte:window onpointerdown={dismissGroup} onkeydown={groupKeydown} />
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

  <div class="groups ok-scroll" class:home-groups={current?.id === 'home'} hidden={collapsed} id="ribbon-panel" role="tabpanel" aria-labelledby="ribbon-tab-{current?.id}">
    {#if current?.title === 'Video Format'}
      <VideoFormatRibbon />
    {:else}
    {#if current?.id === 'playback'}<MediaPlaybackRibbon />{/if}
    {#if current?.id === 'view'}<ViewRibbon />{/if}
    {#if current?.id === 'slideShow'}<SlideShowRibbon />{/if}
    {#each current?.groups ?? [] as group (group.title)}
      <div class="group" class:font-group={current?.id === 'home' && group.title === 'Font'} class:paragraph-group={current?.id === 'home' && group.title === 'Paragraph'}>
        <div class="group-items">
          {#if current?.id === 'home' && group.title === 'Arrange'}<ArrangeMenu />
          {:else if current?.id === 'home' && group.title === 'Font'}<FontRibbon />
          {:else if current?.id === 'design' && group.title === 'Background'}
            <BackgroundStyles />
          {:else}
          {#each group.items as item (item.id + (item.label ?? ''))}
            {@const cap = capabilityById.get(item.id)}
            {#if current?.id === 'home' && item.id === 'setParagraphAlignment'}<ParagraphAlignment />
            {:else if current?.id === 'home' && item.id === 'setParagraphLineSpacing'}<LineSpacingMenu />
            {:else}
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
            {/if}
          {/each}
          {/if}
        </div>
        <div class="group-title">{t(group.title)}</div>
      </div>
    {/each}
    {#if current?.id === 'home'}
      <div class="compact-groups" role="toolbar" aria-label={t('Home ribbon groups')}>
        {#each current.groups as group (group.title)}
          {#if group.title !== 'Font'}
          <button class="group-menu-trigger" class:paragraph-trigger={group.title === 'Paragraph'} aria-haspopup="menu" aria-expanded={openGroup === group.title} onclick={() => openGroup = openGroup === group.title ? null : group.title}>
            <span>{t(group.title)}</span><span aria-hidden="true">⌄</span>
          </button>
          {/if}
        {/each}
      </div>
      {#if openGroup}
        {@const group = current.groups.find(item => item.title === openGroup)}
        {#if group}
          <div class="group-menu" role="menu" tabindex="-1" aria-label={t(group.title)}>
            <div class="group-menu-items">
              {#if group.title === 'Arrange'}<ArrangeMenu />
              {:else if group.title === 'Font'}<FontRibbon />
              {:else}
                {#each group.items as item (item.id + (item.label ?? ''))}
                  {@const cap = capabilityById.get(item.id)}
                  {#if item.id === 'setParagraphAlignment'}<ParagraphAlignment />
                  {:else if item.id === 'setParagraphLineSpacing'}<LineSpacingMenu />
                  {:else}
                    <button class="cmd" role="menuitem" disabled={!editor.canRun(item.id)} title={tip(item.id)} aria-label={item.label ? t(item.label) : cap ? capLabel(cap) : item.id} onclick={() => { openGroup = null; editor.runOrPrompt(item.id, item.preset ?? {}); }}>
                      <span class="icon"><Icon name={item.icon ?? 'dot'} /></span>
                      <span class="cmd-label">{item.compactLabel ? t(item.compactLabel) : item.label ? t(item.label) : cap ? capLabel(cap) : item.id}</span>
                    </button>
                  {/if}
                {/each}
              {/if}
            </div>
          </div>
        {/if}
      {/if}
    {/if}
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
  .compact-groups, .group-menu { display: none; }
  @media (max-width: 1600px) {
    .groups.home-groups { position: relative; overflow: visible; }
    .groups.home-groups > .group { display: none; }
    .groups.home-groups > .group.font-group { display: flex; }
    .groups.home-groups .compact-groups { display: flex; flex-wrap: wrap; width: 100%; gap: 4px; }
    .group-menu-trigger { display: inline-flex; align-items: center; justify-content: center; gap: 5px; min-width: 76px; padding: 8px 7px; border: 1px solid var(--ok-border); border-radius: var(--ok-radius); background: var(--ok-panel); color: var(--ok-text); font: inherit; font-size: 11px; cursor: pointer; }
    .group-menu-trigger:hover, .group-menu-trigger[aria-expanded='true'] { background: var(--ok-hover); border-color: var(--ok-accent); }
    .group-menu { position: absolute; z-index: 400; display: block; left: 6px; right: 6px; top: calc(100% - 2px); padding: 7px; border: 1px solid var(--ok-border); border-radius: 6px; background: var(--ok-panel); box-shadow: var(--ok-shadow-lg); }
    .group-menu-items { display: flex; flex-wrap: wrap; align-items: center; gap: 3px; }
  }
  @media (min-width: 1100px) and (max-width: 1600px) {
    .groups.home-groups > .group.paragraph-group { display: flex; }
    .paragraph-trigger { display: none; }
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
