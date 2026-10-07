<script lang="ts">
  // The ribbon. Renders the tab/group/command layout from config.ts. Contextual
  // tabs (Shape Format, Table) only appear when the matching selection is
  // active, mirroring PowerPoint. Buttons dispatch through runOrPrompt, so a
  // command needing arguments opens its (auto-generated or bespoke) dialog.
  import { getShapeChartSpec, getShapeMedia, getShapeMediaPlayback, isTableShape } from '@office-kit/pptx';
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
  import ReviewRibbon from './ReviewRibbon.svelte';
  import AnimationsRibbon from './AnimationsRibbon.svelte';
  import ShapeFormatRibbon from './ShapeFormatRibbon.svelte';
  import HomeRibbon from './HomeRibbon.svelte';
  import DesignRibbon from './DesignRibbon.svelte';
  import InsertRibbon from './InsertRibbon.svelte';
  import VideoFormatRibbon from './VideoFormatRibbon.svelte';
  import { t, capLabel } from '../i18n/i18n.svelte.ts';
  import { downloadPptx } from '../core/download.ts';

  const editor = getEditor();
  const doc = editor.doc;

  let activeTab = $state('home');
  let collapsed = $state(false);
  let shareOpen = $state(false);
  // The dev shell owns the Agents task pane and reports its state; null until
  // a host announces one, so the button never shows in the standalone editor.
  let agentsOpen = $state<boolean | null>(null);
  let agentsButton = $state<HTMLButtonElement>();
  function onHostMessage(event: MessageEvent) {
    if (window.parent === window || event.origin !== window.location.origin || event.source !== window.parent) return;
    const data: unknown = event.data;
    if (!data || typeof data !== 'object' || !('type' in data) || data.type !== 'host-panes' || !('agents' in data) || typeof data.agents !== 'boolean') return;
    agentsOpen = data.agents;
    if ('focus' in data && data.focus === true) agentsButton?.focus();
  }
  function toggleAgents() {
    window.parent.postMessage({ type: 'editor-command', action: 'agents' }, window.location.origin);
  }
  async function sendCopy() {
    shareOpen = false;
    try {
      await downloadPptx(editor.doc);
    } catch (err) {
      editor.toast('error', `${t('Save failed')}: ${(err as Error).message}`);
    }
  }

  const visibleTabs = $derived.by<RibbonTab[]>(() => {
    const sel = doc.selection;
    doc.version;
    const shapes = editor.selectedShapes();
    const media = shapes.length === 1 && getShapeMediaPlayback(shapes[0]!) !== null;
    const kind = shapes.length === 1 ? getShapeMedia(shapes[0]!)?.kind : undefined;
    const chart = shapes.length === 1 && getShapeChartSpec(shapes[0]!) !== null;
    // PowerPoint shows Table Design and Layout for a selected table as well as
    // for cells being edited, and replaces Shape Format with them.
    const table = sel.kind === 'cell' || (shapes.length === 1 && isTableShape(shapes[0]!));
    return RIBBON.filter((t) => {
      if (!t.contextual) return true;
      if (t.contextual === 'master') return editor.masterView;
      if (t.contextual === 'media') return media;
      if (t.contextual === 'chart') return chart;
      if (t.contextual === 'shape') return sel.kind === 'shape' && !table;
      if (t.contextual === 'cell' || t.contextual === 'table') return table;
      return false;
    }).map(tab => tab.id !== 'shape' ? tab : kind === 'video' ? { ...tab, title: 'Video Format' } : chart ? { ...tab, title: 'Format' } : tab);
  });

  // Entering Slide Master view opens its tab, as PowerPoint does.
  $effect(() => {
    if (editor.masterView) activeTab = 'slideMaster';
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

<svelte:window onmessage={onHostMessage} onpointerdown={(event) => { if (shareOpen && !(event.target as Element).closest?.('.share-anchor')) shareOpen = false; }} onkeydown={(event) => { if (shareOpen && event.key === 'Escape') shareOpen = false; }} />

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

    <!-- Mac PowerPoint ends the tab row with Comments and Share; task-pane
         toggles such as Copilot sit beside them, and so does Agents. -->
    <div class="actions">
      {#if agentsOpen !== null}
        <button bind:this={agentsButton} class="agents" aria-pressed={agentsOpen} title={t('Agents')} onclick={toggleAgents}><svg width="16" height="16" viewBox="0 0 16 16" aria-hidden="true"><path d="M8 1.5 9.4 6.6 14.5 8 9.4 9.4 8 14.5 6.6 9.4 1.5 8 6.6 6.6Z" /></svg><span>{t('Agents')}</span></button>
      {/if}
      <button class="comments" aria-label={t('Comments')} aria-pressed={editor.activeDialog === 'addSlideComment'} disabled={!editor.doc.currentSlide} onclick={() => { if (editor.activeDialog === 'addSlideComment') editor.activeDialog = null; else editor.runOrPrompt('addSlideComment'); }}><Icon name="comment" size={16} /><span>{t('Comments')}</span></button>
      <div class="share-anchor">
        <button class="share" aria-label={t('Share')} aria-haspopup="menu" aria-expanded={shareOpen} onclick={() => (shareOpen = !shareOpen)}><Icon name="share" size={16} /><span>{t('Share')}</span><span aria-hidden="true">⌄</span></button>
        {#if shareOpen}
          <div class="share-menu" role="menu" aria-label={t('Share')}>
            <button role="menuitem" title={t('Sharing with people needs OneDrive or SharePoint.')} disabled>{t('Share with People...')}</button>
            <button role="menuitem" title={t('Sharing with people needs OneDrive or SharePoint.')} disabled>{t('Copy Link')}</button>
            <hr />
            <button role="menuitem" onclick={sendCopy}>{t('Send a Copy (PowerPoint Presentation)')}</button>
          </div>
        {/if}
      </div>
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
    {#if current?.id === 'review'}<ReviewRibbon />{/if}
    {#if current?.id === 'transitions'}<TransitionsRibbon />{/if}
    {#if current?.id === 'animations'}<AnimationsRibbon />{/if}
    {#if current?.id === 'home'}<HomeRibbon />{/if}
    {#if current?.id === 'insert'}<InsertRibbon />{/if}
    {#if current?.id === 'design'}<DesignRibbon />{/if}
    {#if current?.id === 'shape'}<ShapeFormatRibbon />{/if}
    {#each current?.groups ?? [] as group (group.title)}
      <div class="group" role="group" aria-label={t(group.title)}>
        <div class="group-items">
          {#each group.items as item (item.id + (item.label ?? ''))}
            {@const cap = capabilityById.get(item.id)}
            <button
              class="cmd"
              disabled={item.unavailable ? true : item.enabled ? !item.enabled(editor) : !editor.canRun(item.id)}
              title={item.unavailable ? t(item.unavailable) : cap ? tip(item.id) : item.label ? t(item.label) : item.id}
              aria-label={item.label ? t(item.label) : cap ? capLabel(cap) : item.id}
              onclick={(event) => (item.run ? item.run(editor, event.currentTarget) : editor.runOrPrompt(item.id, item.preset ?? {}))}
            >
              <span class="icon"><Icon name={item.icon ?? 'dot'} size={32} /></span>
              <span class="cmd-label">{item.compactLabel ? t(item.compactLabel) : item.label ? t(item.label) : cap ? capLabel(cap) : item.id}</span>
            </button>
          {/each}
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
  .actions { display: flex; align-items: center; gap: 6px; flex: none; margin-left: 8px; }
  .actions button { display: flex; align-items: center; gap: 5px; height: 26px; padding: 0 10px; font: inherit; font-size: 12px; color: var(--ok-text); border: 1px solid var(--ok-border-strong); border-radius: 6px; background: var(--ok-panel); cursor: pointer; }
  .actions button:hover:not(:disabled) { background: var(--ok-hover); }
  .actions button:disabled { opacity: 0.4; cursor: default; }
  .actions .agents svg { fill: none; stroke: currentColor; stroke-width: 1.2; stroke-linejoin: round; }
  .actions .agents[aria-pressed='true'] { background: var(--ok-selected); border-color: var(--ok-selected-border); }
  .actions .agents[aria-pressed='true'] svg { fill: var(--ok-accent); stroke: var(--ok-accent); }
  .actions .comments[aria-pressed='true'] { background: var(--ok-selected); border-color: var(--ok-selected-border); }
  /* PowerPoint's Share button is the one filled accent control in the window. */
  .actions .share { color: #fff; border-color: var(--ok-accent); background: var(--ok-accent); }
  .actions .share:hover:not(:disabled) { background: var(--ok-accent); filter: brightness(1.08); }
  .share-anchor { position: relative; }
  .share-menu { position: absolute; top: calc(100% + 4px); right: 0; z-index: 500; display: flex; flex-direction: column; min-width: 260px; padding: 4px; border: 1px solid var(--ok-border); border-radius: 6px; background: var(--ok-panel); box-shadow: var(--ok-shadow-lg); }
  .share-menu button { height: auto; padding: 6px 10px; border: none; background: none; color: var(--ok-text); text-align: left; filter: none; }
  .share-menu hr { width: 100%; border: none; border-top: 1px solid var(--ok-border); margin: 4px 0; }
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
