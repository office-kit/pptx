<script lang="ts">
  import { eventTarget } from '../core/dom-root.ts';
  import MenuBar from './MenuBar.svelte';
  import { getEditor } from '../core/context.ts';
  import { downloadPptx } from '../core/download.ts';
  import { t, getLocale, setLocale, LOCALES, type Locale } from '../i18n/i18n.svelte.ts';
  import type { Attachment } from 'svelte/attachments';
  import ProposalBar from './ProposalBar.svelte';

  let { onsave, autoSave = false, compact = false, status }: { onsave?: () => Promise<void>; autoSave?: boolean; compact?: boolean; status?: HTMLElement } = $props();
  const editor = getEditor();
  const doc = editor.doc;
  let fileInput = $state<HTMLInputElement>();
  let moreOpen = $state(false);

  async function onOpen(e: Event) {
    const input = e.currentTarget as HTMLInputElement;
    const file = input.files?.[0];
    if (!file) return;
    try {
      const buf = new Uint8Array(await file.arrayBuffer());
      await doc.loadBytes(buf, file.name);
      if (onsave) doc.dirty = true;
      editor.toast('info', `${t('Opened')} ${file.name}`);
    } catch (err) {
      editor.toast('error', `${t('Open failed')}: ${(err as Error).message}`);
    }
    input.value = '';
  }

  function newPresentation() {
    doc.resetBlank();
    if (onsave) doc.dirty = true;
  }

  // The host's element stays in the host's DOM: inside a shadow root it is
  // slotted, so the page's own styles still reach it.
  const STATUS_SLOT = 'office-kit-status';
  function hostStatus(element: HTMLElement): Attachment<HTMLElement> {
    return (container) => {
      const root = container.getRootNode();
      if (root instanceof ShadowRoot) {
        const slot = document.createElement('slot');
        slot.name = STATUS_SLOT;
        container.append(slot);
        element.slot = STATUS_SLOT;
        root.host.append(element);
      } else {
        container.append(element);
      }
      return () => element.remove();
    };
  }

  async function onSave() {
    try {
      const version = await downloadPptx(doc);
      if (!onsave) doc.markSaved(version);
      editor.toast('info', t('Saved .pptx'));
    } catch (err) {
      editor.toast('error', `${t('Save failed')}: ${(err as Error).message}`);
    }
  }
</script>

<svelte:window onpointerdown={(event) => { if (moreOpen && !(eventTarget(event) as Element).closest?.('.more-anchor')) moreOpen = false; }} onkeydown={(event) => { if (moreOpen && event.key === 'Escape') moreOpen = false; }} />

<div class="topbar" class:compact>
  <div class="brand">
    <span class="mark">◈</span>
    <span class="name">@office-kit/pptx</span>
    <span class="tag">{t('Editor')}</span>
  </div>

  <MenuBar file={{ save: () => void (onsave ?? onSave)(), download: () => void onSave(), open: () => fileInput?.click(), newPresentation }} />

  <div class="quick">
    {#if autoSave}
      <!-- Mac PowerPoint's AutoSave switch leads the title bar. -->
      <label class="autosave"><span>{t('AutoSave')}</span><input type="checkbox" role="switch" bind:checked={editor.autoSave} /></label>
    {/if}
    <!-- Mac PowerPoint's Quick Access Toolbar: icons, with the rest under ⋯. -->
    <button class="ok-btn qat" title={t('Save as .pptx')} aria-label={t('Save')} onclick={onsave ?? onSave}><svg viewBox="0 0 16 16" aria-hidden="true"><path d="M2.5 2.5h9l2 2v9h-11z M5 2.5v4h6v-4 M5 13.5v-4h6v4" /></svg></button>
    <button class="ok-btn qat" title={t('Undo (Ctrl+Z)')} aria-label={t('Undo')} disabled={!doc.canUndo} onclick={() => doc.undo()}><svg viewBox="0 0 16 16" aria-hidden="true"><path d="M5.5 3 2.5 6l3 3 M2.5 6H10a3.5 3.5 0 0 1 0 7H7" /></svg></button>
    <button class="ok-btn qat" title={t('Redo (Ctrl+Y)')} aria-label={t('Redo')} disabled={!doc.canRedo} onclick={() => doc.redo()}><svg viewBox="0 0 16 16" aria-hidden="true"><path d="M10.5 3l3 3-3 3 M13.5 6H6a3.5 3.5 0 0 0 0 7h3" /></svg></button>
    <div class="more-anchor">
      <button class="ok-btn qat" title={t('More Commands')} aria-label={t('More Commands')} aria-haspopup="menu" aria-expanded={moreOpen} onclick={() => (moreOpen = !moreOpen)}>⋯</button>
      {#if moreOpen}
        <div class="more-menu" role="menu" aria-label={t('More Commands')}>
          <button role="menuitem" class="ok-btn" title={t('New')} onclick={() => { moreOpen = false; newPresentation(); }}>{t('New')}</button>
          <button role="menuitem" class="ok-btn" title={t('Open .pptx')} onclick={() => { moreOpen = false; fileInput?.click(); }}>{t('Open')}</button>
          {#if onsave}<button role="menuitem" class="ok-btn" onclick={() => { moreOpen = false; void onSave(); }}>{t('Download')}</button>{/if}
        </div>
      {/if}
    </div>
  </div>

  <div class="filename">
    {doc.fileName}{#if doc.dirty}<span class="dot" title={t('Unsaved changes')}> ●</span>{/if}
  </div>

  <div class="right">
    <label class="lang" title={t('Language')}>
      <select value={getLocale()} onchange={(e) => setLocale((e.currentTarget as HTMLSelectElement).value as Locale)}>
        {#each LOCALES as l (l.id)}
          <option value={l.id}>{l.label}</option>
        {/each}
      </select>
    </label>
    <!-- Sits where Mac PowerPoint's Search box does; it searches every command. -->
    <button class="ok-btn palette-btn" onclick={() => editor.togglePalette(true)} title={t('Search every command (⌘?)')}>
      <svg width="13" height="13" viewBox="0 0 14 14" aria-hidden="true"><circle cx="6" cy="6" r="4.5" /><path d="M9.5 9.5L13 13" /></svg>
      {t('Search (⌘?)')}
    </button>
  </div>

  {#if status}<div class="host-status" {@attach hostStatus(status)}></div>{/if}
  {#if editor.proposal}<ProposalBar proposal={editor.proposal} />{/if}

  <input
    bind:this={fileInput}
    type="file"
    accept=".pptx,application/vnd.openxmlformats-officedocument.presentationml.presentation"
    style="display:none"
    onchange={onOpen}
  />
</div>

<style>
  .topbar {
    display: flex;
    align-items: center;
    gap: 16px;
    min-height: 40px;
    flex-wrap: wrap;
    padding: 5px 12px;
    /* Mac PowerPoint's title bar uses the window chrome color, not the accent. */
    background: var(--ok-ribbon);
    color: var(--ok-text);
  }
  .topbar.compact {
    flex-wrap: nowrap;
    gap: 8px;
    min-height: 32px;
    padding: 2px 8px;
    overflow-x: auto;
  }
  .topbar:has(:global(.conflict)) { flex-wrap: wrap; overflow-x: hidden; }
  .topbar.compact .tag { display: none; }
  .topbar.compact .brand { display: none; }
  .topbar.compact .quick { flex: none; }
  .topbar.compact .quick :global(.ok-btn),
  .topbar.compact .right :global(.palette-btn) { font-size: 11px; padding: 3px 6px; }
  .topbar.compact .right { flex: none; }
  .brand {
    display: flex;
    align-items: baseline;
    gap: 8px;
  }
  .mark {
    font-size: 16px;
  }
  .name {
    font-weight: 600;
    font-size: 13px;
  }
  .tag {
    font-size: 11px;
    opacity: 0.85;
  }
  .quick {
    display: flex;
    align-items: center;
    gap: 2px;
  }
  .autosave { display: flex; align-items: center; gap: 6px; margin-right: 6px; font-size: 12px; white-space: nowrap; cursor: pointer; }
  .autosave input { appearance: none; position: relative; width: 28px; height: 16px; margin: 0; border-radius: 8px; background: var(--ok-border-strong); cursor: pointer; transition: background 0.15s; }
  .autosave input::after { content: ''; position: absolute; top: 2px; left: 2px; width: 12px; height: 12px; border-radius: 50%; background: #fff; transition: transform 0.15s; }
  .autosave input:checked { background: var(--ok-accent); }
  .autosave input:checked::after { transform: translateX(12px); }
  .autosave input:focus-visible { outline: 2px solid var(--ok-accent); outline-offset: 2px; }
  .qat { padding: 3px 5px; }
  .qat svg { width: 16px; height: 16px; fill: none; stroke: currentColor; stroke-width: 1.3; stroke-linecap: round; stroke-linejoin: round; }
  .more-anchor { position: relative; }
  .more-menu { position: absolute; top: 100%; left: 0; z-index: 500; display: flex; flex-direction: column; min-width: 160px; padding: 4px; border: 1px solid var(--ok-border); border-radius: 6px; background: var(--ok-panel); box-shadow: var(--ok-shadow-lg); }
  .more-menu .ok-btn { justify-content: flex-start; }
  .quick :global(.ok-btn:hover) {
    border-color: transparent;
  }
  .filename {
    flex: 1;
    min-width: 60px;
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
    text-align: center;
    font-size: 13px;
    font-weight: 600;
  }
  .dot {
    color: var(--ok-text-3);
  }
  .right {
    margin-left: auto;
    display: flex;
    align-items: center;
    gap: 8px;
  }
  .lang select {
    background: var(--ok-panel);
    color: var(--ok-text);
    border: 1px solid var(--ok-border-strong);
    border-radius: var(--ok-radius);
    font: inherit;
    font-size: 12px;
    padding: 3px 6px;
    cursor: pointer;
  }
  .lang select option {
    color: initial;
  }
  .right :global(.palette-btn) {
    min-width: 180px;
    color: var(--ok-text-2);
    background: var(--ok-panel);
    border-color: var(--ok-border-strong);
    border-radius: 6px;
    font-size: 12px;
  }
  .right :global(.palette-btn svg) {
    fill: none;
    stroke: currentColor;
    stroke-width: 1.4;
  }
  .host-status {
    display: flex;
    align-items: center;
    min-width: 0;
  }
  .topbar.compact .host-status {
    display: contents;
  }
  .topbar.compact .right :global(.palette-btn) {
    min-width: 0;
  }
</style>
