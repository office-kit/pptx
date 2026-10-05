<script lang="ts">
  import ViewMenu from './ViewMenu.svelte';
  import EditMenu from './EditMenu.svelte';
  import { getEditor } from '../core/context.ts';
  import { t, getLocale, setLocale, LOCALES, type Locale } from '../i18n/i18n.svelte.ts';
  import type { Snippet } from 'svelte';

  let { onsave, compact = false, status }: { onsave?: () => Promise<void>; compact?: boolean; status?: Snippet } = $props();
  const editor = getEditor();
  const doc = editor.doc;
  let fileInput = $state<HTMLInputElement>();

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

  async function onSave() {
    try {
      const version = doc.version;
      const bytes = await doc.toBytes();
      const blob = new Blob([bytes as BlobPart], {
        type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = doc.fileName.endsWith('.pptx') ? doc.fileName : `${doc.fileName}.pptx`;
      a.click();
      URL.revokeObjectURL(url);
      if (!onsave) doc.markSaved(version);
      editor.toast('info', t('Saved .pptx'));
    } catch (err) {
      editor.toast('error', `${t('Save failed')}: ${(err as Error).message}`);
    }
  }
</script>

<div class="topbar" class:compact>
  <div class="brand">
    <span class="mark">◈</span>
    <span class="name">@office-kit/pptx</span>
    <span class="tag">{t('Editor')}</span>
  </div>

  <div class="quick">
    <button class="ok-btn" title={t('New')} onclick={() => { doc.resetBlank(); if (onsave) doc.dirty = true; }}>{t('New')}</button>
    <button class="ok-btn" title={t('Open .pptx')} onclick={() => fileInput?.click()}>{t('Open')}</button>
    <button class="ok-btn" title={t('Save as .pptx')} onclick={onsave ?? onSave}>{t('Save')}</button>
    {#if onsave}<button class="ok-btn" onclick={onSave}>{t('Download')}</button>{/if}
    <span class="sep"></span>
    <button class="ok-btn" title={t('Undo (Ctrl+Z)')} disabled={!doc.canUndo} onclick={() => doc.undo()}>↶</button>
    <button class="ok-btn" title={t('Redo (Ctrl+Y)')} disabled={!doc.canRedo} onclick={() => doc.redo()}>↷</button>
  </div>

  <div class="filename">
    {doc.fileName}{#if doc.dirty}<span class="dot" title={t('Unsaved changes')}> ●</span>{/if}
  </div>

  <div class="right">
    <EditMenu />
    <ViewMenu />
    <label class="lang" title={t('Language')}>
      <select value={getLocale()} onchange={(e) => setLocale((e.currentTarget as HTMLSelectElement).value as Locale)}>
        {#each LOCALES as l (l.id)}
          <option value={l.id}>{l.label}</option>
        {/each}
      </select>
    </label>
    <button class="ok-btn palette-btn" onclick={() => editor.togglePalette(true)} title={t('Command palette (Ctrl+K)')}>
      ⌘K · {t('All capabilities')}
    </button>
  </div>

  {#if status}<div class="host-status">{@render status()}</div>{/if}

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
    background: var(--ok-accent);
    color: #fff;
  }
  .topbar.compact {
    flex-wrap: nowrap;
    gap: 8px;
    min-height: 32px;
    padding: 2px 8px;
    overflow-x: auto;
  }
  .topbar.compact:has(:global(.conflict)) { flex-wrap: wrap; overflow-x: hidden; }
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
  .quick :global(.ok-btn) {
    color: #fff;
  }
  .quick :global(.ok-btn:hover) {
    background: rgba(255, 255, 255, 0.16);
    border-color: transparent;
  }
  .quick :global(.ok-btn:disabled) {
    color: rgba(255, 255, 255, 0.4);
  }
  .sep {
    width: 1px;
    height: 20px;
    background: rgba(255, 255, 255, 0.3);
    margin: 0 4px;
  }
  .filename {
    flex: 1;
    min-width: 60px;
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
    text-align: center;
    font-size: 12px;
    opacity: 0.95;
  }
  .dot {
    color: #ffd7c9;
  }
  .right {
    margin-left: auto;
    display: flex;
    align-items: center;
    gap: 8px;
  }
  .lang select {
    background: rgba(255, 255, 255, 0.12);
    color: #fff;
    border: 1px solid rgba(255, 255, 255, 0.4);
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
    color: #fff;
    border-color: rgba(255, 255, 255, 0.4);
    font-size: 12px;
  }
  .right :global(.palette-btn:hover) {
    background: rgba(255, 255, 255, 0.16);
  }
  .host-status {
    display: flex;
    align-items: center;
    min-width: 0;
  }
  .topbar.compact .host-status {
    display: contents;
  }
  .topbar.compact .host-status :global(.save-status) {
    flex: 0 1 auto;
    min-height: 24px;
  }
  .topbar.compact .host-status :global(.conflict) {
    flex-basis: 100%;
  }
</style>
