<script lang="ts">
  import { getEditor } from '../core/context.ts';
  import { t, getLocale, setLocale, LOCALES, type Locale } from '../i18n/i18n.svelte.ts';

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
      editor.toast('info', `${t('Opened')} ${file.name}`);
    } catch (err) {
      editor.toast('error', `${t('Open failed')}: ${(err as Error).message}`);
    }
    input.value = '';
  }

  async function onSave() {
    try {
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
      doc.dirty = false;
      editor.toast('info', t('Saved .pptx'));
    } catch (err) {
      editor.toast('error', `${t('Save failed')}: ${(err as Error).message}`);
    }
  }
</script>

<div class="topbar">
  <div class="brand">
    <span class="mark">◈</span>
    <span class="name">@office-kit/pptx</span>
    <span class="tag">{t('Editor')}</span>
  </div>

  <div class="quick">
    <button class="ok-btn" title={t('New')} onclick={() => doc.resetBlank()}>{t('New')}</button>
    <button class="ok-btn" title={t('Open .pptx')} onclick={() => fileInput?.click()}>{t('Open')}</button>
    <button class="ok-btn" title={t('Save as .pptx')} onclick={onSave}>{t('Save')}</button>
    <span class="sep"></span>
    <button class="ok-btn" title={t('Undo (Ctrl+Z)')} disabled={!doc.canUndo} onclick={() => doc.undo()}>↶</button>
    <button class="ok-btn" title={t('Redo (Ctrl+Y)')} disabled={!doc.canRedo} onclick={() => doc.redo()}>↷</button>
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
    <button class="ok-btn palette-btn" onclick={() => editor.togglePalette(true)} title={t('Command palette (Ctrl+K)')}>
      ⌘K · {t('All capabilities')}
    </button>
  </div>

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
    height: 40px;
    padding: 0 12px;
    background: var(--ok-accent);
    color: #fff;
  }
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
    text-align: center;
    font-size: 12px;
    opacity: 0.95;
  }
  .dot {
    color: #ffd7c9;
  }
  .right {
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
</style>
