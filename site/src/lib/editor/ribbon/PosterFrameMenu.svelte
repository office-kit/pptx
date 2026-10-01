<script lang="ts">
  import { tick } from 'svelte';
  import { getShapeMedia, setShapeImage } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { getMediaPreview } from '../core/media-preview.svelte.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import Icon from '../ui/Icon.svelte';
  const editor = getEditor();
  const doc = editor.doc;
  const preview = getMediaPreview(editor);
  let open = $state(false);
  let busy = $state(false);
  let trigger: HTMLButtonElement;
  let menu = $state<HTMLDivElement>();
  function close(restore = true) { open = false; if (restore) trigger.focus(); }
  async function show() {
    if (open) { close(); return; }
    open = true; await tick();
    menu?.querySelector<HTMLButtonElement>('button:not(:disabled)')?.focus();
  }
  async function currentFrame() {
    close();
    const shape = editor.selectedShapes()[0];
    if (!shape || getShapeMedia(shape)?.kind !== 'video') return;
    const version = doc.version;
    const selection = doc.selection;
    busy = true;
    try {
      const bytes = await preview.captureFrame();
      if (doc.version !== version || doc.selection !== selection) return;
      doc.transact(t('Poster Frame'), () => setShapeImage(shape, bytes));
      preview.showPoster();
    } catch (cause) {
      editor.toast('error', cause instanceof Error ? cause.message : String(cause));
    } finally { busy = false; }
  }
  function fromFile() { close(); preview.showPoster(); editor.runOrPrompt('setShapeImage'); }
  function place(node: HTMLElement) {
    const bounds = trigger.getBoundingClientRect();
    const margin = 8;
    node.style.left = `${Math.max(margin, Math.min(bounds.left, innerWidth - node.offsetWidth - margin))}px`;
    node.style.top = `${Math.min(bounds.bottom, innerHeight - node.offsetHeight - margin)}px`;
  }
  function keys(event: KeyboardEvent) {
    event.stopPropagation();
    if (event.key === 'Escape') { event.preventDefault(); close(); return; }
    if (event.key === 'Tab') { close(false); return; }
    const items = [...menu!.querySelectorAll<HTMLButtonElement>('button:not(:disabled)')];
    const index = items.indexOf(event.target as HTMLButtonElement);
    if (!['ArrowUp', 'ArrowDown', 'Home', 'End'].includes(event.key)) return;
    event.preventDefault();
    items[event.key === 'Home' ? 0 : event.key === 'End' ? items.length - 1 : (index + (event.key === 'ArrowUp' ? -1 : 1) + items.length) % items.length]?.focus();
  }
</script>
<svelte:window onpointerdown={event => { if (open && !menu?.contains(event.target as Node) && !trigger.contains(event.target as Node)) close(false); }} onblur={() => { if (open) close(false); }} onresize={() => { if (open) close(false); }} />
<button class="trigger" bind:this={trigger} disabled={busy} aria-label={t('Poster Frame')} aria-haspopup="menu" aria-expanded={open} onclick={show}><Icon name="image" /><span>{t('Poster Frame')} ▾</span></button>
{#if open}
  <div class="menu" role="menu" aria-label={t('Poster Frame')} tabindex="-1" bind:this={menu} use:place onkeydown={keys}>
    <button role="menuitem" disabled={!preview.state.showVideoFrame} onclick={currentFrame}>{t('Current Frame')}</button>
    <button role="menuitem" onclick={fromFile}>{t('Image from File...')}</button>
  </div>
{/if}
<style>
  .trigger { display:flex; flex-direction:column; align-items:center; justify-content:center; gap:3px; padding:4px 8px; background:transparent; border:1px solid transparent; border-radius:var(--ok-radius); color:var(--ok-text); font:inherit; font-size:11px; cursor:pointer; }
  .trigger:hover { background:var(--ok-hover); border-color:var(--ok-border); }
  .trigger:disabled { opacity:.4; }
  .menu { position:fixed; z-index:400; padding:5px; border:1px solid var(--ok-border); border-radius:6px; background:var(--ok-panel); color:var(--ok-text); box-shadow:var(--ok-shadow-lg); }
  .menu button { display:block; width:100%; border:0; border-radius:4px; padding:6px 12px; background:transparent; color:inherit; font:inherit; text-align:left; }
  .menu button:disabled { opacity:.4; }
  .menu button:hover:not(:disabled), .menu button:focus-visible { background:var(--ok-accent); color:white; }
</style>
