<script lang="ts">
  import { tick } from 'svelte';
  import { getDrawingGuidesVisible, getSnapToGrid, setSnapToGrid } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  const editor = getEditor();
  const doc = editor.doc;
  let open = $state(false);
  let submenu = $state<'grid' | 'zoom' | null>(null);
  let root: HTMLDivElement;
  let trigger: HTMLButtonElement;
  let fullScreen = $state(false);
  const drawing = $derived.by(() => { doc.version; return editor.view.drawing ?? getDrawingGuidesVisible(doc.pres) ?? false; });
  const snapping = $derived.by(() => { doc.version; return getSnapToGrid(doc.pres) ?? false; });
  function close(restore = true) { open = false; submenu = null; if (restore) trigger?.focus(); }
  function choose(action: () => void) { close(); action(); }
  function toggle(key: 'smart' | 'drawing' | 'grid') {
    editor.view.save({ grid: editor.view.grid, smart: editor.view.smart, drawing, [key]: !(key === 'drawing' ? drawing : editor.view[key]) });
  }
  async function focusFirst() { await tick(); root?.querySelector<HTMLButtonElement>('[role="menu"] button')?.focus(); }
  function keys(event: KeyboardEvent) {
    if (!open) return;
    if (event.key === 'Escape') { event.preventDefault(); event.stopPropagation(); close(); return; }
    const target = event.target as HTMLElement;
    if (event.key === 'ArrowRight' && target.dataset.submenu) {
      event.preventDefault(); submenu = target.dataset.submenu as 'grid' | 'zoom';
      void tick().then(() => root.querySelector<HTMLButtonElement>('.submenu button')?.focus());
    } else if (event.key === 'ArrowLeft' && target.closest('.submenu')) {
      event.preventDefault(); const previous = submenu; submenu = null;
      root.querySelector<HTMLButtonElement>(`[data-submenu="${previous}"]`)?.focus();
    } else if (event.key === 'ArrowDown' || event.key === 'ArrowUp' || event.key === 'Home' || event.key === 'End') {
      event.preventDefault();
      const menu = target.closest('[role="menu"]') ?? root.querySelector('[role="menu"]');
      const buttons = [...(menu?.querySelectorAll<HTMLButtonElement>(':scope > button:not(:disabled), :scope > .branch > button') ?? [])];
      const index = buttons.indexOf(target as HTMLButtonElement);
      buttons[event.key === 'Home' ? 0 : event.key === 'End' ? buttons.length - 1 : (index + (event.key === 'ArrowUp' ? -1 : 1) + buttons.length) % buttons.length]?.focus();
    } else if (event.key === 'Tab') close(false);
  }
  async function toggleFullScreen() {
    try { if (document.fullscreenElement) await document.exitFullscreen(); else await document.documentElement.requestFullscreen(); }
    catch (error) { editor.toast('error', String(error)); }
  }
</script>
<svelte:window onpointerdown={event => { if (open && !root.contains(event.target as Node)) close(false); }} onkeydown={keys} onfullscreenchange={() => fullScreen = !!document.fullscreenElement} />
<div class="view-menu" bind:this={root}>
  <button class="ok-btn" bind:this={trigger} aria-haspopup="menu" aria-expanded={open} onclick={() => { open = !open; submenu = null; if (open) void focusFirst(); }}>{t('View')}</button>
  {#if open}
    <div class="menu" role="menu" aria-label={t('View')}>
      {#each [{ mode: 'normal' as const, label: 'Normal', key: '⌘1' }, { mode: 'sorter' as const, label: 'Slide Sorter', key: '⌘2' }] as item}
        <button role="menuitemradio" aria-label={t(item.label)} aria-checked={editor.viewMode === item.mode} onclick={() => choose(() => editor.setViewMode(item.mode))}><span>{editor.viewMode === item.mode ? '✓' : ''}</span>{t(item.label)}<kbd>{item.key}</kbd></button>
      {/each}
      <hr />
      <button role="menuitemcheckbox" aria-checked={editor.ribbonVisible} onclick={() => choose(() => editor.ribbonVisible = !editor.ribbonVisible)}><span>{editor.ribbonVisible ? '✓' : ''}</span>{t('Ribbon')}<kbd>⌥⌘R</kbd></button>
      <div class="branch">
        <button role="menuitem" aria-haspopup="menu" aria-expanded={submenu === 'grid'} aria-label={t('Grid and Guides')} data-submenu="grid" onclick={() => submenu = 'grid'} onpointerenter={() => submenu = 'grid'}><span></span>{t('Grid and Guides')}<kbd>›</kbd></button>
        {#if submenu === 'grid'}
          <div class="menu submenu" role="menu" aria-label={t('Grid and Guides')}>
            {#each [{ key: 'smart' as const, label: 'Smart Guides', value: editor.view.smart }, { key: 'drawing' as const, label: 'Guides', value: drawing }, { key: 'grid' as const, label: 'Gridlines', value: editor.view.grid }] as item}
              <button role="menuitemcheckbox" aria-checked={item.value} onclick={() => choose(() => toggle(item.key))}><span>{item.value ? '✓' : ''}</span>{t(item.label)}</button>
            {/each}
            <hr />
            <button role="menuitemcheckbox" aria-checked={snapping} onclick={() => choose(() => doc.transact(t('Snap to Grid'), () => setSnapToGrid(doc.pres, !snapping)))}><span>{snapping ? '✓' : ''}</span>{t('Snap to Grid')}</button>
            <hr />
            <button role="menuitem" onclick={() => choose(() => editor.activeDialog = 'gridOptions')}><span></span>{t('Grid Options...')}</button>
          </div>
        {/if}
      </div>
      <div class="branch">
        <button role="menuitem" aria-haspopup="menu" aria-expanded={submenu === 'zoom'} aria-label={t('Zoom')} data-submenu="zoom" onclick={() => submenu = 'zoom'} onpointerenter={() => submenu = 'zoom'}><span></span>{t('Zoom')}<kbd>›</kbd></button>
        {#if submenu === 'zoom'}<div class="menu submenu" role="menu" aria-label={t('Zoom')}>
          <button role="menuitem" onclick={() => choose(() => editor.zoomFit())}><span></span>{t('Fit to Window')}</button>
          <button role="menuitem" onclick={() => choose(() => editor.zoomIn())}><span></span>{t('Zoom In')}</button>
          <button role="menuitem" onclick={() => choose(() => editor.zoomOut())}><span></span>{t('Zoom Out')}</button>
        </div>{/if}
      </div>
      <hr />
      <button role="menuitem" disabled={!document.fullscreenEnabled} onclick={() => choose(() => { void toggleFullScreen(); })}><span></span>{t(fullScreen ? 'Exit Full Screen' : 'Enter Full Screen')}</button>
    </div>
  {/if}
</div>
<style>
  .view-menu { position: relative; }
  .menu { position: absolute; top: 100%; left: 0; z-index: 120; min-width: 225px; padding: 5px; border: 1px solid var(--ok-border); border-radius: 6px; background: var(--ok-panel); color: var(--ok-text); box-shadow: var(--ok-shadow-lg); }
  .branch { position: relative; }
  .submenu { left: 100%; top: -5px; }
  .menu button { display: flex; align-items: center; gap: 6px; border: 0; border-radius: 4px; background: transparent; color: inherit; width: 100%; padding: 4px 7px; text-align: left; white-space: nowrap; font: inherit; font-size: 13px; }
  .menu button:hover:not(:disabled), .menu button:focus-visible { background: var(--ok-accent); color: white; outline: none; }
  .menu button:disabled { opacity: .45; }
  .menu button span { width: 12px; }
  kbd { margin-left: auto; font: inherit; padding-left: 25px; }
  hr { border: 0; border-top: 1px solid var(--ok-border); margin: 5px 8px; }
</style>
