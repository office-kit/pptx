<script lang="ts">
  import { tick } from 'svelte';
  import { menuItemForKey } from '../core/menubar-shortcuts.ts';
  import { nativeMenus } from '../ui/menubar-commands.ts';
  import { ALIGN_ITEMS, ROTATE_ITEMS, rotateSelection } from './arrange-actions.ts';
  import { getEditor } from '../core/context.ts';
  import { selectedShapeIds } from '../core/selection.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import Icon from '../ui/Icon.svelte';
  let { compact = false, onchoose }: { compact?: boolean; onchoose?: () => void } = $props();
  const editor = getEditor();
  const doc = editor.doc;
  let open = $state(false);
  let branch = $state<'align' | 'rotate' | null>(null);
  let trigger: HTMLButtonElement;
  let menu = $state<HTMLDivElement>();
  const count = $derived(selectedShapeIds(doc.selection).length);
  const locked = $derived(editor.selectionLocked());
  const toSlide = $derived(count < 2 || editor.alignmentReference === 'slide');
  const order = [{ id: 'bringShapeToFront', label: 'Bring to Front' }, { id: 'sendShapeToBack', label: 'Send to Back' }, { id: 'bringShapeForward', label: 'Bring Forward' }, { id: 'sendShapeBackward', label: 'Send Backward' }];
  function close(restore = true) { open = false; branch = null; if (restore) trigger.focus(); }
  function choose(action: () => void) { close(); onchoose?.(); action(); }
  async function show() { open = !open; branch = null; if (open) { await tick(); menu?.querySelector<HTMLButtonElement>('button:not(:disabled)')?.focus(); } }
  function place(node: HTMLElement, nested = false) {
    function position() {
      const anchor = nested ? node.parentElement!.getBoundingClientRect() : trigger.getBoundingClientRect();
      const margin = 8;
      let x = nested ? anchor.right : anchor.left;
      if (nested && x + node.offsetWidth > innerWidth - margin) x = anchor.left - node.offsetWidth;
      node.style.left = `${Math.max(margin, Math.min(x, innerWidth - node.offsetWidth - margin))}px`;
      node.style.top = `${Math.max(margin, Math.min(nested ? anchor.top : anchor.bottom, innerHeight - node.offsetHeight - margin))}px`;
    }
    position();
    window.addEventListener('resize', position);
    return { destroy() { window.removeEventListener('resize', position); } };
  }
  function keys(event: KeyboardEvent) {
    if (!event.isComposing && menuItemForKey(nativeMenus(), event)) { close(false); return; }
    event.stopPropagation();
    if (event.key === 'Escape') { event.preventDefault(); close(); return; }
    if (event.key === 'Tab') { close(false); return; }
    const target = event.target as HTMLElement;
    if (event.key === 'ArrowRight' && target.dataset.branch) {
      event.preventDefault(); branch = target.dataset.branch as 'align' | 'rotate';
      void tick().then(() => menu?.querySelector<HTMLButtonElement>('.submenu button:not(:disabled)')?.focus());
    } else if (event.key === 'ArrowLeft' && target.closest('.submenu')) {
      event.preventDefault(); const previous = branch; branch = null;
      menu?.querySelector<HTMLButtonElement>(`[data-branch="${previous}"]`)?.focus();
    } else if (['ArrowUp', 'ArrowDown', 'Home', 'End'].includes(event.key)) {
      event.preventDefault();
      const items = [...(target.closest('[role="menu"]')?.querySelectorAll<HTMLButtonElement>(':scope > button:not(:disabled), :scope > .branch > button:not(:disabled)') ?? [])];
      const index = items.indexOf(target as HTMLButtonElement);
      items[event.key === 'Home' ? 0 : event.key === 'End' ? items.length - 1 : (index + (event.key === 'ArrowUp' ? -1 : 1) + items.length) % items.length]?.focus();
    }
  }
</script>
<svelte:window onpointerdown={event => { if (open && !menu?.contains(event.target as Node) && !trigger.contains(event.target as Node)) close(false); }} onblur={() => { if (open) close(false); }} />
<button class="trigger" class:compact bind:this={trigger} aria-label={t('Arrange')} aria-haspopup="menu" aria-expanded={open} onclick={show}>{#if compact}<span>{t('Arrange')} ▾</span>{:else}<span class="icon-row"><Icon name="arrange" size={32} /><span aria-hidden="true">▾</span></span><span>{t('Arrange')}</span>{/if}</button>
{#if open}
  <div class="menu" role="menu" aria-label={t('Arrange')} tabindex="-1" bind:this={menu} use:place onkeydown={keys}>
    <div class="heading">{t('Reorder Objects')}</div>
    <button role="menuitem" disabled={editor.reorderMembers().length < 2} onclick={() => choose(() => editor.activeDialog = 'reorderObjects')}>{t('Reorder Overlapping Objects')}</button>
    <hr /><div class="heading">{t('Reorder Objects')}</div>
    {#each order as item}<button role="menuitem" disabled={!editor.canRun(item.id)} onclick={() => choose(() => editor.invoke(item.id))}>{t(item.label)}</button>{/each}
    <hr /><div class="heading">{t('Group Objects')}</div>
    {#each [{ id: 'groupShapes', label: 'Group' }, { id: 'ungroupShapes', label: 'Ungroup' }] as item}<button role="menuitem" disabled={!editor.canRun(item.id)} onclick={() => choose(() => editor.invoke(item.id))}>{t(item.label)}</button>{/each}
    <button role="menuitem" disabled={!editor.canRegroup()} onclick={() => choose(() => editor.regroupSelection())}>{t('Regroup')}</button>
    <hr /><div class="heading">{t('Position Objects')}</div>
    <div class="branch">
      <button role="menuitem" data-branch="align" aria-label={t('Align')} aria-haspopup="menu" aria-expanded={branch === 'align'} disabled={!count || locked} onpointerenter={() => branch = 'align'} onclick={() => branch = 'align'}>{t('Align')}<span>›</span></button>
      {#if branch === 'align' && count}<div class="menu submenu" role="menu" aria-label={t('Align')} use:place={true}>
        {#each ALIGN_ITEMS as item}<button role="menuitem" onclick={() => choose(() => editor.alignSelection(item.value, toSlide ? 'slide' : 'selection'))}>{t(item.label)}</button>{/each}
        <hr />
        <button role="menuitem" disabled={count === 0 || (!toSlide && count < 3)} onclick={() => choose(() => editor.distributeSelection('horizontal'))}>{t('Distribute Horizontally')}</button>
        <button role="menuitem" disabled={count === 0 || (!toSlide && count < 3)} onclick={() => choose(() => editor.distributeSelection('vertical'))}>{t('Distribute Vertically')}</button>
        <hr />
        <button role="menuitemradio" aria-checked={toSlide} onclick={() => choose(() => editor.alignmentReference = 'slide')}>{t('Align to Slide')}<span>{toSlide ? '✓' : ''}</span></button>
        <button role="menuitemradio" disabled={count < 2} aria-checked={!toSlide} onclick={() => choose(() => editor.alignmentReference = 'selection')}>{t('Align Selected Objects')}<span>{!toSlide ? '✓' : ''}</span></button>
      </div>{/if}
    </div>
    <div class="branch">
      <button role="menuitem" data-branch="rotate" aria-label={t('Rotate')} aria-haspopup="menu" aria-expanded={branch === 'rotate'} disabled={!count || locked} onpointerenter={() => branch = 'rotate'} onclick={() => branch = 'rotate'}>{t('Rotate')}<span>›</span></button>
      {#if branch === 'rotate' && count}<div class="menu submenu" role="menu" aria-label={t('Rotate')} use:place={true}>
        {#each ROTATE_ITEMS as item}<button role="menuitem" onclick={() => choose(() => rotateSelection(editor, item.action))}>{t(item.label)}</button>{/each}
        <hr /><button role="menuitem" disabled={!editor.canRun('setShapeRotation')} onclick={() => choose(() => editor.showRotationOptions())}>{t('More Rotation Options...')}</button>
      </div>{/if}
    </div>
    <hr /><button role="menuitemcheckbox" aria-label={t('Selection Pane...')} aria-checked={editor.selectionPaneVisible} onclick={() => choose(() => { editor.setViewMode('normal'); editor.selectionPaneVisible = !editor.selectionPaneVisible; })}>{t('Selection Pane...')}<span>{editor.selectionPaneVisible ? '✓' : ''}</span></button>
  </div>
{/if}
<style>
  .trigger { display: flex; flex-direction: column; align-items: center; gap: 3px; padding: 4px; background: transparent; border: 1px solid transparent; border-radius: var(--ok-radius); color: var(--ok-text); font: inherit; font-size: 11px; cursor: pointer; }
  .trigger:not(.compact) { align-self: stretch; justify-content: flex-start; gap: 2px; min-width: 50px; padding: 4px 2px; }
  .icon-row { display: flex; align-items: center; gap: 1px; font-size: 10px; }
  .trigger:hover { background: var(--ok-hover); border-color: var(--ok-border); }
  .trigger.compact { flex-direction: row; justify-content: center; min-width: 76px; padding: 8px 7px; border-color: var(--ok-border); background: var(--ok-panel); }
  .trigger.compact:hover, .trigger.compact[aria-expanded='true'] { background: var(--ok-hover); border-color: var(--ok-accent); }
  .menu { position: fixed; z-index: 400; min-width: 210px; max-height: calc(100dvh - 16px); overflow-y: auto; padding: 5px; border: 1px solid var(--ok-border); border-radius: 6px; background: var(--ok-panel); color: var(--ok-text); box-shadow: var(--ok-shadow-lg); }
  .menu button { display: flex; justify-content: space-between; gap: 16px; width: 100%; border: 0; border-radius: 4px; padding: 5px 10px; background: transparent; color: inherit; font: inherit; font-size: 12px; text-align: left; white-space: nowrap; }
  .menu button:hover:not(:disabled), .menu button:focus-visible { background: var(--ok-accent); color: white; outline: none; }
  .menu button:disabled { opacity: .4; }
  .heading { padding: 5px 10px; font-size: 11px; color: var(--ok-text-3); }
  hr { border: 0; border-top: 1px solid var(--ok-border); margin: 5px; }
</style>
