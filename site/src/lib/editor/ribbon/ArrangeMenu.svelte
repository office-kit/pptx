<script lang="ts">
  import { tick } from 'svelte';
  import { arrangeShortcut } from '../core/arrange-shortcuts.ts';
  import { getSlideShapes, getShapeId, getShapeRotation, getShapeFlip, setShapeRotation, setShapeFlip } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { selectedShapeIds } from '../core/selection.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import Icon from '../ui/Icon.svelte';
  const editor = getEditor();
  const doc = editor.doc;
  let open = $state(false);
  let branch = $state<'align' | 'rotate' | null>(null);
  let trigger: HTMLButtonElement;
  let menu = $state<HTMLDivElement>();
  const count = $derived(selectedShapeIds(doc.selection).length);
  const toSlide = $derived(count < 2 || editor.alignmentReference === 'slide');
  const order = [{ id: 'bringShapeToFront', label: 'Bring to Front' }, { id: 'sendShapeToBack', label: 'Send to Back' }, { id: 'bringShapeForward', label: 'Bring Forward' }, { id: 'sendShapeBackward', label: 'Send Backward' }];
  const alignment = [{ value: 'left', label: 'Align Left' }, { value: 'center', label: 'Align Center' }, { value: 'right', label: 'Align Right' }, { value: 'top', label: 'Align Top' }, { value: 'middle', label: 'Align Middle' }, { value: 'bottom', label: 'Align Bottom' }] as const;
  function close(restore = true) { open = false; branch = null; if (restore) trigger.focus(); }
  function choose(action: () => void) { close(); action(); }
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
    if (!event.isComposing && arrangeShortcut(event)) { close(false); return; }
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
  function transform(action: 'right' | 'left' | 'horizontal' | 'vertical') {
    if (!doc.currentSlide) return;
    const ids = new Set(selectedShapeIds(doc.selection));
    const shapes = getSlideShapes(doc.currentSlide).filter(shape => ids.has(getShapeId(shape)));
    if (!shapes.length) return;
    doc.transact(t('Rotate'), () => {
      for (const shape of shapes) {
        if (action === 'right' || action === 'left') setShapeRotation(shape, getShapeRotation(shape) + (action === 'right' ? 90 : -90));
        else setShapeFlip(shape, { [action]: !getShapeFlip(shape)?.[action] });
      }
    });
  }
</script>
<svelte:window onpointerdown={event => { if (open && !menu?.contains(event.target as Node) && !trigger.contains(event.target as Node)) close(false); }} onblur={() => { if (open) close(false); }} />
<button class="trigger" bind:this={trigger} aria-label={t('Arrange')} aria-haspopup="menu" aria-expanded={open} onclick={show}><Icon name="front" /><span>{t('Arrange')} ▾</span></button>
{#if open}
  <div class="menu" role="menu" aria-label={t('Arrange')} tabindex="-1" bind:this={menu} use:place onkeydown={keys}>
    <div class="heading">{t('Reorder Objects')}</div>
    {#each order as item}<button role="menuitem" disabled={!editor.canRun(item.id)} onclick={() => choose(() => editor.invoke(item.id))}>{t(item.label)}</button>{/each}
    <hr /><div class="heading">{t('Group Objects')}</div>
    {#each [{ id: 'groupShapes', label: 'Group' }, { id: 'ungroupShapes', label: 'Ungroup' }] as item}<button role="menuitem" disabled={!editor.canRun(item.id)} onclick={() => choose(() => editor.invoke(item.id))}>{t(item.label)}</button>{/each}
    <button role="menuitem" disabled={!editor.canRegroup()} onclick={() => choose(() => editor.regroupSelection())}>{t('Regroup')}</button>
    <hr /><div class="heading">{t('Position Objects')}</div>
    <div class="branch">
      <button role="menuitem" data-branch="align" aria-label={t('Align')} aria-haspopup="menu" aria-expanded={branch === 'align'} disabled={!count} onpointerenter={() => branch = 'align'} onclick={() => branch = 'align'}>{t('Align')}<span>›</span></button>
      {#if branch === 'align' && count}<div class="menu submenu" role="menu" aria-label={t('Align')} use:place={true}>
        {#each alignment as item}<button role="menuitem" onclick={() => choose(() => editor.alignSelection(item.value, toSlide ? 'slide' : 'selection'))}>{t(item.label)}</button>{/each}
        <hr />
        <button role="menuitem" disabled={count === 0 || (!toSlide && count < 3)} onclick={() => choose(() => editor.distributeSelection('horizontal'))}>{t('Distribute Horizontally')}</button>
        <button role="menuitem" disabled={count === 0 || (!toSlide && count < 3)} onclick={() => choose(() => editor.distributeSelection('vertical'))}>{t('Distribute Vertically')}</button>
        <hr />
        <button role="menuitemradio" aria-checked={toSlide} onclick={() => choose(() => editor.alignmentReference = 'slide')}>{t('Align to Slide')}<span>{toSlide ? '✓' : ''}</span></button>
        <button role="menuitemradio" disabled={count < 2} aria-checked={!toSlide} onclick={() => choose(() => editor.alignmentReference = 'selection')}>{t('Align Selected Objects')}<span>{!toSlide ? '✓' : ''}</span></button>
      </div>{/if}
    </div>
    <div class="branch">
      <button role="menuitem" data-branch="rotate" aria-label={t('Rotate')} aria-haspopup="menu" aria-expanded={branch === 'rotate'} disabled={!count} onpointerenter={() => branch = 'rotate'} onclick={() => branch = 'rotate'}>{t('Rotate')}<span>›</span></button>
      {#if branch === 'rotate' && count}<div class="menu submenu" role="menu" aria-label={t('Rotate')} use:place={true}>
        {#each [{ action: 'right', label: 'Rotate Right 90°' }, { action: 'left', label: 'Rotate Left 90°' }, { action: 'vertical', label: 'Flip Vertical' }, { action: 'horizontal', label: 'Flip Horizontal' }] as item}<button role="menuitem" onclick={() => choose(() => transform(item.action as 'right' | 'left' | 'vertical' | 'horizontal'))}>{t(item.label)}</button>{/each}
        <hr /><button role="menuitem" disabled={!editor.canRun('setShapeRotation')} onclick={() => choose(() => editor.showRotationOptions())}>{t('More Rotation Options...')}</button>
      </div>{/if}
    </div>
    <hr /><button role="menuitemcheckbox" aria-label={t('Selection Pane...')} aria-checked={editor.selectionPaneVisible} onclick={() => choose(() => { editor.setViewMode('normal'); editor.selectionPaneVisible = !editor.selectionPaneVisible; })}>{t('Selection Pane...')}<span>{editor.selectionPaneVisible ? '✓' : ''}</span></button>
  </div>
{/if}
<style>
  .trigger { display: flex; flex-direction: column; align-items: center; gap: 3px; padding: 4px; background: transparent; border: 1px solid transparent; border-radius: var(--ok-radius); color: var(--ok-text); font: inherit; font-size: 11px; cursor: pointer; }
  .trigger:hover { background: var(--ok-hover); border-color: var(--ok-border); }
  .menu { position: fixed; z-index: 400; min-width: 210px; max-height: calc(100dvh - 16px); overflow-y: auto; padding: 5px; border: 1px solid var(--ok-border); border-radius: 6px; background: var(--ok-panel); color: var(--ok-text); box-shadow: var(--ok-shadow-lg); }
  .menu button { display: flex; justify-content: space-between; gap: 16px; width: 100%; border: 0; border-radius: 4px; padding: 5px 10px; background: transparent; color: inherit; font: inherit; font-size: 12px; text-align: left; white-space: nowrap; }
  .menu button:hover:not(:disabled), .menu button:focus-visible { background: var(--ok-accent); color: white; outline: none; }
  .menu button:disabled { opacity: .4; }
  .heading { padding: 5px 10px; font-size: 11px; color: var(--ok-text-3); }
  hr { border: 0; border-top: 1px solid var(--ok-border); margin: 5px; }
</style>
