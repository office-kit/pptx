<script lang="ts">
  import { eventTarget } from '../core/dom-root.ts';
  // The master tabs' Slide Size ▾ button: Standard, Widescreen and Page Setup.
  import { tick } from 'svelte';
  import { getSlideSize, setSlideSize, SLIDE_SIZE_16_9, SLIDE_SIZE_4_3, type SlideSize } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import Icon from '../ui/Icon.svelte';
  import { captionLines } from './caption.ts';
  import { placeBelowTrigger } from './place-menu.ts';

  const editor = getEditor();
  const doc = editor.doc;
  let open = $state(false);
  let anchor = $state<HTMLDivElement>();
  const slideSize = $derived.by(() => { doc.version; return getSlideSize(doc.pres); });
  const same = (size: SlideSize) => !!slideSize && slideSize.width === size.width && slideSize.height === size.height;
  async function toggle() {
    open = !open;
    if (open) {
      await tick();
      anchor?.querySelector<HTMLElement>('.menu [aria-checked="true"], .menu button')?.focus();
    }
  }
  function size(value: SlideSize) {
    open = false;
    // PowerPoint asks whether to maximize or ensure fit; Ensure Fit never
    // pushes content off the slide.
    doc.transact(t('Slide Size'), () => setSlideSize(doc.pres, value, { content: 'fit' }));
  }
</script>

<svelte:window onpointerdown={(event) => { if (open && !anchor?.contains(eventTarget(event) as Node)) open = false; }} onkeydown={(event) => { if (open && event.key === 'Escape') { event.preventDefault(); event.stopPropagation(); open = false; } }} />

<div class="anchor" bind:this={anchor}>
  <button class="big" aria-haspopup="menu" aria-expanded={open} onclick={toggle}><span class="icon-row"><Icon name="resize" size={32} /><span class="arrow" aria-hidden="true">⌄</span></span><span class="caption">{captionLines(t('Slide Size'))}</span></button>
  {#if open}
    <div class="menu" role="menu" use:placeBelowTrigger aria-label={t('Slide Size')}>
      <button class="action" role="menuitemradio" aria-checked={same(SLIDE_SIZE_4_3)} onclick={() => size(SLIDE_SIZE_4_3)}>{t('Standard (4:3)')}</button>
      <button class="action" role="menuitemradio" aria-checked={same(SLIDE_SIZE_16_9)} onclick={() => size(SLIDE_SIZE_16_9)}>{t('Widescreen (16:9)')}</button>
      <hr />
      <button class="action" role="menuitem" disabled={!editor.canRun('setSlideSize')} onclick={() => { open = false; editor.runOrPrompt('setSlideSize'); }}>{t('Page Setup...')}</button>
    </div>
  {/if}
</div>

<style>
  .anchor { position: relative; display: flex; }
  button { font: inherit; color: var(--ok-text); background: none; border: 1px solid transparent; border-radius: var(--ok-radius); cursor: pointer; }
  button:hover:not(:disabled) { background: var(--ok-hover); }
  button:disabled { opacity: 0.4; cursor: default; }
  .big { display: flex; flex-direction: column; align-items: center; justify-content: flex-start; gap: 2px; min-width: 50px; padding: 4px 1px; font-size: 11px; line-height: 1.15; }
  .caption { white-space: pre-line; text-align: center; }
  .icon-row { display: flex; align-items: center; gap: 1px; }
  .arrow { font-size: 10px; }
  .menu { position: fixed; z-index: 400; display: flex; flex-direction: column; min-width: 150px; padding: 4px 0; border: 1px solid var(--ok-border); border-radius: 6px; background: var(--ok-panel); box-shadow: var(--ok-shadow-lg); }
  .action { display: flex; align-items: center; width: 100%; height: 24px; padding: 0 20px; border: 0; border-radius: 0; text-align: left; font-size: 12px; white-space: nowrap; }
  .action[aria-checked='true']::before { content: '✓'; margin-left: -14px; margin-right: 4px; }
  .menu hr { width: 100%; border: none; border-top: 1px solid var(--ok-border); margin: 5px 0; }
</style>
