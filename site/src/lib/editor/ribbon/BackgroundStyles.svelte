<script lang="ts">
  import { tick } from 'svelte';
  import { getSlideMasterBackgroundStyles, setSlideMasterBackgroundStyle, getSlideBackground, clearSlideBackground, getSlides, type SlideMasterBackgroundStyle } from '@office-kit/pptx';
  import { selectedSlideIndices } from '../core/selection.ts';
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import Icon from '../ui/Icon.svelte';
  const editor = getEditor();
  const doc = editor.doc;
  let open = $state(false);
  let trigger: HTMLButtonElement;
  let menu = $state<HTMLDivElement>();
  let error = $state('');
  const styles = $derived.by(() => { doc.version; return open && doc.currentSlide ? getSlideMasterBackgroundStyles(doc.currentSlide) : []; });
  const slides = $derived.by(() => { doc.version; const all = getSlides(doc.pres); return selectedSlideIndices(doc.selection).flatMap(index => all[index] ? [all[index]!] : []); });
  const canReset = $derived(slides.some(slide => getSlideBackground(slide).kind !== 'inherit'));
  function close(restore = true) { open = false; if (restore) trigger.focus(); }
  async function show() {
    if (open) { close(); return; }
    error = ''; open = true; await tick();
    (menu?.querySelector<HTMLButtonElement>('[aria-checked="true"]') ?? menu?.querySelector<HTMLButtonElement>('button:not(:disabled)'))?.focus();
  }
  function apply(style: number) {
    if (!doc.currentSlide) return;
    try { doc.transact(t('Background Styles'), () => setSlideMasterBackgroundStyle(doc.currentSlide!, style)); close(); }
    catch (cause) { error = cause instanceof Error ? cause.message : String(cause); }
  }
  function reset() {
    try { doc.transact(t('Reset Slide Background'), () => { for (const slide of slides) clearSlideBackground(slide); }); close(); }
    catch (cause) { error = cause instanceof Error ? cause.message : String(cause); }
  }
  function preview(style: SlideMasterBackgroundStyle): string {
    const gradient = style.gradient;
    if (!gradient) return style.fill.kind === 'solid' ? `${style.fill.color}${Math.round((style.fill.opacity ?? 1) * 255).toString(16).padStart(2, '0')}` : 'transparent';
    const stops = [...gradient.stops].sort((a, b) => a.offset - b.offset).map(stop => `${stop.resolvedColor ?? stop.color}${Math.round((stop.opacity ?? 1) * 255).toString(16).padStart(2, '0')} ${stop.offset * 100}%`).join(', ');
    const focus = gradient.focus;
    const x = focus ? (focus.left + 1 - focus.right) / 2 : .5;
    const y = focus ? (focus.top + 1 - focus.bottom) / 2 : .5;
    return gradient.path ? `radial-gradient(ellipse farthest-side at ${x * 100}% ${y * 100}%, ${stops})` : `linear-gradient(${(gradient.angleDeg ?? 0) + 90}deg, ${stops})`;
  }
  function place(node: HTMLElement) {
    const bounds = trigger.getBoundingClientRect();
    node.style.left = `${Math.max(8, Math.min(bounds.left, innerWidth - node.offsetWidth - 8))}px`;
    node.style.top = `${Math.max(8, Math.min(bounds.bottom, innerHeight - node.offsetHeight - 8))}px`;
  }
  function keys(event: KeyboardEvent) {
    event.stopPropagation();
    if (event.key === 'Escape') { event.preventDefault(); close(); return; }
    if (event.key === 'Tab') { close(false); return; }
    const items = [...menu!.querySelectorAll<HTMLButtonElement>('button:not(:disabled)')];
    const index = items.indexOf(event.target as HTMLButtonElement);
    const offsets: Record<string, number> = { ArrowLeft: -1, ArrowRight: 1, ArrowUp: index < styles.length ? -4 : -1, ArrowDown: index < styles.length ? 4 : 1 };
    if (!(event.key in offsets) && event.key !== 'Home' && event.key !== 'End') return;
    event.preventDefault();
    items[event.key === 'Home' ? 0 : event.key === 'End' ? items.length - 1 : Math.max(0, Math.min(items.length - 1, index + offsets[event.key]!))]?.focus();
  }
</script>
<svelte:window onpointerdown={event => { if (open && !menu?.contains(event.target as Node) && !trigger.contains(event.target as Node)) close(false); }} onblur={() => { if (open) close(false); }} onresize={() => { if (open) close(false); }} />
<button class="trigger" bind:this={trigger} disabled={!doc.currentSlide} aria-label={t('Background Styles')} aria-haspopup="menu" aria-expanded={open} onclick={show}><Icon name="background" /><span>{t('Background Styles')} ▾</span></button>
{#if open}
  <div class="menu" role="menu" aria-label={t('Background Styles')} tabindex="-1" bind:this={menu} use:place onkeydown={keys}>
    <div class="gallery" role="group" aria-label={t('Background Styles')}>
      {#each styles as style}
        <button class="preset" role="menuitemradio" aria-label={`${t('Style')} ${style.style}`} title={`${t('Style')} ${style.style}`} aria-checked={style.selected} onclick={() => apply(style.style)}><span class="swatch" style:background={preview(style)}></span></button>
      {/each}
    </div>
    <hr />
    <button class="action" role="menuitem" onclick={() => { close(); editor.showBackgroundFormat(); }}>{t('Format Background...')}</button>
    <button class="action" role="menuitem" disabled={!canReset} onclick={reset}>{t('Reset Slide Background')}</button>
    {#if error}<p role="alert">{error}</p>{/if}
  </div>
{/if}
<style>
  .trigger { display: flex; flex-direction: column; align-items: center; gap: 3px; padding: 4px; background: transparent; border: 1px solid transparent; border-radius: var(--ok-radius); color: var(--ok-text); font: inherit; font-size: 11px; cursor: pointer; }
  .trigger:hover { background: var(--ok-hover); border-color: var(--ok-border); }
  .trigger:disabled { opacity: .4; }
  .menu { position: fixed; z-index: 400; padding: 5px; max-height: calc(100dvh - 16px); overflow-y: auto; border: 1px solid var(--ok-border); border-radius: 6px; background: var(--ok-panel); color: var(--ok-text); box-shadow: var(--ok-shadow-lg); }
  .gallery { display: grid; grid-template-columns: repeat(4, 58px); gap: 3px; }
  .preset { padding: 4px; background: transparent; border: 1px solid transparent; border-radius: 3px; }
  .swatch { display: block; width: 48px; height: 32px; border: 1px solid var(--ok-border); }
  .preset:hover, .preset:focus-visible, .preset[aria-checked=true] { background: var(--ok-hover); border-color: var(--ok-accent); }
  .action { display: block; width: 100%; border: 0; border-radius: 4px; padding: 5px 10px; background: transparent; color: inherit; font: inherit; font-size: 12px; text-align: left; }
  .action:hover:not(:disabled), .action:focus-visible { background: var(--ok-accent); color: white; }
  .action:disabled { opacity: .4; }
  hr { border: 0; border-top: 1px solid var(--ok-border); margin: 5px; }
  p { max-width: 240px; color: var(--ok-danger); }
</style>
