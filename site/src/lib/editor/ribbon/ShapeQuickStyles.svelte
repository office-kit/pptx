<script lang="ts">
  import { tick } from 'svelte';
  import { getShapeKind, type SlideData } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { applyShapeQuickStyle, quickStyleColors, themeQuickStyles, presetQuickStyles, type ShapeQuickStyle, type QuickStyleColor } from '../core/shape-quick-styles.ts';
  import { shapeQuickStyleSwatches } from '../core/shape-quick-style-swatches.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  let { compact = false, inline = false }: { compact?: boolean; inline?: boolean } = $props();
  const editor = getEditor();
  const doc = editor.doc;
  let open = $state(false);
  let trigger = $state<HTMLButtonElement>();
  let menu = $state<HTMLDivElement>();
  let images = $state<ReadonlyMap<string, string>>(new Map());
  let error = $state('');
  let swatchSlide = $state.raw<SlideData>();
  let swatchVersion = $state(-1);
  const inlineItems = [...themeQuickStyles, ...presetQuickStyles].flatMap(style =>
    quickStyleColors.map(color => ({ style, color })),
  );
  let inlinePage = $state(0);
  const inlinePageCount = Math.ceil(inlineItems.length / 3);
  const inlineVisible = $derived(inlineItems.slice(inlinePage * 3, inlinePage * 3 + 3));
  const disabled = $derived.by(() => {
    doc.version;
    const shapes = editor.selectedShapes();
    return editor.selectionLocked() || !shapes.length || shapes.some(shape => !['shape', 'connector'].includes(getShapeKind(shape)));
  });
  function close(restore = true) { open = false; if (restore && !inline) trigger!.focus(); }
  function loadImages() {
    const slide = doc.currentSlide;
    const version = doc.version;
    if (!slide || (swatchSlide === slide && swatchVersion === version)) return;
    error = '';
    let succeeded = false;
    try {
      images = shapeQuickStyleSwatches(slide);
      succeeded = true;
    } catch (cause) { error = cause instanceof Error ? cause.message : String(cause); }
    finally {
      if (succeeded) { swatchSlide = slide; swatchVersion = version; }
    }
  }
  function pageInline(delta: number) {
    inlinePage = Math.max(0, Math.min(inlinePageCount - 1, inlinePage + delta));
  }
  async function show() {
    if (open) { close(); return; }
    if (!doc.currentSlide || disabled) return;
    error = '';
    loadImages();
    open = true;
    await tick();
    menu?.querySelector<HTMLButtonElement>('button')?.focus();
  }
  $effect(() => {
    doc.currentSlide;
    doc.version;
    if (inline && !disabled) loadImages();
  });
  function apply(style: ShapeQuickStyle, color: QuickStyleColor) {
    if (disabled) return;
    try {
      doc.transact(t('Quick Styles'), () => { for (const shape of editor.selectedShapes()) applyShapeQuickStyle(shape, style, color); });
      close();
    } catch (cause) { error = cause instanceof Error ? cause.message : String(cause); }
  }
  function place(node: HTMLElement) {
    const bounds = trigger!.getBoundingClientRect();
    node.style.left = `${Math.max(8, Math.min(bounds.left, innerWidth - node.offsetWidth - 8))}px`;
    node.style.top = `${Math.max(8, Math.min(bounds.bottom, innerHeight - node.offsetHeight - 8))}px`;
  }
  function keys(event: KeyboardEvent) {
    event.stopPropagation();
    if (event.key === 'Escape') { event.preventDefault(); close(); return; }
    if (event.key === 'Tab') { close(false); return; }
    const offsets: Record<string, number> = { ArrowLeft: -1, ArrowRight: 1, ArrowUp: -7, ArrowDown: 7 };
    if (!(event.key in offsets) && event.key !== 'Home' && event.key !== 'End') return;
    event.preventDefault();
    const items = [...menu!.querySelectorAll<HTMLButtonElement>('button')];
    const index = items.indexOf(event.target as HTMLButtonElement);
    items[event.key === 'Home' ? 0 : event.key === 'End' ? items.length - 1 : Math.max(0, Math.min(items.length - 1, index + offsets[event.key]!))]?.focus();
  }
</script>
<svelte:window onpointerdown={event => { if (open && !menu?.contains(event.target as Node) && !trigger!.contains(event.target as Node)) close(false); }} onblur={() => { if (open) close(false); }} onresize={() => { if (open) close(false); }} />
  {#if inline}
    <div class="inline-gallery" role="group" aria-label={t('Quick Styles')}>
    <button class="inline-more" class:hidden={inlinePage === 0} {disabled} aria-label={t('Previous Quick Styles')} onclick={() => pageInline(-1)}>‹</button>
    <div class="inline-items">
    {#each inlineVisible as item}
      <button class="inline-item" {disabled} aria-label={`${t(item.style)} - ${t(item.color === 'dk1' ? 'Dark 1' : `Accent ${item.color.slice(-1)}`)}`} title={`${t(item.style)} - ${t(item.color === 'dk1' ? 'Dark 1' : `Accent ${item.color.slice(-1)}`)}`} onclick={() => apply(item.style, item.color)}>
        {#if images.has(`${item.style}:${item.color}`)}<img src={images.get(`${item.style}:${item.color}`)} alt="" />{:else}<span>Abc</span>{/if}
      </button>
    {/each}
    {#each Array(3 - inlineVisible.length) as _}
      <span class="inline-item placeholder" aria-hidden="true"></span>
    {/each}
    </div>
    <button class="inline-more" bind:this={trigger} class:hidden={inlinePage === inlinePageCount - 1} {disabled} aria-label={t('Next Quick Styles')} onclick={() => pageInline(1)}>›</button>
    {#if error}<p class="inline-error" role="alert">{error}</p>{/if}
  </div>
{:else}
  <button class="trigger" class:compact bind:this={trigger} {disabled} aria-label={t('Quick Styles')} aria-haspopup="menu" aria-expanded={open} onclick={show}>{#if compact}<span>{t('Quick Styles')} ▾</span>{:else}<span class="icon-row"><span class="sample" aria-hidden="true">Abc</span><span aria-hidden="true">▾</span></span><span>{t('Quick Styles')}</span>{/if}</button>
{/if}
{#if open}
  <div class="menu" role="menu" aria-label={t('Quick Styles')} tabindex="-1" bind:this={menu} use:place onkeydown={keys}>
    {#each [{ title: 'Theme Styles', rows: themeQuickStyles }, { title: 'Presets', rows: presetQuickStyles }] as group}
      <div class="heading">{t(group.title)}</div>
      <div class="gallery" role="group" aria-label={t(group.title)}>
        {#each group.rows as style}{#each quickStyleColors as color}
          <button role="menuitem" aria-label={`${t(style)} - ${t(color === 'dk1' ? 'Dark 1' : `Accent ${color.slice(-1)}`)}`} title={`${t(style)} - ${t(color === 'dk1' ? 'Dark 1' : `Accent ${color.slice(-1)}`)}`} onclick={() => apply(style, color)}>
            {#if images.has(`${style}:${color}`)}<img src={images.get(`${style}:${color}`)} alt="" />{:else}<span>Abc</span>{/if}
          </button>
        {/each}{/each}
      </div>
    {/each}
    {#if error}<p role="alert">{error}</p>{/if}
  </div>
{/if}
<style>
  .trigger { display:flex; flex-direction:column; align-items:center; gap:3px; padding:4px; background:transparent; border:1px solid transparent; border-radius:var(--ok-radius); color:var(--ok-text); font:inherit; font-size:11px; cursor:pointer; }
  .trigger:not(.compact) > span:last-child { max-width:46px; text-align:center; line-height:1.15; }
  .trigger:not(.compact) { align-self:stretch; justify-content:flex-start; gap:2px; min-width:50px; padding:4px 2px; }
  .icon-row { display:flex; align-items:center; gap:1px; font-size:10px; }
  .trigger:hover { background:var(--ok-hover); border-color:var(--ok-border); }
  .trigger:disabled { opacity:.4; }
  .sample { display:flex; align-items:center; justify-content:center; width:34px; height:30px; border:1px solid currentColor; border-radius:3px; font-size:13px; }
  .compact { flex-direction:row; }
  /* PowerPoint's in-ribbon strip: 18 × 58 pt arrows around three 58 pt
     swatches in a 174 pt frame. */
  .inline-gallery { display:flex; align-items:flex-start; flex:none; height:60px; }
  .inline-items { display:flex; align-items:center; justify-content:space-around; width:174px; height:58px; margin-top:1px; box-sizing:border-box; border:1px solid var(--ok-border); border-radius:3px; background:white; }
  .inline-item, .inline-more { width:54px; height:50px; padding:0; border:1px solid transparent; border-radius:2px; background:white; color:var(--ok-text); cursor:pointer; }
  .inline-item:hover:not(:disabled), .inline-item:focus-visible, .inline-more:hover:not(:disabled), .inline-more:focus-visible { outline:2px solid var(--ok-accent); outline-offset:-2px; }
  .inline-item:disabled, .inline-more:disabled { opacity:.4; cursor:default; }
  .inline-item img { width:100%; height:100%; display:block; object-fit:contain; }
  .inline-item.placeholder { visibility:hidden; pointer-events:none; }
  .inline-more { width:18px; height:58px; margin-top:1px; background:none; font-size:14px; line-height:1; }
  .inline-more:hover:not(:disabled) { background:var(--ok-hover); outline:none; }
  .inline-more.hidden { visibility:hidden; }
  .inline-error { flex-basis:100%; max-width:280px; color:var(--ok-danger); }
  .menu { position:fixed; z-index:400; padding:5px; max-height:calc(100dvh - 16px); overflow-y:auto; border:1px solid var(--ok-border); border-radius:6px; background:var(--ok-panel); color:var(--ok-text); box-shadow:var(--ok-shadow-lg); }
  .heading { padding:4px 6px; font-size:12px; font-weight:600; }
  .gallery { display:grid; grid-template-columns:repeat(7, 40px); gap:2px; }
  .gallery button { padding:0; width:40px; height:30px; border:1px solid transparent; background:white; border-radius:2px; }
  .gallery button:hover, .gallery button:focus-visible { outline:2px solid var(--ok-accent); outline-offset:-2px; }
  img { width:100%; height:100%; display:block; }
  p { max-width:280px; color:var(--ok-danger); }
</style>
