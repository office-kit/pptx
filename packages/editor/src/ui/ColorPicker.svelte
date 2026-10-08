<script lang="ts">
  import { eventTarget } from '../core/dom-root.ts';
  import { tick } from 'svelte';
  import { asColor, getPresentationTheme, type Color, type ColorTransform } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { paletteSwatches, sameColorTransforms, type PaletteSwatch } from '../core/theme-palette.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import type { TextureId } from '../core/textures.ts';
  import TextureGallery from './TextureGallery.svelte';

  let { label, value, resolvedColor, disabled = false, compact = false, glyph, showThemeShades = false, selectedColorTransforms = [], automatic, automaticSelected = false, texture, choose }: {
    label: string;
    /** Only a ⌄ arrow, as the reference desktop app's Shape Fill / Shape Outline split buttons show. */
    compact?: boolean;
    /** With `compact`, a letter drawn over a bar of the current color, like the reference desktop app's Font Color. */
    glyph?: string;
    automatic?: () => void;
    automaticSelected?: boolean;
    value?: string;
    resolvedColor?: string;
    disabled?: boolean;
    showThemeShades?: boolean;
    selectedColorTransforms?: readonly ColorTransform[];
    /** Shape Fill's Texture ▸ submenu; `disabled` greys it out where textures do not apply. */
    texture?: { disabled?: boolean; choose: (id: TextureId) => void; more: () => void };
    choose: (color: Color, colorTransforms?: readonly ColorTransform[]) => void;
  } = $props();
  const editor = getEditor();
  const theme = $derived.by(() => { editor.doc.version; return getPresentationTheme(editor.doc.pres); });
  const colors = $derived(paletteSwatches(theme, showThemeShades));
  const paint = $derived(resolvedColor ?? colors.find(color => color.color.toLowerCase() === value?.replace(/^scheme:/, '').toLowerCase())?.paint ?? value);
  let open = $state(false);
  let trigger: HTMLButtonElement;
  let custom: HTMLInputElement;
  let menu = $state<HTMLDivElement>();
  let textureItem = $state<HTMLButtonElement>();
  let textureOpen = $state(false);
  let submenu = $state<HTMLDivElement>();
  function close(restore = true) { open = false; textureOpen = false; if (restore) trigger.focus(); }
  function select(color: string) {
    const parsed = asColor(color);
    if (parsed && !disabled && !trigger.matches(':disabled')) choose(parsed);
    close();
  }
  function selectTheme(swatch: PaletteSwatch): void {
    const parsed = asColor(swatch.color);
    if (parsed && !disabled && !trigger.matches(':disabled')) choose(parsed, swatch.transforms);
    close();
  }
  async function show() {
    if (open) { close(); return; }
    open = true;
    await tick();
    (menu?.querySelector<HTMLButtonElement>('[aria-checked="true"]') ?? menu?.querySelector<HTMLButtonElement>('button'))?.focus();
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
    if (event.key === 'ArrowRight' && event.target === textureItem) { event.preventDefault(); textureOpen = true; return; }
    const offsets: Record<string, number> = { ArrowLeft: -1, ArrowRight: 1, ArrowUp: -10, ArrowDown: 10 };
    if (!(event.key in offsets) && event.key !== 'Home' && event.key !== 'End') return;
    event.preventDefault();
    const items = [...menu!.querySelectorAll<HTMLButtonElement>('button')];
    const index = items.indexOf(event.target as HTMLButtonElement);
    items[event.key === 'Home' ? 0 : event.key === 'End' ? items.length - 1 : (index + offsets[event.key]! + items.length) % items.length]?.focus();
  }
  $effect(() => { value; disabled; editor.doc.selection; open = false; });
</script>

<svelte:window onpointerdown={event => { if (open && !menu?.contains(eventTarget(event) as Node) && !submenu?.contains(eventTarget(event) as Node) && !trigger.contains(eventTarget(event) as Node)) close(false); }} onblur={() => { if (open) close(false); }} onresize={() => { if (open) close(false); }} />
<button type="button" class={compact ? 'trigger compact' : 'ok-input trigger'} bind:this={trigger} aria-label={label} aria-haspopup="menu" aria-expanded={open} {disabled} onclick={show}>{#if !compact}<span class="swatch" style:background={paint ?? 'transparent'}></span>{:else if glyph}<span class="glyph" aria-hidden="true">{glyph}<span class="swatch" style:background={paint ?? 'transparent'}></span></span>{/if}<span>▾</span></button>
<input class="custom" type="color" bind:this={custom} aria-label={`${label}: ${t('More Colors...')}`} tabindex="-1" {disabled} value={paint?.startsWith('#') ? paint : '#000000'} onchange={event => select(event.currentTarget.value)} />
{#if open}
  <div class="palette" role="menu" aria-label={label} tabindex="-1" bind:this={menu} use:place onkeydown={keys}>
    {#if automatic}
      <button type="button" class="more" role="menuitemradio" aria-checked={automaticSelected} onclick={() => { if (!disabled) automatic?.(); close(); }}>{t('Automatic')}</button>
    {/if}
    {#each [true, false] as isTheme}
      {#if !isTheme || theme}
        <div class="heading">{t(isTheme ? 'Theme Colors' : 'Standard Colors')}</div>
        <div class="colors" role="group" aria-label={t(isTheme ? 'Theme Colors' : 'Standard Colors')}>
          {#each colors.filter(color => color.theme === isTheme) as color}
            {@const label = `${t(color.name)}${color.shade ? `, ${t(color.shade)} ${color.shadePercent}%` : ''}`}
            <button type="button" role="menuitemradio" aria-label={label} title={label} aria-checked={value?.replace(/^scheme:/, '').toLowerCase() === color.color.toLowerCase() && sameColorTransforms(selectedColorTransforms, color.transforms)} style:background={color.paint} onclick={() => color.theme ? selectTheme(color) : select(color.color)}></button>
          {/each}
        </div>
      {/if}
    {/each}
    <button type="button" class="more" role="menuitem" onclick={() => { close(); custom.click(); }}>{t('More Colors...')}</button>
    {#if texture}
      <button type="button" class="more submenu-item" role="menuitem" aria-haspopup="menu" aria-expanded={textureOpen} disabled={texture.disabled} bind:this={textureItem} onclick={() => { textureOpen = true; }} onpointerenter={() => { if (!texture.disabled) textureOpen = true; }}><span>{t('Texture')}</span><span aria-hidden="true">▸</span></button>
    {/if}
  </div>
  {#if texture && textureOpen && textureItem}
    <div class="submenu" bind:this={submenu}>
      <TextureGallery label={t('Texture')} anchor={textureItem} side="right"
        choose={id => { close(); texture.choose(id); }}
        more={() => { close(); texture.more(); }}
        close={() => { textureOpen = false; }} />
    </div>
  {/if}
{/if}

<style>
  .trigger { display: inline-flex; align-items: center; justify-content: space-between; gap: 6px; width: 54px; padding: 3px 5px; }
  .trigger.compact { width: auto; padding: 1px 3px; font: inherit; color: var(--ok-text); background: none; border: 1px solid transparent; border-radius: var(--ok-radius); cursor: pointer; }
  .trigger.compact:hover:not(:disabled) { background: var(--ok-hover); }
  .glyph { display: inline-flex; flex-direction: column; align-items: center; font-size: 13px; font-weight: 600; line-height: 13px; }
  .glyph .swatch { width: 14px; height: 3px; border: none; border-radius: 0; }
  .swatch { display: block; width: 24px; height: 18px; border: 1px solid var(--ok-border); }
  .custom { position: fixed; opacity: 0; pointer-events: none; width: 1px; height: 1px; }
  .palette { position: fixed; z-index: 400; padding: 6px; background: var(--ok-panel); border: 1px solid var(--ok-border); border-radius: 6px; box-shadow: var(--ok-shadow-lg); }
  .heading { font-size: 11px; margin: 3px 2px 6px; }
  .colors { display: grid; grid-template-columns: repeat(10, 18px); gap: 3px; margin-bottom: 10px; }
  .colors button { width: 18px; height: 18px; padding: 0; border: 1px solid var(--ok-border); }
  .colors button:hover, .colors button:focus-visible, .colors button[aria-checked=true] { outline: 2px solid var(--ok-accent); outline-offset: 1px; }
  .more { display: block; width: 100%; text-align: left; border: 0; border-top: 1px solid var(--ok-border); background: transparent; color: inherit; padding: 6px 2px 2px; font-size: inherit; }
  .more:hover, .more:focus-visible { background: var(--ok-hover); }
  .more:disabled { color: var(--ok-muted); background: transparent; }
  .submenu-item { display: flex; justify-content: space-between; border-top: 0; }
  .submenu { display: contents; }
</style>
