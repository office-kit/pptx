<script lang="ts">
  import { tick } from 'svelte';
  import { radialDirections } from '../core/gradient-directions.ts';
  import { t } from '../i18n/i18n.svelte.ts';

  let { angle, disabled, choose, radial = false }: { radial?: boolean; angle: number | undefined; disabled: boolean; choose: (angle: number) => void } = $props();
  const linearDirections = [
    [45, 'Linear Diagonal - Top Left to Bottom Right'],
    [90, 'Linear Down'],
    [135, 'Linear Diagonal - Top Right to Bottom Left'],
    [0, 'Linear Right'],
    [180, 'Linear Left'],
    [315, 'Linear Diagonal - Bottom Left to Top Right'],
    [270, 'Linear Up'],
    [225, 'Linear Diagonal - Bottom Right to Top Left'],
  ] as const;
  const directions = $derived(radial ? radialDirections.map((item, index) => [index, item.label] as const) : linearDirections);
  const columns = $derived(radial ? 5 : 4);
  let open = $state(false);
  let trigger: HTMLButtonElement;
  let menu = $state<HTMLDivElement>();
  function close(restore = true) { open = false; if (restore) trigger.focus(); }
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
    const offsets: Record<string, number> = { ArrowLeft: -1, ArrowRight: 1, ArrowUp: -columns, ArrowDown: columns };
    if (!(event.key in offsets) && event.key !== 'Home' && event.key !== 'End') return;
    event.preventDefault();
    const items = [...menu!.querySelectorAll<HTMLButtonElement>('button')];
    const index = items.indexOf(event.target as HTMLButtonElement);
    items[event.key === 'Home' ? 0 : event.key === 'End' ? items.length - 1 : (index + offsets[event.key]! + items.length) % items.length]?.focus();
  }
  $effect(() => { angle; disabled; radial; open = false; });
</script>

{#snippet swatch(value: number)}
  <span class="swatch" style:background={radial ? `radial-gradient(ellipse farthest-side at ${radialDirections[value]!.x * 100}% ${radialDirections[value]!.y * 100}%, #ecf1fa, #4472c4)` : `linear-gradient(${value + 90}deg, #4472c4, #ecf1fa)`}></span>
{/snippet}
<svelte:window onpointerdown={event => { if (open && !menu?.contains(event.target as Node) && !trigger.contains(event.target as Node)) close(false); }} onblur={() => { if (open) close(false); }} onresize={() => { if (open) close(false); }} />
<div class="field"><span>{t('Direction')}</span><button class="ok-input trigger" bind:this={trigger} aria-label={t('Gradient direction')} aria-haspopup="menu" aria-expanded={open} {disabled} onclick={show}>{#if angle === undefined}<span class="swatch"></span>{:else}{@render swatch(angle)}{/if}<span>▾</span></button></div>
{#if open}
  <div class="gallery" style:grid-template-columns={`repeat(${columns}, 44px)`} role="menu" aria-label={t('Gradient direction')} tabindex="-1" bind:this={menu} use:place onkeydown={keys}>
    {#each directions as [value, label]}
      <button role="menuitemradio" aria-label={t(label)} title={t(label)} aria-checked={angle === value} onclick={() => { if (!disabled) choose(value); close(); }}>{@render swatch(value)}</button>
    {/each}
  </div>
{/if}
<style>
  .field { display: flex; align-items: center; justify-content: space-between; gap: 8px; }
  .trigger { display: flex; align-items: center; justify-content: space-between; width: 78px; padding: 3px 5px; }
  .swatch { display: block; width: 32px; height: 26px; border: 1px solid var(--ok-border); }
  .gallery { position: fixed; z-index: 400; display: grid; grid-template-columns: repeat(4, 44px); gap: 3px; padding: 6px; background: var(--ok-panel); border: 1px solid var(--ok-border); border-radius: 6px; box-shadow: var(--ok-shadow-lg); }
  .gallery button { border: 1px solid transparent; border-radius: 3px; background: transparent; padding: 5px; }
  .gallery button:hover, .gallery button:focus-visible, .gallery button[aria-checked=true] { background: var(--ok-hover); border-color: var(--ok-accent); }
</style>
