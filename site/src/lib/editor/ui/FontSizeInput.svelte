<script lang="ts">
  import { tick } from 'svelte';
  import { t } from '../i18n/i18n.svelte.ts';

  let { value, disabled = false, choose }: { value?: number; disabled?: boolean; choose: (size: number) => void } = $props();
  // The size gallery in Mac PowerPoint includes half-point 10.5 and extends to 96.
  const sizes = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 40, 44, 48, 54, 60, 66, 72, 80, 88, 96];
  let open = $state(false);
  let field: HTMLInputElement;
  let trigger: HTMLButtonElement;
  let menu = $state<HTMLDivElement>();
  function close(restore = true) { open = false; if (restore) field.focus(); }
  function select(size: number) { if (!disabled) choose(size); close(); }
  async function show() {
    if (open) { close(); return; }
    open = true;
    await tick();
    const item = menu?.querySelector<HTMLButtonElement>('[aria-checked="true"]') ?? menu?.querySelector<HTMLButtonElement>('button');
    item?.focus();
    item?.scrollIntoView({ block: 'nearest' });
  }
  function place(node: HTMLElement) {
    const bounds = field.getBoundingClientRect();
    node.style.left = `${Math.max(8, Math.min(bounds.left, innerWidth - node.offsetWidth - 8))}px`;
    node.style.top = `${Math.max(8, Math.min(bounds.bottom, innerHeight - node.offsetHeight - 8))}px`;
  }
  function keys(event: KeyboardEvent) {
    event.stopPropagation();
    if (event.key === 'Escape') { event.preventDefault(); close(); return; }
    if (event.key === 'Tab') { close(false); return; }
    if (!['ArrowUp', 'ArrowDown', 'Home', 'End'].includes(event.key)) return;
    event.preventDefault();
    const items = [...menu!.querySelectorAll<HTMLButtonElement>('button')];
    const index = items.indexOf(event.target as HTMLButtonElement);
    items[event.key === 'Home' ? 0 : event.key === 'End' ? items.length - 1 : (index + (event.key === 'ArrowDown' ? 1 : -1) + items.length) % items.length]?.focus();
  }
  $effect(() => { value; disabled; open = false; });
</script>

<svelte:window onpointerdown={event => { if (open && !menu?.contains(event.target as Node) && !trigger.contains(event.target as Node)) close(false); }} onblur={() => { if (open) close(false); }} onresize={() => { if (open) close(false); }} />
<div class="size-field">
  <input class="ok-input" bind:this={field} aria-label={t('Font size')} type="number" min="1" max="4000" step="0.5" {disabled} value={value ?? ''} placeholder="—" onchange={event => { if (event.currentTarget.value && event.currentTarget.reportValidity()) choose(event.currentTarget.valueAsNumber); }} onkeydown={event => { if (event.altKey && event.key === 'ArrowDown') { event.preventDefault(); void show(); } }} />
  <button class="ok-input" bind:this={trigger} aria-label={t('Font size options')} aria-haspopup="menu" aria-expanded={open} {disabled} onclick={show}>▾</button>
</div>
{#if open}
  <div class="size-menu" role="menu" aria-label={t('Font size')} tabindex="-1" bind:this={menu} use:place onkeydown={keys}>
    {#each sizes as size}
      <button role="menuitemradio" aria-checked={value === size} onclick={() => select(size)}>{size}</button>
    {/each}
  </div>
{/if}

<style>
  .size-field { display: flex; }
  input { width: 43px; padding-right: 0; border-radius: 3px 0 0 3px; appearance: textfield; }
  input::-webkit-inner-spin-button, input::-webkit-outer-spin-button { appearance: none; margin: 0; }
  .size-field button { width: 16px; padding: 0; border-left: 0; border-radius: 0 3px 3px 0; }
  .size-menu { position: fixed; z-index: 400; width: 65px; max-height: min(480px, calc(100vh - 16px)); overflow-y: auto; padding: 3px; background: var(--ok-panel); border: 1px solid var(--ok-border); border-radius: 4px; box-shadow: var(--ok-shadow-lg); }
  .size-menu button { display: block; width: 100%; border: 0; padding: 3px 8px; background: transparent; color: inherit; text-align: left; font: inherit; }
  .size-menu button:hover, .size-menu button:focus-visible, .size-menu button[aria-checked='true'] { background: var(--ok-hover); outline: none; }
</style>
