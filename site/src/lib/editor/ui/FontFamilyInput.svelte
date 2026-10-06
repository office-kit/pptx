<script lang="ts">
  import { tick } from 'svelte';
  import { t } from '../i18n/i18n.svelte.ts';

  let { value, label = t('Font'), disabled = false, ribbon = false, families: extraFamilies = [], choose }: { value?: string; label?: string; disabled?: boolean; ribbon?: boolean; families?: readonly string[]; choose: (font: string) => void } = $props();
  const commonFamilies = [
    'Aptos', 'Aptos Display', 'Arial', 'Calibri', 'Cambria', 'Candara',
    'Consolas', 'Courier New', 'Georgia', 'Helvetica', 'Meiryo', 'MS Gothic',
    'Noto Sans', 'Noto Sans CJK JP', 'Segoe UI', 'Times New Roman', '游ゴシック',
  ];
  const families = $derived([...new Set([...extraFamilies, value ?? '', ...commonFamilies].filter(Boolean))]);
  let open = $state(false);
  let filter = $state('');
  let field: HTMLInputElement;
  let trigger: HTMLButtonElement;
  let menu = $state<HTMLDivElement>();
  const matches = $derived(families.filter((family) => family.toLowerCase().includes(filter.trim().toLowerCase())));
  function close(restore = true) { open = false; if (restore) field.focus(); }
  function select(font: string) { if (!disabled) choose(font); close(); }
  async function show() {
    if (open) { close(); return; }
    filter = '';
    open = true;
    await tick();
    menu?.querySelector<HTMLInputElement>('input')?.focus();
  }
  function place(node: HTMLElement) {
    const bounds = field.getBoundingClientRect();
    node.style.left = `${Math.max(8, Math.min(bounds.left, innerWidth - node.offsetWidth - 8))}px`;
    node.style.top = `${Math.max(8, Math.min(bounds.bottom, innerHeight - node.offsetHeight - 8))}px`;
  }
  function keys(event: KeyboardEvent) {
    event.stopPropagation();
    if (event.target instanceof HTMLInputElement && (event.key === 'Home' || event.key === 'End')) return;
    if (event.key === 'Escape') { event.preventDefault(); close(); return; }
    if (event.key === 'Tab') { close(false); return; }
    if (!['ArrowUp', 'ArrowDown', 'Home', 'End'].includes(event.key)) return;
    event.preventDefault();
    const items = [...menu!.querySelectorAll<HTMLButtonElement>('button')];
    const index = items.indexOf(event.target as HTMLButtonElement);
    const next = event.key === 'Home' ? 0
      : event.key === 'End' ? items.length - 1
      : index < 0 ? (event.key === 'ArrowDown' ? 0 : items.length - 1)
      : (index + (event.key === 'ArrowDown' ? 1 : -1) + items.length) % items.length;
    items[next]?.focus();
  }
</script>

<svelte:window onpointerdown={event => { if (open && !menu?.contains(event.target as Node) && !trigger.contains(event.target as Node)) close(false); }} onblur={() => { if (open) close(false); }} onresize={() => { if (open) close(false); }} />
<div class="font-field" class:ribbon>
  <input class="ok-input font" bind:this={field} aria-label={label} disabled={disabled} value={value ?? ''} placeholder={t('Mixed or inherited')} onchange={event => { const font = event.currentTarget.value.trim(); if (font) choose(font); }} onkeydown={event => { if (event.altKey && event.key === 'ArrowDown') { event.preventDefault(); void show(); } }} />
  <button type="button" class="ok-input" bind:this={trigger} aria-label={t('Font options')} aria-haspopup="menu" aria-expanded={open} {disabled} onclick={show}>▾</button>
</div>
{#if open}
  <div class="font-menu" role="menu" aria-label={label} tabindex="-1" bind:this={menu} use:place onkeydown={keys}>
    <input class="font-filter" aria-label={t('Search fonts')} placeholder={t('Search fonts')} bind:value={filter} />
    {#each matches as family}
      <button type="button" role="menuitemradio" aria-checked={value === family} onclick={() => select(family)}>{family}</button>
    {/each}
    {#if !matches.length}<span class="empty">{t('No matching fonts')}</span>{/if}
  </div>
{/if}

<style>
  .font-field { display: flex; }
  input { width: 110px; border-radius: 3px 0 0 3px; }
  .font-field.ribbon input { width: 145px; }
  .font-field button { width: 16px; padding: 0; border-left: 0; border-radius: 0 3px 3px 0; }
  .font-menu { position: fixed; z-index: 400; width: 210px; max-height: min(480px, calc(100vh - 16px)); overflow-y: auto; padding: 3px; background: var(--ok-panel); border: 1px solid var(--ok-border); border-radius: 4px; box-shadow: var(--ok-shadow-lg); }
  .font-menu button { display: block; width: 100%; border: 0; padding: 3px 8px; background: transparent; color: inherit; text-align: left; font: inherit; }
  .font-menu button:hover, .font-menu button:focus-visible, .font-menu button[aria-checked='true'] { background: var(--ok-hover); outline: none; }
  .font-filter { box-sizing: border-box; width: 100%; margin-bottom: 3px; }
  .empty { display: block; padding: 4px 8px; color: var(--ok-text-2); }
</style>
