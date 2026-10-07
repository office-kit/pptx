<script lang="ts">
  import { tick } from 'svelte';
  import type { TextCase } from '@office-kit/pptx';
  import { t } from '../i18n/i18n.svelte.ts';

  let { disabled = false, onchange }: { disabled?: boolean; onchange: (value: TextCase) => void } = $props();
  const choices: { value: TextCase; label: string }[] = [
    { value: 'sentence', label: 'Sentence case.' },
    { value: 'lower', label: 'lowercase' },
    { value: 'upper', label: 'UPPERCASE' },
    { value: 'title', label: 'Capitalize Each Word' },
    { value: 'toggle', label: 'tOGGLE cASE' },
  ];
  let open = $state(false);
  let trigger = $state<HTMLButtonElement>();
  let menu = $state<HTMLDivElement>();
  function close(restore = true) { open = false; if (restore) trigger?.focus(); }
  async function show() {
    if (open) { close(); return; }
    open = true;
    await tick();
    menu?.querySelector('button')?.focus();
  }
  function place(node: HTMLElement) {
    const bounds = trigger!.getBoundingClientRect();
    node.style.left = `${Math.max(8, Math.min(bounds.left, innerWidth - node.offsetWidth - 8))}px`;
    node.style.top = `${Math.min(bounds.bottom, innerHeight - node.offsetHeight - 8)}px`;
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
  $effect(() => { if (disabled && open) close(false); });
</script>

<svelte:window onpointerdown={event => { if (open && !menu?.contains(event.target as Node) && !trigger?.contains(event.target as Node)) close(false); }} onblur={() => { if (open) close(false); }} onresize={() => { if (open) close(false); }} />
<button class="trigger" type="button" bind:this={trigger} {disabled} aria-label={t('Change Case')} title={t('Change Case')} aria-haspopup="menu" aria-expanded={open} onmousedown={event => event.preventDefault()} onclick={show}>Aa ▾</button>
{#if open}
  <div class="menu" role="menu" aria-label={t('Change Case')} tabindex="-1" bind:this={menu} use:place onkeydown={keys}>
    {#each choices as choice}
      <button type="button" role="menuitem" onclick={() => { close(false); onchange(choice.value); }}>{t(choice.label)}</button>
    {/each}
  </div>
{/if}

<style>
  .trigger { min-width: 38px; height: 26px; padding: 0 4px; border: 1px solid transparent; border-radius: var(--ok-radius); background: transparent; color: inherit; font: inherit; font-size: 11px; cursor: pointer; }
  .trigger:hover:not(:disabled), .trigger:focus-visible { background: var(--ok-hover); outline: none; }
  .trigger:disabled { opacity: .4; cursor: default; }
  .menu { position: fixed; z-index: 400; min-width: 200px; padding: 7px; border: 1px solid var(--ok-border); border-radius: 6px; background: var(--ok-panel); color: var(--ok-text); box-shadow: var(--ok-shadow-lg); }
  .menu > button { display: block; width: 100%; padding: 6px 8px; border: 0; background: transparent; color: inherit; font: inherit; font-size: 12px; text-align: left; cursor: pointer; }
  .menu > button:hover, .menu > button:focus-visible { background: var(--ok-hover); outline: none; }
</style>
