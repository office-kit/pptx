<script lang="ts">
  import { eventTarget } from '../core/dom-root.ts';
  import { tick } from 'svelte';
  import type { TextFormat } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';

  let { formats, disabled = false, onformat }: { formats: TextFormat[]; disabled?: boolean; onformat: (format: TextFormat) => void } = $props();
  const presets = [
    { key: 'Very Tight', value: -300 }, { key: 'Tight', value: -150 },
    { key: 'Normal', value: 0 }, { key: 'Loose', value: 300 }, { key: 'Very Loose', value: 600 },
  ] as const;
  let open = $state(false);
  let trigger = $state<HTMLButtonElement>();
  let menu = $state<HTMLDivElement>();
  const editor = getEditor();
  const commonSpacing = (): number | undefined => {
    if (!formats.length) return undefined;
    const values = formats.map(f => f.spc ?? 0);
    if (!values.every(value => value === values[0])) return undefined;
    return values[0];
  };
  const spacing = $derived(commonSpacing());

  function close(restore = true) { open = false; if (restore) trigger?.focus(); }
  async function show() {
    if (disabled) return;
    if (open) { close(); return; }
    open = true;
    await tick();
    const item = menu?.querySelector<HTMLButtonElement>('[aria-checked="true"]') ?? menu?.querySelector<HTMLButtonElement>('button');
    item?.focus();
  }
  function apply(value: number) { onformat({ spc: value }); close(); }
  function showMore() {
    open = false;
    editor.openFontDialog('character', trigger ?? undefined);
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

<svelte:window onpointerdown={event => { if (open && !menu?.contains(eventTarget(event) as Node) && !trigger?.contains(eventTarget(event) as Node)) close(false); }} onblur={() => { if (open) close(false); }} onresize={() => { if (open) close(false); }} />
<button class="trigger" type="button" bind:this={trigger} {disabled} aria-label={t('Character Spacing')} aria-haspopup="menu" aria-expanded={open} onclick={show}>AV ▾</button>
{#if open}
  <div class="menu" role="menu" aria-label={t('Character Spacing')} tabindex="-1" bind:this={menu} use:place onkeydown={keys}>
    <div class="heading">{t('Character Spacing')}</div>
    {#each presets as preset}
      <button type="button" role="menuitemradio" aria-checked={spacing === preset.value} onclick={() => apply(preset.value)}>{t(preset.key)}</button>
    {/each}
    <button type="button" role="menuitem" onclick={showMore}>{t('More Spacing...')}</button>
  </div>
{/if}
<style>
  .trigger { min-width: 38px; height: 26px; padding: 0 4px; border: 1px solid transparent; border-radius: var(--ok-radius); background: transparent; color: inherit; font: inherit; font-size: 11px; cursor: pointer; }
  .trigger:hover:not(:disabled), .trigger:focus-visible { background: var(--ok-hover); outline: none; }
  .trigger:disabled { opacity: .4; cursor: default; }
  .menu { position: fixed; z-index: 400; min-width: 190px; padding: 7px; border: 1px solid var(--ok-border); border-radius: 6px; background: var(--ok-panel); color: var(--ok-text); box-shadow: var(--ok-shadow-lg); }
  .heading { padding: 2px 4px 6px; color: var(--ok-text-2); font-size: 11px; }
  .menu > button { display: block; width: 100%; padding: 6px 8px; border: 0; background: transparent; color: inherit; font: inherit; font-size: 12px; text-align: left; cursor: pointer; }
  .menu > button:hover, .menu > button:focus-visible, .menu > button[aria-checked='true'] { background: var(--ok-hover); outline: none; }
</style>
