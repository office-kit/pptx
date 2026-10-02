<script lang="ts">
  import { tick } from 'svelte';
  import type { TextFormat } from '@office-kit/pptx';
  import { t } from '../i18n/i18n.svelte.ts';

  let { formats, disabled = false, onformat }: { formats: TextFormat[]; disabled?: boolean; onformat: (format: TextFormat) => void } = $props();
  const presets = [
    { key: 'Very Tight', value: -300 }, { key: 'Tight', value: -150 },
    { key: 'Normal', value: 0 }, { key: 'Loose', value: 300 }, { key: 'Very Loose', value: 600 },
  ] as const;
  let open = $state(false);
  let trigger = $state<HTMLButtonElement>();
  let menu = $state<HTMLDivElement>();
  let dialog = $state<HTMLDialogElement>();
  let mode = $state<'' | 'normal' | 'expanded' | 'condensed'>('expanded');
  let amount = $state<number | undefined>(0);
  let spacingTouched = $state(false);
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
  async function showMore() {
    open = false;
    const current = spacing;
    mode = current === undefined ? '' : current === 0 ? 'normal' : current < 0 ? 'condensed' : 'expanded';
    amount = current === undefined ? undefined : Math.abs(current) / 100;
    spacingTouched = false;
    await tick();
    dialog?.showModal();
    dialog?.querySelector<HTMLElement>('select, input, button')?.focus();
  }
  function cancelMore() { dialog?.close(); trigger?.focus(); }
  function applyMore(event: SubmitEvent) {
    event.preventDefault();
    if (!spacingTouched || !mode) { cancelMore(); return; }
    if (mode === 'normal') onformat({ spc: 0 });
    else {
      if (amount === undefined || !Number.isFinite(amount) || amount < 0 || amount > 1000) return;
      onformat({ spc: Math.round((mode === 'condensed' ? -amount : amount) * 100) });
    }
    dialog?.close();
    trigger?.focus();
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
  $effect(() => { if (disabled) { if (open) close(false); if (dialog?.open) dialog.close(); } });
</script>

<svelte:window onpointerdown={event => { if (open && !menu?.contains(event.target as Node) && !trigger?.contains(event.target as Node)) close(false); }} onblur={() => { if (open) close(false); }} onresize={() => { if (open) close(false); }} />
<button class="trigger" type="button" bind:this={trigger} {disabled} aria-label={t('Character Spacing')} aria-haspopup="menu" aria-expanded={open} onclick={show}>A↔V ▾</button>
{#if open}
  <div class="menu" role="menu" aria-label={t('Character Spacing')} tabindex="-1" bind:this={menu} use:place onkeydown={keys}>
    <div class="heading">{t('Character Spacing')}</div>
    {#each presets as preset}
      <button type="button" role="menuitemradio" aria-checked={spacing === preset.value} onclick={() => apply(preset.value)}>{t(preset.key)}</button>
    {/each}
    <button type="button" role="menuitem" onclick={showMore}>{t('More Spacing...')}</button>
  </div>
{/if}
<dialog bind:this={dialog} aria-label={t('Character Spacing')} title={t('Character Spacing')} onclose={() => trigger?.focus()} onkeydown={event => event.stopPropagation()}>
  <form onsubmit={applyMore}>
    <h2>{t('Character Spacing')}</h2>
    <label>{t('Character spacing mode')}<select bind:value={mode} onchange={() => { spacingTouched = true; if (amount === undefined) amount = 0; }}><option value="" disabled>{t('Mixed')}</option><option value="normal">{t('Normal')}</option><option value="expanded">{t('Expanded')}</option><option value="condensed">{t('Condensed')}</option></select></label>
    <label>{t('By (pt)')}<input type="number" min="0" max="1000" step="0.1" bind:value={amount} disabled={!mode || mode === 'normal'} oninput={() => spacingTouched = true} /></label>
    <footer><button type="button" onclick={cancelMore}>{t('Cancel')}</button><button type="submit">{t('OK')}</button></footer>
  </form>
</dialog>

<style>
  .trigger { min-width: 0; padding: 3px 5px; border: 1px solid var(--ok-border); border-radius: var(--ok-radius); background: transparent; color: inherit; font: inherit; font-size: 11px; cursor: pointer; }
  .trigger:hover:not(:disabled), .trigger:focus-visible { background: var(--ok-hover); outline: none; }
  .trigger:disabled { opacity: .4; cursor: default; }
  .menu { position: fixed; z-index: 400; min-width: 190px; padding: 7px; border: 1px solid var(--ok-border); border-radius: 6px; background: var(--ok-panel); color: var(--ok-text); box-shadow: var(--ok-shadow-lg); }
  .heading { padding: 2px 4px 6px; color: var(--ok-text-2); font-size: 11px; }
  .menu > button { display: block; width: 100%; padding: 6px 8px; border: 0; background: transparent; color: inherit; font: inherit; font-size: 12px; text-align: left; cursor: pointer; }
  .menu > button:hover, .menu > button:focus-visible, .menu > button[aria-checked='true'] { background: var(--ok-hover); outline: none; }
  dialog { width: 320px; border: 1px solid var(--ok-border); border-radius: 7px; background: var(--ok-panel); color: var(--ok-text); padding: 0; box-shadow: var(--ok-shadow-lg); }
  dialog::backdrop { background: #0003; }
  h2 { font-size: 13px; margin: 0; padding: 8px 12px; border-bottom: 1px solid var(--ok-border); }
  label { display: flex; align-items: center; justify-content: space-between; gap: 10px; margin: 12px 16px; font-size: 11px; }
  select, input[type='number'] { width: 110px; padding: 3px; border: 1px solid var(--ok-border); border-radius: 3px; background: var(--ok-input); color: inherit; font: inherit; }
  footer { display: flex; justify-content: flex-end; gap: 7px; padding: 0 16px 14px; }
  footer button { min-width: 65px; padding: 4px 8px; border: 1px solid var(--ok-border); border-radius: 4px; background: transparent; color: inherit; font: inherit; cursor: pointer; }
  footer button[type='submit'] { background: var(--ok-accent); border-color: var(--ok-accent); color: white; }
</style>
