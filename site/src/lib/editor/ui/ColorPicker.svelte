<script lang="ts">
  import { tick } from 'svelte';
  import { asColor, getPresentationTheme, type Color } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';

  let { label, value, resolvedColor, disabled = false, choose }: { label: string; value?: string; resolvedColor?: string; disabled?: boolean; choose: (color: Color) => void } = $props();
  const editor = getEditor();
  const theme = $derived.by(() => { editor.doc.version; return getPresentationTheme(editor.doc.pres); });
  const themeSlots = [
    ['bg1', 'light1', 'Background 1'], ['tx1', 'dark1', 'Text 1'],
    ['bg2', 'light2', 'Background 2'], ['tx2', 'dark2', 'Text 2'],
    ['accent1', 'accent1', 'Accent 1'], ['accent2', 'accent2', 'Accent 2'],
    ['accent3', 'accent3', 'Accent 3'], ['accent4', 'accent4', 'Accent 4'],
    ['accent5', 'accent5', 'Accent 5'], ['accent6', 'accent6', 'Accent 6'],
  ] as const;
  const standard = [
    ['#C00000', 'Dark Red'], ['#FF0000', 'Red'], ['#FFC000', 'Orange'],
    ['#FFFF00', 'Yellow'], ['#92D050', 'Light Green'], ['#00B050', 'Green'],
    ['#00B0F0', 'Light Blue'], ['#0070C0', 'Blue'], ['#002060', 'Dark Blue'], ['#7030A0', 'Purple'],
  ] as const;
  const colors = $derived([
    ...(theme ? themeSlots.filter(([, slot]) => theme![slot]).map(([color, slot, name]) => ({ color, paint: theme![slot], name, theme: true })) : []),
    ...standard.map(([color, name]) => ({ color, paint: color, name, theme: false })),
  ]);
  let open = $state(false);
  let trigger: HTMLButtonElement;
  let custom: HTMLInputElement;
  let menu = $state<HTMLDivElement>();
  function close(restore = true) { open = false; if (restore) trigger.focus(); }
  function select(color: string) {
    const parsed = asColor(color);
    if (parsed && !disabled && !trigger.matches(':disabled')) choose(parsed);
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
    const offsets: Record<string, number> = { ArrowLeft: -1, ArrowRight: 1, ArrowUp: -10, ArrowDown: 10 };
    if (!(event.key in offsets) && event.key !== 'Home' && event.key !== 'End') return;
    event.preventDefault();
    const items = [...menu!.querySelectorAll<HTMLButtonElement>('button')];
    const index = items.indexOf(event.target as HTMLButtonElement);
    items[event.key === 'Home' ? 0 : event.key === 'End' ? items.length - 1 : (index + offsets[event.key]! + items.length) % items.length]?.focus();
  }
  $effect(() => { value; disabled; editor.doc.selection; open = false; });
</script>

<svelte:window onpointerdown={event => { if (open && !menu?.contains(event.target as Node) && !trigger.contains(event.target as Node)) close(false); }} onblur={() => { if (open) close(false); }} onresize={() => { if (open) close(false); }} />
<button type="button" class="ok-input trigger" bind:this={trigger} aria-label={label} aria-haspopup="menu" aria-expanded={open} {disabled} onclick={show}><span class="swatch" style:background={resolvedColor ?? value ?? 'transparent'}></span><span>▾</span></button>
<input class="custom" type="color" bind:this={custom} aria-label={`${label}: ${t('More Colors...')}`} tabindex="-1" {disabled} value={resolvedColor ?? '#000000'} onchange={event => select(event.currentTarget.value)} />
{#if open}
  <div class="palette" role="menu" aria-label={label} tabindex="-1" bind:this={menu} use:place onkeydown={keys}>
    {#each [true, false] as isTheme}
      {#if !isTheme || theme}
        <div class="heading">{t(isTheme ? 'Theme Colors' : 'Standard Colors')}</div>
        <div class="colors" role="group" aria-label={t(isTheme ? 'Theme Colors' : 'Standard Colors')}>
          {#each colors.filter(color => color.theme === isTheme) as color}
            <button type="button" role="menuitemradio" aria-label={t(color.name)} title={t(color.name)} aria-checked={value?.replace(/^scheme:/, '').toLowerCase() === color.color.toLowerCase()} style:background={color.paint} onclick={() => select(color.color)}></button>
          {/each}
        </div>
      {/if}
    {/each}
    <button type="button" class="more" role="menuitem" onclick={() => { close(); custom.click(); }}>{t('More Colors...')}</button>
  </div>
{/if}

<style>
  .trigger { display: inline-flex; align-items: center; justify-content: space-between; gap: 6px; width: 54px; padding: 3px 5px; }
  .swatch { display: block; width: 24px; height: 18px; border: 1px solid var(--ok-border); }
  .custom { position: fixed; opacity: 0; pointer-events: none; width: 1px; height: 1px; }
  .palette { position: fixed; z-index: 400; padding: 6px; background: var(--ok-panel); border: 1px solid var(--ok-border); border-radius: 6px; box-shadow: var(--ok-shadow-lg); }
  .heading { font-size: 11px; margin: 3px 2px 6px; }
  .colors { display: grid; grid-template-columns: repeat(10, 18px); gap: 3px; margin-bottom: 10px; }
  .colors button { width: 18px; height: 18px; padding: 0; border: 1px solid var(--ok-border); }
  .colors button:hover, .colors button:focus-visible, .colors button[aria-checked=true] { outline: 2px solid var(--ok-accent); outline-offset: 1px; }
  .more { display: block; width: 100%; text-align: left; border: 0; border-top: 1px solid var(--ok-border); background: transparent; color: inherit; padding: 6px 2px 2px; font-size: inherit; }
  .more:hover, .more:focus-visible { background: var(--ok-hover); }
</style>
