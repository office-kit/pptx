<script lang="ts">
  import { tick } from 'svelte';
  import { setParagraphLineSpacing, type ParagraphProperties } from '@office-kit/pptx';
  import { editTargetParagraphs, targetParagraphProperties, type ParagraphEdit } from '../core/paragraph-targets.ts';
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import Icon from '../ui/Icon.svelte';
  const editor = getEditor();
  const doc = editor.doc;
  let open = $state(false);
  // The Paragraph dialog belongs to the Home ribbon, which outlives a collapsed
  // Paragraph group's popup.
  let { onoptions }: { onoptions: () => void } = $props();
  let trigger: HTMLButtonElement;
  let menu = $state<HTMLDivElement>();
  const properties = $derived.by<ParagraphProperties[]>(() => { doc.version; return targetParagraphProperties(editor); });
  const enabled = $derived(properties.length > 0 && !editor.selectionLocked());
  const spacing = $derived.by(() => {
    const values = properties.map(p => !p.lineSpacing ? 1 : p.lineSpacing.kind === 'pct' ? p.lineSpacing.value : null);
    return values.every(value => value === values[0]) ? values[0] : null;
  });
  function apply(edit: ParagraphEdit) {
    if (enabled) editTargetParagraphs(editor, edit);
  }
  function close(restore = true) { open = false; if (restore) trigger.focus(); }
  async function show() { open = !open; if (open) { await tick(); (menu?.querySelector<HTMLButtonElement>('[aria-checked="true"]') ?? menu?.querySelector<HTMLButtonElement>('button'))?.focus(); } }
  function place(node: HTMLElement) {
    const bounds = trigger.getBoundingClientRect();
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
  $effect(() => { if (!enabled) open = false; });
</script>
<svelte:window onpointerdown={event => { if (open && !menu?.contains(event.target as Node) && !trigger.contains(event.target as Node)) close(false); }} onblur={() => { if (open) close(false); }} onresize={() => { if (open) close(false); }} />
<button class="trigger ok-btn" bind:this={trigger} aria-label={t('Line spacing')} title={t('Line spacing')} aria-haspopup="menu" aria-expanded={open} disabled={!enabled} onclick={show}><Icon name="line-spacing" /><span>▾</span></button>
{#if open}
  <div class="menu" role="menu" aria-label={t('Line spacing')} tabindex="-1" bind:this={menu} use:place onkeydown={keys}>
    {#each [1, 1.5, 2, 2.5, 3] as value}<button role="menuitemradio" aria-label={value.toFixed(1)} aria-checked={spacing === value} onclick={() => { apply((shape, index) => setParagraphLineSpacing(shape, index, { kind: 'pct', value })); close(); }}><span class="check">{spacing === value ? '✓' : ''}</span>{value.toFixed(1)}</button>{/each}
    <button role="menuitem" onclick={() => { close(false); onoptions(); }}>{t('Line Spacing Options...')}</button>
  </div>
{/if}
<style>
  .trigger { display: flex; align-items: center; gap: 2px; padding: 3px; }
  .menu { position: fixed; z-index: 400; min-width: 170px; padding: 4px; background: var(--ok-panel); color: var(--ok-text); border: 1px solid var(--ok-border); border-radius: 5px; box-shadow: var(--ok-shadow-lg); }
  .menu button { display: flex; gap: 6px; width: 100%; border: 0; padding: 5px 8px; background: transparent; color: inherit; text-align: left; font: inherit; }
  .check { width: 14px; }
  .menu button:hover, .menu button:focus-visible { background: var(--ok-accent); color: white; outline: none; }
</style>
