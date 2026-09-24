<script lang="ts">
  import { tick } from 'svelte';
  import { getShapeKind, getShapePreset, getShapeStrokeArrow, setShapeStrokeArrow } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';

  const editor = getEditor();
  type Arrow = Parameters<typeof setShapeStrokeArrow>[2];
  const types = [['none', 'No Arrow'], ['triangle', 'Arrow'], ['arrow', 'Open Arrow'], ['stealth', 'Stealth Arrow'], ['diamond', 'Diamond Arrow'], ['oval', 'Oval Arrow']] as const;
  const dimensions = ['sm', 'med', 'lg'] as const;
  // Mac PowerPoint orders the gallery by width, then length (Size 2 = sm/med).
  const sizes = dimensions.flatMap(width => dimensions.map(length => ({ width, length })));
  const fields = [
    { end: 'head', property: 'type', label: 'Begin Arrow type' },
    { end: 'head', property: 'size', label: 'Begin Arrow size' },
    { end: 'tail', property: 'type', label: 'End Arrow type' },
    { end: 'tail', property: 'size', label: 'End Arrow size' },
  ] as const;
  type Field = typeof fields[number];
  const shapes = $derived.by(() => { editor.doc.version; return editor.selectedShapes(); });
  const disabled = $derived(!shapes.length || editor.selectionLocked() || shapes.some(shape => getShapeKind(shape) !== 'connector' && getShapePreset(shape) !== 'line'));
  let active = $state<Field | null>(null);
  let trigger: HTMLButtonElement;
  let menu = $state<HTMLDivElement>();
  function read(field: Field) {
    const values = shapes.map(shape => {
      const arrow = getShapeStrokeArrow(shape, field.end);
      return field.property === 'type' ? arrow?.type ?? 'none' : `${arrow?.width ?? 'med'}/${arrow?.length ?? 'med'}`;
    });
    return values.every(value => value === values[0]) ? values[0] : 'mixed';
  }
  function arrow(field: Field): Arrow {
    const value = read(field);
    if (field.property === 'type') return { type: types.find(([type]) => type === value)?.[0] ?? 'none' };
    return { type: getShapeStrokeArrow(shapes[0]!, field.end)?.type ?? 'triangle', ...sizes.find(size => `${size.width}/${size.length}` === value) };
  }
  function close(restore = true) { active = null; if (restore) trigger.focus(); }
  async function show(field: Field, button: HTMLButtonElement) {
    if (active === field) { close(); return; }
    trigger = button; active = field;
    await tick();
    (menu?.querySelector<HTMLButtonElement>('[aria-checked="true"]') ?? menu?.querySelector<HTMLButtonElement>('button'))?.focus();
  }
  function choose(field: Field, update: Partial<Arrow>) {
    if (disabled) { close(); return; }
    editor.doc.transact(t(field.label), () => {
      for (const shape of shapes) setShapeStrokeArrow(shape, field.end, { type: 'none', ...getShapeStrokeArrow(shape, field.end), ...update });
    });
    close();
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
    const offsets: Record<string, number> = { ArrowLeft: -1, ArrowRight: 1, ArrowUp: -3, ArrowDown: 3 };
    if (!(event.key in offsets) && event.key !== 'Home' && event.key !== 'End') return;
    event.preventDefault();
    const items = [...menu!.querySelectorAll<HTMLButtonElement>('button')];
    const index = items.indexOf(event.target as HTMLButtonElement);
    items[event.key === 'Home' ? 0 : event.key === 'End' ? items.length - 1 : (index + offsets[event.key]! + items.length) % items.length]?.focus();
  }
  $effect(() => { editor.doc.selection; editor.doc.version; active = null; });
</script>

{#snippet preview(value: Arrow, end: 'head' | 'tail')}
  {@const width = value.width === 'sm' ? 4 : value.width === 'lg' ? 9 : 6}
  {@const length = value.length === 'sm' ? 7 : value.length === 'lg' ? 17 : 12}
  <svg viewBox="0 0 60 24" width="60" height="24" aria-hidden="true">
    <g transform={end === 'head' ? 'translate(60 0) scale(-1 1)' : undefined}>
      <path d="M5 12H51" fill="none" stroke="currentColor" stroke-width="1.5" />
      {#if value.type === 'oval'}<ellipse cx={52-length/2} cy="12" rx={length/2} ry={width} fill="currentColor" />
      {:else if value.type === 'diamond'}<path d={`M52 12L${52-length/2} ${12-width}L${52-length} 12L${52-length/2} ${12+width}Z`} fill="currentColor" />
      {:else if value.type !== 'none'}<path d={`M${52-length} ${12-width}L52 12L${52-length} ${12+width}${value.type === 'stealth' ? `L${52-length*.7} 12Z` : value.type === 'arrow' ? '' : 'Z'}`} fill={value.type === 'arrow' ? 'none' : 'currentColor'} stroke="currentColor" stroke-width="1.5" />{/if}
    </g>
  </svg>
{/snippet}
<svelte:window onpointerdown={event => { if (active && !menu?.contains(event.target as Node) && !trigger.contains(event.target as Node)) close(false); }} onblur={() => { if (active) close(false); }} onresize={() => { if (active) close(false); }} />
{#each fields as field}
  <div class="field"><span>{t(field.label)}</span><button class="ok-input trigger" aria-label={t(field.label)} aria-haspopup="menu" aria-expanded={active === field} {disabled} onclick={event => show(field, event.currentTarget)}>
    {#if read(field) === 'mixed'}<span>{t('Mixed')}</span>{:else if shapes.length}{@render preview(arrow(field), field.end)}{/if}<span>▾</span>
  </button></div>
{/each}
{#if active}
  {@const field = active}
  <div class="gallery" role="menu" aria-label={t(field.label)} tabindex="-1" bind:this={menu} use:place onkeydown={keys}>
    {#if field.property === 'type'}
      {#each types as [type, label]}<button role="menuitemradio" aria-label={t(label)} title={t(label)} aria-checked={read(field) === type} onclick={() => choose(field, { type })}>{@render preview({ type }, field.end)}</button>{/each}
    {:else}
      {#each sizes as size, index}<button role="menuitemradio" aria-label={`${t('Arrow Size')} ${index+1}`} title={`${t('Arrow Size')} ${index+1}`} aria-checked={read(field) === `${size.width}/${size.length}`} onclick={() => choose(field, size)}>{@render preview({ ...arrow(field), ...size }, field.end)}</button>{/each}
    {/if}
  </div>
{/if}
<style>
  .field { display: flex; align-items: center; justify-content: space-between; gap: 8px; }
  .trigger { display: flex; align-items: center; justify-content: space-between; width: 86px; height: 26px; padding: 0 4px; font: inherit; }
  .gallery { position: fixed; z-index: 400; display: grid; grid-template-columns: repeat(3, 68px); gap: 3px; padding: 6px; background: var(--ok-panel); border: 1px solid var(--ok-border); border-radius: 6px; box-shadow: var(--ok-shadow-lg); }
  .gallery button { border: 1px solid transparent; border-radius: 3px; background: transparent; color: var(--ok-text); padding: 4px; }
  .gallery button:hover, .gallery button:focus-visible, .gallery button[aria-checked=true] { background: var(--ok-hover); border-color: var(--ok-accent); }
</style>
