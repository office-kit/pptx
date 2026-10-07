<script lang="ts">
  import { tick } from 'svelte';
  import { getShapeKind, getShapeStrokeDash, getShapeStrokeCap, getShapeStrokeJoin, getShapeStrokeCompound, getShapeStrokeSketch, type LineSketch } from '@office-kit/pptx';
  import ArrowStyleFields from './ArrowStyleFields.svelte';
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';

  const editor = getEditor();
  const fields = [
    { label: 'Compound type', read: getShapeStrokeCompound, command: 'setShapeStrokeCompound', parameter: 'cmpd', options: [['sng', 'Single'], ['dbl', 'Double'], ['thickThin', 'Thick Thin'], ['thinThick', 'Thin Thick'], ['tri', 'Triple']] },
    { label: 'Dash type', accessibleLabel: 'Outline style', read: getShapeStrokeDash, command: 'setShapeStrokeDash', parameter: 'dash', options: [
    ['solid', 'Solid line'], ['dot', 'Dotted line'], ['dash', 'Dashed line'],
    ['lgDash', 'Long dashed line'], ['dashDot', 'Dash-dot line'],
    ['lgDashDot', 'Long dash-dot line'], ['lgDashDotDot', 'Long dash-dot-dot line'],
    ['sysDash', 'System dashed line'], ['sysDot', 'System dotted line'],
    ['sysDashDot', 'System dash-dot line'], ['sysDashDotDot', 'System dash-dot-dot line'],
  ] },
    { label: 'Cap type', read: getShapeStrokeCap, command: 'setShapeStrokeCap', parameter: 'cap', options: [['flat', 'Flat'], ['rnd', 'Round'], ['sq', 'Square']] },
    { label: 'Join type', read: getShapeStrokeJoin, command: 'setShapeStrokeJoin', parameter: 'join', options: [['round', 'Round'], ['bevel', 'Bevel'], ['miter', 'Miter']] },
  ] as const;
  // Mac PowerPoint's Sketched style menu: None, Curved, Freehand, Scribble.
  // Its Japanese build labels both Freehand and Scribble フリーハンド.
  const sketches = [[null, 'None'], ['curved', 'Curved'], ['freehand', 'Freehand'], ['scribble', 'Scribble']] as const;
  // PowerPoint disables Sketched style for connectors.
  const sketchDisabled = $derived.by(() => {
    editor.doc.version;
    const shapes = editor.selectedShapes();
    return !shapes.length || editor.selectionLocked() || shapes.some(shape => getShapeKind(shape) !== 'shape');
  });
  const sketch = $derived.by((): LineSketch | null | 'mixed' => {
    editor.doc.version;
    const values = new Set(editor.selectedShapes().map(getShapeStrokeSketch));
    return values.size === 1 ? [...values][0]! : 'mixed';
  });
  let sketchOpen = $state(false);
  let sketchTrigger = $state<HTMLButtonElement>();
  let sketchMenu = $state<HTMLDivElement>();
  $effect(() => { editor.doc.selection; editor.doc.version; sketchOpen = false; });
  function closeSketch(restore = true) { sketchOpen = false; if (restore) sketchTrigger?.focus(); }
  async function toggleSketch() {
    if (sketchOpen) { closeSketch(); return; }
    sketchOpen = true;
    await tick();
    (sketchMenu?.querySelector<HTMLButtonElement>('[aria-checked="true"]') ?? sketchMenu?.querySelector<HTMLButtonElement>('button'))?.focus();
  }
  function chooseSketch(value: LineSketch | null) {
    if (!sketchDisabled && value !== sketch) editor.invoke('setShapeStrokeSketch', { sketch: value });
    closeSketch();
  }
  function placeSketch(node: HTMLElement) {
    const bounds = sketchTrigger!.getBoundingClientRect();
    node.style.left = `${Math.max(8, Math.min(bounds.left, innerWidth - node.offsetWidth - 8))}px`;
    node.style.top = `${Math.max(8, Math.min(bounds.bottom, innerHeight - node.offsetHeight - 8))}px`;
  }
  function sketchKeys(event: KeyboardEvent) {
    event.stopPropagation();
    if (event.key === 'Escape') { event.preventDefault(); closeSketch(); return; }
    if (event.key === 'Tab') { closeSketch(false); return; }
    const offsets: Record<string, number> = { ArrowUp: -1, ArrowDown: 1 };
    if (!(event.key in offsets) && event.key !== 'Home' && event.key !== 'End') return;
    event.preventDefault();
    const items = [...sketchMenu!.querySelectorAll<HTMLButtonElement>('button')];
    const index = items.indexOf(event.target as HTMLButtonElement);
    items[event.key === 'Home' ? 0 : event.key === 'End' ? items.length - 1 : (index + offsets[event.key]! + items.length) % items.length]?.focus();
  }
  const values = $derived.by(() => {
    editor.doc.version;
    const shapes = editor.selectedShapes();
    return fields.map(field => {
      const values = new Set(shapes.map(shape => field.read(shape) ?? 'inherit'));
      return values.size === 1 ? [...values][0]! : 'mixed';
    });
  });
</script>

{#snippet sketchPreview(value: LineSketch | null)}
  <svg viewBox="0 0 72 24" width="72" height="24" aria-hidden="true"><path d={value === 'curved' ? 'M2 20Q14 6 70 6' : value === 'freehand' ? 'M2 14C14 10 22 18 36 14S58 10 70 13' : value === 'scribble' ? 'M2 12C8 4 12 20 18 12S28 4 34 12S44 20 50 12S60 4 70 12' : 'M2 12H70'} fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" /></svg>
{/snippet}
<svelte:window onpointerdown={event => { if (sketchOpen && !sketchMenu?.contains(event.target as Node) && !sketchTrigger?.contains(event.target as Node)) closeSketch(false); }} onblur={() => { if (sketchOpen) closeSketch(false); }} onresize={() => { if (sketchOpen) closeSketch(false); }} />
<!-- PowerPoint's Sketched style sits between Width and Compound type. -->
<div class="sketch">
  <span>{t('Sketched style')}</span>
  <button bind:this={sketchTrigger} class="ok-input menu-button" aria-label={t('Sketched style')} aria-haspopup="menu" aria-expanded={sketchOpen} disabled={sketchDisabled} onclick={toggleSketch}>
    {#if sketch === 'mixed'}<span title={t('Mixed')}>–</span>{:else}{@render sketchPreview(sketch)}{/if}<span class="arrow">▾</span>
  </button>
</div>
{#if sketchOpen}
  <div class="sketch-menu" role="menu" aria-label={t('Sketched style')} tabindex="-1" bind:this={sketchMenu} use:placeSketch onkeydown={sketchKeys}>
    {#each sketches as [value, label]}
      <button role="menuitemradio" aria-label={t(label)} title={t(label)} aria-checked={sketch === value} onclick={() => chooseSketch(value)}><span class="check" aria-hidden="true">{sketch === value ? '✓' : ''}</span>{@render sketchPreview(value)}</button>
    {/each}
  </div>
{/if}
{#each fields as field, index}
  <label>
    <span>{t(field.label)}</span>
    <select class="ok-input" aria-label={t('accessibleLabel' in field ? field.accessibleLabel : field.label)} value={values[index]}
      onchange={event => editor.invoke(field.command, { [field.parameter]: event.currentTarget.value })}>
      {#if values[index] === 'inherit' || values[index] === 'mixed'}
        <option value={values[index]} disabled>{t(values[index] === 'mixed' ? 'Mixed' : 'Inherited')}</option>
      {/if}
      {#each field.options as [value, label]}<option {value}>{t(label)}</option>{/each}
    </select>
  </label>
{/each}

<ArrowStyleFields />

<style>
  .sketch { display: flex; align-items: center; justify-content: space-between; gap: 8px; min-height: 26px; }
  /* PowerPoint's 39 × 26 pt gallery buttons. */
  .menu-button { display: flex; align-items: center; justify-content: space-between; box-sizing: border-box; width: 39px; height: 26px; padding: 0 2px 0 4px; }
  .menu-button :global(svg) { width: 24px; height: 10px; flex: none; }
  /* Native rows: 139 × 29 pt with the check mark at the left. */
  .sketch-menu { position: fixed; z-index: 400; display: flex; flex-direction: column; width: 139px; padding: 4px 0; background: var(--ok-panel); border: 1px solid var(--ok-border); border-radius: 6px; box-shadow: var(--ok-shadow-lg); }
  .sketch-menu button { display: flex; align-items: center; gap: 4px; height: 29px; padding: 0 8px; border: 0; background: transparent; color: var(--ok-text); }
  .sketch-menu button:hover, .sketch-menu button:focus-visible { background: var(--ok-hover); }
  .sketch-menu .check { width: 14px; }
  .menu-button .arrow { font-size: 9px; }
  label { display: flex; align-items: center; justify-content: space-between; gap: 8px; min-height: 26px; }
  /* PowerPoint's Format pane pop-up buttons are 112 × 26 pt. */
  select { box-sizing: border-box; width: 112px; height: 26px; font-size: inherit; }
</style>
