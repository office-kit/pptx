<script lang="ts">
  // Mac PowerPoint 16's Draw tab: Draw, Eraser ▾ and Lasso Select; the pen
  // gallery with Add ▾; Ink to Text / Shape / Math; Draw with Trackpad. It
  // has no collapsed layout: at 1200 pt only the gallery narrows.
  import { tick } from 'svelte';
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import type { PenKind } from '../core/ink.svelte.ts';
  import { captionLines } from './caption.ts';
  import { placeBelowTrigger } from './place-menu.ts';

  const editor = getEditor();
  const ink = editor.ink;
  let colorInput = $state<HTMLInputElement>();
  let open = $state<'eraser' | 'add' | null>(null);
  let adding = $state<PenKind>('pen');
  let root = $state<HTMLDivElement>();

  async function toggle(menu: 'eraser' | 'add') {
    open = open === menu ? null : menu;
    if (open) {
      await tick();
      root?.querySelector<HTMLElement>('.menu button:not(:disabled)')?.focus();
    }
  }
  function add(kind: PenKind) {
    open = null;
    adding = kind;
    colorInput?.click();
  }
</script>

<svelte:window
  onpointerdown={(event) => { if (open && !(event.target as Element).closest?.('.draw .anchor')) open = null; }}
  onkeydown={(event) => { if (open && event.key === 'Escape') { event.preventDefault(); event.stopPropagation(); open = null; } }}
/>

<div class="draw" bind:this={root}>
  <section class="cluster" aria-label={t('Tools')}>
    <button class="big" aria-pressed={ink.tool === 'pen'} aria-label={t('Draw')} title={t('Enter Draw mode')} onclick={() => ink.toggle('pen')}>
      <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M5 19 16 4l3 2L8 21l-4 1z" /><path d="M3 22c3-2 6 0 9-2" stroke="#c00000" /></svg>
      <span class="caption">{t('Draw')}</span>
    </button>
    <div class="anchor overlay">
      <button class="big menu-button" aria-pressed={ink.tool === 'eraser'} aria-label={t('Eraser')} onclick={() => ink.toggle('eraser')}>
        <span class="icon-row"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="m4 15 9-10 7 6-9 10H8z" fill="#b48ad8" /><path d="M8 21h12" /></svg><span class="arrow-space" aria-hidden="true"></span></span>
        <span class="caption">{t('Eraser')}</span>
      </button>
      <button class="side" aria-label={t('Eraser options')} aria-haspopup="menu" aria-expanded={open === 'eraser'} onclick={() => toggle('eraser')}>⌄</button>
      {#if open === 'eraser'}
        <div class="menu" role="menu" aria-label={t('Eraser')} use:placeBelowTrigger>
          <!-- The editor's eraser removes each stroke it touches: PowerPoint's Stroke Eraser. -->
          <button role="menuitemradio" aria-checked="true" onclick={() => { open = null; ink.tool = 'eraser'; }}>{t('Stroke Eraser')}</button>
          <button role="menuitemradio" aria-checked="false" disabled title={t('Erasing part of a stroke is not available in this editor yet.')}>{t('Point Eraser')}</button>
          <button role="menuitemradio" aria-checked="false" disabled title={t('Erasing part of a stroke is not available in this editor yet.')}>{t('Segment Eraser')}</button>
        </div>
      {/if}
    </div>
    <button class="big" aria-pressed={ink.tool === 'lasso'} aria-label={t('Lasso Select')} onclick={() => ink.toggle('lasso')}>
      <svg viewBox="0 0 24 24" aria-hidden="true"><ellipse cx="11" cy="10" rx="8" ry="6" stroke-dasharray="2 2" /><path d="M18 17v6M15 20h6" stroke="#2e8b57" /></svg>
      <span class="caption">{captionLines(t('Lasso Select'))}</span>
    </button>
  </section>

  <section class="cluster pens-group" aria-label={t('Pens')}>
    <div class="pens" role="radiogroup" aria-label={t('Pens')}>
      {#each ink.pens as pen (pen.id)}
        <button class="pen" role="radio" aria-checked={ink.tool === 'pen' && ink.pen.id === pen.id} aria-label={t(pen.label)} title={t(pen.label)} onclick={() => ink.choosePen(pen)}>
          <svg viewBox="0 0 20 44" aria-hidden="true">
            <path d="M5 2h10v26l-5 12-5-12z" fill={pen.color} fill-opacity={pen.opacity} />
            <path d="M5 28h10" />
          </svg>
        </button>
      {/each}
    </div>
    <div class="anchor">
      <button class="big menu-button" aria-label={t('Add')} aria-haspopup="menu" aria-expanded={open === 'add'} onclick={() => toggle('add')}>
        <span class="icon-row"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 4v16M4 12h16" stroke="#2e8b57" stroke-width="1.6" /></svg><span class="arrow" aria-hidden="true">⌄</span></span>
        <span class="caption">{t('Add')}</span>
      </button>
      {#if open === 'add'}
        <div class="menu" role="menu" aria-label={t('Add')} use:placeBelowTrigger>
          <button role="menuitem" onclick={() => add('pen')}>{t('Add Pen')}</button>
          <button role="menuitem" onclick={() => add('pencil')}>{t('Add Pencil')}</button>
          <button role="menuitem" onclick={() => add('highlighter')}>{t('Add Highlighter')}</button>
        </div>
      {/if}
    </div>
    <input
      bind:this={colorInput}
      class="color"
      type="color"
      tabindex="-1"
      aria-label={t('Pen Color')}
      onchange={(event) => ink.addPen(event.currentTarget.value as `#${string}`, adding)}
    />
  </section>

  <section class="cluster" aria-label={t('Convert')}>
    <button class="big" aria-label={t('Ink to Text')} disabled title={t('Handwriting recognition is not available in the browser.')}>
      <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M4 9c2-6 6-6 6 0s-4 3-4 0M14 20a3 3 0 1 1 6 0v-6" /></svg>
      <span class="caption">{captionLines(t('Ink to Text'))}</span>
    </button>
    <button class="big" aria-label={t('Ink to Shape')} aria-pressed={ink.toShape} onclick={() => { ink.toShape = !ink.toShape; if (ink.toShape) ink.tool = 'pen'; }}>
      <svg viewBox="0 0 24 24" aria-hidden="true"><rect x="6" y="8" width="14" height="12" /><path d="M3 4c3 2 6-2 9 0" /></svg>
      <span class="caption">{captionLines(t('Ink to Shape'))}</span>
    </button>
    <button class="big" aria-label={t('Ink to Math')} disabled title={t('Handwriting recognition is not available in the browser.')}>
      <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M5 6h12M8 6v12M14 6v10c0 2 2 2 3 1" /></svg>
      <span class="caption">{captionLines(t('Ink to Math'))}</span>
    </button>
  </section>

  <section class="cluster" aria-label={t('Trackpad')}>
    <label class="switch">
      <input type="checkbox" role="switch" checked={ink.tool === 'pen'} onchange={() => ink.toggle('pen')} />
      <span class="caption">{captionLines(t('Draw with Trackpad'))}</span>
    </label>
  </section>
</div>

<style>
  /* The Home tab's metrics (see HomeRibbon.svelte). The pen gallery is a
     180 × 60 pt box 18 pt inside its group; Draw with Trackpad is 62 pt wide. */
  .draw { display: flex; align-items: stretch; min-width: 0; width: 100%; height: 72px; }
  .cluster { position: relative; display: flex; align-items: center; padding: 0 9px; border-right: 1px solid var(--ok-border); flex: none; }
  .cluster:first-child { padding-left: 3px; }
  .cluster:last-child { border-right: none; }
  .pens-group { padding-left: 27px; }
  button { font: inherit; color: var(--ok-text); background: none; border: 1px solid transparent; border-radius: var(--ok-radius); cursor: pointer; }
  button:hover:not(:disabled) { background: var(--ok-hover); }
  button:disabled { opacity: 0.4; cursor: default; }
  button[aria-pressed='true'], button[aria-checked='true'] { background: var(--ok-selected); border-color: var(--ok-selected-border); }
  .big { display: flex; flex-direction: column; align-items: center; justify-content: flex-start; align-self: stretch; gap: 2px; min-width: 38px; padding: 4px 1px; font-size: 11px; line-height: 1.15; }
  .big.menu-button { min-width: 50px; }
  .caption { white-space: pre-line; text-align: center; }
  .icon-row { display: flex; align-items: center; gap: 1px; }
  .arrow { font-size: 10px; }
  .anchor { position: relative; display: flex; align-self: stretch; }
  /* The ▾ sits over the space beside the icon, as in PowerPoint's 50 pt menu buttons. */
  .arrow-space { width: 9px; }
  .overlay > .side { position: absolute; top: 12px; right: 2px; width: 12px; height: 22px; padding: 0; font-size: 12px; }
  svg { width: 32px; height: 32px; flex-shrink: 0; stroke: currentColor; fill: none; stroke-width: 1.1; }
  .pens { display: flex; align-items: center; gap: 4px; box-sizing: border-box; width: 180px; height: 60px; margin-right: 19px; padding: 0 8px; overflow-x: auto; border: 1px solid var(--ok-border); border-radius: var(--ok-radius); background: var(--ok-panel); }
  .pen { display: flex; align-items: center; justify-content: center; flex: none; width: 36px; height: 52px; padding: 0; }
  .pen svg { width: 20px; height: 44px; }
  .menu { position: fixed; z-index: 400; display: flex; flex-direction: column; min-width: 180px; padding: 4px; border: 1px solid var(--ok-border); border-radius: 6px; background: var(--ok-panel); box-shadow: var(--ok-shadow-lg); }
  .menu button { display: flex; align-items: center; height: 24px; padding: 0 10px; text-align: left; font-size: 12px; white-space: nowrap; }
  .menu button[aria-checked='true'] { background: none; border-color: transparent; }
  .menu button[aria-checked='true']::before { content: '✓'; margin-left: -4px; margin-right: 4px; }
  .color { position: absolute; width: 1px; height: 1px; opacity: 0; pointer-events: none; }
  .switch { display: flex; flex-direction: column; align-items: center; align-self: stretch; justify-content: flex-start; gap: 4px; width: 62px; padding-top: 6px; font-size: 11px; line-height: 1.15; text-align: center; cursor: pointer; }
  .switch input { appearance: none; position: relative; width: 36px; height: 20px; margin: 0; border: 1px solid var(--ok-border); border-radius: 10px; background: var(--ok-hover); cursor: pointer; }
  .switch input::before { content: ''; position: absolute; top: 2px; left: 2px; width: 14px; height: 14px; border-radius: 50%; background: var(--ok-muted, #888); transition: left 0.12s; }
  .switch input:checked { background: var(--ok-accent); border-color: var(--ok-accent); }
  .switch input:checked::before { left: 18px; background: #fff; }
</style>
