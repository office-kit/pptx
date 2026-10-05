<script lang="ts">
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';

  const editor = getEditor();
  const ink = editor.ink;
  let colorInput = $state<HTMLInputElement>();
</script>

<div class="group">
  <button aria-pressed={ink.tool === 'pen'} onclick={() => ink.toggle('pen')}>
    <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M5 19 16 4l3 2L8 21l-4 1z" /><path d="M3 22c3-2 6 0 9-2" stroke="#c00000" /></svg>
    {t('Draw')}
  </button>
  <button aria-pressed={ink.tool === 'eraser'} onclick={() => ink.toggle('eraser')}>
    <svg viewBox="0 0 24 24" aria-hidden="true"><path d="m4 15 9-10 7 6-9 10H8z" fill="#b48ad8" /><path d="M8 21h12" /></svg>
    {t('Eraser')}
  </button>
  <button aria-pressed={ink.tool === 'lasso'} onclick={() => ink.toggle('lasso')}>
    <svg viewBox="0 0 24 24" aria-hidden="true"><ellipse cx="11" cy="10" rx="8" ry="6" stroke-dasharray="2 2" /><path d="M18 17v6M15 20h6" stroke="#2e8b57" /></svg>
    {t('Lasso Select')}
  </button>
</div>
<div class="group pens" role="radiogroup" aria-label={t('Pens')}>
  {#each ink.pens as pen (pen.id)}
    <button
      class="pen"
      role="radio"
      aria-checked={ink.tool === 'pen' && ink.pen.id === pen.id}
      aria-label={t(pen.label)}
      title={t(pen.label)}
      onclick={() => ink.choosePen(pen)}
    >
      <svg viewBox="0 0 20 44" aria-hidden="true">
        <path d="M5 2h10v26l-5 12-5-12z" fill={pen.color} fill-opacity={pen.opacity} />
        <path d="M5 28h10" />
      </svg>
    </button>
  {/each}
</div>
<div class="group">
  <button onclick={() => colorInput?.click()}>
    <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 4v16M4 12h16" stroke="#2e8b57" stroke-width="1.6" /></svg>
    {t('Add Pen')}
  </button>
  <input
    bind:this={colorInput}
    class="color"
    type="color"
    tabindex="-1"
    aria-label={t('Pen Color')}
    onchange={(event) => ink.addPen(event.currentTarget.value as `#${string}`)}
  />
</div>
<div class="group">
  <button disabled title={t('Handwriting recognition is not available in the browser.')}>
    <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M4 9c2-6 6-6 6 0s-4 3-4 0M14 20a3 3 0 1 1 6 0v-6" /></svg>
    {t('Ink to Text')}
  </button>
  <button aria-pressed={ink.toShape} onclick={() => { ink.toShape = !ink.toShape; if (ink.toShape) ink.tool = 'pen'; }}>
    <svg viewBox="0 0 24 24" aria-hidden="true"><rect x="6" y="8" width="14" height="12" /><path d="M3 4c3 2 6-2 9 0" /></svg>
    {t('Ink to Shape')}
  </button>
  <button disabled title={t('Handwriting recognition is not available in the browser.')}>
    <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M5 6h12M8 6v12M14 6v10c0 2 2 2 3 1" /></svg>
    {t('Ink to Math')}
  </button>
</div>
<div class="group">
  <label class="switch">
    <input type="checkbox" role="switch" checked={ink.tool === 'pen'} onchange={() => ink.toggle('pen')} />
    <span>{t('Draw with Trackpad')}</span>
  </label>
</div>

<style>
  .group { display: flex; align-items: center; flex-shrink: 0; gap: 2px; padding: 0 8px; border-right: 1px solid var(--ok-border); }
  .pens { margin: 4px 4px; padding: 0 6px; border: 1px solid var(--ok-border); border-radius: var(--ok-radius); background: var(--ok-panel-2, transparent); }
  button { display: flex; flex-direction: column; align-items: center; justify-content: flex-start; gap: 2px; min-width: 52px; max-width: 96px; min-height: 66px; padding: 3px 4px; font: inherit; font-size: 11px; line-height: 1.2; color: var(--ok-text); background: transparent; border: 1px solid transparent; border-radius: var(--ok-radius); cursor: pointer; }
  button:not(:disabled):hover, button[aria-pressed='true'], button[aria-checked='true'] { background: var(--ok-hover); border-color: var(--ok-border); }
  button:disabled { opacity: 0.45; cursor: default; }
  .pen { min-width: 40px; justify-content: center; }
  .pen svg { width: 20px; height: 44px; }
  svg { width: 32px; height: 32px; flex-shrink: 0; stroke: currentColor; fill: none; stroke-width: 1.1; }
  .color { position: absolute; width: 1px; height: 1px; opacity: 0; pointer-events: none; }
  .switch input { appearance: none; position: relative; width: 36px; height: 20px; margin: 4px 0; border: 1px solid var(--ok-border); border-radius: 10px; background: var(--ok-hover); cursor: pointer; }
  .switch input::before { content: ''; position: absolute; top: 2px; left: 2px; width: 14px; height: 14px; border-radius: 50%; background: var(--ok-muted, #888); transition: left 0.12s; }
  .switch input:checked { background: var(--ok-accent); border-color: var(--ok-accent); }
  .switch input:checked::before { left: 18px; background: #fff; }
  .switch { display: flex; flex-direction: column; align-items: center; gap: 4px; max-width: 96px; font-size: 11px; text-align: center; cursor: pointer; }
</style>
