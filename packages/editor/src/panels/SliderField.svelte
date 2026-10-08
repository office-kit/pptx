<script lang="ts">
  import { t } from '../i18n/i18n.svelte.ts';

  // The reference desktop app's slider row: an 82 pt label, a 113 pt slider and a 66 pt
  // spin box. An empty box means the effect is off or the selection is mixed.
  let { label, name, value, min, max, step = 1, unit, disabled = false, slider = true, title, apply }: {
    label: string;
    /** Accessible name; defaults to the label. */
    name?: string;
    value: number | undefined;
    min: number;
    max: number;
    step?: number;
    unit: string;
    disabled?: boolean;
    /** Rows such as Width / Height in 3-D Format have only the spin box. */
    slider?: boolean;
    title?: string;
    apply: (value: number) => void;
  } = $props();
  const shown = $derived(value === undefined ? '' : String(Math.round(value * 100) / 100));
  function commit(input: HTMLInputElement) {
    if (disabled || !input.reportValidity() || !Number.isFinite(input.valueAsNumber)) { input.value = shown; return; }
    apply(Math.min(max, Math.max(min, input.valueAsNumber)));
  }
</script>

<div class="slider-field" {title}>
  <span class="label">{label}</span>
  <div class="controls">
    {#if slider}<input type="range" {min} {max} {step} value={value ?? min} aria-label={name ?? label} aria-valuetext={value === undefined ? t('None') : `${shown} ${unit}`} {disabled} onchange={(event) => commit(event.currentTarget)} />{/if}
    <span class="number"><input class="ok-input" type="number" {min} {max} step="any" value={shown} aria-label={name ?? label} {disabled} onchange={(event) => commit(event.currentTarget)} /><span>{unit}</span></span>
  </div>
</div>

<style>
  .slider-field { display: flex; align-items: center; gap: 4px; min-height: 26px; }
  .label { flex: none; min-width: 82px; white-space: nowrap; }
  .controls { display: flex; flex: 1; gap: 4px; align-items: center; justify-content: flex-end; min-width: 0; }
  input[type=range] { flex: 1; width: 0; min-width: 0; accent-color: var(--ok-accent); }
  .number { display: flex; align-items: center; gap: 2px; width: 66px; flex: none; }
  .number input { box-sizing: border-box; width: 46px; height: 26px; min-width: 0; padding: 2px 4px; font-size: inherit; }
  .number > span { color: var(--ok-text-2); font-size: 11px; }
</style>
