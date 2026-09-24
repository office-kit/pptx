<script lang="ts">
  import { getShapeFill, getShapeStroke, getShapeFillOpacity, getShapeStrokeOpacity, setShapeFill, setShapeStroke } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';

  let { paint }: { paint: 'fill' | 'line' } = $props();
  const editor = getEditor();
  const shapes = $derived.by(() => { editor.doc.version; return editor.selectedShapes(); });
  const disabled = $derived(!shapes.length || shapes.some(shape => (paint === 'fill' ? getShapeFill(shape) : getShapeStroke(shape)).kind !== 'solid'));
  const value = $derived.by(() => {
    const values = shapes.map(shape => Math.round((1 - ((paint === 'fill' ? getShapeFillOpacity(shape) : getShapeStrokeOpacity(shape)) ?? 1)) * 100_000) / 1000);
    return values.every(value => value === values[0]) ? values[0] : undefined;
  });
  const label = $derived(paint === 'fill' ? 'Fill transparency' : 'Line transparency');
  function apply(input: HTMLInputElement) {
    if (disabled) return;
    if (!input.reportValidity() || !Number.isFinite(input.valueAsNumber)) { input.value = value === undefined ? '' : String(value); return; }
    const opacity = 1 - input.valueAsNumber / 100;
    editor.doc.transact(t(label), () => {
      for (const shape of shapes) {
        if (paint === 'fill') setShapeFill(shape, { opacity });
        else setShapeStroke(shape, { opacity });
      }
    });
  }
</script>
<div class="transparency">
  <span>{t('Transparency')}</span>
  <div class="controls">
    <input type="range" min="0" max="100" step="1" value={value ?? 0} aria-label={t(label)} aria-valuetext={value === undefined ? t('Mixed') : `${value}%`} {disabled} onchange={event => apply(event.currentTarget)} />
    <span class="number"><input class="ok-input" type="number" min="0" max="100" step="any" value={value ?? ''} placeholder={value === undefined ? t('Mixed') : undefined} aria-label={t(label)} {disabled} onchange={event => apply(event.currentTarget)} /><span>%</span></span>
  </div>
</div>
<style>
  .transparency { display: flex; flex-direction: column; gap: 4px; }
  .controls { display: flex; gap: 8px; align-items: center; }
  input[type=range] { flex: 1; width: 0; min-width: 0; accent-color: var(--ok-accent); }
  .number { display: flex; align-items: center; gap: 3px; width: 72px; }
  .number input { width: 52px; min-width: 0; padding: 2px 4px; font-size: inherit; }
</style>
