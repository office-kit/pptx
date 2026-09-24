<script lang="ts">
  import { getShapeGradientFillEffective, type ReadGradientFill, type ReadGradientStop } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import GradientDirection from './GradientDirection.svelte';

  const editor = getEditor();
  const gradient = $derived.by(() => {
    editor.doc.version;
    const shapes = editor.selectedShapes();
    if (shapes.length !== 1) return null;
    return getShapeGradientFillEffective(editor.doc.pres, shapes[0]!);
  });
  let selected = $state(0);
  $effect(() => { editor.doc.selection; selected = 0; });
  const selectedIndex = $derived(Math.min(selected, (gradient?.stops.length ?? 1) - 1));
  let dragging = $state<{ index: number; offset: number; left: number; width: number } | null>(null);
  $effect(() => { editor.doc.version; editor.doc.selection; dragging = null; });
  const visibleStops = $derived(gradient?.stops.map((item, index) => dragging?.index === index ? { ...item, offset: dragging.offset } : item) ?? []);
  const stop = $derived(visibleStops[selectedIndex]);
  const locked = $derived(editor.selectionLocked());
  const stopColor = (stop: ReadGradientStop) => stop.resolvedColor ?? (/^#[\da-f]{6}$/i.test(stop.color) ? stop.color : '#000000');
  const track = $derived(gradient ? `linear-gradient(to right, ${[...visibleStops].sort((a, b) => a.offset - b.offset).map(stop => `${stopColor(stop)} ${stop.offset * 100}%`).join(', ')})` : '');

  function startDrag(event: PointerEvent, index: number) {
    if (locked || event.button !== 0) return;
    const target = event.currentTarget as HTMLButtonElement;
    const bounds = target.parentElement!.getBoundingClientRect();
    selected = index;
    dragging = { index, offset: gradient!.stops[index]!.offset, left: bounds.left, width: bounds.width };
    target.focus();
    target.setPointerCapture(event.pointerId);
    event.preventDefault();
  }
  function moveDrag(event: PointerEvent) {
    if (!dragging) return;
    dragging.offset = Math.max(0, Math.min(1, Math.round((event.clientX - dragging.left) / dragging.width * 1000) / 1000));
  }
  function finishDrag() {
    if (!dragging || !gradient) return;
    const { offset, index } = dragging;
    dragging = null;
    if (offset !== gradient.stops[index]!.offset) editStop({ offset });
  }
  function apply(patch: Partial<ReadGradientFill>) {
    if (!gradient || locked) return;
    editor.invoke('setShapeGradientFill', { options: { ...gradient, ...patch } });
  }
  function editStop(patch: Partial<ReadGradientStop>) {
    if (!gradient) return;
    apply({ stops: gradient.stops.map((stop, index) => index === selectedIndex ? { ...stop, ...patch } : stop) });
  }
  function numeric(input: HTMLInputElement, field: 'offset' | 'brightness' | 'opacity') {
    if (!stop) return;
    const current = field === 'opacity' ? 1 - (stop.opacity ?? 1) : stop[field] ?? 0;
    if (!input.reportValidity() || !Number.isFinite(input.valueAsNumber)) {
      input.value = String(current * 100); return;
    }
    editStop({ [field]: field === 'opacity' ? 1 - input.valueAsNumber / 100 : input.valueAsNumber / 100 });
  }
  function angle(input: HTMLInputElement) {
    if (!input.reportValidity() || !Number.isFinite(input.valueAsNumber)) { input.value = String(gradient?.angleDeg ?? 90); return; }
    apply({ angleDeg: input.valueAsNumber });
  }
  function addStop() {
    if (!gradient || !stop) return;
    const ordered = [...gradient.stops].sort((a, b) => a.offset - b.offset);
    const next = ordered.find(item => item.offset > stop.offset);
    const neighbor = next ?? (stop.offset === 1 ? ordered.slice().reverse().find(item => item.offset < stop.offset) : undefined);
    const offset = (stop.offset + (neighbor?.offset ?? 1)) / 2;
    const left = stopColor(stop).slice(1);
    const right = stopColor(neighbor ?? stop).slice(1);
    const color = '#' + [0, 2, 4].map(index => Math.round((parseInt(left.slice(index, index + 2), 16) + parseInt(right.slice(index, index + 2), 16)) / 2).toString(16).padStart(2, '0')).join('');
    const opacity = ((stop.opacity ?? 1) + (neighbor?.opacity ?? stop.opacity ?? 1)) / 2;
    const stops = [...gradient.stops, { offset, color, opacity }];
    apply({ stops });
    selected = stops.length - 1;
  }
  function removeStop() {
    if (!gradient || gradient.stops.length <= 2) return;
    apply({ stops: gradient.stops.filter((_, index) => index !== selectedIndex) });
    selected = Math.max(0, selectedIndex - 1);
  }
</script>

{#if gradient && stop}
  <fieldset disabled={locked} class="gradient-fields">
    <label class="field"><span>{t('Type')}</span>
      <select class="ok-input" aria-label={t('Gradient type')} value={gradient.path ?? 'linear'} onchange={event => {
        const path = event.currentTarget.value;
        if (path === 'linear' || path === 'circle' || path === 'rect' || path === 'shape') apply({ path });
      }}>
        <option value="linear">{t('Linear')}</option><option value="circle">{t('Radial')}</option>
        <option value="rect">{t('Rectangular')}</option><option value="shape">{t('Path')}</option>
      </select>
    </label>
    {#if gradient.path === undefined || gradient.path === 'linear'}
      <GradientDirection angle={gradient.angleDeg ?? 90} disabled={locked} choose={angleDeg => apply({ angleDeg, scaled: true })} />
    {/if}
    <label class="field"><span>{t('Angle')}</span><span class="number">
      <input class="ok-input" type="number" min="0" max="359.9" step="any" aria-label={t('Gradient angle')} value={gradient.angleDeg ?? 90} disabled={gradient.path !== undefined && gradient.path !== 'linear'} onchange={event => angle(event.currentTarget)} />°
    </span></label>
    <span>{t('Gradient stops')}</span>
    <div class="stops" role="group" aria-label={t('Gradient stops')} style:background={track}>
      {#each visibleStops as item, index}
        <button type="button" class:selected={selectedIndex === index} class="stop" style:left={`${item.offset * 100}%`} style:background={stopColor(item)} aria-label={`${t('Gradient stop')} ${index + 1}`} aria-pressed={selectedIndex === index} title={`${Math.round(item.offset * 100)}%`} onclick={() => selected = index} onpointerdown={event => startDrag(event, index)} onpointermove={moveDrag} onpointerup={finishDrag} onpointercancel={() => dragging = null} onlostpointercapture={() => dragging = null} onkeydown={event => {
          if (event.key === 'Escape' && dragging) { event.preventDefault(); event.stopPropagation(); dragging = null; return; }
          if (event.key !== 'ArrowLeft' && event.key !== 'ArrowRight') return;
          event.preventDefault(); event.stopPropagation();
          selected = index;
          editStop({ offset: Math.max(0, Math.min(1, item.offset + (event.key === 'ArrowLeft' ? -0.01 : 0.01))) });
        }}></button>
      {/each}
    </div>
    <div class="stop-actions"><button type="button" class="ok-btn" aria-label={t('Add gradient stop')} title={t('Add gradient stop')} onclick={addStop}>+</button><button type="button" class="ok-btn" aria-label={t('Remove gradient stop')} title={t('Remove gradient stop')} disabled={gradient.stops.length <= 2} onclick={removeStop}>−</button></div>
    <label class="field"><span>{t('Color')}</span><input type="color" aria-label={t('Gradient stop color')} value={stopColor(stop)} onchange={event => editStop({ color: event.currentTarget.value, brightness: 0 })} /></label>
    <label class="field"><span>{t('Position')}</span><span class="number"><input class="ok-input" type="number" min="0" max="100" step="any" aria-label={t('Gradient stop position')} value={Math.round(stop.offset * 100000) / 1000} onchange={event => numeric(event.currentTarget, 'offset')} />%</span></label>
    {#each [{ field: 'opacity', label: 'Transparency', accessible: 'Gradient stop transparency', min: 0, value: (1 - (stop.opacity ?? 1)) * 100 }, { field: 'brightness', label: 'Brightness', accessible: 'Gradient stop brightness', min: -100, value: (stop.brightness ?? 0) * 100 }] as control}
      <div class="amount"><span>{t(control.label)}</span><div class="amount-controls">
        <input type="range" min={control.min} max="100" step="1" aria-label={t(control.accessible)} value={control.value} onchange={event => numeric(event.currentTarget, control.field === 'opacity' ? 'opacity' : 'brightness')} />
        <span class="number"><input class="ok-input" type="number" min={control.min} max="100" step="any" aria-label={t(control.accessible)} value={Math.round(control.value * 1000) / 1000} onchange={event => numeric(event.currentTarget, control.field === 'opacity' ? 'opacity' : 'brightness')} />%</span>
      </div></div>
    {/each}
    <label class="rotate"><input type="checkbox" checked={gradient.rotateWithShape !== false} onchange={event => apply({ rotateWithShape: event.currentTarget.checked })} />{t('Rotate with shape')}</label>
  </fieldset>
{:else}
  <p>{t('Select one shape to edit gradient stops.')}</p>
{/if}

<style>
  .gradient-fields { border: 0; padding: 0; margin: 0; min-width: 0; display: flex; flex-direction: column; gap: 10px; }
  .field { display: flex; align-items: center; justify-content: space-between; gap: 8px; }
  .field select { width: 140px; }
  .number { display: flex; align-items: center; gap: 3px; width: 78px; }
  .number input { width: 58px; min-width: 0; padding: 2px 4px; font-size: inherit; }
  .stops { height: 18px; position: relative; margin: 2px 7px 10px; border: 1px solid var(--ok-border); }
  .stop { touch-action: none; position: absolute; top: 10px; transform: translateX(-50%); width: 12px; height: 18px; border: 1px solid #555; border-radius: 2px; padding: 0; }
  .stop.selected { outline: 2px solid var(--ok-accent); outline-offset: 2px; z-index: 1; }
  .stop-actions { display: flex; justify-content: flex-end; gap: 6px; }
  .amount { display: flex; flex-direction: column; gap: 4px; }
  .amount-controls { display: flex; align-items: center; gap: 8px; }
  input[type=range] { flex: 1; min-width: 0; width: 0; accent-color: var(--ok-accent); }
  .rotate { display: flex; align-items: center; gap: 6px; }
</style>
