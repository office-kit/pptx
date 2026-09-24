<script lang="ts">
  import { getShapeImageFillLayout, setShapeImageFillLayout, getShapeImageOpacity, setShapeImageOpacity, setShapeImageFill, getShapeKind, pt, type ImageFillLayout, type ImageTileAlignment, type ImageTileFlip, type SlideShapeData } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';

  let { visible = true }: { visible?: boolean } = $props();
  const editor = getEditor();
  const doc = editor.doc;
  const shapes = $derived.by(() => { doc.version; return editor.selectedShapes(); });
  const layouts = $derived(shapes.map(getShapeImageFillLayout));
  const tiled = $derived(layouts.length > 0 && layouts.every(value => value?.mode === 'tile'));
  const mixedMode = $derived(layouts.some(value => value?.mode === 'tile') && !tiled);
  let input = $state<HTMLInputElement>();
  let loading = $state(false);
  const locked = $derived(editor.selectionLocked() || loading);
  let error = $state('');
  const alignments: [ImageTileAlignment, string][] = [['tl', 'Top left'], ['t', 'Top'], ['tr', 'Top right'], ['l', 'Left'], ['ctr', 'Center'], ['r', 'Right'], ['bl', 'Bottom left'], ['b', 'Bottom'], ['br', 'Bottom right']];
  const flips: [ImageTileFlip, string][] = [['none', 'None'], ['x', 'Horizontal'], ['y', 'Vertical'], ['xy', 'Both']];
  const tileFields = [['offsetX', 'Offset X'], ['offsetY', 'Offset Y'], ['scaleX', 'Scale X'], ['scaleY', 'Scale Y']] as const;
  const stretchFields = [['left', 'Offset left'], ['right', 'Offset right'], ['top', 'Offset top'], ['bottom', 'Offset bottom']] as const;
  function common<T>(values: T[]): T | undefined { return values.every(value => value === values[0]) ? values[0] : undefined; }
  const transparency = $derived(common(shapes.map(shape => Math.round((1 - (getShapeImageOpacity(shape) ?? 1)) * 100000) / 1000)));
  const rotation = $derived(common(layouts.map(value => value?.rotateWithShape ?? true)));
  const alignment = $derived(common(layouts.map(value => value?.mode === 'tile' ? value.alignment ?? 'tl' : undefined)));
  const flip = $derived(common(layouts.map(value => value?.mode === 'tile' ? value.flip ?? 'none' : undefined)));

  function apply(label: string, edit: (shape: SlideShapeData) => void) {
    if (locked || !shapes.length) return;
    try { doc.transact(t(label), () => { for (const shape of shapes) edit(shape); }); error = ''; }
    catch (cause) { error = cause instanceof Error ? cause.message : String(cause); }
  }
  function layout(edit: (value: ImageFillLayout) => ImageFillLayout) {
    apply('Picture or texture fill', shape => {
      const value = getShapeImageFillLayout(shape);
      if (value) setShapeImageFillLayout(shape, edit(value));
    });
  }
  function numericValue(field: string) {
    return common(layouts.map(value => {
      if (!value) return undefined;
      if (value.mode === 'tile') {
        if (field === 'offsetX' || field === 'offsetY') return (value[field] ?? 0) / pt(1);
        if (field === 'scaleX' || field === 'scaleY') return (value[field] ?? 1) * 100;
      } else if (field === 'left' || field === 'right' || field === 'top' || field === 'bottom') return (value[field] ?? 0) * 100;
      return undefined;
    }));
  }
  function numeric(element: HTMLInputElement, field: string) {
    if (!element.reportValidity() || !Number.isFinite(element.valueAsNumber)) {
      element.value = String(numericValue(field) ?? ''); return;
    }
    const value = element.valueAsNumber;
    layout(current => {
      if (current.mode === 'tile') {
        if (field === 'offsetX' || field === 'offsetY') return { ...current, [field]: pt(value) };
        if (field === 'scaleX' || field === 'scaleY') return { ...current, [field]: value / 100 };
      } else if (field === 'left' || field === 'right' || field === 'top' || field === 'bottom') return { ...current, [field]: value / 100 };
      return current;
    });
  }
  export function chooseImage() { if (!locked) input?.click(); }
  async function upload(event: Event) {
    const element = event.currentTarget;
    if (!(element instanceof HTMLInputElement)) return;
    const file = element.files?.[0];
    if (!file || locked || shapes.some(shape => getShapeKind(shape) !== 'shape')) return;
    const targets = [...shapes], presentation = doc.pres, version = doc.version, selection = doc.selection;
    loading = true; error = '';
    try {
      const bytes = new Uint8Array(await file.arrayBuffer());
      if (doc.pres !== presentation || doc.version !== version || doc.selection !== selection || editor.selectionLocked()) {
        error = t('The selection changed. Choose the picture again.'); return;
      }
      doc.transact(t('Picture or texture fill'), () => { for (const target of targets) setShapeImageFill(target, bytes); });
    } catch (cause) { error = cause instanceof Error ? cause.message : String(cause); }
    finally { loading = false; element.value = ''; }
  }
</script>

<input type="file" accept="image/*" hidden bind:this={input} onchange={upload} aria-label={t('Picture source')} />
{#if visible}
  <fieldset disabled={locked}>
    <span>{t('Picture source')}</span>
    <button class="ok-btn" disabled={shapes.some(shape => getShapeKind(shape) !== 'shape')} onclick={chooseImage}>{t('Insert...')}</button>
    <span>{t('Transparency')}</span>
    <div class="transparency">
      <input type="range" min="0" max="100" value={transparency ?? 0} aria-label={t('Picture transparency')} aria-valuetext={transparency === undefined ? t('Mixed') : `${transparency}%`} onchange={event => { const value = event.currentTarget.valueAsNumber; apply('Picture transparency', shape => setShapeImageOpacity(shape, 1 - value / 100)); }} />
      <label class="number"><input class="ok-input" type="number" min="0" max="100" step="any" value={transparency ?? ''} placeholder={t('Mixed')} aria-label={t('Picture transparency')} onchange={event => { const element = event.currentTarget; if (element.reportValidity() && Number.isFinite(element.valueAsNumber)) { const value = element.valueAsNumber; apply('Picture transparency', shape => setShapeImageOpacity(shape, 1 - value / 100)); } else element.value = String(transparency ?? ''); }} />%</label>
    </div>
    <label class="check"><input type="checkbox" checked={tiled} indeterminate={mixedMode} onchange={event => { const tile = event.currentTarget.checked; layout(current => ({ mode: tile ? 'tile' : 'stretch', rotateWithShape: current.rotateWithShape })); }} />{t('Tile picture as texture')}</label>
    {#each tiled ? tileFields : stretchFields as [field, label]}
      <label class="row"><span>{t(label)}</span><span class="number"><input class="ok-input" type="number" min={field.startsWith('scale') ? 0 : tiled ? -1584 : -100000} max={field.startsWith('scale') ? 100 : tiled ? 1584 : 100000} step="any" value={numericValue(field) ?? ''} placeholder={t('Mixed')} aria-label={t(label)} onchange={event => numeric(event.currentTarget, field)} />{field.startsWith('offset') ? 'pt' : '%'}</span></label>
    {/each}
    {#if tiled}
      <label class="row">{t('Alignment')}<select class="ok-input" value={alignment ?? ''} onchange={event => { const value = alignments.find(([key]) => key === event.currentTarget.value)?.[0]; if (value) layout(current => current.mode === 'tile' ? { ...current, alignment: value } : current); }}>
        {#if alignment === undefined}<option value="" disabled>{t('Mixed')}</option>{/if}
        {#each alignments as [value, label]}<option {value}>{t(label)}</option>{/each}
      </select></label>
      <label class="row">{t('Mirror type')}<select class="ok-input" value={flip ?? ''} onchange={event => { const value = flips.find(([key]) => key === event.currentTarget.value)?.[0]; if (value) layout(current => current.mode === 'tile' ? { ...current, flip: value } : current); }}>
        {#if flip === undefined}<option value="" disabled>{t('Mixed')}</option>{/if}
        {#each flips as [value, label]}<option {value}>{t(label)}</option>{/each}
      </select></label>
    {/if}
    <label class="check"><input type="checkbox" checked={rotation ?? false} indeterminate={rotation === undefined} onchange={event => { const value = event.currentTarget.checked; layout(current => ({ ...current, rotateWithShape: value })); }} />{t('Rotate with shape')}</label>
  </fieldset>
{/if}
{#if error}<p role="alert">{error}</p>{/if}

<style>
  fieldset { border: 0; padding: 0; margin: 0; display: grid; gap: 8px; min-width: 0; font-size: 11px; }
  .row, .check, .number, .transparency { display: flex; align-items: center; gap: 6px; }
  .row { justify-content: space-between; }
  .row select { max-width: 130px; }
  .number input { width: 64px; min-width: 0; }
  .transparency input[type=range] { flex: 1; min-width: 0; width: 0; accent-color: var(--ok-accent); }
  [role=alert] { color: #bf3131; font-size: 11px; }
</style>
