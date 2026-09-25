<script lang="ts">
  import { onMount, untrack } from 'svelte';
  import { emu, getShapeKind, getShapeImageBytes, getShapeImageFormat, getShapeImageCrop, getShapeBounds, setShapeBounds, setShapeImageCrop } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { selectedShapeIds } from '../core/selection.ts';
  import { t } from '../i18n/i18n.svelte.ts';

  const editor = getEditor();
  const doc = editor.doc;
  const selection = untrack(() => doc.selection);
  const version = untrack(() => doc.version);
  const geometryLocked = untrack(() => editor.selectionLocked());
  const ids = selectedShapeIds(selection);
  const shape = ids.length === 1 ? untrack(() => doc.shapeById(selection.slideIndex, ids[0]!)) : null;
  const picture = shape && getShapeKind(shape) === 'picture' ? shape : null;
  const original = picture ? getShapeImageCrop(picture) : null;
  const bounds = picture ? getShapeBounds(picture) : null;
  const bytes = picture ? getShapeImageBytes(picture) : null;
  const format = picture ? getShapeImageFormat(picture) : null;
  type Rect = { left: number; top: number; right: number; bottom: number };
  type Handle = 'move' | 'n' | 'ne' | 'e' | 'se' | 's' | 'sw' | 'w' | 'nw';
  const minimum = 0.001;
  const full: Rect = { left: 0, top: 0, right: 1, bottom: 1 };
  let rect = $state<Rect>({ left: original?.left ?? 0, top: original?.top ?? 0, right: 1 - (original?.right ?? 0), bottom: 1 - (original?.bottom ?? 0) });
  let dialog: HTMLDialogElement;
  let surface = $state<HTMLDivElement>();
  let url = $state('');
  let loaded = $state(false);
  let dimensions = $state({ width: 1, height: 1 });
  let error = $state('');
  let gesture: { pointer: number; handle: Handle; x: number; y: number; width: number; height: number; original: Rect } | null = null;
  const handles: { id: Handle; label: string; x: number; y: number }[] = [
    { id: 'nw', label: 'Crop top left', x: 0, y: 0 }, { id: 'n', label: 'Crop top edge', x: 50, y: 0 },
    { id: 'ne', label: 'Crop top right', x: 100, y: 0 }, { id: 'e', label: 'Crop right edge', x: 100, y: 50 },
    { id: 'se', label: 'Crop bottom right', x: 100, y: 100 }, { id: 's', label: 'Crop bottom edge', x: 50, y: 100 },
    { id: 'sw', label: 'Crop bottom left', x: 0, y: 100 }, { id: 'w', label: 'Crop left edge', x: 0, y: 50 },
  ];
  const valid = $derived(rect.left >= 0 && rect.top >= 0 && rect.right <= 1 && rect.bottom <= 1 && rect.right > rect.left && rect.bottom > rect.top);
  const width = $derived(rect.right - rect.left);
  const height = $derived(rect.bottom - rect.top);
  let ratioKey = $state('free');
  const presets = [
    { value: 'free', label: 'Free crop', ratio: 0 },
    { value: 'source', label: 'Original image ratio', ratio: 0 },
    { value: '1:1', label: 'Square (1:1)', ratio: 1 },
    ...[[16, 9], [4, 3], [3, 2], [9, 16], [3, 4], [2, 3]].map(([w, h]) => ({ value: `${w}:${h}`, label: `${w}:${h}`, ratio: w! / h! })),
  ];
  const ratio = $derived(ratioKey === 'source' ? dimensions.width / dimensions.height : presets.find(preset => preset.value === ratioKey)?.ratio ?? 0);
  const sourceRatio = $derived(ratio * dimensions.height / dimensions.width);
  const originalAspect = bounds && bounds.w > 0 && bounds.h > 0 ? bounds.w / bounds.h : 1;
  const aspect = $derived(ratio || originalAspect);
  function chooseRatio() {
    if (!ratio) return;
    const start = valid ? rect : full;
    const w = Math.min(start.right - start.left, (start.bottom - start.top) * sourceRatio);
    const h = w / sourceRatio;
    const cx = (start.left + start.right) / 2, cy = (start.top + start.bottom) / 2;
    rect = { left: cx - w / 2, right: cx + w / 2, top: cy - h / 2, bottom: cy + h / 2 };
  }
  onMount(() => {
    dialog.showModal();
    if (!picture) { error = t('Select one image to crop.'); return; }
    if (!bytes || !format) { error = t('The image could not be read'); return; }
    const next = URL.createObjectURL(new Blob([new Uint8Array(bytes)], { type: format === 'svg' ? 'image/svg+xml' : `image/${format}` }));
    url = next;
    return () => URL.revokeObjectURL(next);
  });
  function adjust(start: Rect, handle: Handle, dx: number, dy: number) {
    const next = { ...start };
    if (handle === 'move') {
      dx = Math.max(-start.left, Math.min(1 - start.right, dx));
      dy = Math.max(-start.top, Math.min(1 - start.bottom, dy));
      next.left += dx; next.right += dx; next.top += dy; next.bottom += dy;
    } else if (ratio) {
      const horizontal = handle.includes('w') ? -1 : handle.includes('e') ? 1 : 0;
      const vertical = handle.includes('n') ? -1 : handle.includes('s') ? 1 : 0;
      // Opposite corners stay fixed; edge handles grow symmetrically on the other axis.
      const ax = horizontal < 0 ? start.right : horizontal > 0 ? start.left : (start.left + start.right) / 2;
      const ay = vertical < 0 ? start.bottom : vertical > 0 ? start.top : (start.top + start.bottom) / 2;
      const maxW = horizontal < 0 ? ax : horizontal > 0 ? 1 - ax : 2 * Math.min(ax, 1 - ax);
      const maxH = vertical < 0 ? ay : vertical > 0 ? 1 - ay : 2 * Math.min(ay, 1 - ay);
      const dw = dx * horizontal, dh = dy * vertical * sourceRatio;
      const delta = Math.abs(dw) >= Math.abs(dh) ? dw : dh;
      const limit = Math.min(maxW, maxH * sourceRatio);
      const w = Math.min(limit, Math.max(Math.min(limit, minimum * Math.max(1, sourceRatio)), start.right - start.left + delta));
      const h = w / sourceRatio;
      next.left = horizontal < 0 ? ax - w : horizontal > 0 ? ax : ax - w / 2;
      next.right = next.left + w;
      next.top = vertical < 0 ? ay - h : vertical > 0 ? ay : ay - h / 2;
      next.bottom = next.top + h;
    } else {
      if (handle.includes('w')) next.left = Math.max(0, Math.min(start.right - minimum, start.left + dx));
      if (handle.includes('e')) next.right = Math.min(1, Math.max(start.left + minimum, start.right + dx));
      if (handle.includes('n')) next.top = Math.max(0, Math.min(start.bottom - minimum, start.top + dy));
      if (handle.includes('s')) next.bottom = Math.min(1, Math.max(start.top + minimum, start.bottom + dy));
    }
    rect = { left: Math.max(0, next.left), top: Math.max(0, next.top), right: Math.min(1, next.right), bottom: Math.min(1, next.bottom) };
  }
  function begin(event: PointerEvent, handle: Handle) {
    if (event.button !== 0 || !valid) return;
    event.preventDefault();
    const box = surface!.getBoundingClientRect();
    gesture = { pointer: event.pointerId, handle, x: event.clientX, y: event.clientY, width: box.width, height: box.height, original: { ...rect } };
    if (event.currentTarget instanceof HTMLElement) { event.currentTarget.focus(); event.currentTarget.setPointerCapture(event.pointerId); }
  }
  function move(event: PointerEvent) {
    if (!gesture || gesture.pointer !== event.pointerId) return;
    adjust(gesture.original, gesture.handle, (event.clientX - gesture.x) / gesture.width, (event.clientY - gesture.y) / gesture.height);
  }
  function end(event: PointerEvent, cancel = false) {
    if (!gesture || gesture.pointer !== event.pointerId) return;
    if (cancel) rect = gesture.original;
    gesture = null;
  }
  function key(event: KeyboardEvent, handle: Handle) {
    if (!valid || !['ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown'].includes(event.key)) return;
    event.preventDefault();
    event.stopPropagation();
    const step = event.shiftKey ? 0.1 : 0.01;
    adjust(rect, handle, event.key === 'ArrowLeft' ? -step : event.key === 'ArrowRight' ? step : 0, event.key === 'ArrowUp' ? -step : event.key === 'ArrowDown' ? step : 0);
  }
  function submit(event: SubmitEvent) {
    event.preventDefault();
    if (!picture || !loaded || !valid) return;
    if (doc.version !== version || doc.selection !== selection) { error = t('The document changed. Select the image again.'); return; }
    try {
      const crop = { left: rect.left, top: rect.top, right: 1 - rect.right, bottom: 1 - rect.bottom };
      doc.transact(t('Crop image'), () => {
        setShapeImageCrop(picture, Object.values(crop).every(value => value === 0) ? null : crop);
        if (ratio && bounds && !geometryLocked) {
          const w = Math.min(bounds.w, bounds.h * ratio), h = w / ratio;
          setShapeBounds(picture, { ...bounds, x: emu(Math.round(bounds.x + (bounds.w - w) / 2)), y: emu(Math.round(bounds.y + (bounds.h - h) / 2)), w: emu(Math.max(1, Math.round(w))), h: emu(Math.max(1, Math.round(h))) });
        }
      });
      editor.closeDialog();
    } catch (cause) { error = cause instanceof Error ? cause.message : String(cause); }
  }
</script>

<dialog bind:this={dialog} aria-label={t('Crop image')} onclose={() => editor.closeDialog()}>
  <form onsubmit={submit}>
    <header><strong>{t('Crop image')}</strong><button class="ok-btn" type="button" aria-label={t('Close')} onclick={() => editor.closeDialog()}>✕</button></header>
    <p>{t('Drag the edges to crop or drag the selection to move it. Arrow keys adjust by 1%; hold Shift for 10%.')}</p>
    <label class="ratio">{t('Crop aspect ratio')}<select aria-label={t('Crop aspect ratio')} value={ratioKey} onchange={event => { ratioKey = event.currentTarget.value; chooseRatio(); }} disabled={!loaded || geometryLocked}>{#each presets as preset}<option value={preset.value}>{t(preset.label)}</option>{/each}</select></label>
    {#if ratio}<p>{t('The image frame will match this ratio and stay centered.')}</p>{/if}
    {#if url}
      <div class="workspace">
        <div bind:this={surface} class="crop-surface" style:aspect-ratio={`${dimensions.width} / ${dimensions.height}`} style:width={`min(100%, ${380 * dimensions.width / dimensions.height}px)`}>
          <img class="source" src={url} alt={t('Selected image preview')} draggable="false" onload={event => { if (!(event.currentTarget instanceof HTMLImageElement)) return; dimensions = { width: event.currentTarget.naturalWidth, height: event.currentTarget.naturalHeight }; loaded = true; }} onerror={() => { loaded = false; error = t('The image could not be read'); }} />
          {#if loaded && valid}
            <div class="shade" style:clip-path={`polygon(evenodd, 0 0, 100% 0, 100% 100%, 0 100%, 0 0, ${rect.left * 100}% ${rect.top * 100}%, ${rect.left * 100}% ${rect.bottom * 100}%, ${rect.right * 100}% ${rect.bottom * 100}%, ${rect.right * 100}% ${rect.top * 100}%, ${rect.left * 100}% ${rect.top * 100}%)`}></div>
            <div class="crop-box" style:left={`${rect.left * 100}%`} style:top={`${rect.top * 100}%`} style:width={`${width * 100}%`} style:height={`${height * 100}%`}>
              <button class="move" type="button" aria-label={t('Move crop selection')} onpointerdown={event => begin(event, 'move')} onpointermove={move} onpointerup={event => end(event)} onpointercancel={event => end(event, true)} onkeydown={event => key(event, 'move')}></button>
              {#each handles as handle}<button class="handle" type="button" aria-label={t(handle.label)} style:left={`${handle.x}%`} style:top={`${handle.y}%`} style:cursor={`${handle.id}-resize`} onpointerdown={event => begin(event, handle.id)} onpointermove={move} onpointerup={event => end(event)} onpointercancel={event => end(event, true)} onkeydown={event => key(event, handle.id)}></button>{/each}
            </div>
          {/if}
        </div>
      </div>
      {#if loaded && valid}<div class="result"><span>{t('Cropped image preview')}</span><div class="preview" style:aspect-ratio={aspect} style:width={`min(180px, ${100 * aspect}px)`}><img src={url} alt={t('Cropped image preview')} style:width={`${100 / width}%`} style:height={`${100 / height}%`} style:left={`${-rect.left / width * 100}%`} style:top={`${-rect.top / height * 100}%`} /></div></div>{/if}
    {/if}
    {#if !valid}<p role="alert">{t('Reset crop to adjust this image visually.')}</p>{/if}
    {#if error}<p role="alert">{error}</p>{/if}
    <footer><button class="ok-btn" type="button" onclick={() => { ratioKey = 'free'; rect = { ...full }; }}>{t('Reset crop')}</button><span></span><button class="ok-btn" type="button" onclick={() => editor.closeDialog()}>{t('Cancel')}</button><button class="ok-btn primary" type="submit" disabled={!loaded || !valid}>{t('Apply')}</button></footer>
  </form>
</dialog>

<style>
  dialog { width: min(740px, 90vw); max-height: 90vh; overflow: auto; padding: 18px; border: 1px solid var(--ok-border); border-radius: var(--ok-radius-lg); background: var(--ok-panel); color: var(--ok-text); }
  dialog::backdrop { background: #0006; }
  form { display: grid; gap: 16px; }
  header, footer { display: flex; align-items: center; gap: 12px; }
  header { justify-content: space-between; }
  footer span { flex: 1; }
  p, .result { font-size: 12px; margin: 0; }
  .ratio { display: flex; align-items: center; gap: 12px; font-size: 12px; }
  .ratio select { padding: 5px 8px; background: var(--ok-panel); color: var(--ok-text); border: 1px solid var(--ok-border); border-radius: 4px; }
  .workspace { display: flex; justify-content: center; padding: 12px; background: var(--ok-bg); }
  .crop-surface { position: relative; }
  .source { display: block; width: 100%; height: 100%; user-select: none; }
  .crop-box { position: absolute; box-sizing: border-box; border: 2px solid white; outline: 1px solid #111; }
  .shade { position: absolute; inset: 0; background: #0007; pointer-events: none; }
  .move { position: absolute; inset: 0; width: 100%; height: 100%; border: 0; background: transparent; cursor: move; touch-action: none; }
  .handle { position: absolute; width: 14px; height: 14px; padding: 0; transform: translate(-50%, -50%); border: 2px solid #111; background: white; touch-action: none; }
  .move:focus-visible, .handle:focus-visible { outline: 3px solid var(--ok-accent); outline-offset: 2px; }
  .result { display: flex; align-items: center; justify-content: center; gap: 14px; }
  .preview { position: relative; overflow: hidden; background: var(--ok-bg); }
  .preview img { position: absolute; max-width: none; }
  [role='alert'] { color: #bf3131; }
</style>
