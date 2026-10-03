<script lang="ts">
  import type { ParagraphTabStop } from '@office-kit/pptx';
  import type { TabStopEdit } from '../core/paragraph-tabs.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  let { area, stage, zoom, text = null, onindent, ontabs }: {
    area: HTMLElement; stage: HTMLElement; zoom: number;
    text?: { left: number; first: number; scale: number; tabStops: readonly ParagraphTabStop[] } | null;
    ontabs?: (edits: TabStopEdit[]) => void;
    onindent?: (kind: 'first' | 'hanging' | 'left', delta: number) => void;
  } = $props();
  const tabKinds = [
    { alignment: 'left', label: 'Left tab', path: 'M7 3v8h6' },
    { alignment: 'center', label: 'Center tab', path: 'M7 3v8M2 11h10' },
    { alignment: 'right', label: 'Right tab', path: 'M7 3v8H1' },
    { alignment: 'decimal', label: 'Decimal tab', path: 'M7 3v8M2 11h10M11 5h1' },
  ] as const;
  let tabKind = $state(0);
  let tabDrag = $state<{ stop: ParagraphTabStop; position: number; remove: boolean; pointer: number } | null>(null);
  function tabPosition(clientX: number) {
    return Math.max(0, Math.min(51206400, Math.round((clientX - root.getBoundingClientRect().left - textBounds.x) / emuToPixel)));
  }
  function addTab(event: PointerEvent) {
    if (event.button !== 0 || !text) return;
    event.preventDefault();
    ontabs?.([{ kind: 'set', stop: { positionEmu: tabPosition(event.clientX), alignment: tabKinds[tabKind]!.alignment } }]);
  }
  function moveTab(event: PointerEvent) {
    if (!tabDrag || event.pointerId !== tabDrag.pointer) return;
    const top = root.getBoundingClientRect().top;
    tabDrag.position = tabPosition(event.clientX);
    tabDrag.remove = event.clientY < top - thickness || event.clientY > top + thickness * 2;
  }
  function finishTab(event: PointerEvent) {
    if (!tabDrag || event.pointerId !== tabDrag.pointer) return;
    moveTab(event);
    const change = tabDrag;
    tabDrag = null;
    if (!change.remove && change.position === change.stop.positionEmu) return;
    const edits: TabStopEdit[] = [{ kind: 'clear', positionEmu: change.stop.positionEmu }];
    if (!change.remove) edits.push({ kind: 'set', stop: { ...change.stop, positionEmu: change.position } });
    ontabs?.(edits);
  }
  let textBounds = $state({ x: 0, y: 0, scale: 1 });
  let drag = $state<{ kind: 'first' | 'hanging' | 'left'; start: number; delta: number; pointer: number } | null>(null);
  const emuToPixel = $derived(zoom / 9525 * (text?.scale ?? 1) * textBounds.scale);
  const firstPosition = $derived(textBounds.x + ((text?.left ?? 0) + (text?.first ?? 0) + (drag && drag.kind !== 'hanging' ? drag.delta : 0)) * emuToPixel);
  const leftPosition = $derived(textBounds.x + ((text?.left ?? 0) + (drag && drag.kind !== 'first' ? drag.delta : 0)) * emuToPixel);
  const handles = [{ kind: 'first', label: 'First line indent' }, { kind: 'hanging', label: 'Hanging indent' }, { kind: 'left', label: 'Left indent' }] as const;
  function startIndent(event: PointerEvent, kind: 'first' | 'hanging' | 'left') {
    if (event.button !== 0) return;
    event.preventDefault(); event.stopPropagation();
    (event.currentTarget as HTMLElement).setPointerCapture(event.pointerId);
    drag = { kind, start: event.clientX, delta: 0, pointer: event.pointerId };
  }
  function moveIndent(event: PointerEvent) {
    if (!drag || event.pointerId !== drag.pointer || !text) return;
    let delta = Math.round((event.clientX - drag.start) / emuToPixel);
    // The paragraph margin is nonnegative; first-line offsets may hang left.
    if (drag.kind !== 'first') delta = Math.max(-text.left, delta);
    const limit = 51206400;
    if (drag.kind === 'first') delta = Math.max(-limit - text.first, Math.min(limit - text.first, delta));
    else {
      delta = Math.min(limit - text.left, delta);
      if (drag.kind === 'hanging') delta = Math.max(text.first - limit, Math.min(text.first + limit, delta));
    }
    drag.delta = delta;
  }
  $effect(() => {
    if (!drag && !tabDrag) return;
    function cancel(event: KeyboardEvent) {
      if (event.key !== 'Escape') return;
      event.preventDefault(); event.stopPropagation(); drag = null; tabDrag = null;
    }
    document.addEventListener('keydown', cancel, true);
    return () => document.removeEventListener('keydown', cancel, true);
  });
  function finishIndent(event: PointerEvent) {
    if (!drag || event.pointerId !== drag.pointer) return;
    moveIndent(event);
    const change = drag;
    drag = null;
    if (change.delta) onindent?.(change.kind, change.delta);
  }
  let root: HTMLDivElement;
  let bounds = $state({ x: 0, y: 0, width: 0, height: 0, slideWidth: 0, slideHeight: 0 });
  const thickness = 22;
  const pixelsPerCentimeter = $derived(96 * zoom / 2.54);
  // Keep subdivisions legible when zoomed out without allocating off-screen ticks.
  const interval = $derived(pixelsPerCentimeter >= 24 ? 0.25 : pixelsPerCentimeter >= 12 ? 0.5 : pixelsPerCentimeter >= 6 ? 1 : 2);
  const labelInterval = $derived(pixelsPerCentimeter >= 24 ? 1 : pixelsPerCentimeter >= 12 ? 2 : pixelsPerCentimeter >= 6 ? 5 : 10);
  function ticks(offset: number, length: number, viewport: number, center = offset + length / 2) {
    const first = Math.ceil((Math.max(thickness, offset) - center) / pixelsPerCentimeter / interval);
    const last = Math.floor((Math.min(viewport, offset + length) - center) / pixelsPerCentimeter / interval);
    return Array.from({ length: Math.max(0, last - first + 1) }, (_, i) => {
      const value = (first + i) * interval;
      return { value, position: center + value * pixelsPerCentimeter, major: Number.isInteger(value), labeled: value % labelInterval === 0 };
    });
  }
  const horizontal = $derived(ticks(bounds.x, bounds.slideWidth, bounds.width, text ? textBounds.x : undefined));
  const vertical = $derived(ticks(bounds.y, bounds.slideHeight, bounds.height, text ? textBounds.y : undefined));
  $effect(() => {
    text;
    function measure() {
      const origin = root.getBoundingClientRect();
      const slide = stage.getBoundingClientRect();
      const input = text ? area.querySelector<HTMLElement>('.inline-edit') : null;
      if (input) {
        const rect = input.getBoundingClientRect();
        const style = getComputedStyle(input);
        const scaleX = rect.width / input.offsetWidth;
        const scaleY = rect.height / input.offsetHeight;
        textBounds = { x: rect.left - origin.left + parseFloat(style.paddingLeft) * scaleX, y: rect.top - origin.top + parseFloat(style.paddingTop) * scaleY, scale: scaleX };
      }
      bounds = { x: slide.left - origin.left, y: slide.top - origin.top, width: area.clientWidth + thickness, height: area.clientHeight + thickness, slideWidth: slide.width, slideHeight: slide.height };
    }
    const observer = new ResizeObserver(measure);
    observer.observe(area);
    observer.observe(stage);
    area.addEventListener('scroll', measure, { passive: true });
    measure();
    return () => { observer.disconnect(); area.removeEventListener('scroll', measure); };
  });
</script>

<div class="rulers" bind:this={root}>
  <svg class="horizontal" role="img" aria-label={t('Horizontal ruler (centimeters)')} width={bounds.width} height={thickness}>
    <rect x={bounds.x} y="0" width={bounds.slideWidth} height={thickness} />
    {#each horizontal as mark}
      <line data-value={mark.value} x1={mark.position} x2={mark.position} y1={mark.major ? 16 : 18} y2={thickness} />
      {#if mark.labeled}<text x={mark.position} y="12" text-anchor="middle">{Math.abs(mark.value)}</text>{/if}
    {/each}
  </svg>
  <svg class="vertical" role="img" aria-label={t('Vertical ruler (centimeters)')} width={thickness} height={bounds.height}>
    <rect x="0" y={bounds.y} width={thickness} height={bounds.slideHeight} />
    {#each vertical as mark}
      <line data-value={mark.value} y1={mark.position} y2={mark.position} x1={mark.major ? 16 : 18} x2={thickness} />
      {#if mark.labeled}<text transform={`translate(9 ${mark.position}) rotate(-90)`} text-anchor="middle">{Math.abs(mark.value)}</text>{/if}
    {/each}
  </svg>
  {#if text}
    <button class="tab-track" style:left={`${Math.max(thickness, textBounds.x)}px`} style:width={`${Math.max(0, Math.min(bounds.width, bounds.x + bounds.slideWidth) - Math.max(thickness, textBounds.x))}px`} aria-label={t('Add tab stop')} onpointerdown={addTab}></button>
    {#each text.tabStops as stop (stop.positionEmu)}
      {@const moving = tabDrag?.stop.positionEmu === stop.positionEmu}
      <button class="tab-stop" class:removing={moving && tabDrag?.remove} aria-label={`${t('Tab stop')} ${stop.positionEmu / 360000} cm`} title={`${t(tabKinds.find(kind => kind.alignment === stop.alignment)!.label)}: ${Math.round(stop.positionEmu / 3600) / 100} cm`}
        style:left={`${textBounds.x + (moving ? tabDrag!.position : stop.positionEmu) * emuToPixel}px`}
        onpointerdown={event => {
          if (event.button !== 0) return;
          event.preventDefault(); event.stopPropagation();
          event.currentTarget.setPointerCapture(event.pointerId);
          tabDrag = { stop, position: stop.positionEmu, remove: false, pointer: event.pointerId };
        }}
        onpointermove={moveTab} onpointerup={finishTab} onpointercancel={() => tabDrag = null} onlostpointercapture={() => tabDrag = null}
        onkeydown={event => {
          if (event.key === 'Delete' || event.key === 'Backspace') {
            event.preventDefault(); event.stopPropagation(); ontabs?.([{ kind: 'clear', positionEmu: stop.positionEmu }]);
          } else if (event.key === 'ArrowLeft' || event.key === 'ArrowRight') {
            event.preventDefault(); event.stopPropagation();
            ontabs?.([{ kind: 'clear', positionEmu: stop.positionEmu }, { kind: 'set', stop: { ...stop, positionEmu: Math.max(0, Math.min(51206400, stop.positionEmu + (event.key === 'ArrowLeft' ? -36000 : 36000))) } }]);
          }
        }}><svg viewBox="0 0 14 14" aria-hidden="true"><path d={tabKinds.find(kind => kind.alignment === stop.alignment)!.path} /></svg></button>
    {/each}
    {#each handles as handle}
      <button class="indent {handle.kind}" aria-label={t(handle.label)} title={t(handle.label)}
        style:left={`${handle.kind === 'first' ? firstPosition : leftPosition}px`}
        onpointerdown={event => startIndent(event, handle.kind)}
        onpointermove={moveIndent} onpointerup={finishIndent} onpointercancel={() => drag = null}
        onlostpointercapture={() => drag = null}
        onkeydown={event => {
          if (event.key === 'ArrowLeft' || event.key === 'ArrowRight') {
            event.preventDefault(); event.stopPropagation();
            onindent?.(handle.kind, (event.key === 'ArrowLeft' ? -1 : 1) * 36000);
          }
        }}></button>
    {/each}
  {/if}
  {#if text}
    <button class="corner tab-selector" aria-label={t(tabKinds[tabKind]!.label)} title={t(tabKinds[tabKind]!.label)} onpointerdown={event => event.preventDefault()} onclick={() => tabKind = (tabKind + 1) % tabKinds.length}><svg viewBox="0 0 14 14" aria-hidden="true"><path d={tabKinds[tabKind]!.path} /></svg></button>
  {:else}<div class="corner"></div>{/if}
</div>

<style>
  .rulers { position: absolute; inset: 0; overflow: hidden; pointer-events: none; z-index: 2; }
  svg { position: absolute; left: 0; top: 0; background: var(--ok-canvas-bg); overflow: hidden; }
  rect, .corner { fill: var(--ok-panel); background: var(--ok-panel); }
  line { stroke: var(--ok-text-2); stroke-width: 1; }
  text { fill: var(--ok-text-2); font: 10px sans-serif; }
  .tab-track { position: absolute; top: 0; height: 22px; padding: 0; border: 0; background: transparent; pointer-events: auto; }
  .tab-stop { position: absolute; top: 8px; width: 14px; height: 14px; margin-left: -7px; padding: 0; border: 0; background: transparent; pointer-events: auto; touch-action: none; cursor: ew-resize; }
  .tab-stop.removing { opacity: 0.3; }
  .tab-stop svg, .tab-selector svg { position: static; width: 14px; height: 14px; background: transparent; }
  .tab-stop path, .tab-selector path { stroke: var(--ok-text); stroke-width: 2; fill: none; }
  .tab-selector { pointer-events: auto; padding: 3px; border: 1px solid var(--ok-border); }
  .indent { position: absolute; margin: 0 0 0 -5px; padding: 0; width: 10px; height: 8px; border: 1px solid var(--ok-text-2); background: var(--ok-panel); pointer-events: auto; touch-action: none; cursor: ew-resize; }
  .indent.first { top: 0; clip-path: polygon(0 0,100% 0,100% 40%,50% 100%,0 40%); }
  .indent.hanging { top: 9px; clip-path: polygon(50% 0,100% 60%,100% 100%,0 100%,0 60%); }
  .indent.left { top: 17px; height: 5px; }
  .corner { position: absolute; left: 0; top: 0; width: 22px; height: 22px; }
</style>
