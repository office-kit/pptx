<script lang="ts">
  import type { ParagraphTabStop } from '@office-kit/pptx';
  import type { TabStopEdit } from '../core/paragraph-tabs.ts';
  import { MAX_INDENT_EMU, rulerAxis, type IndentHandle, type RulerAxis, type RulerChange, type TextFlow } from '../core/ruler.ts';
  import { elementScreenMatrix } from './element-matrix.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  let { area, stage, zoom, text = null, onindent, ontabs, onpreview }: {
    area: HTMLElement; stage: HTMLElement; zoom: number;
    /** The last selected paragraph's indentation and tabs; `scale` is the autofit factor. Null outside text editing, where the rulers measure the slide from its centre with no markers. */
    text?: { left: number; first: number; scale: number; flow: TextFlow; tabStops: readonly ParagraphTabStop[] } | null;
    ontabs?: (edits: TabStopEdit[]) => void;
    onindent?: (kind: IndentHandle, delta: number) => void;
    /** The in-progress gesture, so text can reflow before it is committed; null ends it. */
    onpreview?: (change: RulerChange | null) => void;
  } = $props();
  const tabKinds = [
    { alignment: 'left', label: 'Left tab', path: 'M7 3v8h6' },
    { alignment: 'center', label: 'Center tab', path: 'M7 3v8M2 11h10' },
    { alignment: 'right', label: 'Right tab', path: 'M7 3v8H1' },
    { alignment: 'decimal', label: 'Decimal tab', path: 'M7 3v8M2 11h10M11 5h1' },
  ] as const;
  const keyStep = 36000;
  let tabKind = $state(0);
  let tabDrag = $state<{ stop: ParagraphTabStop; position: number; remove: boolean; pointer: number } | null>(null);
  let axis = $state<RulerAxis>({ axis: 'x', origin: { x: 0, y: 0 }, scale: 1 });
  const vertical = $derived(axis.axis === 'y');
  /** Signed screen pixels per EMU along the text's inline axis. */
  const emuToPixel = $derived(axis.scale * (text?.scale ?? 1) / 9525);
  const along = (event: PointerEvent) => (vertical ? event.clientY : event.clientX);
  const across = (event: PointerEvent) => {
    const rect = root.getBoundingClientRect();
    return vertical ? event.clientX - rect.left : event.clientY - rect.top;
  };
  const position = (emu: number) => axis.origin[axis.axis] + emu * emuToPixel;
  function tabPosition(event: PointerEvent) {
    const rect = root.getBoundingClientRect();
    const start = vertical ? rect.top : rect.left;
    return Math.max(0, Math.min(MAX_INDENT_EMU, Math.round((along(event) - start - axis.origin[axis.axis]) / emuToPixel)));
  }
  function addTab(event: PointerEvent) {
    if (event.button !== 0 || !text) return;
    event.preventDefault();
    ontabs?.([{ kind: 'set', stop: { positionEmu: tabPosition(event), alignment: tabKinds[tabKind]!.alignment } }]);
  }
  function tabEdits(drag: NonNullable<typeof tabDrag>): TabStopEdit[] {
    return drag.remove
      ? [{ kind: 'clear', positionEmu: drag.stop.positionEmu }]
      : [{ kind: 'move', fromEmu: drag.stop.positionEmu, toEmu: drag.position }];
  }
  function moveTab(event: PointerEvent) {
    if (!tabDrag || event.pointerId !== tabDrag.pointer) return;
    const cross = across(event);
    const next = { position: tabPosition(event), remove: cross < -thickness || cross > thickness * 2 };
    if (next.position === tabDrag.position && next.remove === tabDrag.remove) return;
    Object.assign(tabDrag, next);
    onpreview?.({ kind: 'tabs', edits: tabEdits(tabDrag) });
  }
  function finishTab(event: PointerEvent) {
    if (!tabDrag || event.pointerId !== tabDrag.pointer) return;
    moveTab(event);
    const change = tabDrag;
    tabDrag = null;
    onpreview?.(null);
    if (!change.remove && change.position === change.stop.positionEmu) return;
    ontabs?.(tabEdits(change));
  }
  let drag = $state<{ kind: IndentHandle; start: number; delta: number; pointer: number } | null>(null);
  const firstPosition = $derived(position((text?.left ?? 0) + (text?.first ?? 0) + (drag && drag.kind !== 'hanging' ? drag.delta : 0)));
  const leftPosition = $derived(position((text?.left ?? 0) + (drag && drag.kind !== 'first' ? drag.delta : 0)));
  const handles = [{ kind: 'first', label: 'First line indent' }, { kind: 'hanging', label: 'Hanging indent' }, { kind: 'left', label: 'Left indent' }] as const;
  function startIndent(event: PointerEvent, kind: IndentHandle) {
    if (event.button !== 0) return;
    event.preventDefault(); event.stopPropagation();
    (event.currentTarget as HTMLElement).setPointerCapture(event.pointerId);
    drag = { kind, start: along(event), delta: 0, pointer: event.pointerId };
  }
  function moveIndent(event: PointerEvent) {
    if (!drag || event.pointerId !== drag.pointer || !text) return;
    let delta = Math.round((along(event) - drag.start) / emuToPixel);
    // The paragraph margin is nonnegative; first-line offsets may hang left.
    if (drag.kind !== 'first') delta = Math.max(-text.left, delta);
    if (drag.kind === 'first') delta = Math.max(-MAX_INDENT_EMU - text.first, Math.min(MAX_INDENT_EMU - text.first, delta));
    else {
      delta = Math.min(MAX_INDENT_EMU - text.left, delta);
      if (drag.kind === 'hanging') delta = Math.max(text.first - MAX_INDENT_EMU, Math.min(text.first + MAX_INDENT_EMU, delta));
    }
    if (delta === drag.delta) return;
    drag.delta = delta;
    onpreview?.({ kind: 'indent', handle: drag.kind, delta });
  }
  function cancelDrag() {
    if (!drag && !tabDrag) return;
    drag = null; tabDrag = null;
    onpreview?.(null);
  }
  $effect(() => {
    if (!drag && !tabDrag) return;
    function cancel(event: KeyboardEvent) {
      if (event.key !== 'Escape') return;
      event.preventDefault(); event.stopPropagation(); cancelDrag();
    }
    document.addEventListener('keydown', cancel, true);
    return () => document.removeEventListener('keydown', cancel, true);
  });
  function finishIndent(event: PointerEvent) {
    if (!drag || event.pointerId !== drag.pointer) return;
    moveIndent(event);
    const change = drag;
    drag = null;
    onpreview?.(null);
    if (change.delta) onindent?.(change.kind, change.delta);
  }
  /** EMU change for an arrow key, in the direction the marker moves on screen. */
  function arrowStep(event: KeyboardEvent): number | null {
    const [back, forward] = vertical ? ['ArrowUp', 'ArrowDown'] : ['ArrowLeft', 'ArrowRight'];
    if (event.key !== back && event.key !== forward) return null;
    event.preventDefault(); event.stopPropagation();
    return (event.key === forward ? 1 : -1) * Math.sign(emuToPixel) * keyStep;
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
  const horizontal = $derived(ticks(bounds.x, bounds.slideWidth, bounds.width, text ? axis.origin.x : undefined));
  const verticalTicks = $derived(ticks(bounds.y, bounds.slideHeight, bounds.height, text ? axis.origin.y : undefined));
  // The tab track runs from the paragraph start to the slide edge it faces.
  const track = $derived.by(() => {
    const start = vertical ? bounds.y : bounds.x;
    const end = start + (vertical ? bounds.slideHeight : bounds.slideWidth);
    const viewport = vertical ? bounds.height : bounds.width;
    const origin = axis.origin[axis.axis];
    const [low, high] = emuToPixel > 0 ? [origin, end] : [start, origin];
    return { start: Math.max(thickness, low), length: Math.max(0, Math.min(viewport, high) - Math.max(thickness, low)) };
  });
  $effect(() => {
    text;
    function measure() {
      const origin = root.getBoundingClientRect();
      const slide = stage.getBoundingClientRect();
      const input = text ? area.querySelector<HTMLElement>('.inline-edit') : null;
      if (input && text) {
        const style = getComputedStyle(input);
        const top = input.clientTop + parseFloat(style.paddingTop);
        // The inline start of the first line: top-right for vertical-rl.
        const start = style.writingMode === 'vertical-rl'
          ? { x: input.clientLeft + input.clientWidth - parseFloat(style.paddingRight), y: top }
          : { x: input.clientLeft + parseFloat(style.paddingLeft), y: top };
        const inline = style.writingMode.startsWith('vertical') ? { x: 0, y: 1 } : { x: 1, y: 0 };
        const matrix = elementScreenMatrix(input);
        const a = matrix.transformPoint(new DOMPoint(start.x, start.y));
        const b = matrix.transformPoint(new DOMPoint(start.x + inline.x, start.y + inline.y));
        const box = (input.closest<HTMLElement>('.inline-edit-shell') ?? input).getBoundingClientRect();
        axis = rulerAxis({
          origin: { x: a.x - origin.left, y: a.y - origin.top },
          direction: { x: b.x - a.x, y: b.y - a.y },
          center: { x: box.left + box.width / 2 - origin.left, y: box.top + box.height / 2 - origin.top },
          flow: text.flow,
        });
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
    {#each verticalTicks as mark}
      <line data-value={mark.value} y1={mark.position} y2={mark.position} x1={mark.major ? 16 : 18} x2={thickness} />
      {#if mark.labeled}<text transform={`translate(9 ${mark.position}) rotate(-90)`} text-anchor="middle">{Math.abs(mark.value)}</text>{/if}
    {/each}
  </svg>
  {#if text}
    <!-- Markers are laid out along a horizontal strip; the vertical ruler
         transposes it so the outer edge and marker order stay the same. -->
    <div class="text-axis" class:vertical data-axis={axis.axis}>
      <button class="tab-track" style:left={`${track.start}px`} style:width={`${track.length}px`} aria-label={t('Add tab stop')} onpointerdown={addTab}></button>
      {#each text.tabStops as stop (stop.positionEmu)}
        {@const moving = tabDrag?.stop.positionEmu === stop.positionEmu}
        <button class="tab-stop" class:removing={moving && tabDrag?.remove} aria-label={`${t('Tab stop')} ${stop.positionEmu / 360000} cm`} title={`${t(tabKinds.find(kind => kind.alignment === stop.alignment)!.label)}: ${Math.round(stop.positionEmu / 3600) / 100} cm`}
          style:left={`${position(moving ? tabDrag!.position : stop.positionEmu)}px`}
          onpointerdown={event => {
            if (event.button !== 0) return;
            event.preventDefault(); event.stopPropagation();
            event.currentTarget.setPointerCapture(event.pointerId);
            tabDrag = { stop, position: stop.positionEmu, remove: false, pointer: event.pointerId };
          }}
          onpointermove={moveTab} onpointerup={finishTab} onpointercancel={cancelDrag} onlostpointercapture={cancelDrag}
          onkeydown={event => {
            if (event.key === 'Delete' || event.key === 'Backspace') {
              event.preventDefault(); event.stopPropagation(); ontabs?.([{ kind: 'clear', positionEmu: stop.positionEmu }]);
              return;
            }
            const step = arrowStep(event);
            if (step !== null) ontabs?.([{ kind: 'move', fromEmu: stop.positionEmu, toEmu: Math.max(0, Math.min(MAX_INDENT_EMU, stop.positionEmu + step)) }]);
          }}><svg viewBox="0 0 14 14" aria-hidden="true"><path d={tabKinds.find(kind => kind.alignment === stop.alignment)!.path} /></svg></button>
      {/each}
      {#each handles as handle}
        <button class="indent {handle.kind}" aria-label={t(handle.label)} title={t(handle.label)}
          style:left={`${handle.kind === 'first' ? firstPosition : leftPosition}px`}
          onpointerdown={event => startIndent(event, handle.kind)}
          onpointermove={moveIndent} onpointerup={finishIndent} onpointercancel={cancelDrag}
          onlostpointercapture={cancelDrag}
          onkeydown={event => {
            const step = arrowStep(event);
            if (step !== null) onindent?.(handle.kind, step);
          }}></button>
      {/each}
    </div>
    <button class="corner tab-selector" aria-label={t(tabKinds[tabKind]!.label)} title={t(tabKinds[tabKind]!.label)} onpointerdown={event => event.preventDefault()} onclick={() => tabKind = (tabKind + 1) % tabKinds.length}><svg viewBox="0 0 14 14" aria-hidden="true"><path d={tabKinds[tabKind]!.path} /></svg></button>
  {:else}<div class="corner"></div>{/if}
</div>

<style>
  .rulers { position: absolute; inset: 0; overflow: hidden; pointer-events: none; z-index: 2; }
  svg { position: absolute; left: 0; top: 0; background: var(--ok-canvas-bg); overflow: hidden; }
  rect, .corner { fill: var(--ok-panel); background: var(--ok-panel); }
  line { stroke: var(--ok-text-2); stroke-width: 1; }
  text { fill: var(--ok-text-2); font: 10px sans-serif; }
  .text-axis { position: absolute; left: 0; top: 0; width: 0; height: 22px; }
  /* Transpose: strip x becomes screen y and the strip's outer edge stays outermost. */
  .text-axis.vertical { transform: matrix(0, 1, 1, 0, 0, 0); transform-origin: 0 0; }
  .tab-track { position: absolute; top: 0; height: 22px; padding: 0; border: 0; background: transparent; pointer-events: auto; }
  .tab-stop { position: absolute; top: 8px; width: 14px; height: 14px; margin-left: -7px; padding: 0; border: 0; background: transparent; pointer-events: auto; touch-action: none; cursor: ew-resize; }
  .vertical .tab-stop, .vertical .indent { cursor: ns-resize; }
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
