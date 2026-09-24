<script lang="ts">
  import { tick } from 'svelte';
  import { getSlideSize, setDrawingGuides, type DrawingGuide } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  let { guides, scaleX, scaleY }: { guides: readonly DrawingGuide[]; scaleX: number; scaleY: number } = $props();
  const editor = getEditor();
  const doc = editor.doc;
  let drag = $state<{ id: number; pointer: number; origin: number; start: number; position: number; version: number } | null>(null);
  let menu = $state<{ id: number; x: number; y: number } | null>(null);
  let colorMenu = $state(false);
  let menuElement = $state<HTMLDivElement>();
  const colors = [['Gray', '#888888'], ['Red', '#c43c3c'], ['Orange', '#e48312'], ['Yellow', '#c4a300'], ['Green', '#25833a'], ['Blue', '#2873c4'], ['Purple', '#8042a8']];
  function edit(next: readonly DrawingGuide[]) {
    doc.transact(t('Edit guides'), () => setDrawingGuides(doc.pres, next));
    menu = null;
  }
  function add(axis: 'x' | 'y') {
    const size = getSlideSize(doc.pres);
    if (!size) return;
    const id = Math.max(0, ...guides.map(item => item.id)) + 1;
    const position = (axis === 'x' ? size.width : size.height) / 2;
    edit([...guides, { id, axis, position: position + 91440 * guides.filter(item => item.axis === axis).length, color: '#888888' }]);
  }
  function context(event: MouseEvent, guide: DrawingGuide) {
    event.preventDefault(); event.stopPropagation();
    menu = { id: guide.id, x: Math.max(0, Math.min(event.clientX, window.innerWidth - 330)), y: Math.max(0, Math.min(event.clientY, window.innerHeight - 240)) };
    colorMenu = false;
    void tick().then(() => menuElement?.querySelector<HTMLButtonElement>('button')?.focus());
  }
  function menuKeys(event: KeyboardEvent) {
    const target = event.target as HTMLElement;
    if (event.key === 'ArrowUp' || event.key === 'ArrowDown') {
      event.preventDefault(); event.stopPropagation();
      const buttons = [...(target.closest('[role="menu"]')?.querySelectorAll<HTMLButtonElement>(':scope > button') ?? [])];
      const index = buttons.indexOf(target as HTMLButtonElement);
      buttons[(index + (event.key === 'ArrowUp' ? -1 : 1) + buttons.length) % buttons.length]?.focus();
    }
  }
  function start(event: PointerEvent, guide: DrawingGuide) {
    if (event.button !== 0) return;
    event.preventDefault();
    event.stopPropagation();
    (event.currentTarget as HTMLElement).setPointerCapture(event.pointerId);
    drag = { id: guide.id, pointer: event.pointerId, origin: guide.position, start: guide.axis === 'x' ? event.clientX : event.clientY, position: guide.position, version: doc.version };
  }
  function move(event: PointerEvent, guide: DrawingGuide) {
    if (!drag || drag.pointer !== event.pointerId || drag.id !== guide.id) return;
    const cursor = guide.axis === 'x' ? event.clientX : event.clientY;
    drag.position = drag.origin + (cursor - drag.start) / (guide.axis === 'x' ? scaleX : scaleY);
  }
  function finish(event: PointerEvent, guide: DrawingGuide) {
    if (!drag || drag.pointer !== event.pointerId) return;
    move(event, guide);
    const current = drag;
    drag = null;
    if (current.version !== doc.version || Math.abs(current.position - current.origin) < 1) return;
    doc.transact(t('Move guide'), () => setDrawingGuides(doc.pres, guides.map(item => item.id === current.id ? { ...item, position: current.position } : item)));
  }
  function keydown(event: KeyboardEvent, guide: DrawingGuide) {
    if (event.key === 'Escape') { drag = null; return; }
    const backward = guide.axis === 'x' ? 'ArrowLeft' : 'ArrowUp';
    const forward = guide.axis === 'x' ? 'ArrowRight' : 'ArrowDown';
    if (![backward, forward, 'Delete', 'Backspace'].includes(event.key)) return;
    event.preventDefault(); event.stopPropagation();
    doc.transact(t('Edit guides'), () => setDrawingGuides(doc.pres,
      event.key === 'Delete' || event.key === 'Backspace'
        ? guides.filter(item => item.id !== guide.id)
        : guides.map(item => item.id === guide.id ? { ...item, position: item.position + (event.key === backward ? -1 : 1) * (event.shiftKey ? 91440 : 9144) } : item)));
  }
</script>
<svelte:window onpointerdown={event => { if (menu && !menuElement?.contains(event.target as Node)) menu = null; }} onkeydown={event => { if (event.key === 'Escape' && menu) menu = null; if (event.key === 'Escape' && drag) { event.preventDefault(); event.stopPropagation(); drag = null; } }} />
{#each guides as guide (guide.id)}
  {@const position = drag?.id === guide.id ? drag.position : guide.position}
  <button class="drawing-guide" class:vertical={guide.axis === 'x'} aria-label={t(guide.axis === 'x' ? 'Vertical guide' : 'Horizontal guide')}
    style={guide.axis === 'x' ? `left:${position * scaleX - 3}px;top:0;height:100%;--guide-color:${guide.color}` : `top:${position * scaleY - 3}px;left:0;width:100%;--guide-color:${guide.color}`}
    oncontextmenu={event => context(event, guide)} onpointerdown={event => start(event, guide)} onpointermove={event => move(event, guide)} onpointerup={event => finish(event, guide)} onpointercancel={() => drag = null} onlostpointercapture={() => drag = null} onkeydown={event => keydown(event, guide)}>
    <span></span>
  </button>
{/each}
{#if menu}
  <div class="guide-menu" role="menu" tabindex="-1" onkeydown={menuKeys} aria-label={t('Guides')} bind:this={menuElement} style="left:{menu.x}px;top:{menu.y}px" oncontextmenu={event => event.preventDefault()}>
    <button role="menuitem" onclick={() => add('x')}>{t('Add Vertical Guide')}</button>
    <button role="menuitem" onclick={() => add('y')}>{t('Add Horizontal Guide')}</button>
    <button role="menuitem" onclick={() => edit(guides.filter(item => item.id !== menu?.id))}>{t('Delete')}</button>
    <button role="menuitem" aria-haspopup="menu" aria-expanded={colorMenu} onclick={() => { colorMenu = !colorMenu; if (colorMenu) void tick().then(() => menuElement?.querySelector<HTMLButtonElement>('.colors button')?.focus()); }}>{t('Color')}</button>
    {#if colorMenu}<div class="colors" role="menu" aria-label={t('Color')}>
      {#each colors as [label, color]}
        <button role="menuitem" onclick={() => edit(guides.map(item => item.id === menu?.id ? { ...item, color: color! } : item))}><span style="background:{color}"></span>{t(label!)}</button>
      {/each}
    </div>{/if}
  </div>
{/if}
<style>
  .guide-menu { position: fixed; z-index: 150; pointer-events: auto; min-width: 185px; padding: 5px; border: 1px solid var(--ok-border); border-radius: 6px; background: var(--ok-panel); color: var(--ok-text); box-shadow: var(--ok-shadow-lg); }
  .guide-menu button { display: flex; align-items: center; gap: 8px; width: 100%; padding: 5px 8px; border: 0; border-radius: 4px; background: transparent; color: inherit; text-align: left; font: inherit; font-size: 12px; }
  .guide-menu button:hover, .guide-menu button:focus-visible { background: var(--ok-accent); color: white; }
  .colors { position: absolute; left: 100%; bottom: 0; min-width: 110px; border: 1px solid var(--ok-border); padding: 5px; background: var(--ok-panel); border-radius: 6px; }
  .colors span { width: 10px; height: 10px; border: 1px solid #888; }

  .drawing-guide { position: absolute; z-index: 5; pointer-events: auto; background: transparent; padding: 3px 0; border: 0; height: 7px; cursor: ns-resize; touch-action: none; }
  .drawing-guide span { display: block; border-top: 1px dashed var(--guide-color); }
  .drawing-guide.vertical { width: 7px; padding: 0 3px; cursor: ew-resize; }
  .drawing-guide.vertical span { height: 100%; border-top: 0; border-left: 1px dashed var(--guide-color); }
  .drawing-guide:focus-visible { outline: 1px solid var(--ok-accent); }
</style>
