<script lang="ts">
  import { onMount, untrack } from 'svelte';
  import { getShapeId, getShapeName } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  const editor = getEditor();
  const doc = editor.doc;
  const initial = untrack(() => ({ version: doc.version, svg: doc.currentSvg, members: [...editor.reorderMembers()].reverse() }));
  const names = new Map(initial.members.map(shape => [getShapeId(shape), getShapeName(shape)]));
  let order = $state(initial.members.map(getShapeId));
  let previews = $state(new Map<number, string>());
  let active = $state<number | null>(null);
  let dialog: HTMLDialogElement;
  let stage: HTMLDivElement;
  let dragging: { id: number; startX: number; index: number } | null = null;
  onMount(() => {
    const source = new DOMParser().parseFromString(initial.svg, 'image/svg+xml').documentElement;
    const objects = new Map([...source.querySelectorAll('[data-pptx-shape-id]')].map(node => [Number(node.getAttribute('data-pptx-shape-id')), node]));
    const definitions = new Map([...source.querySelectorAll('defs [id]')].map(node => [node.id, node]));
    const definitionMarkup = new Map<string, string>();
    const serializer = new XMLSerializer();
    const images = new Map<number, string>();
    for (const id of order) {
      const object = objects.get(id);
      if (!object) continue;
      const root = source.cloneNode(false);
      let layer = object.cloneNode(true);
      // Ancestors carry the transforms of nested groups.
      for (let parent = object.parentElement; parent && parent !== source; parent = parent.parentElement) {
        const wrapper = parent.cloneNode(false);
        wrapper.appendChild(layer);
        layer = wrapper;
      }
      const references = (markup: string): string[] => [...markup.matchAll(/url\(#([^)]*)\)|(?:xlink:)?href="#([^"]*)"/g)].map(match => match[1] ?? match[2]!);
      const pending = references(serializer.serializeToString(layer));
      const included = new Set([...object.querySelectorAll('[id]')].map(node => node.id));
      const defs = source.ownerDocument.createElementNS(source.namespaceURI, 'defs');
      while (pending.length) {
        const reference = pending.pop()!;
        if (included.has(reference)) continue;
        included.add(reference);
        const definition = definitions.get(reference);
        if (!definition) continue;
        let markup = definitionMarkup.get(reference);
        if (markup === undefined) {
          markup = serializer.serializeToString(definition);
          definitionMarkup.set(reference, markup);
        }
        pending.push(...references(markup));
        defs.appendChild(definition.cloneNode(true));
      }
      root.appendChild(defs);
      root.appendChild(layer);
      images.set(id, `data:image/svg+xml;charset=utf-8,${encodeURIComponent(serializer.serializeToString(root))}`);
    }
    previews = images;
    dialog.showModal();
  });
  function move(id: number, target: number) {
    const from = order.indexOf(id);
    const to = Math.max(0, Math.min(target, order.length - 1));
    if (from < 0 || from === to) return;
    const next = [...order];
    next.splice(from, 1);
    next.splice(to, 0, id);
    order = next;
  }
  function start(event: PointerEvent, id: number) {
    if (event.button !== 0) return;
    active = id;
    dragging = { id, startX: event.clientX, index: order.indexOf(id) };
    (event.currentTarget as HTMLButtonElement).setPointerCapture(event.pointerId);
  }
  function drag(event: PointerEvent) {
    if (!dragging) return;
    const spacing = stage.clientWidth * .7 / Math.max(order.length, 2);
    move(dragging.id, dragging.index + Math.round((event.clientX - dragging.startX) / spacing));
  }
  function key(event: KeyboardEvent, id: number) {
    if (!['ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown', 'Home', 'End'].includes(event.key)) return;
    event.preventDefault();
    active = id;
    move(id, event.key === 'Home' ? 0 : event.key === 'End' ? order.length - 1 : order.indexOf(id) + (['ArrowLeft', 'ArrowUp'].includes(event.key) ? -1 : 1));
  }
  function confirm() {
    if (doc.version === initial.version) editor.reorderSelection(order);
    editor.closeDialog();
  }
</script>
<dialog bind:this={dialog} aria-label={t('Reorder Overlapping Objects')} onclose={() => editor.closeDialog()} onkeydown={event => { if (event.key === 'Enter' && (event.target === dialog || (event.target as HTMLElement).closest('.layer'))) { event.preventDefault(); confirm(); } }}>
  <div class="stage" bind:this={stage}>
    {#each order as id, index (id)}
      <button class="layer" class:active={active === id} aria-label={`${names.get(id)} ${t('at position')} ${index + 1}`} style:left={`${15 + index * 70 / Math.max(order.length, 2)}%`} style:z-index={order.length - index} style:width={`${Math.min(48, 100 / Math.max(order.length, 2))}%`} onpointerdown={event => start(event, id)} onpointermove={drag} onpointerup={() => dragging = null} onpointercancel={() => dragging = null} onkeydown={event => key(event, id)}>
        {#if previews.has(id)}<img src={previews.get(id)} alt="" draggable="false" />{/if}
        <span class="number">{index + 1}</span>
      </button>
    {/each}
  </div>
  <footer><button onclick={() => editor.closeDialog()}>{t('Cancel')}</button><button class="confirm" onclick={confirm}>{t('OK')}</button></footer>
</dialog>
<style>
  dialog { width: 100vw; height: 100dvh; max-width: none; max-height: none; margin: 0; padding: 0; border: 0; background: #0c0c0c; color: white; overflow: hidden; }
  dialog::backdrop { background: #0c0c0c; }
  .stage { position: absolute; inset: 0 0 45px; perspective: 1400px; }
  .layer { position: absolute; top: 23%; height: 55%; padding: 0; border: 5px solid #70b8d6aa; background: #6aaccc33; transform: rotateY(55deg); transform-origin: left center; cursor: grab; touch-action: none; transition: left .16s ease; }
  .layer.active, .layer:focus-visible { background: #389ac9aa; border-color: #70b8d6; outline: none; }
  .layer:active { cursor: grabbing; }
  .layer img { width: 100%; height: 100%; object-fit: fill; pointer-events: none; }
  .number { position: absolute; left: 18px; bottom: 10px; color: #b9e2f1; font-size: 60px; font-weight: 300; pointer-events: none; }
  footer { position: absolute; right: 18px; bottom: 12px; display: flex; gap: 10px; }
  footer button { min-width: 65px; border: 0; border-radius: 4px; background: #333; color: white; padding: 3px 10px; font-size: 12px; }
  footer .confirm { background: #b4440c; }
  @media (prefers-reduced-motion: reduce) { .layer { transition: none; } }
</style>
