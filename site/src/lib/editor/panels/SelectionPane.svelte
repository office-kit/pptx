<script lang="ts">
  import { onDestroy } from 'svelte';
  import { getGroupChildren, getShapeId, getShapeName, isShapeHidden, renameShape, setShapeHidden, setShapeZIndex, type SlideShapeData } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { selectedShapeIds, topLevelShapes } from '../core/selection.ts';
  import { t } from '../i18n/i18n.svelte.ts';

  const editor = getEditor();
  const doc = editor.doc;
  interface Row { shape: SlideShapeData; id: number; name: string; hidden: boolean; parent: number | null; depth: number; group: boolean; }
  let expanded = $state<number[]>([]);
  const expandedIds = $derived(new Set(expanded));
  let renaming = $state<number | null>(null);
  let name = $state('');
  let anchor = $state<number | null>(null);
  let dragging = $state<number | null>(null);
  let insertion = $state<{ id: number; after: boolean } | null>(null);
  let list: HTMLDivElement;
  const SCROLL_EDGE = 32;
  const SCROLL_SPEED = 700;
  const MAX_FRAME_MS = 50;
  let scrollSpeed = 0;
  let scrollFrame = 0;
  let previousFrame = 0;

  function stopScroll() {
    cancelAnimationFrame(scrollFrame);
    scrollFrame = 0; scrollSpeed = 0; previousFrame = 0;
  }
  function scroll(time: number) {
    const elapsed = previousFrame ? Math.min(MAX_FRAME_MS, time - previousFrame) : 0;
    previousFrame = time;
    list.scrollTop += scrollSpeed * elapsed / 1000;
    scrollFrame = requestAnimationFrame(scroll);
  }
  function edgeScroll(event: DragEvent) {
    if (dragging === null) return;
    const bounds = list.getBoundingClientRect();
    if (event.clientX < bounds.left || event.clientX > bounds.right || event.clientY < bounds.top || event.clientY > bounds.bottom) { stopScroll(); return; }
    const top = Math.max(0, (SCROLL_EDGE - (event.clientY - bounds.top)) / SCROLL_EDGE);
    const bottom = Math.max(0, (SCROLL_EDGE - (bounds.bottom - event.clientY)) / SCROLL_EDGE);
    scrollSpeed = (bottom - top) * SCROLL_SPEED;
    if (!scrollSpeed) stopScroll();
    else if (!scrollFrame) scrollFrame = requestAnimationFrame(scroll);
  }
  function endDrag() { dragging = null; insertion = null; stopScroll(); }
  onDestroy(stopScroll);
  const selected = $derived(new Set(selectedShapeIds(doc.selection)));
  const allRows = $derived.by(() => {
    doc.version;
    const rows: Row[] = [];
    function visit(shapes: readonly SlideShapeData[], parent: number | null, depth: number) {
      for (const shape of [...shapes].reverse()) {
        const id = getShapeId(shape);
        const children = getGroupChildren(shape);
        rows.push({ shape, id, name: getShapeName(shape), hidden: isShapeHidden(shape), parent, depth, group: children.length > 0 });
        visit(children, id, depth + 1);
      }
    }
    if (doc.currentSlide) visit(topLevelShapes(doc.currentSlide), null, 0);
    return rows;
  });
  const allHidden = $derived(allRows.length > 0 && allRows.every(row => row.hidden));
  const rows = $derived.by(() => {
    const visible = new Set<number>();
    return allRows.filter(row => {
      if (row.parent !== null && (!visible.has(row.parent) || !expandedIds.has(row.parent))) return false;
      visible.add(row.id);
      return true;
    });
  });

  function choose(row: Row, event: MouseEvent | KeyboardEvent) {
    const parents = new Map(allRows.map(item => [item.id, item.parent]));
    const sameScope = [...selected].every(id => parents.get(id) === row.parent);
    const start = rows.findIndex(item => item.id === anchor && item.parent === row.parent);
    if (event.shiftKey && start >= 0) {
      const end = rows.indexOf(row);
      doc.selection = { kind: 'shape', slideIndex: doc.selection.slideIndex, shapeIds: rows.slice(Math.min(start, end), Math.max(start, end) + 1).filter(item => item.parent === row.parent).map(item => item.id) };
    } else {
      doc.selectShape(doc.selection.slideIndex, row.id, (event.metaKey || event.ctrlKey) && sameScope);
      anchor = row.id;
    }
  }
  function visibility(row: Row) {
    doc.transact(t(row.hidden ? 'Show object' : 'Hide object'), () => setShapeHidden(row.shape, !row.hidden));
  }
  function showAll(hidden: boolean) {
    const changed = allRows.filter(row => row.hidden !== hidden);
    if (changed.length) doc.transact(t(hidden ? 'Hide All' : 'Show All'), () => { for (const row of changed) setShapeHidden(row.shape, hidden); });
  }
  function finishRename(cancel = false) {
    const row = allRows.find(row => row.id === renaming);
    renaming = null;
    if (!cancel && row && row.name !== name) doc.transact(t('Rename object'), () => renameShape(row.shape, name));
  }
  function focusName(input: HTMLInputElement) { input.focus(); input.select(); }
  function dragStart(event: DragEvent, row: Row) {
    dragging = row.id;
    doc.selectShape(doc.selection.slideIndex, row.id);
    if (event.dataTransfer) {
      event.dataTransfer.effectAllowed = 'move';
      event.dataTransfer.setData('text/plain', row.name);
    }
  }
  function dragOver(event: DragEvent, row: Row) {
    const source = allRows.find(item => item.id === dragging);
    if (!source || source.id === row.id || source.parent !== row.parent) { insertion = null; return; }
    event.preventDefault();
    if (event.dataTransfer) event.dataTransfer.dropEffect = 'move';
    const bounds = (event.currentTarget as HTMLElement).getBoundingClientRect();
    insertion = { id: row.id, after: event.clientY > bounds.top + bounds.height / 2 };
  }
  function drop(event: DragEvent, row: Row) {
    const source = allRows.find(item => item.id === dragging);
    if (!source || !insertion || insertion.id !== row.id || source.parent !== row.parent) return;
    event.preventDefault();
    const siblings = allRows.filter(item => item.parent === row.parent && item.id !== source.id);
    const index = siblings.indexOf(row) + (insertion.after ? 1 : 0);
    const original = allRows.filter(item => item.parent === row.parent).indexOf(source);
    endDrag();
    if (index !== original) doc.transact(t('Reorder object'), () => setShapeZIndex(source.shape, siblings.length - index));
  }
  function key(event: KeyboardEvent, row: Row) {
    if (event.key === 'F2') {
      event.preventDefault(); event.stopPropagation(); name = row.name; renaming = row.id;
    } else if (['ArrowUp', 'ArrowDown', 'Home', 'End'].includes(event.key)) {
      event.preventDefault(); event.stopPropagation();
      const index = rows.indexOf(row);
      const next = event.key === 'Home' ? 0 : event.key === 'End' ? rows.length - 1 : Math.max(0, Math.min(rows.length - 1, index + (event.key === 'ArrowUp' ? -1 : 1)));
      const target = rows[next];
      if (target) { choose(target, event); (event.currentTarget as HTMLElement).closest('.objects')?.querySelector<HTMLButtonElement>(`[data-object-id="${target.id}"] .name`)?.focus(); }
    } else if (['ArrowLeft', 'ArrowRight'].includes(event.key)) {
      event.preventDefault(); event.stopPropagation();
      if (row.group) expanded = event.key === 'ArrowRight' ? [...new Set([...expanded, row.id])] : expanded.filter(id => id !== row.id);
    }
  }
</script>

<svelte:window ondragover={edgeScroll} ondragend={endDrag} ondrop={endDrag} />

<section class="panel" aria-label={t('Selection Pane')}>
  <header><strong>{t('Selection Pane')}</strong><button aria-label={t('Close Selection Pane')} onclick={() => editor.selectionPaneVisible = false}>×</button></header>
  <div class="actions"><button aria-label={t(allHidden ? 'Show All' : 'Hide All')} title={t(allHidden ? 'Show All' : 'Hide All')} disabled={!allRows.length} onclick={() => showAll(!allHidden)}><svg viewBox="0 0 20 16" aria-hidden="true"><path d="M1 8Q10 -3 19 8Q10 19 1 8Z"/><circle cx="10" cy="8" r="3"/>{#if allHidden}<path d="m2 1 16 14"/>{/if}</svg></button></div>
  <div class="objects ok-scroll" bind:this={list}>
    {#each rows as row (row.id)}
      <div class="row" class:selected={selected.has(row.id)} data-object-id={row.id} style:padding-left={`${row.depth * 14}px`}>
        {#if row.group}<button class="expand" aria-label={`${t(expandedIds.has(row.id) ? 'Collapse' : 'Expand')} ${row.name}`} aria-expanded={expandedIds.has(row.id)} onclick={() => expanded = expandedIds.has(row.id) ? expanded.filter(id => id !== row.id) : [...expanded, row.id]}>{expandedIds.has(row.id) ? '▾' : '▸'}</button>{:else}<span class="spacer"></span>{/if}
        {#if renaming === row.id}
          <input use:focusName aria-label={t('Object name')} bind:value={name} onblur={() => finishRename()} onkeydown={event => { event.stopPropagation(); if (event.key === 'Enter' || event.key === 'Escape') { event.preventDefault(); finishRename(event.key === 'Escape'); } }} />
        {:else}
          <button class="name" class:insert-before={insertion?.id === row.id && !insertion.after} class:insert-after={insertion?.id === row.id && insertion.after} draggable="true" ondragstart={event => dragStart(event, row)} ondragover={event => dragOver(event, row)} ondragleave={() => insertion = null} ondrop={event => drop(event, row)} ondragend={endDrag} title={row.name} aria-pressed={selected.has(row.id)} onclick={event => choose(row, event)} ondblclick={() => { name = row.name; renaming = row.id; }} onkeydown={event => key(event, row)}>{row.name}</button>
        {/if}
        <button class="visibility" aria-label={`${t(row.hidden ? 'Show object' : 'Hide object')}: ${row.name}`} aria-pressed={!row.hidden} onclick={() => visibility(row)}><svg viewBox="0 0 20 16" aria-hidden="true"><path d="M1 8Q10 -3 19 8Q10 19 1 8Z"/><circle cx="10" cy="8" r="3"/>{#if row.hidden}<path d="m2 1 16 14"/>{/if}</svg></button>
      </div>
    {/each}
  </div>
  <footer><button disabled={!editor.canRun('bringShapeForward')} onclick={() => editor.invoke('bringShapeForward')}>{t('Bring Forward')}</button><button disabled={!editor.canRun('sendShapeBackward')} onclick={() => editor.invoke('sendShapeBackward')}>{t('Send Backward')}</button></footer>
</section>

<style>
  .panel { display: flex; flex-direction: column; min-width: 0; min-height: 0; background: var(--ok-panel); border-left: 1px solid var(--ok-border); padding: 10px; }
  header { display: flex; align-items: center; justify-content: space-between; margin-bottom: 12px; }
  header button { font-size: 20px; }
  .actions { display: flex; justify-content: flex-end; gap: 8px; margin-bottom: 10px; }
  .objects { flex: 1; overflow: auto; min-height: 0; }
  footer { display: flex; gap: 6px; padding-top: 10px; }
  footer button { flex: 1; border-color: var(--ok-border); }
  .row { display: flex; align-items: center; min-height: 28px; }
  .row.selected { background: var(--ok-hover); }
  button { font: inherit; font-size: 12px; color: var(--ok-text); background: transparent; border: 1px solid transparent; border-radius: var(--ok-radius); padding: 4px; cursor: pointer; }
  button:hover { background: var(--ok-hover); }
  button:disabled { opacity: .4; cursor: default; }
  .name { flex: 1; min-width: 0; text-align: left; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
  .name.insert-before { border-top-color: var(--ok-accent); }
  .name.insert-after { border-bottom-color: var(--ok-accent); }
  .expand, .spacer { flex: 0 0 20px; }
  .visibility { flex: 0 0 26px; }
  svg { display: block; width: 18px; height: 16px; fill: none; stroke: currentColor; stroke-width: 1.3; }
  input { width: 0; min-width: 0; flex: 1; font: inherit; padding: 4px; }
</style>
