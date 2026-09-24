<script lang="ts">
  import { getGroupChildren, getShapeId, getShapeName, isShapeHidden, renameShape, setShapeHidden, type SlideShapeData } from '@office-kit/pptx';
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

<section class="panel" aria-label={t('Selection Pane')}>
  <header><strong>{t('Selection Pane')}</strong><button aria-label={t('Close Selection Pane')} onclick={() => editor.selectionPaneVisible = false}>×</button></header>
  <div class="actions"><button disabled={!allRows.length} onclick={() => showAll(false)}>{t('Show All')}</button><button disabled={!allRows.length} onclick={() => showAll(true)}>{t('Hide All')}</button></div>
  <div class="objects ok-scroll">
    {#each rows as row (row.id)}
      <div class="row" class:selected={selected.has(row.id)} data-object-id={row.id} style:padding-left={`${row.depth * 14}px`}>
        {#if row.group}<button class="expand" aria-label={`${t(expandedIds.has(row.id) ? 'Collapse' : 'Expand')} ${row.name}`} aria-expanded={expandedIds.has(row.id)} onclick={() => expanded = expandedIds.has(row.id) ? expanded.filter(id => id !== row.id) : [...expanded, row.id]}>{expandedIds.has(row.id) ? '▾' : '▸'}</button>{:else}<span class="spacer"></span>{/if}
        {#if renaming === row.id}
          <input use:focusName aria-label={t('Object name')} bind:value={name} onblur={() => finishRename()} onkeydown={event => { event.stopPropagation(); if (event.key === 'Enter' || event.key === 'Escape') { event.preventDefault(); finishRename(event.key === 'Escape'); } }} />
        {:else}
          <button class="name" title={row.name} aria-pressed={selected.has(row.id)} onclick={event => choose(row, event)} ondblclick={() => { name = row.name; renaming = row.id; }} onkeydown={event => key(event, row)}>{row.name}</button>
        {/if}
        <button class="visibility" aria-label={`${t(row.hidden ? 'Show object' : 'Hide object')}: ${row.name}`} aria-pressed={!row.hidden} onclick={() => visibility(row)}><svg viewBox="0 0 20 16" aria-hidden="true"><path d="M1 8Q10 -3 19 8Q10 19 1 8Z"/><circle cx="10" cy="8" r="3"/>{#if row.hidden}<path d="m2 1 16 14"/>{/if}</svg></button>
      </div>
    {/each}
  </div>
</section>

<style>
  .panel { display: flex; flex-direction: column; min-width: 0; min-height: 0; background: var(--ok-panel); border-left: 1px solid var(--ok-border); padding: 10px; }
  header { display: flex; align-items: center; justify-content: space-between; margin-bottom: 12px; }
  header button { font-size: 20px; }
  .actions { display: flex; gap: 8px; margin-bottom: 10px; }
  .objects { overflow: auto; min-height: 0; }
  .row { display: flex; align-items: center; min-height: 28px; }
  .row.selected { background: var(--ok-hover); }
  button { font: inherit; font-size: 12px; color: var(--ok-text); background: transparent; border: 1px solid transparent; border-radius: var(--ok-radius); padding: 4px; cursor: pointer; }
  button:hover { background: var(--ok-hover); }
  button:disabled { opacity: .4; cursor: default; }
  .name { flex: 1; min-width: 0; text-align: left; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
  .expand, .spacer { flex: 0 0 20px; }
  .visibility { flex: 0 0 26px; }
  svg { display: block; width: 18px; height: 16px; fill: none; stroke: currentColor; stroke-width: 1.3; }
  input { width: 0; min-width: 0; flex: 1; font: inherit; padding: 4px; }
</style>
