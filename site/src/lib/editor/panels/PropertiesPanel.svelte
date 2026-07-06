<script lang="ts">
  // The right-hand format pane. It is generated from the manifest: given the
  // current selection it lists *every* capability that can act on it, grouped
  // by category. That makes the panel exhaustive by construction — the moment a
  // capability applies to the selection, it shows up here. Bespoke inline
  // editors (see BespokeSections) enrich the most common ones; the rest render
  // as a labelled row that opens the auto-generated dialog.
  import { getEditor } from '../core/context.ts';
  import {
    capabilities,
    CATEGORY_LABELS,
    CATEGORY_ORDER,
    type CategoryId,
    type ResolvedCapability,
  } from '../manifest/index.ts';
  import BespokeSections from './BespokeSections.svelte';
  import { t, capLabel, catLabel, getLocale } from '../i18n/i18n.svelte.ts';

  const editor = getEditor();
  const doc = editor.doc;

  // "run" (no args) / "N arg(s)" — English pluralizes, Japanese doesn't.
  function argLabel(paramCount: number): string {
    const n = paramCount - 1; // first param is the operand
    if (n <= 0) return t('run');
    return getLocale() === 'ja' ? `${n} 引数` : `${n} arg${n > 1 ? 's' : ''}`;
  }

  const applicable = $derived.by<ResolvedCapability[]>(() => {
    doc.version;
    doc.selection;
    return capabilities.filter((c) => c.canvas && editor.canRun(c.id));
  });

  const grouped = $derived.by(() => {
    const map = new Map<CategoryId, ResolvedCapability[]>();
    for (const cap of applicable) {
      const arr = map.get(cap.category) ?? [];
      arr.push(cap);
      map.set(cap.category, arr);
    }
    return CATEGORY_ORDER.filter((c) => map.has(c)).map((c) => ({
      category: c,
      label: CATEGORY_LABELS[c],
      items: map.get(c)!,
    }));
  });

  const selLabel = $derived.by(() => {
    const sel = doc.selection;
    if (sel.kind === 'shape')
      return sel.shapeIds.length > 1 ? `${sel.shapeIds.length} ${t('shapes')}` : t('Shape');
    if (sel.kind === 'cell') return t('Table cell');
    return t('Slide');
  });

  // Which category sections are expanded.
  let open = $state<Record<string, boolean>>({});
  function toggle(c: string) {
    open = { ...open, [c]: !open[c] };
  }
</script>

<div class="panel ok-scroll">
  <div class="panel-head">
    <span class="dot"></span>
    <strong>{selLabel}</strong>
    <span class="count">{applicable.length} {t('actions')}</span>
  </div>

  <!-- Hand-tuned quick controls for the common properties. -->
  <BespokeSections />

  <div class="all">
    <div class="all-title">{t('All applicable capabilities')}</div>
    {#each grouped as g (g.category)}
      <section class="cat">
        <button class="cat-head" onclick={() => toggle(g.category)}>
          <span class="chev" class:open={open[g.category]}>▸</span>
          {catLabel(g.label)}
          <span class="n">{g.items.length}</span>
        </button>
        {#if open[g.category]}
          <div class="cat-items">
            {#each g.items as cap (cap.id)}
              <button class="row" title={cap.id} onclick={() => editor.runOrPrompt(cap.id)}>
                <span class="row-label">{capLabel(cap)}</span>
                <span class="row-args">{argLabel(cap.params.length)}</span>
              </button>
            {/each}
          </div>
        {/if}
      </section>
    {/each}
  </div>
</div>

<style>
  .panel {
    background: var(--ok-panel);
    border-left: 1px solid var(--ok-border);
    overflow-y: auto;
    display: flex;
    flex-direction: column;
  }
  .panel-head {
    display: flex;
    align-items: center;
    gap: 8px;
    padding: 10px 12px;
    border-bottom: 1px solid var(--ok-border);
    position: sticky;
    top: 0;
    background: var(--ok-panel);
    z-index: 1;
  }
  .dot {
    width: 8px;
    height: 8px;
    border-radius: 50%;
    background: var(--ok-accent);
  }
  .count {
    margin-left: auto;
    font-size: 11px;
    color: var(--ok-text-3);
  }
  .all {
    padding: 8px;
  }
  .all-title {
    font-size: 11px;
    text-transform: uppercase;
    letter-spacing: 0.04em;
    color: var(--ok-text-3);
    padding: 6px 4px;
  }
  .cat-head {
    display: flex;
    align-items: center;
    gap: 6px;
    width: 100%;
    text-align: left;
    background: none;
    border: none;
    padding: 6px 4px;
    font: inherit;
    font-size: 12px;
    cursor: pointer;
    color: var(--ok-text);
    border-radius: var(--ok-radius);
  }
  .cat-head:hover {
    background: var(--ok-hover);
  }
  .chev {
    transition: transform 0.12s;
    color: var(--ok-text-3);
    font-size: 10px;
  }
  .chev.open {
    transform: rotate(90deg);
  }
  .n {
    margin-left: auto;
    font-size: 10px;
    color: var(--ok-text-3);
    background: var(--ok-bg);
    border-radius: 10px;
    padding: 0 7px;
  }
  .cat-items {
    padding: 2px 0 6px 18px;
    display: flex;
    flex-direction: column;
    gap: 1px;
  }
  .row {
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 8px;
    width: 100%;
    text-align: left;
    background: none;
    border: 1px solid transparent;
    border-radius: var(--ok-radius);
    padding: 5px 8px;
    font: inherit;
    font-size: 12px;
    cursor: pointer;
  }
  .row:hover {
    background: var(--ok-hover);
    border-color: var(--ok-border);
  }
  .row-args {
    font-size: 10px;
    color: var(--ok-text-3);
  }
</style>
