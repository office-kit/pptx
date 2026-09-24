<script lang="ts">
  import { getShapeKind } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import {
    capabilities,
    CATEGORY_LABELS,
    CATEGORY_ORDER,
    type CategoryId,
    type ResolvedCapability,
  } from '../manifest/index.ts';
  import Icon from '../ui/Icon.svelte';
  import SlideSection from './SlideSection.svelte';
  import LayoutSection from './LayoutSection.svelte';
  import ChartSection from './ChartSection.svelte';
  import TableSection from './TableSection.svelte';
  import ImageSection from './ImageSection.svelte';
  import ArrangeSection from './ArrangeSection.svelte';
  import AnimationSection from './AnimationSection.svelte';
  import ParagraphSection from './ParagraphSection.svelte';
  import BespokeSections from './BespokeSections.svelte';
  import { t, capLabel, catLabel, getLocale } from '../i18n/i18n.svelte.ts';

  const editor = getEditor();
  const doc = editor.doc;
  const formatTabs = [
    { id: 'paint', label: 'Fill & Line', icon: 'fill' },
    { id: 'effects', label: 'Effects', icon: 'shadow' },
    { id: 'size', label: 'Size & Properties', icon: 'resize' },
  ] as const;
  const isShape = $derived.by(() => {
    doc.version;
    const shapes = editor.selectedShapes();
    return shapes.length > 0 && shapes.every((shape) =>
      ['shape', 'connector', 'group'].includes(getShapeKind(shape)),
    );
  });
  function tabKeys(event: KeyboardEvent) {
    if (!['ArrowLeft', 'ArrowRight', 'Home', 'End'].includes(event.key)) return;
    event.preventDefault();
    event.stopPropagation();
    const index = formatTabs.findIndex((tab) => tab.id === editor.formatPaneTab);
    const next = event.key === 'Home'
      ? 0
      : event.key === 'End'
        ? formatTabs.length - 1
        : (index + (event.key === 'ArrowLeft' ? -1 : 1) + formatTabs.length) % formatTabs.length;
    editor.formatPaneTab = formatTabs[next]!.id;
    (event.currentTarget as HTMLElement).querySelectorAll<HTMLButtonElement>('[role="tab"]')[next]?.focus();
  }

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
      const tab = cap.category === 'fill' || cap.category === 'stroke'
        ? 'paint'
        : cap.category === 'effect' ? 'effects' : 'size';
      if (isShape && editor.formatPaneTab !== tab) continue;
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
    if (sel.kind === 'shape') return isShape ? t('Format Shape') : t('Shape');
    if (sel.kind === 'cell') return t('Table cell');
    return t('Slide');
  });

  // Which category sections are expanded.
  let open = $state<Record<string, boolean>>({});
  function toggle(c: string) {
    open = { ...open, [c]: !open[c] };
  }
</script>

<div class="panel ok-scroll" hidden={!editor.propertiesPaneVisible}>
  <div class="panel-head">
    <strong>{selLabel}</strong>
    {#if isShape}
      <button class="close-pane" aria-label={t('Close Format Shape')} title={t('Close Format Shape')} onclick={(event) => {
        editor.propertiesPaneVisible = false;
        event.currentTarget.closest('.ok-shell')?.querySelector<HTMLElement>('.hit.selected')?.focus({ preventScroll: true });
      }}>×</button>
    {/if}
  </div>

  {#if isShape}
    <div
      class="format-tabs"
      role="tablist"
      tabindex="-1"
      aria-label={t('Format Shape')}
      onkeydown={tabKeys}
    >
      {#each formatTabs as tab}
        <button
          role="tab"
          id="format-tab-{tab.id}"
          aria-label={t(tab.label)}
          title={t(tab.label)}
          aria-selected={editor.formatPaneTab === tab.id}
          aria-controls="format-panel"
          tabindex={editor.formatPaneTab === tab.id ? 0 : -1}
          onclick={() => editor.formatPaneTab = tab.id}
        >
          <Icon name={tab.icon} size={24} />
        </button>
      {/each}
    </div>
  {/if}
  <div
    id="format-panel"
    role={isShape ? 'tabpanel' : undefined}
    aria-labelledby={isShape ? `format-tab-${editor.formatPaneTab}` : undefined}
  >
    <SlideSection />
    <LayoutSection />
    <div hidden={isShape && editor.formatPaneTab !== 'size'}>
      <ChartSection />
      <TableSection />
      <ImageSection />
    </div>
    <BespokeSections tab={isShape ? editor.formatPaneTab : 'all'} />
    <div hidden={isShape && editor.formatPaneTab !== 'size'}>
      <ParagraphSection />
      <ArrangeSection />
      <AnimationSection />
    </div>

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
  [hidden] {
    display: none;
  }
  .close-pane {
    margin-left: auto;
    padding: 0 4px;
    border: none;
    background: none;
    color: var(--ok-text-2);
    font-size: 20px;
    line-height: 20px;
    cursor: pointer;
  }
  .format-tabs {
    display: flex;
    gap: 8px;
    padding: 8px 12px;
    background: var(--ok-panel);
    position: sticky;
    top: 37px;
    z-index: 1;
  }
  .format-tabs button {
    display: grid;
    place-items: center;
    width: 38px;
    height: 34px;
    border: 1px solid transparent;
    background: none;
    color: var(--ok-text-2);
    border-radius: var(--ok-radius);
    cursor: pointer;
  }
  .format-tabs button:hover {
    background: var(--ok-hover);
  }
  .format-tabs button[aria-selected='true'] {
    border-color: var(--ok-border-strong);
    background: var(--ok-hover);
    color: var(--ok-accent);
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
