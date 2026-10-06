<script lang="ts">
  import { untrack } from 'svelte';
  import { getShapeChartSpec, getShapeFillEffective, getShapeKind, getShapeMedia } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import {
    capabilities,
    CATEGORY_LABELS,
    CATEGORY_ORDER,
    type CategoryId,
    type ResolvedCapability,
  } from '../manifest/index.ts';
  import Icon from '../ui/Icon.svelte';
  import BackgroundSection from './BackgroundSection.svelte';
  import SlideSection from './SlideSection.svelte';
  import LayoutSection from './LayoutSection.svelte';
  import ChartSection from './ChartSection.svelte';
  import TableSection from './TableSection.svelte';
  import ImageSection from './ImageSection.svelte';
  import VideoSection from './VideoSection.svelte';
  import ArrangeSection from './ArrangeSection.svelte';
  import AnimationSection from './AnimationSection.svelte';
  import ParagraphSection from './ParagraphSection.svelte';
  import BespokeSections from './BespokeSections.svelte';
  import { t, capLabel, catLabel, getLocale } from '../i18n/i18n.svelte.ts';

  const editor = getEditor();
  const doc = editor.doc;
  const baseFormatTabs = [
    { id: 'paint', label: 'Fill & Line', icon: 'fill' },
    { id: 'effects', label: 'Effects', icon: 'shadow' },
    { id: 'size', label: 'Size & Properties', icon: 'resize' },
  ] as const;
  const videoTab = { id: 'video', label: 'Video', icon: 'video' } as const;
  const selectedVideo = $derived.by(() => {
    doc.version;
    const sel = doc.selection;
    if (sel.kind !== 'shape' || sel.shapeIds.length !== 1) return false;
    const shape = doc.shapeById(sel.slideIndex, sel.shapeIds[0]!);
    return !!shape && getShapeMedia(shape)?.kind === 'video';
  });
  const formatTabs = $derived(selectedVideo ? [...baseFormatTabs, videoTab] : baseFormatTabs);
  const isShape = $derived.by(() => {
    if (editor.propertiesPaneMode === 'background') return false;
    doc.version;
    const shapes = editor.selectedShapes();
    return selectedVideo || (shapes.length > 0 && shapes.every((shape) =>
      ['shape', 'connector', 'group'].includes(getShapeKind(shape)),
    ));
  });
  // Like PowerPoint's Format pane, it follows the selection: an object shows
  // its format, the slide shows Format Background.
  const objectSelected = $derived(doc.selection.kind === 'shape' || doc.selection.kind === 'cell');
  let lastObjectSelected = untrack(() => objectSelected);
  $effect(() => {
    if (objectSelected === lastObjectSelected) return;
    lastObjectSelected = objectSelected;
    untrack(() => { editor.propertiesPaneMode = objectSelected ? 'selection' : 'background'; });
  });
  $effect(() => {
    if (isShape && !formatTabs.some((tab) => tab.id === editor.formatPaneTab)) editor.formatPaneTab = 'paint';
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

  const backgroundPane = $derived(editor.propertiesPaneMode === 'background' || !objectSelected);
  const selLabel = $derived.by(() => {
    if (backgroundPane) return t('Format Background');
    const sel = doc.selection;
    if (selectedVideo) return t('Format Video');
    doc.version;
    const shape = sel.kind === 'shape' && sel.shapeIds.length === 1 ? doc.shapeById(sel.slideIndex, sel.shapeIds[0]!) : null;
    if (shape && getShapeChartSpec(shape)) return t('Format Chart Area');
    if (shape && getShapeKind(shape) === 'picture') return t('Format Picture');
    // PowerPoint retitles Format Shape once the shape is filled with a picture or texture.
    const shapes = editor.selectedShapes();
    if (shapes.length > 0 && shapes.every(target => getShapeKind(target) === 'shape' && getShapeFillEffective(doc.pres, target).kind === 'image')) return t('Format Picture');
    return t('Format Shape');
  });

  // Which category sections are expanded.
  let open = $state<Record<string, boolean>>({});
  function toggle(c: string) {
    open = { ...open, [c]: !open[c] };
  }
</script>

{#snippet selectionSections()}
  <SlideSection />
  <LayoutSection />
  {#if selectedVideo && editor.formatPaneTab === 'video'}
    <VideoSection />
  {/if}
  <div hidden={isShape && editor.formatPaneTab !== 'size'}>
    <ChartSection />
    <TableSection />
    <ImageSection />
  </div>
  {#if !selectedVideo || editor.formatPaneTab !== 'video'}
    <BespokeSections tab={isShape && editor.formatPaneTab !== 'video' ? editor.formatPaneTab : 'all'} />
  {/if}
  <div hidden={isShape && editor.formatPaneTab !== 'size'}>
    <ParagraphSection />
    <ArrangeSection />
    <AnimationSection />
  </div>

  {#if !selectedVideo || editor.formatPaneTab !== 'video'}
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
  {/if}{/snippet}

<div class="panel ok-scroll" class:background-pane={backgroundPane} hidden={!editor.propertiesPaneVisible}>
  <div class="panel-head">
    <strong>{selLabel}</strong>
          <button class="close-pane" aria-label={t(backgroundPane ? 'Close Format Background' : selectedVideo ? 'Close Format Video' : 'Close Format Shape')} title={t(backgroundPane ? 'Close Format Background' : selectedVideo ? 'Close Format Video' : 'Close Format Shape')} onclick={(event) => {
        editor.propertiesPaneVisible = false;
        const shell = event.currentTarget.closest('.ok-shell');
        const target = shell?.querySelector<HTMLElement>(backgroundPane ? '.format-background-trigger' : '.hit.selected');
        target?.focus({ preventScroll: true });
      }}>×</button>
  </div>

  {#if isShape}
    <div
      class="format-tabs"
      role="tablist"
      tabindex="-1"
      aria-label={t(selectedVideo ? 'Format Video' : 'Format Shape')}
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
    {#if backgroundPane}
      <BackgroundSection extra={objectSelected ? undefined : selectionSections} />
    {:else}
      {@render selectionSections()}

    {/if}
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
  .background-pane { overflow: hidden; min-height: 0; }
  .background-pane .panel-head { flex-shrink: 0; }
  .background-pane #format-panel { flex: 1; min-height: 0; display: flex; flex-direction: column; }
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
