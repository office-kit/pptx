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
  import EffectSections from './EffectSections.svelte';
  import TextFillSections from './TextFillSections.svelte';
  import TextBoxSection from './TextBoxSection.svelte';
  import { textEffects } from './effects-model.ts';
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
  const textTabs = [
    { id: 'textFill', label: 'Text Fill & Outline', icon: 'font-color' },
    { id: 'textEffects', label: 'Text Effects', icon: 'glow' },
    { id: 'textbox', label: 'Textbox', icon: 'textbox' },
  ] as const;
  // Text Options apply to the text of shapes that can hold it.
  const textOptionsAvailable = $derived.by(() => {
    doc.version;
    return !selectedVideo && editor.selectedShapes().some((shape) => getShapeKind(shape) === 'shape');
  });
  const textOptions = $derived(textOptionsAvailable && editor.formatPaneOptions === 'text');
  const activeTabs = $derived<readonly { id: string; label: string; icon: string }[]>(textOptions ? textTabs : formatTabs);
  const activeTab = $derived(textOptions ? editor.formatPaneTextTab : editor.formatPaneTab);
  function selectTab(id: string) {
    if (textOptions) editor.formatPaneTextTab = id as typeof editor.formatPaneTextTab;
    else editor.formatPaneTab = id as typeof editor.formatPaneTab;
  }
  const isShape = $derived.by(() => {
    if (editor.propertiesPaneMode === 'background') return false;
    doc.version;
    const shapes = editor.selectedShapes();
    return selectedVideo || (shapes.length > 0 && shapes.every((shape) =>
      ['shape', 'connector', 'group'].includes(getShapeKind(shape)),
    ));
  });
  // Like the reference desktop app's Format pane, it follows the selection: an object shows
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
    const index = activeTabs.findIndex((tab) => tab.id === activeTab);
    const next = event.key === 'Home'
      ? 0
      : event.key === 'End'
        ? activeTabs.length - 1
        : (index + (event.key === 'ArrowLeft' ? -1 : 1) + activeTabs.length) % activeTabs.length;
    selectTab(activeTabs[next]!.id);
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
    // The reference desktop app retitles Format Shape once the shape is filled with a picture or texture.
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

  {#if isShape && !selectedVideo}
    <!-- The reference desktop app's Shape Options / Text Options switch above the categories. -->
    <div class="options-switch" role="radiogroup" aria-label={t('Format Shape options')}>
      {#each [['shape', 'Shape Options'], ['text', 'Text Options']] as const as [value, label]}
        <button role="radio" aria-checked={(textOptions ? 'text' : 'shape') === value} disabled={value === 'text' && !textOptionsAvailable}
          title={value === 'text' && !textOptionsAvailable ? t('Select a shape that can hold text.') : undefined}
          onclick={() => editor.formatPaneOptions = value}>{t(label)}</button>
      {/each}
    </div>
  {/if}
  {#if isShape}
    <div
      class="format-tabs"
      role="tablist"
      tabindex="-1"
      aria-label={t(selectedVideo ? 'Format Video' : 'Format Shape')}
      onkeydown={tabKeys}
    >
      {#each activeTabs as tab (tab.id)}
        <button
          role="tab"
          id="format-tab-{tab.id}"
          aria-label={t(tab.label)}
          title={t(tab.label)}
          aria-selected={activeTab === tab.id}
          aria-controls="format-panel"
          tabindex={activeTab === tab.id ? 0 : -1}
          onclick={() => selectTab(tab.id)}
        >
          <Icon name={tab.icon} size={24} />
        </button>
      {/each}
    </div>
  {/if}
  <div
    id="format-panel"
    role={isShape ? 'tabpanel' : undefined}
    aria-labelledby={isShape ? `format-tab-${activeTab}` : undefined}
  >
    {#if backgroundPane}
      <BackgroundSection extra={objectSelected ? undefined : selectionSections} />
    {:else if textOptions}
      <div class="text-options">
        {#if editor.formatPaneTextTab === 'textFill'}
          <TextFillSections />
        {:else if editor.formatPaneTextTab === 'textEffects'}
          <EffectSections target={textEffects} />
        {:else}
          <TextBoxSection sectionId="textbox" />
        {/if}
      </div>
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
    /* The reference desktop app's (Mac) pane title bar is 27 pt tall. */
    min-height: 27px;
    box-sizing: border-box;
    padding: 3px 12px;
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
  /* The reference desktop app (Mac): a 278 × 26 pt segmented switch, 14 pt below the title. */
  .options-switch {
    display: flex;
    gap: 1px;
    margin: 14px 11px 0 10px;
    padding: 1px;
    border-radius: 6px;
    background: var(--ok-hover);
  }
  .options-switch button {
    flex: 1;
    height: 24px;
    border: none;
    border-radius: 5px;
    background: none;
    color: var(--ok-text);
    font: inherit;
    font-size: 12px;
    cursor: pointer;
  }
  .options-switch button[aria-checked='true'] { background: var(--ok-panel); box-shadow: 0 0 0 1px var(--ok-border-strong); font-weight: 600; }
  .options-switch button:disabled { opacity: 0.45; cursor: default; }
  .text-options { padding: 8px 10px; display: flex; flex-direction: column; }
  .format-tabs {
    display: flex;
    gap: 0;
    padding: 8px 10px 0;
    background: var(--ok-panel);
    position: sticky;
    top: 27px;
    z-index: 1;
  }
  .format-tabs button {
    display: grid;
    place-items: center;
    width: 42px;
    height: 42px;
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
