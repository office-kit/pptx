<script lang="ts">
  // Slide Master view's left pane, measured from Mac PowerPoint 16: a 232 pt
  // list; each master is a 119 pt cell numbered at the left with a 193 × 109 pt
  // image, and its layouts follow as 75 pt cells on an 87 pt pitch, their
  // 124 × 71 pt images indented to x = 90.
  import { tick } from 'svelte';
  import { getSlideLayoutName, getSlideLayoutPartName, getSlideSize } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { masterGroups } from '../core/master-geometry.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import MasterSlide from './MasterSlide.svelte';

  const editor = getEditor();
  const doc = editor.doc;
  const groups = $derived.by(() => { doc.version; return masterGroups(doc.pres); });
  const size = $derived.by(() => { doc.version; return getSlideSize(doc.pres) ?? { width: 12192000, height: 6858000 }; });
  const MASTER_IMAGE_WIDTH = 193;
  const LAYOUT_IMAGE_WIDTH = 124;
  let pane = $state<HTMLElement>();

  // Each cell is identified by its layout part, or the master's part for the master.
  const cells = $derived(groups.flatMap((group) => [{ master: group.partName, layout: null as string | null }, ...group.layouts.map((layout) => ({ master: group.partName, layout: getSlideLayoutPartName(layout) }))]));
  const selectedIndex = $derived(cells.findIndex((cell) => doc.layoutTarget?.partName === cell.layout && (cell.layout !== null || editor.selectedMaster === cell.master)));

  // Undo can take away the selected master or layout (an inserted one);
  // the selection then falls back to the first master, as the pane would after a delete.
  $effect(() => {
    if (selectedIndex === -1 && cells[0]) select(cells[0].master, null);
  });

  function select(master: string, layout: string | null) {
    editor.selectMasterCell(master, layout);
  }
  async function keys(event: KeyboardEvent) {
    if (!['ArrowUp', 'ArrowDown', 'Home', 'End'].includes(event.key)) return;
    event.preventDefault();
    const index = event.key === 'Home' ? 0 : event.key === 'End' ? cells.length - 1 : Math.max(0, Math.min(cells.length - 1, selectedIndex + (event.key === 'ArrowUp' ? -1 : 1)));
    const cell = cells[index]!;
    select(cell.master, cell.layout);
    await tick();
    pane?.querySelector<HTMLElement>('[aria-current="true"]')?.focus();
  }
</script>

<nav class="master-pane ok-scroll" bind:this={pane} aria-label={t('Slide Master Pane')}>
  {#each groups as group, m (group.partName)}
    {@const representative = group.layouts[0]}
    {#if representative}
      {@const current = editor.selectedMaster === group.partName && doc.layoutTarget?.partName === null}
      <button class="cell master" aria-label="{t('Slide Master')} {m + 1}" aria-current={current ? 'true' : undefined} tabindex={current || (selectedIndex < 0 && m === 0) ? 0 : -1} onclick={() => select(group.partName, null)} onkeydown={keys}>
        <span class="num">{m + 1}</span>
        <span class="image" style:width="{MASTER_IMAGE_WIDTH}px" style:aspect-ratio="{size.width} / {size.height}"><MasterSlide pres={doc.pres} layout={representative} master version={doc.version} pixelWidth={MASTER_IMAGE_WIDTH} /></span>
      </button>
      {#each group.layouts as layout (getSlideLayoutPartName(layout))}
        {@const partName = getSlideLayoutPartName(layout)}
        {@const selected = doc.layoutTarget?.partName === partName}
        <button class="cell layout" aria-label={getSlideLayoutName(layout)} title={getSlideLayoutName(layout)} data-layout={partName} aria-current={selected ? 'true' : undefined} tabindex={selected ? 0 : -1} onclick={() => select(group.partName, partName)} onkeydown={keys}>
          <span class="image" style:width="{LAYOUT_IMAGE_WIDTH}px" style:aspect-ratio="{size.width} / {size.height}"><MasterSlide pres={doc.pres} {layout} version={doc.version} pixelWidth={LAYOUT_IMAGE_WIDTH} /></span>
        </button>
      {/each}
    {/if}
  {/each}
</nav>

<style>
  .master-pane { box-sizing: border-box; width: 100%; min-height: 0; overflow-y: auto; scrollbar-gutter: stable; background: var(--ok-panel-2); border-right: 1px solid var(--ok-border); }
  /* Cells grow with the slide's aspect ratio; at 16:9 they are 119 and 75 pt. */
  .cell { position: relative; display: block; box-sizing: border-box; width: 232px; border: 0; background: transparent; color: var(--ok-text-2); font: inherit; cursor: pointer; }
  .cell.master { padding: 7px 0 3px 20px; margin-bottom: 6px; }
  /* Mac PowerPoint keeps 75 pt layout cells at 16:9 although the image is 70 pt tall. */
  .cell.layout { min-height: 75px; padding: 2px 0 2px 90px; margin-bottom: 12px; }
  .num { position: absolute; left: 6px; top: 6px; font-size: 13px; }
  .image { display: block; box-sizing: border-box; border: 1px solid var(--ok-border-strong); background: #fff; overflow: hidden; }
  .cell:hover .image { outline: 3px solid var(--ok-border-strong); outline-offset: 1px; }
  /* Mac PowerPoint rings the selected master or layout as it rings a selected thumbnail. */
  .cell[aria-current='true'] .image { outline: 3px solid var(--ok-accent); outline-offset: 1px; }
  .cell:focus-visible { outline: none; }
  .cell:focus-visible .image { outline: 3px solid var(--ok-selected-border); outline-offset: 1px; }
</style>
