<script lang="ts">
  // The contextual tabs' Arrange group. Expanded it is PowerPoint's row of
  // Bring Forward ▾, Send Backward ▾, Selection Pane, Reorder Objects ▾,
  // Align ▾ and Group ▾ / Rotate ▾ (Table Layout: Align, Group and Rotate as
  // labelled rows, without Selection Pane and Reorder Objects); collapsed it
  // is the single Arrange ▾ menu the Home tab uses.
  import { getEditor } from '../core/context.ts';
  import { selectedShapeIds } from '../core/selection.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import Icon from '../ui/Icon.svelte';
  import ArrangeMenu from './ArrangeMenu.svelte';
  import { captionLines } from './caption.ts';
  import MenuButton from './MenuButton.svelte';
  import { ALIGN_ITEMS, ROTATE_ITEMS, rotateSelection } from './arrange-actions.ts';

  let { collapsed, table = false }: { collapsed: boolean; table?: boolean } = $props();
  const editor = getEditor();
  const doc = editor.doc;
  const count = $derived(selectedShapeIds(doc.selection).length);
  const locked = $derived(editor.selectionLocked());
  const toSlide = $derived(count < 2 || editor.alignmentReference === 'slide');
</script>

{#snippet order(forward: boolean)}
  {#each forward ? [['bringShapeForward', 'Bring Forward'], ['bringShapeToFront', 'Bring to Front']] : [['sendShapeBackward', 'Send Backward'], ['sendShapeToBack', 'Send to Back']] as [id, label] (id)}
    <button role="menuitem" disabled={!editor.canRun(id!)} onclick={() => editor.invoke(id!)}>{t(label!)}</button>
  {/each}
{/snippet}

{#snippet align()}
  {#each ALIGN_ITEMS as item (item.value)}<button role="menuitem" onclick={() => editor.alignSelection(item.value, toSlide ? 'slide' : 'selection')}>{t(item.label)}</button>{/each}
  <hr />
  <button role="menuitem" disabled={!toSlide && count < 3} onclick={() => editor.distributeSelection('horizontal')}>{t('Distribute Horizontally')}</button>
  <button role="menuitem" disabled={!toSlide && count < 3} onclick={() => editor.distributeSelection('vertical')}>{t('Distribute Vertically')}</button>
  <hr />
  <button role="menuitemradio" aria-checked={toSlide} onclick={() => (editor.alignmentReference = 'slide')}>{t('Align to Slide')}<span>{toSlide ? '✓' : ''}</span></button>
  <button role="menuitemradio" aria-checked={!toSlide} disabled={count < 2} onclick={() => (editor.alignmentReference = 'selection')}>{t('Align Selected Objects')}<span>{toSlide ? '' : '✓'}</span></button>
{/snippet}

{#snippet group()}
  <button role="menuitem" disabled={!editor.canRun('groupShapes')} onclick={() => editor.invoke('groupShapes')}>{t('Group')}</button>
  <button role="menuitem" disabled={!editor.canRegroup()} onclick={() => editor.regroupSelection()}>{t('Regroup')}</button>
  <button role="menuitem" disabled={!editor.canRun('ungroupShapes')} onclick={() => editor.invoke('ungroupShapes')}>{t('Ungroup')}</button>
{/snippet}

{#snippet rotate()}
  {#each ROTATE_ITEMS as item (item.action)}<button role="menuitem" onclick={() => rotateSelection(editor, item.action)}>{t(item.label)}</button>{/each}
  <hr />
  <button role="menuitem" disabled={!editor.canRun('setShapeRotation')} onclick={() => editor.showRotationOptions()}>{t('More Rotation Options...')}</button>
{/snippet}

<section class="ctx-group" aria-label={t('Arrange')}>
  {#if collapsed}
    <ArrangeMenu />
  {:else}
    <MenuButton look="big" icon="forward" label={t('Bring Forward')} disabled={!count || locked}>{@render order(true)}</MenuButton>
    <MenuButton look="big" icon="backward" label={t('Send Backward')} disabled={!count || locked}>{@render order(false)}</MenuButton>
    {#if table}
      <div class="ctx-rows">
        <MenuButton look="row" icon="align-objects" label={t('Align')} disabled={!count || locked}>{@render align()}</MenuButton>
        <MenuButton look="row" icon="group" label={t('Group')} disabled={!editor.canRun('groupShapes') && !editor.canRun('ungroupShapes') && !editor.canRegroup()}>{@render group()}</MenuButton>
        <MenuButton look="row" icon="rotate" label={t('Rotate')} disabled={!editor.canRun('setShapeRotation')}>{@render rotate()}</MenuButton>
      </div>
    {:else}
      <button class="ctx-big" aria-label={t('Selection Pane')} aria-pressed={editor.selectionPaneVisible} title={t('Display the Selection Pane')} onclick={() => { editor.setViewMode('normal'); editor.selectionPaneVisible = !editor.selectionPaneVisible; }}><span class="ctx-icon-row"><Icon name="selection-pane" size={32} /></span><span class="ctx-caption">{captionLines(t('Selection Pane'))}</span></button>
      <MenuButton look="big" icon="arrange" label={t('Reorder Objects')}>
        <button role="menuitem" disabled={editor.reorderMembers().length < 2} onclick={() => (editor.activeDialog = 'reorderObjects')}>{t('Reorder Overlapping Objects')}</button>
      </MenuButton>
      <MenuButton look="big" icon="align-objects" label={t('Align')} disabled={!count || locked}>{@render align()}</MenuButton>
      <div class="ctx-rows tools">
        <MenuButton look="tool" icon="group" label={t('Group')} disabled={!editor.canRun('groupShapes') && !editor.canRun('ungroupShapes') && !editor.canRegroup()}>{@render group()}</MenuButton>
        <MenuButton look="tool" icon="rotate" label={t('Rotate')} disabled={!editor.canRun('setShapeRotation')}>{@render rotate()}</MenuButton>
      </div>
    {/if}
  {/if}
</section>
