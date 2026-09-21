<script lang="ts">
  import './ui/tokens.css';
  import { parseTableClipboard } from './core/table-clipboard.ts';
  import { t } from './i18n/i18n.svelte.ts';
  import { untrack, type Snippet } from 'svelte';
  import { EditorController } from './core/controller.svelte.ts';
  import { setEditor } from './core/context.ts';
  import TopBar from './ui/TopBar.svelte';
  import Ribbon from './ribbon/Ribbon.svelte';
  import SlideNavigator from './ui/SlideNavigator.svelte';
  import SlideCanvas from './canvas/SlideCanvas.svelte';
  import PropertiesPanel from './panels/PropertiesPanel.svelte';
  import StatusBar from './ui/StatusBar.svelte';
  import CommandPalette from './ui/CommandPalette.svelte';
  import CommandDialog from './ui/CommandDialog.svelte';
  import CropDialog from './ui/CropDialog.svelte';
  import ImageDialog from './ui/ImageDialog.svelte';
  import ChartDialog from './ui/ChartDialog.svelte';
  import FindReplaceDialog from './ui/FindReplaceDialog.svelte';
  import NotesDialog from './ui/NotesDialog.svelte';
  import TransitionDialog from './ui/TransitionDialog.svelte';
  import SlideSizeDialog from './ui/SlideSizeDialog.svelte';
  import TableDialog from './ui/TableDialog.svelte';
  import ContextMenu from './ui/ContextMenu.svelte';
  import ToastStack from './ui/ToastStack.svelte';

  let { editor: initialEditor = new EditorController(), onsave, status }: {
    editor?: EditorController;
    onsave?: () => Promise<void>;
    status?: Snippet;
  } = $props();
  const editor = untrack(() => initialEditor);
  setEditor(editor);
  const doc = editor.doc;

  const NUDGE = 18288; // 0.02in in EMU
  const NUDGE_BIG = 182880; // 0.2in

  function onKeydown(e: KeyboardEvent) {
    if (e.isComposing) return;
    if (editor.activeDialog) {
      if (e.key === 'Escape') { e.preventDefault(); editor.closeDialog(); }
      return;
    }
    const mod = e.ctrlKey || e.metaKey;
    const target = e.target as HTMLElement;
    const typing =
      target?.isContentEditable || /^(INPUT|TEXTAREA|SELECT)$/.test(target?.tagName ?? '');

    if (mod && e.key.toLowerCase() === 's' && onsave) {
      e.preventDefault();
      void onsave();
      return;
    }
    if (mod && e.key.toLowerCase() === 'k') {
      e.preventDefault();
      editor.togglePalette();
      return;
    }
    if (mod && ['f', 'h'].includes(e.key.toLowerCase())) {
      e.preventDefault();
      editor.runOrPrompt('replaceTextInPresentation');
      return;
    }
    if (typing || e.defaultPrevented) return;

    const hasShapes = doc.selection.kind === 'shape' || doc.selection.kind === 'cell';

    if (mod && e.key.toLowerCase() === 'z' && !e.shiftKey) {
      e.preventDefault();
      doc.undo();
    } else if (mod && (e.key.toLowerCase() === 'y' || (e.key.toLowerCase() === 'z' && e.shiftKey))) {
      e.preventDefault();
      doc.redo();
    } else if (mod && e.key.toLowerCase() === 'a') {
      e.preventDefault();
      editor.selectAllShapes();
    } else if (mod && e.key.toLowerCase() === 'd') {
      e.preventDefault();
      editor.duplicateSelection();
    } else if (mod && e.key.toLowerCase() === 'g') {
      e.preventDefault();
      const command = e.shiftKey ? 'ungroupShapes' : 'groupShapes';
      if (editor.canRun(command)) editor.invoke(command);
    } else if (mod && e.key.toLowerCase() === 'c') {
      if (doc.selection.kind !== 'cell') editor.copySelection();
    } else if (mod && e.key.toLowerCase() === 'x') {
      if (doc.selection.kind !== 'cell') editor.cutSelection();
    } else if (mod && e.key.toLowerCase() === 'v') {
      if (doc.selection.kind !== 'cell') editor.paste();
    } else if (mod && e.key === '=') {
      e.preventDefault();
      editor.zoomIn();
    } else if (mod && e.key === '-') {
      e.preventDefault();
      editor.zoomOut();
    } else if (mod && e.key === '0') {
      e.preventDefault();
      editor.zoomFit();
    } else if ((e.key === 'Delete' || e.key === 'Backspace') && (hasShapes || doc.selection.kind === 'slide')) {
      e.preventDefault();
      if (doc.selection.kind === 'cell') editor.clearCellText();
      else editor.deleteSelection();
    } else if (e.key.startsWith('Arrow') && hasShapes) {
      e.preventDefault();
      if (doc.selection.kind === 'cell') {
        editor.moveCellSelection(e.key === 'ArrowUp' ? -1 : e.key === 'ArrowDown' ? 1 : 0, e.key === 'ArrowLeft' ? -1 : e.key === 'ArrowRight' ? 1 : 0, e.shiftKey);
        return;
      }
      const d = e.shiftKey ? NUDGE_BIG : NUDGE;
      if (e.key === 'ArrowLeft') editor.nudge(-d, 0);
      else if (e.key === 'ArrowRight') editor.nudge(d, 0);
      else if (e.key === 'ArrowUp') editor.nudge(0, -d);
      else if (e.key === 'ArrowDown') editor.nudge(0, d);
    } else if (e.key === 'Escape') {
      if (editor.contextMenu) editor.closeContextMenu();
      else if (editor.paletteOpen) editor.togglePalette(false);
      else if (editor.activeDialog) editor.closeDialog();
      else if (doc.selection.kind === 'cell') doc.selectShape(doc.selection.slideIndex, doc.selection.shapeId);
      else doc.clearShapeSelection();
    }
  }
  function onCellClipboard(event: ClipboardEvent) {
    const target = event.target as HTMLElement;
    if (event.defaultPrevented || editor.activeDialog || doc.selection.kind !== 'cell' ||
      target?.isContentEditable || /^(INPUT|TEXTAREA|SELECT)$/.test(target?.tagName ?? '') || !event.clipboardData) return;
    if (event.type === 'paste') {
      if (!event.clipboardData.types.includes('text/plain')) return;
      event.preventDefault();
      const values = parseTableClipboard(event.clipboardData.getData('text/plain'));
      if (values) editor.pasteCellValues(values);
      else editor.toast('error', t('The clipboard table text is malformed'));
    } else {
      const text = event.type === 'cut' ? editor.cutSelection() : editor.copySelection();
      if (text === undefined) return;
      event.preventDefault();
      event.clipboardData.setData('text/plain', text);
    }
  }

</script>

<svelte:window on:keydown={onKeydown} on:copy={onCellClipboard} on:cut={onCellClipboard} on:paste={onCellClipboard} />

<div class="ok-editor ok-shell">
  <TopBar {onsave} />
  {#if status}<div class="host-status">{@render status()}</div>{/if}
  <Ribbon />
  <div class="ok-body">
    <SlideNavigator />
    <SlideCanvas />
    <PropertiesPanel />
  </div>
  <StatusBar />

  {#if editor.paletteOpen}
    <CommandPalette />
  {/if}
  {#if editor.activeDialog}
    {#if editor.activeDialog === 'addSlideImage' || editor.activeDialog === 'setShapeImage'}
      {#key editor.activeDialog}<ImageDialog replace={editor.activeDialog === 'setShapeImage'} />{/key}
    {:else if editor.activeDialog === 'setShapeImageCrop'}
      <CropDialog />
    {:else if editor.activeDialog === 'addSlideChart' || editor.activeDialog === 'setChartSpec'}
      {#key editor.activeDialog}<ChartDialog edit={editor.activeDialog === 'setChartSpec'} />{/key}
    {:else if editor.activeDialog === 'replaceTextInPresentation'}
      <FindReplaceDialog />
    {:else if editor.activeDialog === 'setSlideNotes'}
      <NotesDialog />
    {:else if editor.activeDialog === 'setSlideTransition'}
      <TransitionDialog />
    {:else if editor.activeDialog === 'setSlideSize'}
      <SlideSizeDialog />
    {:else if editor.activeDialog === 'addSlideTable'}
      <TableDialog />
    {:else}
      <CommandDialog id={editor.activeDialog} />
    {/if}
  {/if}
  {#if editor.contextMenu}
    <ContextMenu />
  {/if}
  <ToastStack />
</div>

<style>
  .ok-shell {
    position: fixed;
    inset: 0;
    z-index: 50;
    display: grid;
    grid-template-rows: auto auto minmax(0, 1fr) auto;
    background: var(--ok-bg);
    overflow: hidden;
  }
  .ok-shell > :global(*) { min-width: 0; }
  @media (max-width: 1100px) {
    .ok-shell { --ok-nav-w: 120px; --ok-panel-w: 230px; }
  }
  .ok-shell:has(.host-status) {
    grid-template-rows: auto auto auto minmax(0, 1fr) auto;
  }
  .ok-body {
    display: grid;
    grid-template-columns: var(--ok-nav-w) minmax(0, 1fr) var(--ok-panel-w);
    min-height: 0;
    overflow: hidden;
  }
</style>
