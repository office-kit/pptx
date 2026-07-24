<script lang="ts">
  import './ui/tokens.css';
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
  import ContextMenu from './ui/ContextMenu.svelte';
  import ToastStack from './ui/ToastStack.svelte';

  const editor = new EditorController();
  setEditor(editor);
  const doc = editor.doc;

  const NUDGE = 18288; // 0.02in in EMU
  const NUDGE_BIG = 182880; // 0.2in

  function onKeydown(e: KeyboardEvent) {
    const mod = e.ctrlKey || e.metaKey;
    const target = e.target as HTMLElement;
    const typing =
      target?.isContentEditable || /^(INPUT|TEXTAREA|SELECT)$/.test(target?.tagName ?? '');

    if (mod && e.key.toLowerCase() === 'k') {
      e.preventDefault();
      editor.togglePalette();
      return;
    }
    if (typing) return;

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
    } else if (mod && e.key.toLowerCase() === 'c') {
      editor.copySelection();
    } else if (mod && e.key.toLowerCase() === 'x') {
      editor.cutSelection();
    } else if (mod && e.key.toLowerCase() === 'v') {
      editor.paste();
    } else if (mod && e.key === '=') {
      e.preventDefault();
      editor.zoomIn();
    } else if (mod && e.key === '-') {
      e.preventDefault();
      editor.zoomOut();
    } else if (mod && e.key === '0') {
      e.preventDefault();
      editor.zoomFit();
    } else if ((e.key === 'Delete' || e.key === 'Backspace') && hasShapes) {
      e.preventDefault();
      editor.deleteSelection();
    } else if (e.key.startsWith('Arrow') && hasShapes) {
      e.preventDefault();
      const d = e.shiftKey ? NUDGE_BIG : NUDGE;
      if (e.key === 'ArrowLeft') editor.nudge(-d, 0);
      else if (e.key === 'ArrowRight') editor.nudge(d, 0);
      else if (e.key === 'ArrowUp') editor.nudge(0, -d);
      else if (e.key === 'ArrowDown') editor.nudge(0, d);
    } else if (e.key === 'Escape') {
      if (editor.contextMenu) editor.closeContextMenu();
      else if (editor.paletteOpen) editor.togglePalette(false);
      else if (editor.activeDialog) editor.closeDialog();
      else doc.clearShapeSelection();
    }
  }
</script>

<svelte:window on:keydown={onKeydown} />

<div class="ok-editor ok-shell">
  <TopBar />
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
    <CommandDialog id={editor.activeDialog} />
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
    grid-template-rows: auto auto 1fr auto;
    background: var(--ok-bg);
    overflow: hidden;
  }
  .ok-body {
    display: grid;
    grid-template-columns: var(--ok-nav-w) 1fr var(--ok-panel-w);
    min-height: 0;
    overflow: hidden;
  }
</style>
