<script lang="ts">
  import { parseTableClipboard } from './core/table-clipboard.ts';
  import { menuItemForKey } from './core/menubar-shortcuts.ts';
  import { nativeMenus } from './ui/menubar-commands.ts';
  import { getLocale, t } from './i18n/i18n.svelte.ts';
  import { untrack } from 'svelte';
  import { EditorController } from './core/controller.svelte.ts';
  import { setEditor } from './core/context.ts';
  import { eventTarget } from './core/dom-root.ts';
  import TopBar from './ui/TopBar.svelte';
  import Ribbon from './ribbon/Ribbon.svelte';
  import SlideNavigator from './ui/SlideNavigator.svelte';
  import ThumbnailPane from './ui/ThumbnailPane.svelte';
  import SlideCanvas from './canvas/SlideCanvas.svelte';
  import PropertiesPanel from './panels/PropertiesPanel.svelte';
  import SelectionPane from './panels/SelectionPane.svelte';
  import StatusBar from './ui/StatusBar.svelte';
  import CommandPalette from './ui/CommandPalette.svelte';
  import CommandDialog from './ui/CommandDialog.svelte';
  import ReorderObjectsDialog from './ui/ReorderObjectsDialog.svelte';
  import ZoomDialog from './ui/ZoomDialog.svelte';
  import GridOptionsDialog from './ui/GridOptionsDialog.svelte';
  import CompressPicturesDialog from './ui/CompressPicturesDialog.svelte';
  import CropDialog from './ui/CropDialog.svelte';
  import ImageDialog from './ui/ImageDialog.svelte';
  import ChartDialog from './ui/ChartDialog.svelte';
  import FindReplaceDialog from './ui/FindReplaceDialog.svelte';
  import LinkDialog from './ui/LinkDialog.svelte';
  import CommentsDialog from './ui/CommentsDialog.svelte';
  import NewSlideDialog from './ui/NewSlideDialog.svelte';
  import NotesPane from './ui/NotesPane.svelte';
  import TransitionDialog from './ui/TransitionDialog.svelte';
  import ShowPropertiesDialog from './ui/ShowPropertiesDialog.svelte';
  import SlideSizeDialog from './ui/SlideSizeDialog.svelte';
  import HeaderFooterDialog from './ui/HeaderFooterDialog.svelte';
  import SymbolPicker from './ui/SymbolPicker.svelte';
  import ShapeGallery from './ui/ShapeGallery.svelte';
  import WordArtGallery from './ui/WordArtGallery.svelte';
  import { insertWordArt } from './core/insert-objects.ts';
  import NotesPageView from './ui/NotesPageView.svelte';
  import SlideMasterPane from './ui/SlideMasterPane.svelte';
  import MasterCanvas from './canvas/MasterCanvas.svelte';
  import PageMasterView from './ui/PageMasterView.svelte';
  import ReadingView from './ui/ReadingView.svelte';
  import { isSlideEditingView } from './core/view-modes.ts';
  import RehearsalDialog from './ui/RehearsalDialog.svelte';
  import TableDialog from './ui/TableDialog.svelte';
  import CustomShowsDialog from './ui/CustomShowsDialog.svelte';
  import FontDialog from './ui/FontDialog.svelte';
  import ParagraphDialog from './ui/ParagraphDialog.svelte';
  import { editTargetParagraphs, targetParagraphProperties } from './core/paragraph-targets.ts';
  import ContextMenu from './ui/ContextMenu.svelte';
  import ToastStack from './ui/ToastStack.svelte';

  let { editor: initialEditor = new EditorController(), onsave, status, compactHost = false, autoSave = false, embedded = false }: {
    editor?: EditorController;
    onsave?: () => Promise<void>;
    /** The host's own element for the title bar (its save status, say). */
    status?: HTMLElement;
    /** A slimmer title bar, for a host with little room. */
    compactHost?: boolean;
    /** Saves through `onsave` shortly after each edit while the title bar's AutoSave switch is on. */
    autoSave?: boolean;
    /** Fill the containing element instead of the window. */
    embedded?: boolean;
  } = $props();
  const editor = untrack(() => initialEditor);
  setEditor(editor);
  // Long enough that a burst of quick edits saves once.
  const AUTO_SAVE_DELAY = 700;
  $effect(() => {
    const doc = editor.doc;
    doc.committedVersion;
    if (!autoSave || !onsave || !editor.autoSave || !doc.dirty || doc.liveEditing) return;
    const save = onsave;
    const timer = setTimeout(() => void save(), AUTO_SAVE_DELAY);
    return () => clearTimeout(timer);
  });
  // The preview page around the frame follows the slide being edited.
  $effect(() => {
    if (!editor.hostFrame) return;
    window.parent.postMessage(
      {
        type: 'editor-focus',
        slide: editor.doc.selection.slideIndex,
        count: editor.doc.slides.length,
        dirty: editor.doc.dirty,
        locale: getLocale(),
      },
      window.location.origin,
    );
  });
  let shell = $state<HTMLElement>();
  $effect(() => {
    if (!shell) return;
    const node = shell;
    const activate = () => editor.activate();
    editor.attachShell(node);
    node.addEventListener('pointerdown', activate, true);
    node.addEventListener('focusin', activate, true);
    return () => {
      node.removeEventListener('pointerdown', activate, true);
      node.removeEventListener('focusin', activate, true);
      editor.attachShell(undefined);
    };
  });
  const doc = editor.doc;
  const commentsOpen = $derived(['addSlideComment', 'setCommentText', 'removeSlideComment'].includes(editor.activeDialog ?? ''));
  $effect(() => {
    doc.selection;
    untrack(() => editor.completeFormatPainter());
  });
  const slideEditing = $derived(isSlideEditingView(editor.viewMode));
  const navigationWidth = $derived(editor.viewMode === 'outline' ? editor.outlineWidth ?? 360 : editor.thumbnailWidth);

  const NUDGE = 18288; // 0.02in in EMU
  const NUDGE_BIG = 182880; // 0.2in

  function onKeydown(e: KeyboardEvent) {
    if (e.isComposing || !editor.ownsEvent(e)) return;
    if (editor.activeDialog) {
      if (e.key === 'Escape') { e.preventDefault(); editor.closeDialog(); }
      return;
    }
    const mod = e.ctrlKey || e.metaKey;
    const target = eventTarget(e) as HTMLElement;
    const typing =
      target?.isContentEditable || /^(INPUT|TEXTAREA|SELECT)$/.test(target?.tagName ?? '');

    // Menu-bar shortcuts (⌘T, ⇧⌘N, ⌥⌘G, ⌘1 …) run in MenuBar; this keymap
    // keeps undo, the clipboard, Select All, deletion, nudging and Escape.
    if (typing || e.defaultPrevented) return;
    if (e.key === 'ContextMenu' || (e.shiftKey && e.key === 'F10')) {
      e.preventDefault();
      const bounds = target.getBoundingClientRect();
      editor.openContextMenu(bounds.left, bounds.bottom);
      return;
    }

    const hasShapes = doc.selection.kind === 'shape' || doc.selection.kind === 'cell';

    // Undo, Repeat, the clipboard and Select All keep their keys from the
    // menu bar's table, so they follow PowerPoint's shortcuts exactly.
    const menuKey = menuItemForKey(nativeMenus(), e)?.id;
    if (menuKey === 'edit/undo') {
      e.preventDefault();
      doc.undo();
    } else if (menuKey === 'edit/repeat' || (mod && e.shiftKey && !e.altKey && e.code === 'KeyZ')) {
      e.preventDefault();
      doc.redo();
    } else if (menuKey === 'edit/select-all') {
      e.preventDefault();
      editor.selectAll();
    } else if (menuKey === 'edit/copy') {
      if (doc.selection.kind !== 'cell') editor.copySelection();
    } else if (menuKey === 'edit/cut') {
      if (doc.selection.kind !== 'cell') editor.cutSelection();
    } else if (menuKey === 'edit/paste') {
      if (doc.selection.kind !== 'cell') editor.paste();
    } else if (mod && e.key === '0') {
      e.preventDefault();
      editor.zoomFit();
    } else if ((e.key === 'Delete' || e.key === 'Backspace') && (hasShapes || doc.selection.kind === 'slide')) {
      e.preventDefault();
      editor.deleteSelection();
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
      // A ribbon menu or collapsed group that closed on this Escape claims it;
      // window listeners run in mount order, so this one may run first.
      if (e.defaultPrevented || (target instanceof Element && target.closest('[role="menu"], .group-popup'))) return;
      if (editor.contextMenu) editor.closeContextMenu();
      else if (editor.paletteOpen) editor.togglePalette(false);
      else if (editor.activeDialog) editor.closeDialog();
      else if (doc.selection.kind === 'cell') doc.selectShape(doc.selection.slideIndex, doc.selection.shapeId);
      else if (!editor.exitGroup()) doc.clearShapeSelection();
    }
  }
  function onCellClipboard(event: ClipboardEvent) {
    const target = eventTarget(event) as HTMLElement;
    if (!editor.ownsEvent(event) || event.defaultPrevented || editor.activeDialog || doc.selection.kind !== 'cell' ||
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

  // The preview page reports a finished rehearsal (Slide Show ▸ Rehearse Timings).
  function onParentMessage(event: MessageEvent) {
    if (!editor.hostFrame || event.origin !== window.location.origin || event.source !== window.parent || event.source === window) return;
    if (event.data?.type !== 'rehearsal-timings' || !Array.isArray(event.data.timings)) return;
    editor.rehearsalTimings = event.data.timings.filter(
      (item: unknown): item is { slide: number; ms: number } =>
        typeof item === 'object' && item !== null && Number.isInteger((item as { slide: unknown }).slide) && Number.isFinite((item as { ms: unknown }).ms),
    );
    editor.activeDialog = 'rehearsal';
  }
</script>

<svelte:window on:message={onParentMessage} on:storage={(event) => { if (event.key === null || event.key === 'office-guide-settings') editor.view.reload(); }} on:keydown={onKeydown} on:copy={onCellClipboard} on:cut={onCellClipboard} on:paste={onCellClipboard} />

<div bind:this={shell} class="ok-editor ok-shell" class:embedded class:compact-host={compactHost} style:--ok-nav-w={navigationWidth === null ? undefined : `${navigationWidth}px`}>
  <TopBar {onsave} {status} {autoSave} compact={compactHost} />
  <div>{#if editor.ribbonVisible}<Ribbon />{/if}</div>
  <div class="ok-body" class:sorter={!slideEditing && editor.viewMode !== 'slideMaster'} class:master={editor.viewMode === 'slideMaster'} class:thumbnails-hidden={slideEditing && !editor.thumbnailsVisible} class:panel-hidden={!commentsOpen && !editor.selectionPaneVisible && !editor.propertiesPaneVisible}>
    {#if editor.viewMode === 'notesPage'}<NotesPageView />
    {:else if editor.viewMode === 'sorter'}<SlideNavigator mode="sorter" />
    {:else if editor.viewMode === 'slideMaster'}<SlideMasterPane /><MasterCanvas />
    {:else if editor.viewMode === 'handoutMaster' || editor.viewMode === 'notesMaster'}{#key editor.viewMode}<PageMasterView kind={editor.viewMode} />{/key}
    {:else}
      {#if editor.thumbnailsVisible}<ThumbnailPane outline={editor.viewMode === 'outline'} />{/if}
      <div class="slide-workspace"><SlideCanvas />{#if editor.notesVisible && doc.currentSlide}{#key doc.currentSlide}<NotesPane />{/key}{/if}</div>{#if commentsOpen}<CommentsDialog />{:else if editor.selectionPaneVisible}{#key doc.currentSlide}<SelectionPane />{/key}{:else}<PropertiesPanel />{/if}
    {/if}
  </div>
  <StatusBar />

  {#if editor.paletteOpen}
    <CommandPalette />
  {/if}
  {#if editor.activeDialog}
    {#if editor.activeDialog === 'reorderObjects'}
      <ReorderObjectsDialog />
    {:else if editor.activeDialog === 'zoom'}
      <ZoomDialog />
    {:else if editor.activeDialog === 'gridOptions'}
      <GridOptionsDialog />
    {:else if editor.activeDialog === 'compressPictures'}
      <CompressPicturesDialog />
    {:else if editor.activeDialog === 'addSlideImage' || editor.activeDialog === 'setShapeImage'}
      {#key editor.activeDialog}<ImageDialog replace={editor.activeDialog === 'setShapeImage'} />{/key}
    {:else if editor.activeDialog === 'setShapeImageCrop'}
      <CropDialog />
    {:else if editor.activeDialog === 'addSlideChart' || editor.activeDialog === 'setChartSpec'}
      {#key editor.activeDialog}<ChartDialog edit={editor.activeDialog === 'setChartSpec'} />{/key}
    {:else if editor.activeDialog === 'replaceTextInPresentation'}
      <FindReplaceDialog />
    {:else if editor.activeDialog === 'setShapeHyperlink' || editor.activeDialog === 'setTableCellClickAction'}
      <LinkDialog />
    {:else if commentsOpen}
      <!-- Docked in the right pane; views without one float it instead. -->
      {#if !slideEditing}<CommentsDialog floating />{/if}
    {:else if editor.activeDialog === 'addSlide'}
      <NewSlideDialog />
    {:else if editor.activeDialog === 'setSlideTransition'}
      <TransitionDialog />
    {:else if editor.activeDialog === 'showProperties'}
      <ShowPropertiesDialog />
    {:else if editor.activeDialog === 'setSlideSize'}
      <SlideSizeDialog />
    {:else if editor.activeDialog === 'rehearsal'}
      <RehearsalDialog />
    {:else if editor.activeDialog === 'headerFooter'}
      <HeaderFooterDialog />
    {:else if editor.activeDialog === 'addSlideTable'}
      <TableDialog />
    {:else if editor.activeDialog === 'customShows'}
      <CustomShowsDialog />
    {:else if editor.activeDialog === 'font'}
      <FontDialog />
    {:else if editor.activeDialog === 'paragraph'}
      <ParagraphDialog properties={targetParagraphProperties(editor)} apply={(edit) => editTargetParagraphs(editor, edit)} onclose={() => editor.closeDialog()} />
    {:else}
      <CommandDialog id={editor.activeDialog} />
    {/if}
  {/if}
  {#if editor.contextMenu}
    <ContextMenu />
  {/if}
  <SymbolPicker />
  <ShapeGallery />
  {#if editor.wordArtGallery}
    <WordArtGallery
      label={t('WordArt')}
      anchor={editor.wordArtGallery}
      close={() => (editor.wordArtGallery = null)}
      choose={(preset) => {
        editor.wordArtGallery = null;
        insertWordArt(editor, t('WordArt'), t('Your text here'), preset);
      }}
    />
  {/if}
  {#if editor.readingView}<ReadingView />{/if}
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
  /* An embedded editor fills the element it is mounted in. */
  .ok-shell.embedded { position: absolute; z-index: auto; }
  .ok-shell > :global(*) { min-width: 0; }
  @media (max-width: 1100px) {
    .ok-shell { --ok-nav-w: 120px; --ok-panel-w: 230px; }
  }
  .slide-workspace { display: grid; grid-template-rows: minmax(0, 1fr) auto; min-height: 0; min-width: 0; overflow: hidden; }
  .ok-body.thumbnails-hidden { grid-template-columns: minmax(0, 1fr) var(--ok-panel-w); }
  .ok-body.panel-hidden { grid-template-columns: var(--ok-nav-w) minmax(0, 1fr); }
  .ok-body.thumbnails-hidden.panel-hidden { grid-template-columns: minmax(0, 1fr); }
  .ok-body.sorter { grid-template-columns: minmax(0, 1fr); }
  /* Slide Master view: the 249 pt master pane beside the editing area. */
  .ok-body.master { grid-template-columns: 249px minmax(0, 1fr); }
  .ok-body {
    display: grid;
    grid-template-columns: var(--ok-nav-w) minmax(0, 1fr) var(--ok-panel-w);
    min-height: 0;
    overflow: hidden;
  }
</style>
