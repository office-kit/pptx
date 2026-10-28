// What each menu-bar item does, keyed by the ids in core/menubar-native.ts.
// Every item reaches an existing editor command. Items the editor can never
// perform are left out of the menus; items unavailable in the current state
// are disabled with the reason as their tooltip. The same lookup drives both
// the menus and their keyboard shortcuts.

import {
  getShapeKind,
  getShapeText,
  getSnapToGrid,
  isSlideHidden,
  setSlideHidden,
  setSnapToGrid,
} from '@office-kit/pptx';
import type { EditorController } from '../core/controller.svelte.ts';
import { eventTarget } from '../core/dom-root.ts';
import { insertMedia } from '../core/insert-objects.ts';
import { APPLY_TO_DEFAULTS, NATIVE_MENUBAR, type NativeMenu } from '../core/menubar-native.ts';
import { menuItemForKey, parseShortcut, isModifiedChord } from '../core/menubar-shortcuts.ts';
import { pruneMenus } from '../core/menubar-visible.ts';
import {
  alignParagraphs,
  canAlignParagraphs,
  paragraphAlignment,
} from '../core/paragraph-alignment.ts';
import { targetParagraphProperties } from '../core/paragraph-targets.ts';
import {
  addSection,
  removeSection,
  sectionOf,
  sectionRanges,
  UNTITLED_SECTION,
} from '../core/sections.ts';
import { selectedShapeIds, selectedSlideIndices } from '../core/selection.ts';
import type { GalleryShape } from '../core/shape-gallery.ts';
import { getLocale, t } from '../i18n/i18n.svelte.ts';
import { ALIGN_ITEMS, rotateSelection, type RotateAction } from '../ribbon/arrange-actions.ts';

export interface MenuCommand {
  readonly run?: () => void;
  readonly disabled?: boolean;
  /** Why a disabled item is unavailable (an English message key). */
  readonly reason?: string;
  /** A check mark; undefined for items that never show one. */
  readonly checked?: boolean;
  /** One of a set of mutually exclusive choices (View ▸ Normal …). */
  readonly radio?: boolean;
  /** Replaces the native label, already localized (Edit ▸ Undo <edit> …). */
  readonly label?: string;
}

/** What only the title bar can do: it owns the file input and the host's save. */
export interface MenuHost {
  save(): void;
  download(): void;
  open(): void;
  newPresentation(): void;
  /** Opens Format ▸ Columns... (the dialog lives with the menu bar). */
  columns(): void;
  /** Where a gallery opened from the menu bar hangs. */
  anchor(): HTMLElement | null;
  readonly fullScreen: boolean;
  toggleFullScreen(): void;
}

const FEEDBACK = 'help/feedback';

const unavailable = (reason: string): MenuCommand => ({ disabled: true, reason });

// Reasons shared by several items.
const SLIDE_SHOW_HOST = 'The slide show runs in the preview page around the editor.';
const OUTLINE_PASTE = 'Press ⌘V to paste into the outline.';
const SLIDE_TEXT = 'Edit the text on the slide to use this.';

const SHAPES: Readonly<Record<string, GalleryShape>> = {
  'insert/shape/rectangle': 'rect',
  'insert/shape/rounded-rectangle': 'roundRect',
  'insert/shape/triangle': 'triangle',
  'insert/shape/oval': 'ellipse',
  'insert/shape/line': 'line',
};

const CHARTS: Readonly<Record<string, string>> = {
  'insert/chart/column': 'column',
  'insert/chart/bar': 'bar',
  'insert/chart/line': 'line',
  'insert/chart/area': 'area',
  'insert/chart/pie': 'pie',
  'insert/chart/radar': 'radar',
};

const ROTATIONS: Readonly<Record<string, RotateAction>> = {
  'arrange/rotate-or-flip/rotate-left-90': 'left',
  'arrange/rotate-or-flip/rotate-right-90': 'right',
  'arrange/rotate-or-flip/flip-horizontal': 'horizontal',
  'arrange/rotate-or-flip/flip-vertical': 'vertical',
};

const ARRANGE_ALIGN: Readonly<Record<string, (typeof ALIGN_ITEMS)[number]['value']>> = {
  'arrange/align-or-distribute/align-left': 'left',
  'arrange/align-or-distribute/align-center': 'center',
  'arrange/align-or-distribute/align-right': 'right',
  'arrange/align-or-distribute/align-top': 'top',
  'arrange/align-or-distribute/align-middle': 'middle',
  'arrange/align-or-distribute/align-bottom': 'bottom',
};

const ORDER: Readonly<Record<string, string>> = {
  'arrange/bring-to-front': 'bringShapeToFront',
  'arrange/send-to-back': 'sendShapeToBack',
  'arrange/bring-forward': 'bringShapeForward',
  'arrange/send-backward': 'sendShapeBackward',
  'arrange/group': 'groupShapes',
  'arrange/ungroup': 'ungroupShapes',
};

const VIEWS: Readonly<Record<string, EditorController['viewMode']>> = {
  'view/normal': 'normal',
  'view/slide-sorter': 'sorter',
  'view/notes-page': 'notesPage',
  'view/outline-view': 'outline',
  'view/master/slide-master': 'slideMaster',
  'view/master/handout-master': 'handoutMaster',
  'view/master/notes-master': 'notesMaster',
};

const PARAGRAPH_ALIGN: Readonly<Record<string, string>> = {
  'format/alignment/align-left': 'left',
  'format/alignment/center': 'center',
  'format/alignment/align-right': 'right',
  'format/alignment/justify': 'justify',
  'format/alignment/distributed': 'distribute',
};

// Items the editor can never perform: it has no counterpart, or a browser
// page cannot do them. They stay out of the menus.
const HIDDEN = new Set([
  'file/new-from-template',
  'file/close',
  'file/save-as-template',
  'file/export',
  'file/move',
  'file/rename',
  'file/browse-version-history',
  'file/always-open-read-only',
  'file/restrict-permissions',
  'file/passwords',
  'file/print',
  'edit/paste-special',
  'edit/clear',
  'edit/rename-section',
  'edit/find/find-next',
  'edit/find/find-previous',
  'edit/find/replace-fonts',
  'edit/start-dictation',
  'edit/emoji-and-symbols',
  'view/message-bar',
  'view/markup',
  'view/advanced-markup',
  'insert/picture/photo-browser',
  'insert/picture/stock-images',
  'insert/picture/online-pictures',
  'insert/picture/brand-images',
  'insert/audio/audio-browser',
  'insert/audio/record-audio',
  'insert/video/movie-browser',
  'insert/video/stock-videos',
  'insert/video/online-movie',
  'insert/cameo',
  'insert/equation',
  'insert/icons',
  'insert/object',
  'format/bullets-and-numbering',
  'format/animation-painter',
  'format/replace-fonts',
  'format/theme-colors',
  'arrange/reorder-objects',
  'tools/spelling',
  'tools/thesaurus',
  'tools/translate',
  'tools/set-proofing-language',
  'tools/autocorrect-options',
  'tools/add-ins',
  'slide-show/rehearse-with-coach',
  'slide-show/always-use-subtitles',
  'help/clear-application-data',
  'help/check-for-updates',
]);

// Whole submenus (and the Window menu) the editor has no counterpart for.
const HIDDEN_GROUPS = [
  'file/open-recent',
  'file/share',
  'edit/autofill',
  'insert/slides-from',
  'insert/smartart',
  'insert/3d-models',
  'insert/zoom',
  'insert/action-buttons',
  'tools/macro',
  'slide-show/subtitle-settings',
  'window',
];

function isHidden(id: string): boolean {
  if (HIDDEN.has(id)) return true;
  if (id.startsWith('insert/chart/')) return !(id in CHARTS);
  return HIDDEN_GROUPS.some((group) => id === group || id.startsWith(`${group}/`));
}

// Help ▸ Feedback is only there when the host gives the editor a feedback URL
// (the internal dev mount does).
const MENUS = {
  en: pruneMenus(NATIVE_MENUBAR.en, isHidden),
  ja: pruneMenus(NATIVE_MENUBAR.ja, isHidden),
};
const MENUS_WITHOUT_FEEDBACK = {
  en: pruneMenus(MENUS.en, (id) => id === FEEDBACK),
  ja: pruneMenus(MENUS.ja, (id) => id === FEEDBACK),
};

/** The menus the menu bar shows. */
export function visibleMenus(editor: EditorController): readonly NativeMenu[] {
  return (editor.feedbackUrl === undefined ? MENUS_WITHOUT_FEEDBACK : MENUS)[getLocale()];
}

/**
 * Every menu item, for keyboard shortcuts: a hidden item's shortcut still
 * resolves, so the browser does not act on it (⌘P would print the page).
 */
export function nativeMenus(): readonly NativeMenu[] {
  return NATIVE_MENUBAR[getLocale()];
}

/** The command behind a menu-bar item in the editor's current state. */
export function menuCommand(editor: EditorController, host: MenuHost, id: string): MenuCommand {
  const doc = editor.doc;
  const text = editor.inlineTextFormat;
  const slide = doc.currentSlide;
  const shapes = editor.selectedShapes();
  const shapeSelected = !text && doc.selection.kind === 'shape' && shapes.length > 0;
  const objectSelected = !!text || doc.selection.kind === 'cell' || shapes.length > 0;
  const locked = editor.selectionLocked();
  const run = (command: string, args?: Record<string, unknown>): MenuCommand => ({
    disabled: !editor.canRun(command),
    run: () => editor.runOrPrompt(command, args),
  });
  const present = (action: Parameters<EditorController['present']>[0]): MenuCommand =>
    editor.canPresent ? { run: () => editor.present(action) } : unavailable(SLIDE_SHOW_HOST);

  if (isHidden(id)) return { disabled: true };
  if (id in VIEWS) {
    const mode = VIEWS[id]!;
    return {
      radio: true,
      checked: editor.viewMode === mode,
      run: () => editor.setViewMode(mode),
    };
  }
  if (id in ORDER) return { ...run(ORDER[id]!), run: () => editor.invoke(ORDER[id]!) };
  if (id in PARAGRAPH_ALIGN) {
    const value = PARAGRAPH_ALIGN[id]!;
    const enabled = canAlignParagraphs(editor);
    return {
      disabled: !enabled || locked,
      checked: enabled && paragraphAlignment(editor) === value,
      run: () => alignParagraphs(editor, value),
    };
  }
  if (id in ROTATIONS) {
    return {
      disabled: !shapes.length || locked,
      run: () => rotateSelection(editor, ROTATIONS[id]!),
    };
  }
  const toSlide = shapes.length < 2 || editor.alignmentReference === 'slide';
  if (id in ARRANGE_ALIGN) {
    const value = ARRANGE_ALIGN[id]!;
    return {
      disabled: !shapes.length || locked,
      run: () => editor.alignSelection(value, toSlide ? 'slide' : 'selection'),
    };
  }
  if (id in SHAPES) {
    const shape = SHAPES[id]!;
    return { disabled: !editor.canRun('addSlideShape'), run: () => (editor.drawShape = shape) };
  }
  if (id in CHARTS) return run('addSlideChart', { kind: CHARTS[id] });

  const guides = editor.guidesVisible();
  const saveView = (update: Partial<{ grid: boolean; smart: boolean; drawing: boolean }>) =>
    editor.view.save({
      grid: editor.view.grid,
      smart: editor.view.smart,
      drawing: guides,
      ...update,
    });
  const sorter = editor.viewMode === 'sorter';
  const section = slide ? sectionOf(sectionRanges(doc.pres), doc.selection.slideIndex) : null;
  const slides = selectedSlideIndices(doc.selection)
    .map((index) => doc.slideAt(index))
    .filter((item) => item !== null);
  const hidden = slides.length > 0 && slides.every((item) => isSlideHidden(item));
  const hasText =
    !!text ||
    doc.selection.kind === 'cell' ||
    shapes.some((shape) => getShapeKind(shape) === 'shape' && getShapeText(shape).length > 0);
  const paragraphs = targetParagraphProperties(editor).length > 0 && !locked;
  const textShapes = shapes.some((shape) => getShapeKind(shape) === 'shape');
  const canCopy =
    !!text || doc.selection.kind === 'slide' || selectedShapeIds(doc.selection).length > 0;

  switch (id) {
    // File
    case 'file/new-presentation':
      return { run: host.newPresentation };
    case 'file/open':
      return { run: host.open };
    case 'file/save':
      return { run: host.save };
    case 'file/save-as':
      return { run: host.download };
    case 'file/compress-pictures':
      return { run: () => (editor.activeDialog = 'compressPictures') };
    case 'file/page-setup':
      return run('setSlideSize');
    case 'file/properties':
      return run('setCoreProperties');

    // Edit
    case 'edit/undo':
      return {
        disabled: !doc.canUndo,
        label: doc.undoLabel ? `${t('Undo')} ${doc.undoLabel}` : t("Can't Undo"),
        run: () => void doc.undo(),
      };
    case 'edit/repeat':
      return {
        disabled: !doc.canRedo,
        label: doc.redoLabel ? `${t('Redo')} ${doc.redoLabel}` : t("Can't Repeat"),
        run: () => void doc.redo(),
      };
    case 'edit/cut':
    case 'edit/copy': {
      const action = id === 'edit/cut' ? 'cut' : 'copy';
      if (text) {
        const clipboard = text.clipboard;
        return {
          disabled: text.hasSelection === false,
          run: () => (clipboard ? clipboard(action) : document.execCommand(action)),
        };
      }
      return {
        disabled: !canCopy,
        run: () => void (action === 'cut' ? editor.cutSelection() : editor.copySelection()),
      };
    }
    case 'edit/paste':
      if (!text) return { run: () => void editor.paste() };
      return text.clipboard ? { run: () => text.clipboard!('paste') } : unavailable(OUTLINE_PASTE);
    case 'edit/paste-and-match-formatting':
      if (!text) return unavailable('Paste and Match Formatting pastes into text being edited.');
      return text.pastePlain ? { run: text.pastePlain } : unavailable(OUTLINE_PASTE);
    case 'edit/select-all':
      return {
        disabled: !slide,
        run: () => (text ? document.execCommand('selectAll') : editor.selectAll()),
      };
    case 'edit/duplicate':
      return {
        disabled: !(shapeSelected || doc.selection.kind === 'slide'),
        run: () => editor.duplicateSelection(),
      };
    case 'edit/delete-slide':
      return { disabled: !slide, run: () => editor.invoke('removeSlide') };
    case 'edit/remove-section':
      return section
        ? {
            run: () =>
              doc.transact(t('Remove Section'), () => removeSection(doc.pres, section.start)),
          }
        : unavailable('The current slide is not in a section.');
    case 'edit/find':
      return {};
    case 'edit/find/find':
    case 'edit/find/advanced-find':
    case 'edit/find/replace':
      return { run: () => editor.runOrPrompt('replaceTextInPresentation') };
    case 'edit/select-data':
      return run('setChartSpec');
    case 'edit/toggle-drawing':
      return { disabled: !slide, run: () => editor.ink.toggle('pen') };

    // View
    // Without a host viewer, Reading View is the editor's own full-window mode.
    case 'view/reading-view':
      return { run: () => editor.openReadingView() };
    case 'view/presenter-view':
      return present('presenter');
    case 'view/slide-show':
      return present('start');
    case 'view/show-slides':
      return {
        disabled: sorter,
        run: () => (editor.thumbnailsVisible = !editor.thumbnailsVisible),
      };
    case 'view/master':
      return {};
    case 'view/ribbon':
      return {
        checked: editor.ribbonVisible,
        run: () => (editor.ribbonVisible = !editor.ribbonVisible),
      };
    case 'view/header-and-footer':
    case 'insert/header-and-footer':
      return { disabled: !slide, run: () => (editor.activeDialog = 'headerFooter') };
    case 'view/ruler':
      return {
        disabled: sorter,
        checked: editor.view.ruler,
        run: () => editor.view.save({ ruler: !editor.view.ruler }),
      };
    case 'view/grid-and-guides':
    case 'view/zoom':
      return {};
    case 'view/grid-and-guides/smart-guides':
      return {
        disabled: sorter,
        checked: editor.view.smart,
        run: () => saveView({ smart: !editor.view.smart }),
      };
    case 'view/grid-and-guides/guides':
      return { disabled: sorter, checked: guides, run: () => saveView({ drawing: !guides }) };
    case 'view/grid-and-guides/gridlines':
      return {
        disabled: sorter,
        checked: editor.view.grid,
        run: () => saveView({ grid: !editor.view.grid }),
      };
    case 'view/grid-and-guides/snap-to-grid': {
      const snapping = getSnapToGrid(doc.pres) ?? false;
      return {
        disabled: sorter,
        checked: snapping,
        run: () => doc.transact(t('Snap to Grid'), () => setSnapToGrid(doc.pres, !snapping)),
      };
    }
    case 'view/grid-and-guides/grid-options':
      return { disabled: sorter, run: () => (editor.activeDialog = 'gridOptions') };
    case 'view/zoom/fit-to-window':
      return { run: () => editor.zoomFit() };
    case 'view/zoom/zoom-in':
      return { run: () => editor.zoomIn() };
    case 'view/zoom/zoom-out':
      return { run: () => editor.zoomOut() };
    case 'view/zoom/zoom':
      return { run: () => (editor.activeDialog = 'zoom') };
    case 'view/enter-full-screen':
      return {
        disabled: typeof document === 'undefined' || !document.fullscreenEnabled,
        reason: 'The browser does not allow full screen here.',
        label: host.fullScreen ? t('Exit Full Screen') : undefined,
        run: host.toggleFullScreen,
      };

    // Insert
    case 'insert/new-slide':
      return {
        disabled: !editor.canRun('addSlide') && !editor.canRun('addBlankSlide'),
        run: () => editor.addNewSlide(),
      };
    case 'insert/duplicate-slide':
      return { disabled: !slide, run: () => editor.invoke('duplicateSlide') };
    case 'insert/section':
      return {
        disabled: !slide,
        run: () =>
          doc.transact(t('Add Section'), () =>
            addSection(doc.pres, doc.selection.slideIndex, t(UNTITLED_SECTION)),
          ),
      };
    case 'insert/comment':
      return run('addSlideComment');
    case 'insert/text-art':
      return {
        disabled: !slide,
        run: () => {
          const anchor = host.anchor();
          if (anchor) editor.openTextArtGallery(anchor);
        },
      };
    case 'insert/date-and-time':
      return run('setShapeTextField');
    case 'insert/slide-number':
      return run('setShapeTextField', { type: 'slidenum' });
    case 'insert/table':
      return run('addSlideTable');
    case 'insert/chart':
    case 'insert/picture':
    case 'insert/audio':
    case 'insert/video':
    case 'insert/shape':
      return {};
    case 'insert/picture/picture-from-file':
      return run('addSlideImage');
    case 'insert/audio/audio-from-file':
      return { disabled: !slide, run: () => void insertMedia(editor, 'audio', t('Audio')) };
    case 'insert/video/movie-from-file':
      return { disabled: !slide, run: () => void insertMedia(editor, 'video', t('Video')) };
    case 'insert/symbol':
      return text?.insertText
        ? {
            run: () => {
              const anchor = host.anchor();
              if (anchor) editor.openSymbolPicker(anchor);
            },
          }
        : unavailable('Place the cursor in text to insert a symbol.');
    case 'insert/action-settings':
      return run('setShapeClickAction');
    case 'insert/hyperlink':
      if (text) {
        if (!text.hyperlink) return unavailable(SLIDE_TEXT);
        return text.hasSelection
          ? { run: text.hyperlink }
          : unavailable('Select the text to link first.');
      }
      return run('setShapeHyperlink');

    // Format
    case 'format/font':
      return {
        disabled: !hasText,
        run: () => editor.openFontDialog('font', text?.element?.()),
      };
    case 'format/paragraph':
      return { disabled: !paragraphs, run: () => editor.openParagraphDialog(text?.element?.()) };
    case 'format/columns':
      return { disabled: !textShapes || locked, run: host.columns };
    case 'format/alignment':
      return {};
    case 'format/more-options':
      return { disabled: !textShapes && !text, run: () => editor.showTextFormat('textbox') };
    case 'format/pick-up-object-style':
      if (text) return text.pickUpStyle ? { run: text.pickUpStyle } : unavailable(SLIDE_TEXT);
      return { disabled: !shapeSelected, run: () => editor.copyObjectFormat() };
    case 'format/apply-object-style':
      if (text) return text.applyStyle ? { run: text.applyStyle } : unavailable(SLIDE_TEXT);
      return shapeSelected
        ? { run: () => editor.pasteObjectFormat() }
        : { disabled: true, label: APPLY_TO_DEFAULTS[getLocale()] };
    case 'format/slide-background':
      return { disabled: !slide, run: () => editor.showBackgroundFormat() };
    case 'format/crop':
      // The command also accepts picture-filled shapes; Crop is for pictures.
      return shapeSelected && shapes.every((shape) => getShapeKind(shape) === 'picture')
        ? run('setShapeImageCrop')
        : { disabled: true };
    case 'format/format-object':
      return { disabled: !objectSelected, run: () => editor.showShapeFormat() };

    // Arrange
    case 'arrange/reorder-overlapping-objects':
      return {
        disabled: editor.reorderMembers().length < 2,
        run: () => (editor.activeDialog = 'reorderObjects'),
      };
    case 'arrange/regroup':
      return { disabled: !editor.canRegroup(), run: () => editor.regroupSelection() };
    case 'arrange/rotate-or-flip':
    case 'arrange/align-or-distribute':
      return {};
    case 'arrange/rotate-or-flip/more-rotation-options':
      return {
        disabled: !editor.canRun('setShapeRotation'),
        run: () => editor.showRotationOptions(),
      };
    case 'arrange/align-or-distribute/distribute-horizontally':
    case 'arrange/align-or-distribute/distribute-vertically':
      return {
        disabled: !shapes.length || locked || (!toSlide && shapes.length < 3),
        run: () =>
          editor.distributeSelection(id.endsWith('horizontally') ? 'horizontal' : 'vertical'),
      };
    case 'arrange/align-or-distribute/align-to-slide':
      return {
        disabled: !shapes.length,
        checked: toSlide,
        radio: true,
        run: () => (editor.alignmentReference = 'slide'),
      };
    case 'arrange/align-or-distribute/align-selected-objects':
      return {
        disabled: shapes.length < 2,
        checked: !toSlide,
        radio: true,
        run: () => (editor.alignmentReference = 'selection'),
      };
    case 'arrange/selection-pane':
      return {
        run: () => {
          editor.setViewMode('normal');
          editor.selectionPaneVisible = !editor.selectionPaneVisible;
        },
      };

    // Tools
    case 'tools/check-accessibility':
      return { run: () => (editor.accessibilityOpen = true) };

    // Slide Show
    case 'slide-show/play-from-start':
      return present('start');
    case 'slide-show/play-from-current-slide':
      return present('current');
    case 'slide-show/custom-slide-show':
      return doc.customShows.length ? {} : unavailable('The presentation has no custom shows.');
    case 'slide-show/custom-slide-show/custom-shows':
      return { run: () => editor.openCustomShows() };
    case 'slide-show/presenter-view':
      return present('presenter');
    case 'slide-show/rehearse-timings':
      return present('rehearse');
    case 'slide-show/record-slide-show':
      return present('start');
    case 'slide-show/hide-slide':
      return {
        disabled: slides.length === 0,
        checked: hidden,
        run: () =>
          doc.transact(t('Hide Slide'), () => {
            for (const item of slides) setSlideHidden(item, !hidden);
          }),
      };
    case 'slide-show/set-up-show':
      return { run: () => (editor.activeDialog = 'showProperties') };

    // Help
    case 'help/editor-help':
      return { run: () => editor.togglePalette(true) };
    case FEEDBACK: {
      // Not in the menus without a URL (see nativeMenus).
      const url = editor.feedbackUrl;
      return url === undefined
        ? { disabled: true }
        : { run: () => void window.open(url, '_blank', 'noopener') };
    }
  }
  throw new Error(`Menu item ${id} has no command`);
}

// While text is being edited these keys belong to the text: the browser's
// clipboard and undo, and ⌘↩, which ends editing as it always has here.
const TEXT_KEYS = new Set([
  'edit/undo',
  'edit/repeat',
  'edit/cut',
  'edit/copy',
  'edit/paste',
  'edit/select-all',
  'slide-show/play-from-current-slide',
]);

// The window keymap (EditorApp) keeps undo, the clipboard and Select All: the
// clipboard runs through the browser's copy/paste events for table cells.
const WINDOW_KEYS = new Set([
  'edit/undo',
  'edit/repeat',
  'edit/cut',
  'edit/copy',
  'edit/paste',
  'edit/select-all',
]);

// Commands that act on the whole document still work from a text field
// (notes, the outline, a pane's input); the rest would target the slide.
const DOCUMENT_SCOPE = [
  'file/',
  'view/normal',
  'view/slide-sorter',
  'view/notes-page',
  'view/outline-view',
  'view/reading-view',
  'view/slide-show',
  'view/master/',
  'view/ribbon',
  'slide-show/',
  'edit/find/',
  'insert/new-slide',
  'help/',
];

/**
 * Whether the slide's text editor should let a key reach the menu bar: menu
 * shortcuts with ⌘ or ⌃ act on the selected text, except the keys the text
 * itself owns.
 */
export function menuKeyLeavesText(editor: EditorController, event: KeyboardEvent): boolean {
  const item = menuItemForKey(nativeMenus(), event);
  if (!item || TEXT_KEYS.has(item.id)) return false;
  const chord = parseShortcut(item.shortcut!);
  return chord !== null && isModifiedChord(chord);
}

/** Runs the menu item a key event names. Returns true when the key was taken. */
export function runMenuKey(
  editor: EditorController,
  host: MenuHost,
  event: KeyboardEvent,
): boolean {
  if (
    event.isComposing ||
    event.defaultPrevented ||
    editor.activeDialog ||
    !editor.ownsEvent(event)
  )
    return false;
  const item = menuItemForKey(nativeMenus(), event);
  if (!item || WINDOW_KEYS.has(item.id)) return false;
  const origin = eventTarget(event);
  const target = origin instanceof HTMLElement ? origin : null;
  const inText = !!target?.closest('.inline-edit-shell') && !!editor.inlineTextFormat;
  const typing =
    !!target && (target.isContentEditable || /^(INPUT|TEXTAREA|SELECT)$/.test(target.tagName));
  if (inText && TEXT_KEYS.has(item.id)) return false;
  if (typing && !inText && !DOCUMENT_SCOPE.some((scope) => item.id.startsWith(scope))) return false;
  const chord = parseShortcut(item.shortcut!)!;
  const command = menuCommand(editor, host, item.id);
  if (command.disabled || !command.run) {
    // A disabled command still keeps ⌘/⌃ keys from the browser (⌘D would
    // bookmark the page); plain keys go on to whatever has focus.
    if (isModifiedChord(chord)) event.preventDefault();
    return isModifiedChord(chord);
  }
  event.preventDefault();
  command.run();
  return true;
}
