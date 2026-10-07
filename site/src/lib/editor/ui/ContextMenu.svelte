<script lang="ts">
  import { tick } from 'svelte';
  import { getShapeChartSpec, getShapeKind, getShapeMedia, getSnapToGrid, isShapeLocked, isSlideHidden, type SlideShapeData, setParagraphBullet, setSlideHidden, setSnapToGrid, setSlideOutlineCollapsed } from '@office-kit/pptx';
  // Right-click menu. Items adapt to the current selection and dispatch through
  // the controller's actions (which go through the same undoable command path).
  // The menus follow Mac PowerPoint 16's, in its order, with its separators and
  // shortcut hints (native captures 2026-10-07); commands the editor cannot
  // perform are shown disabled with the reason as their tooltip.
  import { getEditor } from '../core/context.ts';
  import type { TextContextMenu } from '../core/controller.svelte.ts';
  import { BULLET_GALLERY, NUMBERING_GALLERY, listStyleApplied, type ListGalleryEntry } from '../core/list-galleries.ts';
  import { editTargetParagraphs, targetParagraphProperties } from '../core/paragraph-targets.ts';
  import { selectedSlideIndices } from '../core/selection.ts';
  import { addSection, UNTITLED_SECTION } from '../core/sections.ts';
  import { canEditTableLines, canMergeTableBlock, deleteTable, deleteTableLines, insertTableLines, mergeSelectedCells, mergedSelectedCell, selectedCellTarget, selectTableLines, splitSelectedCell, type TableAxis } from '../core/table-commands.ts';
  import { t } from '../i18n/i18n.svelte.ts';

  const editor = getEditor();
  const doc = editor.doc;
  const selected = $derived(selectedSlideIndices(doc.selection));
  const firstSelected = $derived(selected[0] ?? 0);
  const menu = $derived(editor.contextMenu!);

  type Item = {
    label: string;
    accel?: string;
    disabled?: boolean;
    checked?: boolean;
    sep?: boolean;
    /** Why a disabled item is unavailable, shown as its tooltip. */
    reason?: string;
    /** Gallery tile preview: the markers of a list style (empty for None). */
    tile?: readonly string[];
  } & ({ run: () => void; children?: never } | { children: Item[]; run?: never });
  const noop = () => {};
  let submenu = $state<string | null>(null);
  function collapse(collapsed: boolean, all: boolean) {
    const slides = all ? doc.slides : selected.map(index => doc.slideAt(index)!);
    doc.transact(t(collapsed ? 'Collapse' : 'Expand'), () => setSlideOutlineCollapsed(slides, collapsed));
  }

  function outlineCollapseItems(): Item[] {
    return [true, false].map(collapsed => ({
      label: collapsed ? 'Collapse' : 'Expand',
      children: [
        { label: collapsed ? 'Collapse' : 'Expand', run: () => collapse(collapsed, false) },
        { label: collapsed ? 'Collapse All' : 'Expand All', run: () => collapse(collapsed, true) },
      ],
    }));
  }

  function slideItems(): Item[] {
    return [
      { label: 'New Slide', run: () => editor.addNewSlide() },
      { label: 'Duplicate Slide', accel: '⌘D', run: () => editor.invoke('duplicateSlide') },
      { label: 'Delete Slide', accel: 'Del', run: () => editor.invoke('removeSlide'), sep: true },
    ];
  }

  // PowerPoint names the Format pane after what is selected.
  function formatLabel(shapes: readonly SlideShapeData[]): string {
    if (shapes.length !== 1) return 'Format Shape...';
    const shape = shapes[0]!;
    if (getShapeMedia(shape)?.kind === 'video') return 'Format Video...';
    if (getShapeChartSpec(shape)) return 'Format Chart Area...';
    if (getShapeKind(shape) === 'picture') return 'Format Picture...';
    return 'Format Shape...';
  }

  function clipboardItems(text: TextContextMenu | undefined): Item[] {
    if (text) return [
      { label: 'Cut', accel: '⌘X', run: text.cut, disabled: !text.hasSelection },
      { label: 'Copy', accel: '⌘C', run: text.copy, disabled: !text.hasSelection },
      { label: 'Paste', accel: '⌘V', run: text.paste, sep: true },
    ];
    return [
      { label: 'Cut', accel: '⌘X', run: () => editor.cutSelection() },
      { label: 'Copy', accel: '⌘C', run: () => editor.copySelection() },
      { label: 'Paste', accel: '⌘V', run: () => editor.paste(), disabled: !editor.hasClipboard(), sep: true },
    ];
  }

  function arrangeItems(): Item[] {
    return [
      {
        label: 'Group',
        children: [
          { label: 'Group', accel: '⌥⌘G', run: () => editor.invoke('groupShapes'), disabled: !editor.canRun('groupShapes') },
          { label: 'Regroup', accel: '⌥⌘J', run: () => editor.regroupSelection(), disabled: !editor.canRegroup(), sep: true },
          { label: 'Ungroup', accel: '⌥⇧⌘G', run: () => editor.invoke('ungroupShapes'), disabled: !editor.canRun('ungroupShapes') },
        ],
      },
      {
        label: 'Bring to Front',
        children: [
          { label: 'Bring to Front', accel: '⇧⌘F', run: () => editor.invoke('bringShapeToFront') },
          { label: 'Bring Forward', accel: '⌥⇧⌘F', run: () => editor.invoke('bringShapeForward') },
        ],
      },
      {
        label: 'Send to Back',
        children: [
          { label: 'Send to Back', accel: '⇧⌘B', run: () => editor.invoke('sendShapeToBack') },
          { label: 'Send Backward', accel: '⌥⇧⌘B', run: () => editor.invoke('sendShapeBackward') },
        ],
      },
    ];
  }

  function reorderItems(): Item[] {
    return [
      { label: 'Reorder Objects', disabled: true, reason: 'The layered stacking view is not available in this editor.', run: noop },
      { label: 'Reorder Overlapping Objects', sep: true, disabled: editor.reorderMembers().length < 2, run: () => editor.activeDialog = 'reorderObjects' },
    ];
  }

  function lockItem(shapes: readonly SlideShapeData[]): Item {
    const locked = shapes.length > 0 && shapes.every((shape) => isShapeLocked(shape));
    return { label: locked ? 'Unlock' : 'Lock', sep: true, disabled: !shapes.length, run: () => editor.lockObjects(shapes, !locked) };
  }

  const newComment = (): Item => ({ label: 'New Comment', accel: '⇧⌘M', run: () => editor.runOrPrompt('addSlideComment') });

  // Font... / Paragraph... / Bullets ▸ / Numbering ▸ act on the selected text
  // range while editing, or on every paragraph of the selected cells.
  function paragraphItems(text: TextContextMenu | undefined): Item[] {
    const paragraphs = targetParagraphProperties(editor);
    const enabled = paragraphs.length > 0 && !editor.selectionLocked();
    const bullets = paragraphs.map((paragraph) => paragraph.bullet);
    const gallery = (label: string, entries: readonly ListGalleryEntry[]): Item => ({
      label,
      disabled: !enabled,
      children: [
        ...entries.map((entry): Item => ({
          label: entry.label,
          tile: entry.markers,
          checked: listStyleApplied(entry, bullets),
          run: () => editTargetParagraphs(editor, (shape, index) => setParagraphBullet(shape, index, entry.style)),
        })),
        { label: 'Bullets and Numbering...', disabled: true, reason: 'The editor has no Bullets and Numbering dialog yet.', run: noop },
      ],
    });
    return [
      { label: 'Font...', accel: '⌘T', run: () => editor.openFontDialog('font', text?.element) },
      { label: 'Paragraph...', accel: '⌥⌘M', disabled: !enabled, run: () => editor.openParagraphDialog(text?.element) },
      gallery('Bullets', BULLET_GALLERY),
      { ...gallery('Numbering', NUMBERING_GALLERY), sep: true },
    ];
  }

  function proofingItems(): Item[] {
    return [
      { label: 'Thesaurus...', accel: '⌃⌥⌘R', disabled: true, reason: 'The thesaurus needs the Microsoft reference service.', run: noop },
      { label: 'Translate...', sep: true, disabled: true, reason: 'Translation needs the Microsoft translation service.', run: noop },
    ];
  }

  // The Format pane's Text Options cover shape text; table cells have none yet.
  function textFormatItems(cell: boolean): Item[] {
    return [
      cell
        ? { label: 'Format Text Effects...', disabled: true, reason: 'The Format pane has no Text Options for table cells yet.', run: noop }
        : { label: 'Format Text Effects...', run: () => editor.showTextFormat('textEffects') },
      { label: 'Format Shape...', accel: '⇧⌘1', sep: true, run: () => editor.showShapeFormat() },
    ];
  }

  // Mac PowerPoint's menu for text being edited in a shape.
  function textItems(text: TextContextMenu): Item[] {
    return [
      ...clipboardItems(text),
      { label: 'Exit Edit Text', run: text.exit },
      ...paragraphItems(text),
      ...proofingItems(),
      ...textFormatItems(false),
      lockItem(editor.selectedShapes()),
      { label: 'Hyperlink...', accel: '⌘K', sep: true, disabled: !text.hasSelection, reason: 'Select the text to link first.', run: text.hyperlink },
      newComment(),
    ];
  }

  // Mac PowerPoint's menu for table cells (a caret or selected cells). Row and
  // column commands end text editing first, as they reshape the table.
  function cellItems(text: TextContextMenu | undefined): Item[] {
    const target = selectedCellTarget(editor);
    const lines = canEditTableLines(editor);
    const linesReason = 'Split merged cells before changing rows or columns';
    const table = (run: () => void) => () => { text?.exit(); run(); };
    const insert = (label: string, axis: TableAxis, after: boolean): Item => ({ label, disabled: !lines, reason: linesReason, run: table(() => insertTableLines(editor, axis, after)) });
    const remove = (label: string, axis: TableAxis): Item => ({ label, disabled: !lines, reason: linesReason, run: table(() => deleteTableLines(editor, axis)) });
    const select = (label: string, axis: TableAxis): Item => ({ label, run: table(() => selectTableLines(editor, axis)) });
    return [
      ...clipboardItems(text),
      ...paragraphItems(text),
      {
        label: 'Insert',
        children: [
          insert('Insert Columns to the Left', 'column', false),
          insert('Insert Columns to the Right', 'column', true),
          insert('Insert Rows Above', 'row', false),
          insert('Insert Rows Below', 'row', true),
        ],
      },
      {
        label: 'Delete',
        children: [
          remove('Delete Columns', 'column'),
          remove('Delete Rows', 'row'),
          { label: 'Delete Table', run: table(() => deleteTable(editor)) },
        ],
      },
      {
        label: 'Select',
        sep: true,
        children: [
          { label: 'Select Table', run: table(() => {
            const selection = doc.selection;
            if (selection.kind === 'cell') doc.selectShape(selection.slideIndex, selection.shapeId);
          }) },
          select('Select Column', 'column'),
          select('Select Row', 'row'),
        ],
      },
      { label: 'Merge Cells', disabled: !target || !canMergeTableBlock(target.cells, target.block), run: table(() => mergeSelectedCells(editor)) },
      { label: 'Split Cells...', sep: true, disabled: !mergedSelectedCell(editor), reason: 'The editor can split merged cells only.', run: table(() => splitSelectedCell(editor)) },
      ...proofingItems(),
      ...textFormatItems(true),
      lockItem(editor.selectedShapes()),
      text
        ? { label: 'Hyperlink...', accel: '⌘K', sep: true, disabled: !text.hasSelection, reason: 'Select the text to link first.', run: text.hyperlink }
        : { label: 'Hyperlink...', accel: '⌘K', sep: true, disabled: !editor.canRun('setTableCellClickAction'), run: () => editor.runOrPrompt('setTableCellClickAction') },
      newComment(),
    ];
  }

  // Mac PowerPoint's picture menu.
  function pictureItems(shapes: readonly SlideShapeData[]): Item[] {
    const online = 'Online pictures need Microsoft 365 services.';
    return [
      ...clipboardItems(undefined),
      {
        label: 'Change Picture',
        sep: true,
        children: [
          { label: 'From a File...', run: () => editor.runOrPrompt('setShapeImage') },
          { label: 'From Stock Images...', disabled: true, reason: online, run: noop },
          { label: 'From Online Sources...', disabled: true, reason: online, run: noop },
          { label: 'From Brand Images...', disabled: true, reason: online, run: noop },
          { label: 'From Icons...', disabled: true, reason: online, run: noop },
          { label: 'From Clipboard...', disabled: true, reason: 'The browser does not expose copied pictures to the editor.', run: noop },
        ],
      },
      ...reorderItems(),
      ...arrangeItems(),
      lockItem(shapes),
      { label: 'Hyperlink...', accel: '⌘K', sep: true, run: () => editor.runOrPrompt('setShapeHyperlink'), disabled: !editor.canRun('setShapeHyperlink') },
      { label: 'Edit Picture', disabled: true, reason: 'The editor has no picture editor yet.', run: noop },
      { label: 'Save as Picture...', sep: true, disabled: true, reason: 'The editor cannot export objects as pictures yet.', run: noop },
      { label: 'View Alt Text...', run: () => editor.runOrPrompt('setShapeDescription'), disabled: !editor.canRun('setShapeDescription') },
      { label: 'Crop', accel: '⇧C', run: () => editor.runOrPrompt('setShapeImageCrop'), disabled: !editor.canRun('setShapeImageCrop') },
      { label: 'Auto Crop', disabled: true, reason: 'Auto Crop needs the Microsoft image service.', run: noop },
      { label: 'Size and Position...', run: () => editor.showShapeFormat('size') },
      { label: 'Format Picture...', accel: '⇧⌘1', sep: true, run: () => editor.showShapeFormat() },
      { label: 'Action Settings...', sep: true, run: () => editor.runOrPrompt('setShapeClickAction'), disabled: !editor.canRun('setShapeClickAction') },
      newComment(),
    ];
  }

  // Mac PowerPoint's object menu.
  function shapeItems(shapes: readonly SlideShapeData[]): Item[] {
    const single = shapes.length === 1 ? shapes[0]! : null;
    const list: Item[] = [...clipboardItems(undefined)];
    if (single && getShapeKind(single) === 'shape') {
      list.push(
        { label: 'Edit Text', run: () => editor.editSelectedText() },
        { label: 'Edit Points', sep: true, disabled: true, reason: 'The editor has no point editor yet.', run: noop },
      );
    }
    list.push(
      ...reorderItems(),
      ...arrangeItems(),
      lockItem(shapes),
      { label: 'Hyperlink...', accel: '⌘K', sep: true, run: () => editor.runOrPrompt('setShapeHyperlink'), disabled: !editor.canRun('setShapeHyperlink') },
      { label: 'Save as Picture...', sep: true, disabled: true, reason: 'The editor cannot export objects as pictures yet.', run: noop },
      { label: 'Translate...', sep: true, disabled: true, reason: 'Translation needs the Microsoft translation service.', run: noop },
      { label: 'View Alt Text...', run: () => editor.runOrPrompt('setShapeDescription'), disabled: !editor.canRun('setShapeDescription') },
      { label: 'Set as Default Shape Style', disabled: true, reason: 'The editor does not keep a default shape style yet.', run: noop },
    );
    if (shapes.length) {
      list.push(
        { label: 'Size and Position...', run: () => editor.showShapeFormat('size') },
        { label: formatLabel(shapes), accel: '⇧⌘1', run: () => editor.showShapeFormat() },
      );
    }
    list[list.length - 1]!.sep = true;
    list.push(
      { label: 'Action Settings...', sep: true, run: () => editor.runOrPrompt('setShapeClickAction'), disabled: !editor.canRun('setShapeClickAction') },
      newComment(),
    );
    return list;
  }

  // Mac PowerPoint's menu for the slide itself (right-clicking its background).
  // It has no Layout or Reset Slide; those are on the Home tab.
  function backgroundItems(): Item[] {
    const slide = doc.currentSlide;
    const hidden = slide !== null && isSlideHidden(slide);
    const guides = editor.guidesVisible();
    const snapping = getSnapToGrid(doc.pres) ?? false;
    const saveView = (update: Partial<{ grid: boolean; smart: boolean; drawing: boolean }>) =>
      editor.view.save({ grid: editor.view.grid, smart: editor.view.smart, drawing: guides, ...update });
    return [
      { label: 'Cut', accel: '⌘X', disabled: true, run: noop },
      { label: 'Copy', accel: '⌘C', disabled: true, run: noop },
      { label: 'Paste', accel: '⌘V', run: () => editor.paste(), disabled: !editor.hasClipboard() },
      { label: 'Paste Special...', accel: '⌃⌘V', sep: true, disabled: true, reason: 'Paste Special needs the system clipboard formats, which the browser does not expose.', run: noop },
      { label: 'New Slide', accel: '⇧⌘N', run: () => editor.addNewSlide() },
      { label: 'Duplicate Slide', accel: '⇧⌘D', disabled: !slide, run: () => editor.invoke('duplicateSlide') },
      { label: 'Delete Slide', sep: true, disabled: !slide, run: () => editor.invoke('removeSlide') },
      { label: 'Hide Slide', checked: hidden, sep: true, disabled: !slide, run: () => doc.transact(t('Hide Slide'), () => setSlideHidden(slide!, !hidden)) },
      { label: 'Ruler', checked: editor.view.ruler, run: () => editor.view.save({ ruler: !editor.view.ruler }) },
      {
        label: 'Grid and Guides',
        children: [
          { label: 'Add Vertical Guide', run: () => editor.addDrawingGuide('x') },
          { label: 'Add Horizontal Guide', run: () => editor.addDrawingGuide('y') },
          { label: 'Delete', sep: true, disabled: true, reason: 'Right-click a guide to delete it.', run: noop },
          { label: 'Smart Guides', checked: editor.view.smart, run: () => saveView({ smart: !editor.view.smart }) },
          { label: 'Guides', accel: '⌃⌥⌘G', checked: guides, run: () => saveView({ drawing: !guides }) },
          { label: 'Gridlines', accel: "⌘'", checked: editor.view.grid, sep: true, run: () => saveView({ grid: !editor.view.grid }) },
          { label: 'Snap to Grid', checked: snapping, sep: true, run: () => doc.transact(t('Snap to Grid'), () => setSnapToGrid(doc.pres, !snapping)) },
          { label: 'Grid Options...', run: () => editor.activeDialog = 'gridOptions' },
        ],
      },
      { label: 'Zoom...', sep: true, run: () => editor.activeDialog = 'zoom' },
      { label: 'Format Background...', sep: true, run: () => editor.showBackgroundFormat() },
      { label: 'Slide Show', accel: '⇧⌘↩', sep: true, disabled: !editor.canPresent, run: () => editor.present('current') },
      { ...newComment(), disabled: !slide },
    ];
  }

  const items = $derived.by<Item[]>(() => {
    doc.version;
    if (menu.outlineText) {
      const actions = menu.outlineText;
      return [
        { label: 'Cut', accel: '⌘X', run: actions.cut, disabled: !actions.hasTextSelection },
        { label: 'Copy', accel: '⌘C', run: actions.copy, disabled: !actions.hasTextSelection },
        { label: 'Paste', accel: '⌘V', run: actions.paste, sep: true },
        ...slideItems(),
        ...outlineCollapseItems(),
        { label: 'Promote', run: actions.promote, disabled: !actions.canPromote },
        { label: 'Demote', run: actions.demote, disabled: !actions.canDemote },
        { label: 'Move Up', run: actions.moveUp, disabled: !actions.canMoveUp },
        { label: 'Move Down', run: actions.moveDown, disabled: !actions.canMoveDown },
        { label: 'Show Formatting', checked: editor.outlineShowFormatting, run: () => editor.outlineShowFormatting = !editor.outlineShowFormatting },
      ];
    }
    if (menu.text && !menu.text.cell) return textItems(menu.text);
    if (doc.selection.kind === 'cell') return cellItems(menu.text);
    if (doc.selection.kind === 'shape') {
      const shapes = editor.selectedShapes();
      return shapes.length === 1 && getShapeKind(shapes[0]!) === 'picture' && !getShapeMedia(shapes[0]!)
        ? pictureItems(shapes)
        : shapeItems(shapes);
    }
    if (doc.selection.kind === 'slide') {
      if (menu.source === 'outline') {
        return [
          { label: 'Cut', accel: '⌘X', run: () => editor.cutSelection() },
          { label: 'Copy', accel: '⌘C', run: () => editor.copySelection() },
          { label: 'Paste', accel: '⌘V', run: () => editor.paste(), disabled: !editor.hasClipboard() },
          ...slideItems(),
          ...outlineCollapseItems(),
          { label: 'Move Up', run: () => editor.invoke('moveSlide', { toIndex: firstSelected - 1 }), disabled: firstSelected === 0 },
          { label: 'Move Down', run: () => editor.invoke('moveSlide', { toIndex: firstSelected + 1 }), disabled: firstSelected >= doc.slides.length - selected.length },
          { label: 'Show Formatting', checked: editor.outlineShowFormatting, run: () => editor.outlineShowFormatting = !editor.outlineShowFormatting },
        ];
      }
      // Mac PowerPoint's thumbnail menu. It has no Layout or Reset Slide;
      // those are on the Home tab.
      const slides = selected.map((index) => doc.slideAt(index)).filter((slide) => slide !== null);
      const hidden = slides.length > 0 && slides.every((slide) => isSlideHidden(slide));
      const anchor = doc.selection.slideIndex;
      return [
        { label: 'Cut', accel: '⌘X', run: () => editor.cutSelection() },
        { label: 'Copy', accel: '⌘C', run: () => editor.copySelection() },
        { label: 'Paste', accel: '⌘V', run: () => editor.paste(), disabled: !editor.hasClipboard() },
        { label: 'Select All', accel: '⌘A', sep: true, run: () => doc.select({ kind: 'slide', slideIndex: anchor, slideIndices: doc.slides.map((_, index) => index), anchorIndex: anchor }) },
        { label: 'New Slide', accel: '⇧⌘N', run: () => editor.addNewSlide() },
        { label: 'Duplicate Slide', accel: '⌘D', run: () => editor.invoke('duplicateSlide') },
        { label: 'Delete Slide', run: () => editor.invoke('removeSlide') },
        { label: 'Add Section', sep: true, run: () => {
          const start = firstSelected;
          doc.transact(t('Add Section'), () => addSection(doc.pres, start, t(UNTITLED_SECTION)));
        } },
        { label: 'Format Background...', sep: true, run: () => editor.showBackgroundFormat() },
        { label: 'Hide Slide', checked: hidden, sep: true, run: () => doc.transact(t('Hide Slide'), () => { for (const slide of slides) setSlideHidden(slide, !hidden); }) },
        { label: 'Zoom...', run: () => editor.activeDialog = 'zoom' },
        { label: 'Slide Show', sep: true, disabled: !editor.canPresent, run: () => editor.present('current') },
        newComment(),
      ];
    }
    if (editor.viewMode === 'sorter' || menu.source === 'outline') {
      const list: Item[] = [
        { label: 'Paste', accel: '⌘V', run: () => editor.paste(), disabled: !editor.hasClipboard() },
        { label: 'Select all', accel: '⌘A', run: () => editor.selectAllShapes() },
      ];
      if (menu.source === 'outline') list.push({ label: 'Show Formatting', checked: editor.outlineShowFormatting, run: () => editor.outlineShowFormatting = !editor.outlineShowFormatting });
      return list;
    }
    return backgroundItems();
  });

  let origin: HTMLElement | null = null;
  let menuNode: HTMLDivElement | null = null;
  function dismiss() {
    editor.closeContextMenu();
    if (origin?.isConnected) origin.focus({ preventScroll: true });
  }
  function onKeydown(event: KeyboardEvent) {
    event.stopPropagation();
    if (event.key === 'Escape') { event.preventDefault(); dismiss(); return; }
    if (event.key === 'Tab') { dismiss(); return; }
    const target = event.target as HTMLElement;
    if (event.key === 'ArrowRight' && target.dataset.submenu) {
      event.preventDefault(); submenu = target.dataset.submenu;
      void tick().then(() => menuNode?.querySelector<HTMLButtonElement>('.submenu button:not(:disabled)')?.focus());
      return;
    }
    if (event.key === 'ArrowLeft' && target.closest('.submenu')) {
      event.preventDefault();
      target.closest('.branch')?.querySelector<HTMLButtonElement>(':scope > button')?.focus();
      submenu = null; return;
    }
    if (!menuNode || !['ArrowDown', 'ArrowUp', 'Home', 'End'].includes(event.key)) return;
    event.preventDefault();
    const activeMenu = target.closest('[role="menu"]') ?? menuNode;
    const buttons = [...activeMenu.querySelectorAll<HTMLButtonElement>(':scope > button:not(:disabled), :scope > .branch > button:not(:disabled), :scope > .gallery > button:not(:disabled)')];
    if (!buttons.length) return;
    const current = buttons.findIndex(button => button === document.activeElement);
    const next = event.key === 'Home' ? 0 : event.key === 'End' ? buttons.length - 1 :
      (current + (event.key === 'ArrowDown' ? 1 : -1) + buttons.length) % buttons.length;
    buttons[next]?.focus();
  }

  function place(node: HTMLDivElement, position: { x: number; y: number }) {
    origin = document.activeElement instanceof HTMLElement ? document.activeElement : null;
    menuNode = node;
    // The text editor keeps focus (and so its caret and selection) while its menu is open.
    if (!editor.contextMenu?.text) node.querySelector<HTMLButtonElement>('button:not(:disabled)')?.focus({ preventScroll: true });
    let current = position;
    const update = () => {
      node.style.left = `${Math.max(8, Math.min(current.x, window.innerWidth - node.offsetWidth - 8))}px`;
      node.style.top = `${Math.max(8, Math.min(current.y, window.innerHeight - node.offsetHeight - 8))}px`;
    };
    update();
    window.addEventListener('resize', update);
    return {
      update(position: { x: number; y: number }) { current = position; update(); },
      destroy() { window.removeEventListener('resize', update); },
    };
  }

  function placeSubmenu(node: HTMLDivElement) {
    const parent = node.parentElement!.getBoundingClientRect();
    const position = () => {
      const margin = 8;
      node.style.left = `${Math.max(margin, parent.right + node.offsetWidth < window.innerWidth - margin ? parent.right : parent.left - node.offsetWidth)}px`;
      node.style.top = `${Math.max(margin, Math.min(parent.top, window.innerHeight - node.offsetHeight - margin))}px`;
    };
    position();
    window.addEventListener('resize', dismiss);
    return { destroy() { window.removeEventListener('resize', dismiss); } };
  }

  // With focus left in the text editor, Escape closes the menu instead of
  // ending the edit; other keys close it and reach the text.
  function onTextKeydown(event: KeyboardEvent) {
    if (!menu.text) return;
    if (event.key === 'Escape') {
      event.preventDefault();
      event.stopPropagation();
    }
    editor.closeContextMenu();
  }

  function activate(item: Item) {
    if (item.disabled) return;
    if (item.children) {
      submenu = item.label;
      // Keyboard users land in the submenu; the text menu never takes focus.
      if (!menu.text) void tick().then(() => menuNode?.querySelector<HTMLButtonElement>('.submenu button:not(:disabled)')?.focus());
      return;
    }
    dismiss();
    item.run();
  }
</script>

<svelte:window
  onkeydowncapture={onTextKeydown}
  onpointerdown={() => editor.closeContextMenu()}
  onblur={() => editor.closeContextMenu()}
/>

<div
  class="ctx ok-editor"
  use:place={menu}
  role="menu"
  tabindex="-1"
  onkeydown={onKeydown}
  onpointerdown={(e) => e.stopPropagation()}
  onmousedown={(e) => { if (menu.text) e.preventDefault(); }}
  oncontextmenu={(e) => e.preventDefault()}
>
  {#each items as item, index (item.label)}
    {#if item.children}
      {@const tiles = item.children.filter((child) => child.tile)}
      {@const rows = item.children.filter((child) => !child.tile)}
      <div class="branch">
        <button class="ctx-item" role="menuitem" tabindex="-1" aria-label={t(item.label)} aria-haspopup="menu" aria-expanded={submenu === item.label} data-submenu={item.label} disabled={item.disabled} onclick={() => activate(item)} onpointerenter={() => submenu = item.disabled ? null : item.label}>
          <span>{t(item.label)}</span><span>›</span>
        </button>
        {#if submenu === item.label}
          <div class="ctx submenu" role="menu" aria-label={t(item.label)} use:placeSubmenu>
            {#if tiles.length}
              <div class="gallery">
                {#each tiles as tile (tile.label)}
                  <button class="tile" class:checked={tile.checked} role="menuitemradio" aria-checked={tile.checked} aria-label={t(tile.label)} title={t(tile.label)} tabindex="-1" onclick={() => activate(tile)}>
                    {#if tile.tile!.length}
                      {#each tile.tile! as marker, line (line)}<span class="tile-line"><span>{marker}</span><span class="rule"></span></span>{/each}
                    {:else}{t(tile.label)}{/if}
                  </button>
                {/each}
              </div>
              <div class="ctx-sep" role="separator"></div>
            {/if}
            {#each rows as child (child.label)}
              <button class="ctx-item" role={child.checked === undefined ? 'menuitem' : 'menuitemcheckbox'} aria-checked={child.checked} aria-label={t(child.label)} title={child.disabled && child.reason ? t(child.reason) : undefined} tabindex="-1" disabled={child.disabled} onclick={() => activate(child)}>
                <span>{child.checked ? '✓ ' : ''}{t(child.label)}</span>
                {#if child.accel}<span class="accel">{child.accel}</span>{/if}
              </button>
              {#if child.sep}<div class="ctx-sep" role="separator"></div>{/if}
            {/each}
          </div>
        {/if}
      </div>
    {:else}
      <button class="ctx-item" role={item.checked === undefined ? "menuitem" : "menuitemcheckbox"} aria-checked={item.checked} aria-label={t(item.label)} title={item.disabled && item.reason ? t(item.reason) : undefined} tabindex="-1" disabled={item.disabled} onclick={() => activate(item)} onpointerenter={() => submenu = null}>
        <span>{item.checked ? "✓ " : ""}{t(item.label)}</span>
        {#if item.accel}<span class="accel">{item.accel}</span>{/if}
      </button>
    {/if}
    {#if item.sep && index < items.length - 1}<div class="ctx-sep" role="separator"></div>{/if}
  {/each}
</div>

<style>
  .ctx {
    position: fixed;
    z-index: 400;
    min-width: min(200px, calc(100vw - 16px));
    max-width: calc(100vw - 16px);
    max-height: calc(100dvh - 16px);
    box-sizing: border-box;
    overflow: auto;
    background: var(--ok-panel);
    border: 1px solid var(--ok-border);
    border-radius: var(--ok-radius-lg);
    box-shadow: var(--ok-shadow-lg);
    padding: 5px;
  }
  /* Mac menus: 24 pt items and 11 pt separators. */
  .ctx-item {
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 24px;
    width: 100%;
    height: 24px;
    text-align: left;
    padding: 0 10px;
    border: none;
    background: none;
    border-radius: var(--ok-radius);
    font: inherit;
    font-size: 13px;
    color: var(--ok-text);
    cursor: pointer;
  }
  .ctx-item:hover:not(:disabled), .ctx-item:focus-visible {
    background: var(--ok-selected);
  }
  .ctx-item:disabled {
    color: var(--ok-text-3);
    cursor: default;
  }
  .ctx-sep {
    height: 1px;
    margin: 5px 10px;
    background: var(--ok-border);
  }
  /* PowerPoint's list galleries: five 90 pt tiles per row. */
  .gallery {
    display: grid;
    grid-template-columns: repeat(5, 90px);
    grid-auto-rows: 90px;
  }
  .tile {
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    gap: 6px;
    border: 1px solid transparent;
    border-radius: var(--ok-radius);
    background: none;
    font: inherit;
    font-size: 13px;
    color: var(--ok-text);
    cursor: pointer;
  }
  .tile:hover, .tile:focus-visible, .tile.checked {
    border-color: var(--ok-accent);
    background: var(--ok-selected);
  }
  .tile-line { display: flex; align-items: center; gap: 6px; width: 54px; }
  .tile-line > span:first-child { min-width: 18px; text-align: right; font-size: 11px; }
  .rule { flex: 1; height: 1px; background: var(--ok-text-3); }
  .accel {
    font-size: 11px;
    color: var(--ok-text-3);
  }
</style>
