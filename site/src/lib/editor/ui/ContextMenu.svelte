<script lang="ts">
  import { tick } from 'svelte';
  import { getShapeKind, getSlideLayout, getSlideLayoutName, getSlideLayoutPartName, getSlideLayouts, setSlideLayout, setSlideOutlineCollapsed } from '@office-kit/pptx';
  // Right-click menu. Items adapt to the current selection and dispatch through
  // the controller's actions (which go through the same undoable command path).
  import { getEditor } from '../core/context.ts';
  import { selectedSlideIndices } from '../core/selection.ts';
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
  } & ({ run: () => void; children?: never } | { children: Item[]; run?: never });
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
      { label: 'New slide', run: () => editor.addNewSlide() },
      { label: 'Duplicate slide', accel: '⌘D', run: () => editor.invoke('duplicateSlide') },
      { label: 'Delete slide', accel: 'Del', run: () => editor.invoke('removeSlide'), sep: true },
    ];
  }

  const hasShapes = $derived(doc.selection.kind === 'shape' || doc.selection.kind === 'cell');

  const items = $derived.by<Item[]>(() => {
    const list: Item[] = [];
    if (menu.outlineText) {
      const actions = menu.outlineText;
      list.push(
        { label: 'Cut', accel: '⌘X', run: actions.cut, disabled: !actions.hasTextSelection },
        { label: 'Copy', accel: '⌘C', run: actions.copy, disabled: !actions.hasTextSelection },
        { label: 'Paste', accel: '⌘V', run: actions.paste, sep: true },
        ...slideItems(),
        ...outlineCollapseItems(),
        { label: 'Promote', run: actions.promote, disabled: !actions.canPromote },
        { label: 'Demote', run: actions.demote, disabled: !actions.canDemote },
        { label: 'Move Up', run: actions.moveUp, disabled: !actions.canMoveUp },
        { label: 'Move Down', run: actions.moveDown, disabled: !actions.canMoveDown },
      );
    } else if (doc.selection.kind === 'cell') {
      list.push(
        { label: 'Cut', accel: '⌘X', run: () => editor.cutSelection() },
        { label: 'Copy', accel: '⌘C', run: () => editor.copySelection() },
        { label: 'Paste', accel: '⌘V', run: () => editor.paste(), disabled: !editor.hasClipboard() },
        { label: 'Clear cell text', accel: 'Del', run: () => editor.deleteSelection(), sep: true },
        { label: 'Select all cells', accel: '⌘A', run: () => editor.selectAll() },
        { label: 'Select table', run: () => {
          const selection = doc.selection;
          if (selection.kind === 'cell') doc.select({ kind: 'shape', slideIndex: selection.slideIndex, shapeIds: [selection.shapeId] });
        } },
      );
    } else if (hasShapes) {
      list.push(
        { label: 'Cut', accel: '⌘X', run: () => editor.cutSelection() },
        { label: 'Copy', accel: '⌘C', run: () => editor.copySelection() },
        { label: 'Paste', accel: '⌘V', run: () => editor.paste(), disabled: !editor.hasClipboard() },
      );
      // PowerPoint's object menu: Edit Text, Group / Bring to Front / Send to
      // Back as submenus, Link, Alt Text, Size and Position, Format, Comment.
      const shapes = editor.selectedShapes();
      const single = shapes.length === 1 ? shapes[0]! : null;
      if (single && getShapeKind(single) === 'shape') list.push({ label: 'Edit Text', sep: true, run: () => editor.editSelectedText() });
      list.push(
        {
          label: 'Group', sep: !single || getShapeKind(single) !== 'shape',
          children: [
            { label: 'Group', run: () => editor.invoke('groupShapes'), disabled: !editor.canRun('groupShapes') },
            { label: 'Ungroup', run: () => editor.invoke('ungroupShapes'), disabled: !editor.canRun('ungroupShapes') },
            { label: 'Regroup', run: () => editor.regroupSelection(), disabled: !editor.canRegroup() },
          ],
        },
        {
          label: 'Bring to Front',
          children: [
            { label: 'Bring to Front', run: () => editor.invoke('bringShapeToFront') },
            { label: 'Bring Forward', run: () => editor.invoke('bringShapeForward') },
          ],
        },
        {
          label: 'Send to Back',
          children: [
            { label: 'Send to Back', run: () => editor.invoke('sendShapeToBack') },
            { label: 'Send Backward', run: () => editor.invoke('sendShapeBackward') },
          ],
        },
        { label: 'Link...', sep: true, run: () => editor.runOrPrompt('setShapeHyperlink'), disabled: !editor.canRun('setShapeHyperlink') },
        { label: 'Edit Alt Text...', run: () => editor.runOrPrompt('setShapeDescription'), disabled: !editor.canRun('setShapeDescription') },
      );
      if (shapes.length && shapes.every((shape) => ['shape', 'connector', 'group'].includes(getShapeKind(shape)))) {
        list.push(
          { label: 'Size and Position...', sep: true, run: () => editor.showShapeFormat('size') },
          { label: 'Format Shape...', run: () => editor.showShapeFormat() },
        );
      }
      list.push({ label: 'New Comment', sep: true, run: () => editor.runOrPrompt('addSlideComment') });
    } else if (doc.selection.kind === 'slide') {
      list.push(
        { label: 'Cut', accel: '⌘X', run: () => editor.cutSelection() },
        { label: 'Copy', accel: '⌘C', run: () => editor.copySelection() },
        { label: 'Paste', accel: '⌘V', run: () => editor.paste(), disabled: !editor.hasClipboard() },
        ...slideItems(),
      );
      if (menu.source === 'outline') {
        list.push(...outlineCollapseItems());
      }
      list.push(
        { label: menu.source === 'outline' ? 'Move Up' : 'Move slide up', run: () => editor.invoke('moveSlide', { toIndex: firstSelected - 1 }), disabled: firstSelected === 0 },
        { label: menu.source === 'outline' ? 'Move Down' : 'Move slide down', run: () => editor.invoke('moveSlide', { toIndex: firstSelected + 1 }), disabled: firstSelected >= doc.slides.length - selected.length },
      );
    } else {
      list.push(
        { label: 'Paste', accel: '⌘V', run: () => editor.paste(), disabled: !editor.hasClipboard() },
        { label: 'Select all', accel: '⌘A', run: () => editor.selectAllShapes() },
      );
    }
    if (!hasShapes && editor.viewMode !== 'sorter' && menu.source !== 'outline') {
      // PowerPoint's menu for the slide itself.
      const slide = doc.currentSlide;
      const current = slide ? getSlideLayoutPartName(getSlideLayout(slide)!) : '';
      list.push(
        {
          label: 'Layout', sep: true,
          children: getSlideLayouts(doc.pres).map(layout => ({
            label: getSlideLayoutName(layout),
            checked: getSlideLayoutPartName(layout) === current,
            run: () => { if (slide) doc.transact(t('Slide layout'), () => setSlideLayout(slide, layout)); },
          })),
        },
        { label: 'Reset Slide', run: () => editor.invoke('resetSlideLayout'), disabled: !editor.canRun('resetSlideLayout') },
        {
          label: 'Grid and Guides', sep: true,
          children: [
            { label: 'Add Vertical Guide', run: () => editor.addDrawingGuide('x') },
            { label: 'Add Horizontal Guide', run: () => editor.addDrawingGuide('y') },
            { label: 'Grid Options...', run: () => editor.activeDialog = 'gridOptions' },
          ],
        },
        { label: 'Format Background...', sep: true, run: () => editor.showBackgroundFormat() },
        { label: 'New Comment', run: () => editor.runOrPrompt('addSlideComment'), disabled: !slide },
      );
    }
    if (menu.source === 'outline') list.push({
      label: 'Show Formatting', checked: editor.outlineShowFormatting,
      run: () => editor.outlineShowFormatting = !editor.outlineShowFormatting,
    });
    return list;
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
    const buttons = [...activeMenu.querySelectorAll<HTMLButtonElement>(':scope > button:not(:disabled), :scope > .branch > button:not(:disabled)')];
    if (!buttons.length) return;
    const current = buttons.findIndex(button => button === document.activeElement);
    const next = event.key === 'Home' ? 0 : event.key === 'End' ? buttons.length - 1 :
      (current + (event.key === 'ArrowDown' ? 1 : -1) + buttons.length) % buttons.length;
    buttons[next]?.focus();
  }

  function place(node: HTMLDivElement, position: { x: number; y: number }) {
    origin = document.activeElement instanceof HTMLElement ? document.activeElement : null;
    menuNode = node;
    node.querySelector<HTMLButtonElement>('button:not(:disabled)')?.focus({ preventScroll: true });
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

  function activate(item: Item) {
    if (item.disabled) return;
    if (item.children) { submenu = item.label; return; }
    dismiss();
    item.run();
  }
</script>

<svelte:window
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
  oncontextmenu={(e) => e.preventDefault()}
>
  {#each items as item (item.label)}
    {#if item.children}
      <div class="branch">
        <button class="ctx-item" class:sep={item.sep} role="menuitem" tabindex="-1" aria-label={t(item.label)} aria-haspopup="menu" aria-expanded={submenu === item.label} data-submenu={item.label} onclick={() => activate(item)} onpointerenter={() => submenu = item.label}>
          <span>{t(item.label)}</span><span>›</span>
        </button>
        {#if submenu === item.label}
          <div class="ctx submenu" role="menu" aria-label={t(item.label)} use:placeSubmenu>
            {#each item.children as child (child.label)}
              <button class="ctx-item" role={child.checked === undefined ? 'menuitem' : 'menuitemradio'} aria-checked={child.checked} tabindex="-1" disabled={child.disabled} onclick={() => activate(child)}>{child.checked ? '✓ ' : ''}{t(child.label)}</button>
            {/each}
          </div>
        {/if}
      </div>
    {:else}
      <button class="ctx-item" class:sep={item.sep} role={item.checked === undefined ? "menuitem" : "menuitemcheckbox"} aria-checked={item.checked} aria-label={item.checked === undefined ? undefined : t(item.label)} tabindex="-1" disabled={item.disabled} onclick={() => activate(item)} onpointerenter={() => submenu = null}>
        <span>{item.checked ? "✓ " : ""}{t(item.label)}</span>
        {#if item.accel}<span class="accel">{item.accel}</span>{/if}
      </button>
    {/if}
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
  .ctx-item {
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 24px;
    width: 100%;
    text-align: left;
    padding: 6px 10px;
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
  .ctx-item.sep {
    margin-bottom: 5px;
    padding-bottom: 9px;
    border-bottom: 1px solid var(--ok-border);
  }
  .accel {
    font-size: 11px;
    color: var(--ok-text-3);
  }
</style>
