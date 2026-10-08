<script lang="ts">
  import { eventTarget } from '../core/dom-root.ts';
  // The reference desktop app's (Mac) menu bar, in the title bar because a web page cannot own
  // the macOS one: File … Help with the reference desktop app's order, separators, submenus,
  // shortcut glyphs and per-selection enabled and checked states (native
  // captures, 2026-10-07; see core/menubar-native.ts). The application menu
  // is left out: its items (Settings, Hide, Quit …) belong to the browser.
  import { tick } from 'svelte';
  import { getEditor } from '../core/context.ts';
  import type { NativeMenuEntry, NativeMenuItem } from '../core/menubar-native.ts';
  import { menuItemForKey } from '../core/menubar-shortcuts.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import { menuCommand, nativeMenus, runMenuKey, type MenuHost } from './menubar-commands.ts';
  import TextColumnsDialog from './TextColumnsDialog.svelte';

  let { file }: { file: Pick<MenuHost, 'save' | 'download' | 'open' | 'newPresentation'> } = $props();
  const editor = getEditor();
  const menus = $derived(nativeMenus(editor));
  let openMenu = $state<string | null>(null);
  /** Open submenus by depth, outermost first. */
  let path = $state<string[]>([]);
  let fullScreen = $state(false);
  let columnsOpen = $state(false);
  let root: HTMLElement;

  const host: MenuHost = {
    save: () => file.save(),
    download: () => file.download(),
    open: () => file.open(),
    newPresentation: () => file.newPresentation(),
    columns: () => (columnsOpen = true),
    // WordArt and Symbol, the Insert items that open a gallery, hang it under Insert.
    anchor: () => root.querySelector<HTMLElement>('[data-menu="insert"]'),
    get fullScreen() { return fullScreen; },
    toggleFullScreen: () => void toggleFullScreen(),
  };
  const command = (id: string) => menuCommand(editor, host, id);
  // While text is being edited the menus leave focus (caret and selection) in
  // the text, as the reference desktop app's do.
  const keepFocus = $derived(!!editor.inlineTextFormat);

  async function toggleFullScreen() {
    try {
      if (document.fullscreenElement) await document.exitFullscreen();
      else await document.documentElement.requestFullscreen();
    } catch (error) {
      editor.toast('error', String(error));
    }
  }

  function close(restore = true) {
    const id = openMenu;
    openMenu = null;
    path = [];
    if (restore && id && !keepFocus) root.querySelector<HTMLElement>(`[data-menu="${id}"]`)?.focus();
  }

  async function show(id: string, focus: 'first' | 'last' | null = keepFocus ? null : 'first') {
    openMenu = id;
    path = [];
    if (!focus) return;
    await tick();
    const items = rows(root.querySelector('.menu[data-depth="0"]'));
    (focus === 'first' ? items[0] : items.at(-1))?.focus();
  }

  function rows(menu: Element | null): HTMLButtonElement[] {
    return [...(menu?.querySelectorAll<HTMLButtonElement>(':scope > .item:not(:disabled), :scope > .branch > .item:not(:disabled)') ?? [])];
  }

  function choose(item: NativeMenuItem) {
    const action = command(item.id);
    if (action.disabled || !action.run) return;
    close();
    action.run();
  }

  function enter(item: NativeMenuItem, depth: number) {
    path = item.children ? [...path.slice(0, depth), item.id] : path.slice(0, depth);
  }

  async function openSubmenu(item: NativeMenuItem, depth: number) {
    path = [...path.slice(0, depth), item.id];
    await tick();
    rows(root.querySelector(`.menu[data-depth="${depth + 1}"]`))[0]?.focus();
  }

  function step(delta: number) {
    const index = menus.findIndex((menu) => menu.id === openMenu);
    void show(menus[(index + delta + menus.length) % menus.length]!.id);
  }

  function onMenuKeydown(event: KeyboardEvent, depth: number) {
    // Shortcuts close the menu and run, as they do over a native menu.
    if (menuItemForKey(menus, event)) return;
    event.stopPropagation();
    const target = event.target as HTMLElement;
    const menu = target.closest('.menu');
    const items = rows(menu);
    const index = items.indexOf(target as HTMLButtonElement);
    if (event.key === 'Escape') {
      event.preventDefault();
      if (depth > 0) {
        const parent = path[depth - 1];
        path = path.slice(0, depth - 1);
        root.querySelector<HTMLElement>(`[data-item="${parent}"]`)?.focus();
      } else close();
    } else if (event.key === 'Tab') close(false);
    else if (event.key === 'ArrowDown' || event.key === 'ArrowUp' || event.key === 'Home' || event.key === 'End') {
      event.preventDefault();
      const next = event.key === 'Home' ? 0 : event.key === 'End' ? items.length - 1
        : (index + (event.key === 'ArrowUp' ? -1 : 1) + items.length) % items.length;
      items[next]?.focus();
    } else if (event.key === 'ArrowRight') {
      event.preventDefault();
      const branch = target.dataset.branch;
      const item = branch ? find(branch) : undefined;
      if (item && !target.hasAttribute('disabled')) void openSubmenu(item, depth);
      else step(1);
    } else if (event.key === 'ArrowLeft') {
      event.preventDefault();
      if (depth > 0) {
        const parent = path[depth - 1];
        path = path.slice(0, depth - 1);
        root.querySelector<HTMLElement>(`[data-item="${parent}"]`)?.focus();
      } else step(-1);
    }
  }

  function onTitleKeydown(event: KeyboardEvent, id: string) {
    if (event.key === 'ArrowDown' || event.key === 'ArrowUp') {
      event.preventDefault();
      void show(id, event.key === 'ArrowDown' ? 'first' : 'last');
    }
  }

  function find(id: string, entries: readonly NativeMenuEntry[] = menus.flatMap((menu) => menu.items)): NativeMenuItem | undefined {
    for (const entry of entries) {
      if (entry === '-') continue;
      if (entry.id === id) return entry;
      const nested = entry.children && find(id, entry.children);
      if (nested) return nested;
    }
    return undefined;
  }

  /** Fixed placement: under the title, or beside the item for a submenu, kept in the window. */
  function place(node: HTMLElement, depth: number) {
    const margin = 8;
    function position() {
      const anchor = (depth === 0 ? node.previousElementSibling : node.parentElement)!.getBoundingClientRect();
      let left = depth === 0 ? anchor.left : anchor.right - 4;
      if (depth > 0 && left + node.offsetWidth > innerWidth - margin) left = anchor.left - node.offsetWidth + 4;
      node.style.left = `${Math.max(margin, Math.min(left, innerWidth - node.offsetWidth - margin))}px`;
      const top = depth === 0 ? anchor.bottom : anchor.top - 5;
      node.style.top = `${Math.max(margin, Math.min(top, innerHeight - node.offsetHeight - margin))}px`;
    }
    position();
    window.addEventListener('resize', position);
    return { destroy() { window.removeEventListener('resize', position); } };
  }
</script>

<svelte:window
  onkeydown={(event) => { if (runMenuKey(editor, host, event)) close(false); }}
  onpointerdown={(event) => { if (openMenu && !root.contains(eventTarget(event) as Node)) close(false); }}
  onblur={() => { if (openMenu) close(false); }}
  onfullscreenchange={() => (fullScreen = !!document.fullscreenElement)}
/>

{#snippet list(entries: readonly NativeMenuEntry[], depth: number, label: string)}
  <div class="menu" role="menu" aria-label={label} data-depth={depth} tabindex="-1" use:place={depth} onkeydown={(event) => onMenuKeydown(event, depth)}>
    {#each entries as entry, index (index)}
      {#if entry === '-'}
        <div class="sep" role="separator"></div>
      {:else}
        {@const action = command(entry.id)}
        {@const name = action.label ?? entry.label}
        {#if entry.children}
          <div class="branch">
            <button
              class="item"
              role="menuitem"
              tabindex="-1"
              aria-label={name}
              aria-haspopup="menu"
              aria-expanded={path[depth] === entry.id}
              data-item={entry.id}
              data-branch={entry.id}
              disabled={action.disabled || entry.children.length === 0}
              title={action.disabled && action.reason ? t(action.reason) : undefined}
              onmousedown={(event) => { if (keepFocus) event.preventDefault(); }}
              onpointerenter={() => enter(entry, depth)}
              onclick={() => enter(entry, depth)}
            ><span class="mark"></span><span class="label">{name}</span><span class="arrow" aria-hidden="true">›</span></button>
            {#if path[depth] === entry.id && entry.children.length}
              {@render list(entry.children, depth + 1, name)}
            {/if}
          </div>
        {:else}
          <button
            class="item"
            role={action.checked === undefined ? 'menuitem' : action.radio ? 'menuitemradio' : 'menuitemcheckbox'}
            aria-checked={action.checked}
            tabindex="-1"
            aria-label={name}
            data-item={entry.id}
            disabled={action.disabled || !action.run}
            title={action.disabled && action.reason ? t(action.reason) : undefined}
            onmousedown={(event) => { if (keepFocus) event.preventDefault(); }}
            onpointerenter={() => enter(entry, depth)}
            onclick={() => choose(entry)}
          ><span class="mark" aria-hidden="true">{action.checked ? (action.bullet ? '•' : '✓') : ''}</span><span class="label">{name}</span>{#if entry.shortcut}<kbd>{entry.shortcut}</kbd>{/if}</button>
        {/if}
      {/if}
    {/each}
  </div>
{/snippet}

<div class="menubar" role="menubar" aria-label={t('Menu bar')} bind:this={root}>
  {#each menus as menu (menu.id)}
    <div class="top">
      <button
        class="title"
        role="menuitem"
        data-menu={menu.id}
        aria-haspopup="menu"
        aria-expanded={openMenu === menu.id}
        onmousedown={(event) => { if (keepFocus) event.preventDefault(); }}
        onpointerenter={() => { if (openMenu && openMenu !== menu.id) void show(menu.id, null); }}
        onclick={() => (openMenu === menu.id ? close(false) : void show(menu.id))}
        onkeydown={(event) => onTitleKeydown(event, menu.id)}
      >{menu.label}</button>
      {#if openMenu === menu.id}
        {@render list(menu.items, 0, menu.label)}
      {/if}
    </div>
  {/each}
</div>
{#if columnsOpen}<TextColumnsDialog onclose={() => (columnsOpen = false)} />{/if}

<style>
  .menubar { display: flex; align-items: center; gap: 0; flex: none; }
  .top { position: relative; }
  .title { padding: 3px 8px; border: 0; border-radius: 4px; background: transparent; color: var(--ok-text); font: inherit; font-size: 13px; white-space: nowrap; cursor: default; }
  .title:hover, .title[aria-expanded='true'] { background: var(--ok-hover); }
  .title:focus-visible { outline: 2px solid var(--ok-accent); outline-offset: -2px; }
  /* Mac menus: 24 pt items and 11 pt separators, 5 pt inset. */
  .menu { position: fixed; z-index: 450; min-width: 220px; max-height: calc(100dvh - 16px); box-sizing: border-box; overflow-y: auto; padding: 5px; border: 1px solid var(--ok-border); border-radius: 6px; background: var(--ok-panel); color: var(--ok-text); box-shadow: var(--ok-shadow-lg); }
  .branch { position: relative; }
  .item { display: flex; align-items: center; width: 100%; height: 24px; padding: 0 10px 0 4px; border: 0; border-radius: 4px; background: transparent; color: inherit; font: inherit; font-size: 13px; text-align: left; white-space: nowrap; cursor: default; }
  .item:hover:not(:disabled), .item:focus-visible, .item[aria-expanded='true'] { background: var(--ok-accent); color: white; outline: none; }
  .item:disabled { color: var(--ok-text-3); }
  .mark { width: 16px; flex: none; text-align: center; }
  .label { flex: 1; }
  kbd, .arrow { margin-left: 24px; font: inherit; opacity: 0.7; }
  .sep { height: 1px; margin: 5px 10px; background: var(--ok-border); }
</style>
