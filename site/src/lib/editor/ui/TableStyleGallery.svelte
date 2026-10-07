<script lang="ts">
  // PowerPoint's expanded Table Styles gallery: every built-in style under
  // Best Match for Document, Light, Medium and Dark, seven to a row, then
  // Clear Table.
  import { onMount } from 'svelte';
  import { BUILTIN_TABLE_STYLES } from '@office-kit/pptx';
  import { TABLE_STYLE_GROUPS, TABLE_STYLES_PER_ROW, tableStyleName } from '../core/table-styles.ts';
  import { getLocale, t } from '../i18n/i18n.svelte.ts';

  let { images, current, choose, clear, close, anchor }: {
    /** Swatch per style GUID. */
    images: ReadonlyMap<string, string>;
    /** GUID of the table's style, marked in the gallery. */
    current: string | null;
    choose: (id: string) => void;
    clear: () => void;
    close: () => void;
    /** The button the gallery opens under. */
    anchor: HTMLElement;
  } = $props();

  let menu = $state<HTMLDivElement>();
  const groups = TABLE_STYLE_GROUPS.map((group) => ({
    ...group,
    styles: BUILTIN_TABLE_STYLES.filter((style) => style.category === group.category),
  }));

  function keys(event: KeyboardEvent) {
    event.stopPropagation();
    if (event.key === 'Escape') { event.preventDefault(); close(); return; }
    if (event.key === 'Tab') { close(); return; }
    const offsets: Record<string, number> = { ArrowLeft: -1, ArrowRight: 1, ArrowUp: -TABLE_STYLES_PER_ROW, ArrowDown: TABLE_STYLES_PER_ROW };
    if (!(event.key in offsets) && event.key !== 'Home' && event.key !== 'End') return;
    event.preventDefault();
    const items = [...menu!.querySelectorAll<HTMLButtonElement>('[role^="menuitem"]')];
    if (event.key === 'Home' || event.key === 'End') { (event.key === 'Home' ? items[0] : items.at(-1))?.focus(); return; }
    const index = items.indexOf(event.target as HTMLButtonElement) + offsets[event.key]!;
    (items[index] ?? (index > 0 ? items.at(-1) : items[0]))?.focus();
  }

  // Fixed, so the ribbon panel's overflow does not clip it; it opens under
  // the in-ribbon strip like PowerPoint's.
  function place(node: HTMLElement) {
    const bounds = anchor.getBoundingClientRect();
    node.style.left = `${Math.max(8, Math.min(bounds.left, innerWidth - node.offsetWidth - 8))}px`;
    node.style.top = `${Math.max(8, Math.min(bounds.top, innerHeight - node.offsetHeight - 8))}px`;
  }
  onMount(() => (menu?.querySelector<HTMLButtonElement>('[aria-checked="true"]') ?? menu?.querySelector<HTMLButtonElement>('[role^="menuitem"]'))?.focus());
</script>

<svelte:window onpointerdown={(event) => { if (!menu?.contains(event.target as Node) && !anchor.contains(event.target as Node)) close(); }} onresize={close} />

<div class="table-style-gallery" role="menu" aria-label={t('Table Styles')} tabindex="-1" bind:this={menu} use:place onkeydown={keys}>
  {#each groups as group (group.category)}
    <div class="heading">{getLocale() === 'ja' ? group.ja : group.en}</div>
    <div class="grid" role="group" aria-label={getLocale() === 'ja' ? group.ja : group.en}>
      {#each group.styles as style (style.id)}
        {@const name = tableStyleName(style.name, getLocale())}
        <button role="menuitemradio" aria-checked={current === style.id} aria-label={name} title={name} onclick={() => choose(style.id)}>
          {#if images.has(style.id)}<img src={images.get(style.id)} alt="" />{/if}
        </button>
      {/each}
    </div>
  {/each}
  <hr />
  <button role="menuitem" class="clear" onclick={clear}>{t('Clear Table')}</button>
</div>

<style>
  .table-style-gallery { position: fixed; z-index: 450; max-height: calc(100dvh - 16px); overflow-y: auto; padding: 4px 6px 6px; border: 1px solid var(--ok-border); border-radius: 6px; background: var(--ok-panel); color: var(--ok-text); box-shadow: var(--ok-shadow-lg); }
  .heading { padding: 6px 2px 4px; font-size: 12px; font-weight: 600; }
  .grid { display: grid; grid-template-columns: repeat(7, 74px); gap: 0; }
  .grid button { width: 74px; height: 56px; padding: 4px 6px; border: 1px solid transparent; border-radius: 2px; background: #fff; cursor: pointer; }
  .grid button:hover, .grid button:focus-visible { outline: 2px solid var(--ok-accent); outline-offset: -2px; }
  .grid button[aria-checked='true'] { border-color: var(--ok-selected-border); background: var(--ok-selected); }
  img { display: block; width: 100%; height: 100%; }
  hr { border: none; border-top: 1px solid var(--ok-border); margin: 6px 0 4px; }
  .clear { width: 100%; padding: 5px 8px; text-align: left; font: inherit; font-size: 12px; white-space: nowrap; color: var(--ok-text); background: none; border: 1px solid transparent; border-radius: 3px; cursor: pointer; }
  .clear:hover, .clear:focus-visible { background: var(--ok-hover); }
</style>
