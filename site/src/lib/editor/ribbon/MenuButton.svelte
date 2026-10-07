<script lang="ts">
  // A contextual-tab command with a ▾ menu, in one of PowerPoint's four sizes:
  // `big` (icon over caption), `row` (22 pt, icon and label), `icon` (36 × 22,
  // icon only) and `tool` (38 × 26, icon only). Choosing a menu item closes it.
  import { tick, type Snippet } from 'svelte';
  import Icon from '../ui/Icon.svelte';
  import { caption } from './caption.ts';
  import { placeBelowTrigger } from './place-menu.ts';

  let {
    label,
    icon,
    look,
    disabled = false,
    title,
    pressed,
    children,
  }: {
    /** Already translated. */
    label: string;
    icon: string;
    look: 'big' | 'row' | 'icon' | 'tool';
    disabled?: boolean;
    /** Tooltip; defaults to the label. */
    title?: string;
    pressed?: boolean;
    children: Snippet;
  } = $props();

  let open = $state(false);
  let trigger = $state<HTMLButtonElement>();
  let menu = $state<HTMLDivElement>();

  function close(restore = true) {
    open = false;
    if (restore) trigger?.focus();
  }
  async function toggle() {
    if (open) { close(); return; }
    open = true;
    await tick();
    menu?.querySelector<HTMLButtonElement>('button:not(:disabled)')?.focus();
  }
  function chosen(event: MouseEvent) {
    const item = (event.target as Element).closest('button');
    if (item && !item.disabled && !item.hasAttribute('aria-haspopup')) close(false);
  }
  function keys(event: KeyboardEvent) {
    event.stopPropagation();
    if (event.key === 'Escape') { event.preventDefault(); close(); return; }
    if (event.key === 'Tab') { close(false); return; }
    if (!['ArrowUp', 'ArrowDown', 'Home', 'End'].includes(event.key)) return;
    event.preventDefault();
    const items = [...menu!.querySelectorAll<HTMLButtonElement>('button:not(:disabled)')];
    const index = items.indexOf(event.target as HTMLButtonElement);
    const next = event.key === 'Home' ? 0 : event.key === 'End' ? items.length - 1 : (index + (event.key === 'ArrowDown' ? 1 : -1) + items.length) % items.length;
    items[next]?.focus();
  }
</script>

<svelte:window
  onpointerdown={(event) => { if (open && !menu?.contains(event.target as Node) && !trigger?.contains(event.target as Node)) close(false); }}
  onresize={() => { if (open) close(false); }}
/>

<button
  bind:this={trigger}
  class="ctx-{look}"
  aria-label={label}
  title={title ?? label}
  aria-haspopup="menu"
  aria-expanded={open}
  aria-pressed={pressed}
  {disabled}
  onclick={toggle}
>
  {#if look === 'big'}
    <span class="ctx-icon-row"><Icon name={icon} size={32} /><span class="ctx-caret" aria-hidden="true">▾</span></span><span class="ctx-caption">{caption(label)}</span>
  {:else if look === 'row'}
    <Icon name={icon} size={16} /><span>{label}</span><span class="ctx-caret" aria-hidden="true">▾</span>
  {:else}
    <Icon name={icon} size={18} /><span class="ctx-caret" aria-hidden="true">▾</span>
  {/if}
</button>
{#if open}
  <!-- svelte-ignore a11y_click_events_have_key_events -->
  <div class="ctx-menu" role="menu" tabindex="-1" aria-label={label} bind:this={menu} use:placeBelowTrigger onkeydown={keys} onclick={chosen}>
    {@render children()}
  </div>
{/if}
