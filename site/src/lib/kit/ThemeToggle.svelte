<script lang="ts">
  import { onMount } from 'svelte';

  const STORAGE_KEY = 'office-kit-theme';

  let dark = $state(false);

  onMount(() => {
    const stored = document.documentElement.dataset.theme;
    dark = stored ? stored === 'dark' : matchMedia('(prefers-color-scheme: dark)').matches;
  });

  function toggle(): void {
    dark = !dark;
    const theme = dark ? 'dark' : 'light';
    document.documentElement.dataset.theme = theme;
    try {
      localStorage.setItem(STORAGE_KEY, theme);
    } catch {
      // Storage is blocked (private mode); the choice still holds for this page view.
    }
  }
</script>

<button
  type="button"
  class="toggle"
  onclick={toggle}
  aria-label={dark ? 'Switch to light theme' : 'Switch to dark theme'}
>
  <svg width="20" height="20" viewBox="0 0 20 20" aria-hidden="true">
    {#if dark}
      <circle cx="10" cy="10" r="3.6" fill="currentColor" />
      <path
        d="M10 1.8v2.4M10 15.8v2.4M1.8 10h2.4M15.8 10h2.4M4.2 4.2l1.7 1.7M14.1 14.1l1.7 1.7M4.2 15.8l1.7-1.7M14.1 5.9l1.7-1.7"
        stroke="currentColor"
        stroke-width="1.6"
        stroke-linecap="round"
      />
    {:else}
      <path d="M16.5 12.2A7 7 0 0 1 7.8 3.5a7 7 0 1 0 8.7 8.7Z" fill="currentColor" />
    {/if}
  </svg>
</button>

<style>
  .toggle {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    width: 40px;
    height: 40px;
    padding: 0;
    border: none;
    border-radius: var(--radius-sm);
    background: transparent;
    color: var(--ink-2);
    cursor: pointer;
  }

  .toggle:hover {
    color: var(--ink);
    background: var(--wash);
  }
</style>
