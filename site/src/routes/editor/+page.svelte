<script lang="ts">
  import { onMount } from 'svelte';

  let target: HTMLElement;
  // Mounted on the client only: the editor is a browser UI with nothing to prerender.
  onMount(() => {
    let destroy: (() => void) | undefined;
    let cancelled = false;
    void import('@office-kit/pptx-editor').then(({ mountEditor }) => {
      if (!cancelled) destroy = mountEditor(target).destroy;
    });
    return () => {
      cancelled = true;
      destroy?.();
    };
  });
</script>

<svelte:head>
  <title>Editor · @office-kit/pptx</title>
  <meta
    name="description"
    content="A ribbon-style presentation editor built entirely on the @office-kit/pptx public API — every authoring capability the library exposes, wired into one UI."
  />
</svelte:head>

<div class="editor" bind:this={target}></div>

<p class="trademark">
  Microsoft and PowerPoint are trademarks of the Microsoft group of companies. This editor is an
  independent project, not affiliated with or endorsed by Microsoft.
</p>

<style>
  /* The editor takes the whole window, as a desktop application does. */
  .editor {
    position: fixed;
    inset: 0;
    z-index: 50;
  }
  .trademark {
    margin: 0;
    padding: 0.5rem 1rem;
    font-size: 0.75rem;
    color: var(--ink-3);
  }
</style>
