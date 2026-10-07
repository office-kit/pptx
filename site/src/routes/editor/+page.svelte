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

<style>
  /* The editor takes the whole window, as a desktop application does. */
  .editor {
    position: fixed;
    inset: 0;
    z-index: 50;
  }
</style>
