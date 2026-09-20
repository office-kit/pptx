<script lang="ts">
  import '../app.css';
  import CodeCopyEnhancer from '$lib/components/CodeCopyEnhancer.svelte';
  import KitFooter from '$lib/kit/KitFooter.svelte';
  import KitHeader from '$lib/kit/KitHeader.svelte';

  type Props = {
    children?: import('svelte').Snippet;
  };

  const { children }: Props = $props();

  const links = [
    { path: '/docs/getting-started', label: 'Docs', section: '/docs' },
    { path: '/docs/recipes', label: 'Recipes' },
    { path: '/api', label: 'API' },
    { path: '/playground', label: 'Playground' },
    { path: '/repl', label: 'REPL' },
  ];
</script>

<a class="skip" href="#main">Skip to content</a>

<KitHeader {links} />

<main id="main">
  {@render children?.()}
</main>

<CodeCopyEnhancer />

<KitFooter {links} />

<style>
  main {
    min-height: calc(100vh - var(--header-h));
  }

  .skip {
    position: absolute;
    left: 1rem;
    top: -4rem;
    z-index: 40;
    padding: 0.6rem 1rem;
    border-radius: var(--radius);
    background: var(--ink);
    color: var(--paper);
  }

  .skip:focus {
    top: 0.75rem;
  }
</style>
