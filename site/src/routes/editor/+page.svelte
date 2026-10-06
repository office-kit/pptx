<script lang="ts">
  import EditorApp from '$lib/editor/EditorApp.svelte';
  import { t } from '$lib/editor/i18n/i18n.svelte.ts';
</script>

<svelte:head>
  <title>Editor · @office-kit/pptx</title>
  <meta
    name="description"
    content="A ribbon-style presentation editor built entirely on the @office-kit/pptx public API — every authoring capability the library exposes, wired into one UI."
  />
</svelte:head>

<EditorApp />

<p class="trademark">
  {t('Microsoft and PowerPoint are trademarks of the Microsoft group of companies. This editor is an independent project, not affiliated with or endorsed by Microsoft.')}
</p>

<style>
  .trademark {
    margin: 0;
    padding: 0.5rem 1rem;
    font-size: 0.75rem;
    color: var(--ink-3);
  }
</style>
