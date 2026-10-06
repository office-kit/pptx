<script lang="ts">
  // The Format pane's Texture ▾ button under Picture or texture fill.
  import type { TextureId } from '../core/textures.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import TextureGallery from './TextureGallery.svelte';

  let { disabled = false, choose }: {
    disabled?: boolean;
    choose: (id: TextureId) => void;
  } = $props();
  let open = $state(false);
  let trigger = $state<HTMLButtonElement>();
  $effect(() => { if (disabled) open = false; });
</script>

<button type="button" class="ok-btn" bind:this={trigger} aria-label={t('Texture')} aria-haspopup="menu" aria-expanded={open} {disabled} onclick={() => { open = !open; }}>{t('Texture')} ▾</button>
{#if open && trigger}
  <TextureGallery label={t('Texture')} anchor={trigger}
    choose={(id) => { open = false; trigger?.focus(); choose(id); }}
    close={() => { open = false; }} />
{/if}
