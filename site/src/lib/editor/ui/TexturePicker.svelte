<script lang="ts">
  // The Format pane's Texture row under Picture or texture fill: the label on
  // the left and, right-aligned like the other value controls, a swatch ▾
  // button. Mac PowerPoint's swatch is a fixed placeholder icon; it does not
  // show the current texture.
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

<div class="texture-row">
  <span aria-hidden="true">{t('Texture')}</span>
  <button type="button" class="trigger" bind:this={trigger} aria-label={t('Texture')} aria-haspopup="menu" aria-expanded={open} {disabled} onclick={() => { open = !open; }}><span class="placeholder" aria-hidden="true"></span><span aria-hidden="true">▾</span></button>
</div>
{#if open && trigger}
  <TextureGallery label={t('Texture')} anchor={trigger}
    choose={(id) => { open = false; trigger?.focus(); choose(id); }}
    close={() => { open = false; }} />
{/if}

<style>
  .texture-row { display: flex; align-items: center; justify-content: space-between; gap: 6px; font-size: 11px; }
  .trigger { display: inline-flex; align-items: center; gap: 3px; padding: 2px 4px; font: inherit; font-size: 9px; color: var(--ok-text); background: var(--ok-panel); border: 1px solid var(--ok-border); border-radius: var(--ok-radius); cursor: pointer; }
  .trigger:hover:not(:disabled), .trigger[aria-expanded=true] { background: var(--ok-hover); }
  .trigger:disabled { opacity: 0.5; cursor: default; }
  /* PowerPoint's placeholder: a small white tile dotted in blue. */
  .placeholder { width: 14px; height: 14px; border: 1px solid #7f9cc9; background-color: #fff; background-image: radial-gradient(circle, #3b73d1 0.9px, transparent 1.2px); background-size: 3.5px 3.5px; }
</style>
