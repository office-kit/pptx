<script lang="ts">
  import { onMount } from 'svelte';
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';

  let { replace = false }: { replace?: boolean } = $props();
  const editor = getEditor();
  let dialog: HTMLDialogElement;
  let file = $state<File | null>(null);
  let url = $state('');
  let dimensions = $state<{ width: number; height: number } | null>(null);
  let error = $state('');
  let busy = $state(false);
  let mounted = true;
  const title = $derived(t(replace ? 'Replace image' : 'Insert image'));

  onMount(() => {
    dialog.showModal();
    return () => { mounted = false; };
  });
  $effect(() => {
    if (!file) { url = ''; return; }
    const next = URL.createObjectURL(file);
    url = next;
    return () => URL.revokeObjectURL(next);
  });
  function choose(event: Event) {
    const input = event.currentTarget;
    if (!(input instanceof HTMLInputElement)) return;
    file = input.files?.[0] ?? null;
    dimensions = null;
    error = '';
  }
  async function submit(event: SubmitEvent) {
    event.preventDefault();
    if (!file || !dimensions || busy) return;
    busy = true;
    const doc = editor.doc;
    const presentation = doc.pres;
    const version = doc.version;
    const selection = doc.selection;
    try {
      const bytes = new Uint8Array(await file.arrayBuffer());
      if (!mounted) return;
      if (doc.pres !== presentation || doc.version !== version || doc.selection !== selection) {
        error = 'The document changed. Select the image again.';
        return;
      }
      if (editor.applyImage(bytes, file.name, dimensions.width, dimensions.height, replace)) editor.closeDialog();
      else error = 'The image could not be inserted';
    } catch {
      error = 'The image could not be read';
    } finally {
      busy = false;
    }
  }
</script>

<dialog bind:this={dialog} aria-label={title} onclose={() => editor.closeDialog()}>
  <form onsubmit={submit}>
    <header><strong>{title}</strong><button type="button" class="ok-btn" aria-label={t('Close')} onclick={() => editor.closeDialog()}>✕</button></header>
    <label class="file-picker"><span>{t('Choose image')}</span><span>{file?.name ?? t('No image selected')}</span><input aria-label={t('Image file')} type="file" accept="image/png,image/jpeg,image/gif,image/webp,image/svg+xml" onchange={choose} disabled={busy} /></label>
    <p>{t(replace ? 'The image keeps its position, crop and appearance.' : 'The image is centered on the slide. Drag its handles to resize it.')}</p>
    {#if url}
      {#key url}
        <img src={url} alt={t('Selected image preview')} onload={(e) => { if (e.currentTarget instanceof HTMLImageElement) dimensions = { width: e.currentTarget.naturalWidth, height: e.currentTarget.naturalHeight }; }} onerror={() => { dimensions = null; error = 'The image could not be read'; }} />
      {/key}
    {/if}
    {#if error}<p role="alert">{t(error)}</p>{/if}
    <footer><button type="button" class="ok-btn" onclick={() => editor.closeDialog()}>{t('Cancel')}</button><button type="submit" class="ok-btn primary" disabled={!dimensions || busy}>{t(busy ? 'Reading image…' : replace ? 'Replace image' : 'Insert image')}</button></footer>
  </form>
</dialog>

<style>
  dialog { width: min(460px, 90vw); max-height: 85vh; padding: 18px; border: 1px solid var(--ok-border); border-radius: var(--ok-radius-lg); background: var(--ok-panel); color: var(--ok-text); box-shadow: var(--ok-shadow-lg); }
  dialog::backdrop { background: #0006; }
  form { display: flex; flex-direction: column; gap: 16px; }
  header, footer { display: flex; align-items: center; justify-content: space-between; gap: 12px; }
  footer { justify-content: flex-end; }
  label { display: grid; gap: 8px; }
  .file-picker { position: relative; padding: 12px; border: 1px solid var(--ok-border); border-radius: var(--ok-radius); cursor: pointer; }
  .file-picker:focus-within { outline: 2px solid var(--ok-accent); }
  .file-picker span:last-of-type { font-size: 12px; color: var(--ok-text-2); overflow-wrap: anywhere; }
  .file-picker input { position: absolute; inset: 0; width: 100%; height: 100%; opacity: 0; cursor: pointer; }
  p { font-size: 12px; color: var(--ok-text-2); margin: 0; }
  img { width: 100%; height: 220px; object-fit: contain; background: var(--ok-bg); }
  [role='alert'] { color: #bf3131; }
</style>
