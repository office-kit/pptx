<script lang="ts">
  import { onMount } from 'svelte';
  import EditorApp from '../EditorApp.svelte';
  import { EditorController } from '../core/controller.svelte.ts';
  import { getLocale, t } from '../i18n/i18n.svelte.ts';

  interface ServerState {
    revision: string;
    previewRevision: number;
    documentHash?: string;
    fileName: string;
    hasEdits: boolean;
    conflict: boolean;
    building: boolean;
    error: string | null;
    available: boolean;
    message?: string;
  }
  const editor = new EditorController();
  const doc = editor.doc;
  let serverState = $state<ServerState>();
  let loaded = $state(false);
  let saving = $state(false);
  let conflict = $state(false);
  let failure = $state('');
  let baseRevision = '';
  let baseHash: string | undefined;
  let refreshing = false;
  let refreshAgain = false;

  async function refresh() {
    if (refreshing || saving) { refreshAgain = true; return; }
    refreshing = true;
    try {
      const response = await fetch('/editor/state');
      if (!response.ok) throw new Error(t('Preview unavailable'));
      const next: ServerState = await response.json();
      serverState = next;
      if (!next.available) return;
      if (loaded && next.documentHash === baseHash) {
        baseRevision = next.revision;
        conflict = next.conflict;
        return;
      }
      if (loaded && doc.dirty) { conflict = true; return; }
      const version = doc.version;
      const deck = await fetch('/editor/document');
      if (!deck.ok) throw new Error(t('Preview unavailable'));
      const bytes = new Uint8Array(await deck.arrayBuffer());
      if (saving || doc.version !== version) { refreshAgain = true; return; }
      if (deck.headers.get('etag') !== next.revision) { refreshAgain = true; return; }
      await doc.loadBytes(bytes, next.fileName);
      baseRevision = next.revision;
      baseHash = next.documentHash;
      loaded = true;
      conflict = next.conflict;
    } catch (cause) {
      failure = cause instanceof Error ? cause.message : String(cause);
    } finally {
      refreshing = false;
      if (refreshAgain && !saving) { refreshAgain = false; void refresh(); }
    }
  }

  async function save(resolveConflict = false) {
    if (saving || !loaded || !serverState || serverState.building || serverState.error) return;
    if (conflict && !resolveConflict) return;
    saving = true;
    failure = '';
    const version = doc.version;
    try {
      const bytes = await doc.toBytes();
      const response = await fetch('/editor/document', {
        method: 'PUT',
        headers: {
          'Content-Type': 'application/octet-stream',
          'If-Match': resolveConflict ? serverState.revision : baseRevision,
          ...(resolveConflict ? { 'X-Editor-Resolve': 'edits' } : {}),
        },
        body: new Blob([new Uint8Array(bytes)]),
      });
      const next: ServerState = await response.json();
      serverState = next;
      if (response.status === 409) { conflict = true; return; }
      if (!response.ok) throw new Error(next.message ?? t('Save failed'));
      baseRevision = next.revision;
      baseHash = next.documentHash;
      doc.markSaved(version);
      conflict = next.conflict;
    } catch (cause) {
      failure = cause instanceof Error ? cause.message : String(cause);
    } finally {
      saving = false;
      void refresh();
    }
  }

  async function useSource() {
    if (!serverState || saving || !confirm(t('Replace editor changes with the current source?'))) return;
    saving = true;
    failure = '';
    const version = doc.version;
    try {
      const response = await fetch('/editor/resolve', {
        method: 'POST', headers: { 'If-Match': serverState.revision },
      });
      if (!response.ok) throw new Error(t('Preview changed. Try again.'));
      // Load the replacement only after its disk transaction has succeeded.
      if (version !== doc.version) { conflict = true; return; }
      doc.dirty = false;
      loaded = false;
      conflict = false;
    } catch (cause) {
      failure = cause instanceof Error ? cause.message : String(cause);
    } finally {
      saving = false;
      void refresh();
    }
  }

  $effect(() => {
    doc.committedVersion;
    if (!loaded || !doc.dirty || doc.liveEditing || saving || conflict || failure || serverState?.building || serverState?.error) return;
    const timer = setTimeout(() => { void save(); }, 700);
    return () => clearTimeout(timer);
  });
  $effect(() => {
    window.parent.postMessage({
      type: 'editor-focus', slide: doc.selection.slideIndex,
      count: doc.slides.length, revision: serverState?.previewRevision ?? 0,
      dirty: doc.dirty, locale: getLocale(),
    }, window.location.origin);
  });
  onMount(() => {
    const events = new EventSource('/events');
    events.onmessage = () => { void refresh(); };
    events.onerror = () => { failure = t('Reconnecting…'); };
    events.onopen = () => { failure = ''; void refresh(); };
    void refresh();
    return () => events.close();
  });
  function beforeUnload(event: BeforeUnloadEvent) {
    if (doc.dirty) { event.preventDefault(); event.returnValue = ''; }
  }
</script>

<svelte:window onbeforeunload={beforeUnload} />
<EditorApp {editor} onsave={() => save()}>
  {#snippet status()}
    <div class="save-status" role="status">
      <span>{saving ? t('Saving…') : !loaded ? t('Loading presentation…') : doc.dirty ? t('Unsaved changes') : t('Saved to this project')}</span>
      {#if serverState?.building}<span>{t('Building source…')}</span>{/if}
      {#if serverState?.hasEdits && !conflict}<button onclick={useSource} disabled={saving || serverState.building}>{t('Use source')}</button>{/if}
    </div>
    {#if conflict}
      <div class="conflict" role="alert">
        <span>{t('The source or saved deck changed. Your edits are still here.')}</span>
        <button onclick={() => save(true)} disabled={saving || serverState?.building || !!serverState?.error}>{t('Keep my edits')}</button>
        <button onclick={useSource} disabled={saving || serverState?.building || !!serverState?.error}>{t('Use source')}</button>
        <a href="/editor/source" download="source.pptx">{t('Download source')}</a>
      </div>
    {/if}
    {#if failure || serverState?.error}
      <div class="conflict" role="alert">{failure || serverState?.error}<button onclick={() => { failure = ''; void refresh(); }}>{t('Retry')}</button></div>
    {/if}
  {/snippet}
</EditorApp>
{#if !loaded}<div class="loading" aria-label={t('Loading presentation…')}></div>{/if}

<style>
  .save-status, .conflict { display:flex; align-items:center; flex-wrap:wrap; gap:10px; padding:5px 12px; font:12px system-ui; color:#304050; background:#eaf3ed; }
  .save-status span:first-child { flex:1; }
  .conflict { background:#fff0cd; color:#583b00; }
  button, a { font:inherit; color:inherit; cursor:pointer; }
  .loading { position:fixed; inset:110px 0 0; z-index:51; background:#ffffff90; }
</style>
