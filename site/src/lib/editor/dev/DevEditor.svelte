<script lang="ts">
  import { DraftStore, type EditorDraft } from './draft-store.ts';
  import { onMount } from 'svelte';
  import EditorApp from '../EditorApp.svelte';
  import { EditorController } from '../core/controller.svelte.ts';
  import { getLocale, t } from '../i18n/i18n.svelte.ts';

  interface ServerState {
    projectId: string;
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
  let drafts: DraftStore | undefined;
  let recovery = $state<EditorDraft[]>([]);
  let draftFailure = $state(false);
  let checkedDrafts = false;
  const draftId = crypto.randomUUID();
  let restoredDraft: EditorDraft | undefined;
  let recovering = $state(false);
  let recoveryFailure = $state(false);
  function openRecovery(node: HTMLDialogElement) { node.showModal(); }
  let draftQueue = Promise.resolve();

  function checkpoint(bytes: Promise<Uint8Array>, version: number) {
    const projectId = serverState?.projectId;
    const id = draftId;
    const hash = baseHash;
    const fileName = doc.fileName;
    // Observe serialization failures immediately, even while an earlier write is pending.
    const snapshot = bytes.then(value => ({ value }), error => ({ error }));
    const operation = draftQueue.then(async () => {
      const result = await snapshot;
      if ('error' in result) throw result.error;
      if (!drafts || !projectId) throw new Error('Recovery storage unavailable');
      await drafts.put({ id, projectId, baseHash: hash, fileName, bytes: result.value, version, updated: Date.now() });
      draftFailure = false;
    });
    draftQueue = operation.catch(() => { draftFailure = true; });
    return operation.then(() => true, () => false);
  }
  async function restoreDraft(draft: EditorDraft) {
    if (recovering) return;
    recovering = true;
    recoveryFailure = false;
    try {
      await doc.loadBytes(draft.bytes, draft.fileName);
      restoredDraft = draft;
      baseHash = draft.baseHash;
      conflict = serverState?.documentHash !== baseHash || !!serverState?.conflict;
      doc.dirty = true;
      recovery = [];
      if (await checkpoint(doc.toBytes(), doc.version)) {
        await drafts?.remove(draft.id, draft.version);
        restoredDraft = undefined;
      }
    } catch { recoveryFailure = true; draftFailure = true; }
    finally { recovering = false; if (refreshAgain) { refreshAgain = false; void refresh(); } }
  }
  async function discardDraft(draft: EditorDraft) {
    try {
      await drafts?.remove(draft.id, draft.version);
      recovery = recovery.filter(item => item.id !== draft.id);
    } catch { recoveryFailure = true; }
  }

  async function refresh() {
    if (refreshing || saving || recovering) { refreshAgain = true; return; }
    refreshing = true;
    try {
      const response = await fetch('/editor/state');
      if (!response.ok) throw new Error(t('Preview unavailable'));
      const next: ServerState = await response.json();
      if (recovering) { refreshAgain = true; return; }
      serverState = next;
      if (!next.available) return;
      if (!checkedDrafts) {
        checkedDrafts = true;
        try { recovery = await drafts?.list(next.projectId) ?? []; }
        catch { draftFailure = true; }
      }
      if (loaded && recovery.length) return;
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
      if (saving || recovering || doc.version !== version) { refreshAgain = true; return; }
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
    if (recovering || recovery.length || saving || !loaded || !serverState || serverState.building || serverState.error) return;
    if (conflict && !resolveConflict) return;
    saving = true;
    failure = '';
    const version = doc.version;
    try {
      const bytes = await doc.toBytes();
      await checkpoint(Promise.resolve(bytes), version);
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
      await draftQueue;
      try {
        await drafts?.remove(draftId, version);
        if (restoredDraft) { await drafts?.remove(restoredDraft.id, restoredDraft.version); restoredDraft = undefined; }
      }
      catch { draftFailure = true; }
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
      await draftQueue;
      try {
        await drafts?.remove(draftId, version);
        if (restoredDraft) { await drafts?.remove(restoredDraft.id, restoredDraft.version); restoredDraft = undefined; }
      } catch { draftFailure = true; }
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
    if (loaded && doc.dirty && !doc.liveEditing && !recovery.length) void checkpoint(doc.toBytes(), doc.version);
  });
  $effect(() => {
    doc.committedVersion;
    if (recovering || recovery.length || !loaded || !doc.dirty || doc.liveEditing || saving || conflict || failure || serverState?.building || serverState?.error) return;
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
    try { drafts = new DraftStore(); } catch { draftFailure = true; }
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
    {#if draftFailure}<div class="conflict" role="alert">{t('Could not save a recovery copy in this browser. Keep this tab open until the project is saved.')}</div>{/if}
    {#if failure || serverState?.error}
      <div class="conflict" role="alert">{failure || serverState?.error}<button onclick={() => { failure = ''; void refresh(); }}>{t('Retry')}</button></div>
    {/if}
  {/snippet}
</EditorApp>
{#if loaded && recovery.length}
  <dialog class="recovery" aria-label={t('Recover unsaved changes')} use:openRecovery oncancel={e => e.preventDefault()} onkeydown={e => e.stopPropagation()}>
      <h2>{t('Recover unsaved changes')}</h2>
      <p>{t('This browser has a recovery copy for this project. Restore it or discard it to continue.')}</p>
      {#if recoveryFailure}<p role="alert">{t('Could not access the recovery copy. Try again.')}</p>{/if}
      {#each recovery as draft (draft.id)}
        <div class="recovery-entry">
          <span>{draft.fileName} — {new Date(draft.updated).toLocaleString(getLocale())}</span>
          <button disabled={recovering} onclick={() => restoreDraft(draft)}>{t('Restore changes')}</button>
          <button disabled={recovering} onclick={() => discardDraft(draft)}>{t('Discard recovery copy')}</button>
        </div>
      {/each}
  </dialog>
{/if}
{#if !loaded}<div class="loading" aria-label={t('Loading presentation…')}></div>{/if}

<style>
  .recovery::backdrop { background:#0005; }
  .recovery { max-width:650px; max-height:80vh; overflow:auto; padding:24px; border:0; background:white; color:#263544; border-radius:12px; box-shadow:0 8px 32px #0003; }
  .recovery h2 { margin-top:0; }
  .recovery-entry { display:flex; flex-wrap:wrap; align-items:center; gap:12px; padding:12px 0; border-top:1px solid #ddd; }
  .recovery-entry span { flex:1; }
  .save-status, .conflict { display:flex; align-items:center; flex-wrap:wrap; gap:10px; padding:5px 12px; font:12px system-ui; color:#304050; background:#eaf3ed; }
  .save-status span:first-child { flex:1; }
  .conflict { background:#fff0cd; color:#583b00; }
  button, a { font:inherit; color:inherit; cursor:pointer; }
  .loading { position:fixed; inset:110px 0 0; z-index:51; background:#ffffff90; }
</style>
