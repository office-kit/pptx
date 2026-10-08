// The editor page (`/editor`) inside the preview: it loads, saves and merges
// the deck through the server's `/editor/*` endpoints, around `mountDevEditor`.
// Source rebuilds reach the editor as proposals from the source, so unsaved
// edits merge with them and only real collisions ask the user.

import type { EditorOptions, ProposeResult } from '@office-kit/pptx-editor';
import { mountDevEditor } from '@office-kit/pptx-editor/internal';
import { DraftStore, type EditorDraft } from './draft-store.ts';

interface ServerState {
  projectId: string;
  revision: string;
  previewRevision: number;
  documentHash?: string;
  sourceHash?: string;
  fileName: string;
  hasEdits: boolean;
  /** The saved edits could not be merged with the current source. */
  conflict: boolean;
  building: boolean;
  error: string | null;
  available: boolean;
  message?: string;
  /** Why a save was refused: `busy` saves are retried after the rebuild. */
  refused?: 'busy' | 'stale' | 'conflict';
  /** The save was merged with source changes; the uploaded deck's revision. */
  uploadRevision?: string;
}

/** The server document the editor's presentation was derived from. */
interface Base {
  revision: string;
  hash: string | undefined;
  /** `null` when the server no longer has it: merging then asks the user. */
  bytes: Uint8Array | null;
}

const JA: Record<string, string> = {
  'Saving…': '保存中…',
  'Loading presentation…': 'プレゼンテーションを読み込み中…',
  'Unsaved changes': '未保存の変更',
  'Saved to this project': 'このプロジェクトに保存済み',
  'Building source…': 'ソースをビルド中…',
  'Use source': 'ソースを使用',
  'Replace editor changes with the current source?':
    'エディターでの変更を現在のソースで置き換えますか？',
  'Preview unavailable': 'プレビューに接続できません',
  'Preview changed. Try again.': 'プレビューが更新されました。もう一度お試しください。',
  'Reconnecting…': '再接続中…',
  Retry: '再試行',
  'Save failed': '保存に失敗しました',
  'Source changed': 'ソースの変更',
  'Recover unsaved changes': '未保存の変更を復元',
  'This browser has a recovery copy for this project. Restore it or discard it to continue.':
    'このプロジェクトの復元用コピーがブラウザーに残っています。復元するか破棄して続行してください。',
  'Could not access the recovery copy. Try again.':
    '復元用コピーにアクセスできませんでした。もう一度お試しください。',
  'Restore changes': '変更を復元',
  'Discard recovery copy': '復元用コピーを破棄',
  'Could not save a recovery copy in this browser. Keep this tab open until the project is saved.':
    'このブラウザーに復元用コピーを保存できませんでした。プロジェクトへの保存が完了するまで、このタブを開いたままにしてください。',
};

const STYLES = `
.dev-status { display: contents; }
.save-status, .dev-status .conflict { display:flex; align-items:center; flex-wrap:wrap; gap:10px; padding:5px 12px; font:12px system-ui; color:#304050; background:#eaf3ed; }
.compact-host .save-status { gap:6px; padding:3px 8px; font-size:11px; }
.save-status span:first-child { flex:1; }
.dev-status .conflict { background:#fff0cd; color:#583b00; }
/* A wrapping status would change the bar's height whenever it updates and
   shift the canvas mid-gesture; the title bar keeps one line. */
.topbar.compact .save-status { flex:0 1 auto; min-height:24px; white-space:nowrap; }
.topbar.compact .dev-status .conflict { flex-basis:100%; }
.dev-status button, .recovery button { font:inherit; color:inherit; cursor:pointer; }
.recovery::backdrop { background:#0005; }
.recovery { max-width:650px; max-height:80vh; overflow:auto; padding:24px; border:0; background:white; color:#263544; border-radius:12px; box-shadow:0 8px 32px #0003; }
.recovery h2 { margin-top:0; }
.recovery-entry { display:flex; flex-wrap:wrap; align-items:center; gap:12px; padding:12px 0; border-top:1px solid #ddd; }
.recovery-entry span { flex:1; }
.loading { position:fixed; inset:110px 0 0; z-index:51; background:#ffffff90; }
@media (prefers-color-scheme: dark) {
  .recovery { background:#2b2b2b; color:#f0f0f0; }
  .recovery-entry { border-top-color:#3d3d3d; }
  .save-status { color:#cfe8d6; background:#23382b; }
  .dev-status .conflict { background:#4a3a12; color:#ffe2a3; }
  .loading { background:#1f1f1f90; }
}`;

const FEEDBACK_URL = 'https://github.com/office-kit/pptx/issues';

const message = (cause: unknown) => (cause instanceof Error ? cause.message : String(cause));
const bytesOf = async (response: Response) => new Uint8Array(await response.arrayBuffer());

function element<K extends keyof HTMLElementTagNameMap>(
  tag: K,
  properties: Partial<HTMLElementTagNameMap[K]> = {},
  ...children: (Node | string)[]
): HTMLElementTagNameMap[K] {
  const node = Object.assign(document.createElement(tag), properties);
  node.append(...children);
  return node;
}

const style = document.createElement('style');
style.textContent = STYLES;
document.head.append(style);

const status = element('div', { className: 'dev-status' });
const options: EditorOptions = { isolate: false, compact: true, autoSave: true, status, onSave };
const editor = mountDevEditor(document.body, options, {
  // Opened on its own rather than in the preview's frame, there is no page to present on.
  previewFrame: window.parent !== window,
  // The dev tool is ours, so its Help ▸ Feedback leads to our issue tracker.
  feedback: { url: FEEDBACK_URL },
});
const t = (key: string) => (editor.locale === 'ja' ? (JA[key] ?? key) : key);

let serverState: ServerState | undefined;
let base: Base = { revision: '', hash: undefined, bytes: null };
let loaded = false;
let saving = false;
let failure = '';
let refreshing = false;
let refreshAgain = false;
// A source proposal is merging or waiting for the user's choice.
let proposing = false;
// The editor asked to save while that was not possible; save once it is.
let deferred = false;
// The next save keeps the editor's deck over a source it could not merge with.
let resolveEdits = false;
// The next document from the server replaces the deck, unsaved edits included.
let replaceNext = false;

let drafts: DraftStore | undefined;
let recovery: EditorDraft[] = [];
let checkedDrafts = false;
let recovering = false;
let recoveryFailure = false;
let draftFailure = false;
const draftId = crypto.randomUUID();
let restoredDraft: EditorDraft | undefined;
let draftQueue = Promise.resolve();
// Orders recovery copies: an older serialization never replaces a newer one.
let changes = 0;
try {
  drafts = new DraftStore();
} catch {
  draftFailure = true;
}

function checkpoint(bytes: Promise<Uint8Array>): Promise<boolean> {
  const version = ++changes;
  const projectId = serverState?.projectId;
  const { hash, revision } = base;
  const fileName = serverState?.fileName ?? 'presentation.pptx';
  // Observe serialization failures immediately, even while an earlier write is pending.
  const snapshot = bytes.then(
    (value) => ({ value }),
    (error: unknown) => ({ error }),
  );
  const operation = draftQueue.then(async () => {
    const result = await snapshot;
    if ('error' in result) throw result.error;
    if (!drafts || !projectId) throw new Error('Recovery storage unavailable');
    await drafts.put({
      id: draftId,
      projectId,
      baseHash: hash,
      baseRevision: revision,
      fileName,
      bytes: result.value,
      version,
      updated: Date.now(),
    });
    draftFailure = false;
  });
  draftQueue = operation.catch(() => {
    draftFailure = true;
    render();
  });
  return operation.then(
    () => true,
    () => false,
  );
}

/** The deck is saved: its recovery copies are no longer needed. */
async function dropDrafts(): Promise<void> {
  const version = changes;
  await draftQueue;
  try {
    await drafts?.remove(draftId, version);
    if (restoredDraft) {
      await drafts?.remove(restoredDraft.id, restoredDraft.version);
      restoredDraft = undefined;
    }
  } catch {
    draftFailure = true;
    render();
  }
}

/** Records the server document the deck now derives from. */
function adopt(state: ServerState, bytes: Uint8Array): void {
  base = { revision: state.revision, hash: state.documentHash, bytes };
}

/** The preview page pairs the slide being edited with the preview revision. */
function updateState(state: ServerState): void {
  serverState = state;
  window.parent.postMessage(
    { type: 'editor-revision', revision: state.previewRevision ?? 0 },
    window.location.origin,
  );
}

async function onSave(bytes: Uint8Array): Promise<boolean> {
  const state = serverState;
  if (
    saving ||
    recovering ||
    recovery.length ||
    !loaded ||
    proposing ||
    !state ||
    state.building ||
    state.error ||
    (state.conflict && !resolveEdits)
  ) {
    deferred = true;
    return false;
  }
  saving = true;
  failure = '';
  render();
  try {
    await checkpoint(Promise.resolve(bytes));
    const response = await fetch('/editor/document', {
      method: 'PUT',
      headers: {
        'Content-Type': 'application/octet-stream',
        'If-Match': resolveEdits ? state.revision : base.revision,
        ...(resolveEdits ? { 'X-Editor-Resolve': 'edits' } : {}),
      },
      body: new Blob([new Uint8Array(bytes)]),
    });
    const next: ServerState = await response.json();
    updateState(next);
    // Busy: retried after the rebuild. Otherwise the server could not merge the
    // upload with a newer deck; the refresh below merges it here instead.
    if (response.status === 409) {
      deferred = true;
      return false;
    }
    if (!response.ok) throw new Error(next.message ?? t('Save failed'));
    resolveEdits = false;
    // An upload the server merged with source changes: it remembers the upload
    // under its own revision, and the refresh merges the result in.
    base = next.uploadRevision
      ? { revision: next.uploadRevision, hash: undefined, bytes }
      : { revision: next.revision, hash: next.documentHash, bytes };
    return true;
  } catch (cause) {
    failure = message(cause);
    deferred = true;
    return false;
  } finally {
    saving = false;
    render();
    void refresh();
  }
}

function saveDeferred(): void {
  if (!deferred || saving || failure || proposing) return;
  deferred = false;
  if (editor.dirty) void editor.save();
}

/** Discards the saved canvas version on the server: the source becomes the deck. */
async function resolveToSource(revision: string): Promise<void> {
  const response = await fetch('/editor/resolve', {
    method: 'POST',
    headers: { 'If-Match': revision },
  });
  if (!response.ok) throw new Error(t('Preview changed. Try again.'));
}

async function proposeSource(
  baseBytes: Uint8Array | null,
  edited: Uint8Array,
): Promise<ProposeResult> {
  proposing = true;
  render();
  try {
    return await editor.propose(baseBytes, edited, { label: t('Source changed'), from: 'source' });
  } finally {
    proposing = false;
    render();
  }
}

/** Keeping the editor's deck saves it; taking the source drops the saved canvas version too. */
async function settle(result: ProposeResult, state: ServerState): Promise<void> {
  if (result.status === 'took-theirs') {
    await resolveToSource(state.revision);
    return;
  }
  if (result.status === 'kept-mine' || state.conflict) {
    resolveEdits = state.conflict;
    deferred = false;
    await editor.save();
  }
}

/** The saved edits could not be merged with the source: the user chooses in the editor. */
async function proposeServerConflict(state: ServerState): Promise<void> {
  const [savedBase, source] = await Promise.all([fetch('/editor/base'), fetch('/editor/source')]);
  if (!source.ok) throw new Error(t('Preview unavailable'));
  if (source.headers.get('etag') !== state.revision) {
    refreshAgain = true;
    return;
  }
  const result = await proposeSource(
    savedBase.ok ? await bytesOf(savedBase) : null,
    await bytesOf(source),
  );
  await settle(result, state);
}

async function refresh(): Promise<void> {
  if (refreshing || saving || recovering || proposing) {
    refreshAgain = true;
    return;
  }
  refreshing = true;
  try {
    const response = await fetch('/editor/state');
    if (!response.ok) throw new Error(t('Preview unavailable'));
    const next: ServerState = await response.json();
    if (recovering) {
      refreshAgain = true;
      return;
    }
    updateState(next);
    render();
    if (!next.available) return;
    if (!checkedDrafts) {
      checkedDrafts = true;
      try {
        recovery = (await drafts?.list(next.projectId)) ?? [];
      } catch {
        draftFailure = true;
      }
      render();
    }
    if (loaded && recovery.length) return;
    if (loaded && !replaceNext && next.documentHash === base.hash) {
      base = { ...base, revision: next.revision };
    } else {
      const deck = await fetch('/editor/document');
      if (!deck.ok) throw new Error(t('Preview unavailable'));
      const bytes = await bytesOf(deck);
      if (saving || recovering || deck.headers.get('etag') !== next.revision) {
        refreshAgain = true;
        return;
      }
      if (!loaded || replaceNext) {
        await editor.open(bytes, { fileName: next.fileName });
        loaded = true;
        replaceNext = false;
        adopt(next, bytes);
      } else {
        const result = await proposeSource(base.bytes, bytes);
        adopt(next, bytes);
        await settle(result, { ...next, conflict: false });
      }
    }
    if (next.conflict && !resolveEdits && !recovery.length) await proposeServerConflict(next);
  } catch (cause) {
    failure = message(cause);
  } finally {
    refreshing = false;
    render();
    if (refreshAgain && !saving) {
      refreshAgain = false;
      void refresh();
    } else saveDeferred();
  }
}

async function useSource(): Promise<void> {
  if (!serverState || saving || proposing) return;
  if (!confirm(t('Replace editor changes with the current source?'))) return;
  saving = true;
  failure = '';
  render();
  const before = changes;
  try {
    await resolveToSource(serverState.revision);
    // Edits made while the request ran merge with the source instead of being replaced.
    replaceNext = before === changes;
  } catch (cause) {
    failure = message(cause);
  } finally {
    saving = false;
    render();
    void refresh();
  }
}

async function restoreDraft(draft: EditorDraft): Promise<void> {
  if (recovering) return;
  recovering = true;
  recoveryFailure = false;
  render();
  try {
    await editor.open(draft.bytes, { fileName: draft.fileName, unsaved: true });
    restoredDraft = draft;
    recovery = [];
    // A copy of an older deck merges with the current one, if the server still has that deck.
    if (draft.baseHash !== serverState?.documentHash) {
      const older = draft.baseRevision
        ? await fetch(`/editor/document?revision=${encodeURIComponent(draft.baseRevision)}`)
        : undefined;
      base = {
        revision: draft.baseRevision ?? '',
        hash: draft.baseHash,
        bytes: older?.ok ? await bytesOf(older) : null,
      };
    }
    if (await checkpoint(editor.snapshot())) {
      await drafts?.remove(draft.id, draft.version);
      restoredDraft = undefined;
    }
  } catch {
    recoveryFailure = true;
    draftFailure = true;
  } finally {
    recovering = false;
    render();
    refreshAgain = false;
    void refresh();
  }
}

async function discardDraft(draft: EditorDraft): Promise<void> {
  try {
    await drafts?.remove(draft.id, draft.version);
    recovery = recovery.filter((item) => item.id !== draft.id);
  } catch {
    recoveryFailure = true;
  }
  render();
  if (!recovery.length) void refresh();
}

const recoveryDialog = element('dialog', { className: 'recovery' });
recoveryDialog.addEventListener('cancel', (event) => event.preventDefault());
// Keys in the dialog must not reach the editor's shortcuts underneath.
recoveryDialog.addEventListener('keydown', (event) => event.stopPropagation());
const loading = element('div', { className: 'loading' });
let rendered = '';

function render(): void {
  const state = serverState;
  const view = {
    locale: editor.locale,
    text: saving
      ? 'Saving…'
      : !loaded
        ? 'Loading presentation…'
        : editor.dirty
          ? 'Unsaved changes'
          : 'Saved to this project',
    building: !!state?.building,
    useSource: !!state?.hasEdits && !state.conflict && !proposing,
    useSourceDisabled: saving || !!state?.building,
    draftFailure,
    failure: failure || state?.error || '',
    recovery: loaded && recovery.length ? recovery.map((draft) => draft.id) : [],
    recovering,
    recoveryFailure,
    loaded,
  };
  // Rebuilding only on a change keeps focus on the status buttons.
  const key = JSON.stringify(view);
  if (key === rendered) return;
  rendered = key;

  const line = element(
    'div',
    { className: 'save-status', role: 'status' },
    element('span', {}, t(view.text)),
  );
  if (view.building) line.append(element('span', {}, t('Building source…')));
  if (view.useSource)
    line.append(
      element(
        'button',
        { disabled: view.useSourceDisabled, onclick: () => void useSource() },
        t('Use source'),
      ),
    );
  status.replaceChildren(line);
  if (view.draftFailure)
    status.append(
      element(
        'div',
        { className: 'conflict', role: 'alert' },
        t(
          'Could not save a recovery copy in this browser. Keep this tab open until the project is saved.',
        ),
      ),
    );
  if (view.failure)
    status.append(
      element(
        'div',
        { className: 'conflict', role: 'alert' },
        view.failure,
        element(
          'button',
          {
            onclick: () => {
              failure = '';
              render();
              void refresh();
            },
          },
          t('Retry'),
        ),
      ),
    );

  if (view.recovery.length) {
    recoveryDialog.ariaLabel = t('Recover unsaved changes');
    recoveryDialog.replaceChildren(
      element('h2', {}, t('Recover unsaved changes')),
      element(
        'p',
        {},
        t(
          'This browser has a recovery copy for this project. Restore it or discard it to continue.',
        ),
      ),
      ...(view.recoveryFailure
        ? [element('p', { role: 'alert' }, t('Could not access the recovery copy. Try again.'))]
        : []),
      ...recovery.map((draft) =>
        element(
          'div',
          { className: 'recovery-entry' },
          element(
            'span',
            {},
            `${draft.fileName} — ${new Date(draft.updated).toLocaleString(view.locale)}`,
          ),
          element(
            'button',
            { disabled: recovering, onclick: () => void restoreDraft(draft) },
            t('Restore changes'),
          ),
          element(
            'button',
            { disabled: recovering, onclick: () => void discardDraft(draft) },
            t('Discard recovery copy'),
          ),
        ),
      ),
    );
    if (!recoveryDialog.isConnected) {
      document.body.append(recoveryDialog);
      recoveryDialog.showModal();
    }
  } else if (recoveryDialog.isConnected) {
    recoveryDialog.close();
    recoveryDialog.remove();
  }
  if (!view.loaded) document.body.append(loading);
  else loading.remove();
  loading.ariaLabel = t('Loading presentation…');
}

editor.on('change', () => {
  if (loaded && editor.dirty && !recovery.length) void checkpoint(editor.snapshot());
  render();
});
editor.on('dirtychange', ({ dirty }) => {
  if (!dirty) void dropDrafts();
  render();
});
editor.on('localechange', render);
window.addEventListener('beforeunload', (event) => {
  if (!editor.dirty) return;
  event.preventDefault();
  event.returnValue = '';
});

await editor.ready;
render();
const events = new EventSource('/events');
events.onmessage = () => void refresh();
events.onerror = () => {
  failure = t('Reconnecting…');
  render();
};
events.onopen = () => {
  failure = '';
  void refresh();
};
void refresh();
