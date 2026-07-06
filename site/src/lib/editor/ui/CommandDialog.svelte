<script lang="ts">
  // Auto-generated argument dialog for any capability. Given a command id it
  // renders one ParamField per parameter (seeded with defaults + any preset
  // args), then dispatches through the controller. This is the fallback that
  // makes every capability usable without a bespoke form.
  import { getEditor } from '../core/context.ts';
  import { capabilityById } from '../manifest/index.ts';
  import ParamField from './ParamField.svelte';
  import { t, capLabel } from '../i18n/i18n.svelte.ts';

  interface Props {
    id: string;
  }
  let { id }: Props = $props();

  const editor = getEditor();
  const cap = $derived(capabilityById.get(id));
  const cmd = $derived(editor.command(id));

  // Editable args, seeded once from defaults + preset.
  let args = $state<Record<string, unknown>>({});
  $effect(() => {
    const seed: Record<string, unknown> = {};
    for (const p of cmd?.params ?? []) {
      if (editor.pendingPreset[p.name] !== undefined) seed[p.name] = editor.pendingPreset[p.name];
      else if (p.default !== undefined) seed[p.name] = coerceDefault(p.default, p.kind);
    }
    args = seed;
  });

  function coerceDefault(raw: string, kind: string): unknown {
    if (kind === 'number' || kind === 'index' || kind === 'emu') {
      const n = Number(raw);
      return Number.isNaN(n) ? undefined : n;
    }
    if (kind === 'boolean') return raw === 'true';
    return raw.replace(/^['"]|['"]$/g, '');
  }

  function submit(e: Event) {
    e.preventDefault();
    editor.invoke(id, args);
    editor.closeDialog();
  }
</script>

<div
  class="backdrop"
  role="presentation"
  onclick={(e) => {
    if (e.target === e.currentTarget) editor.closeDialog();
  }}
>
  <div class="dialog ok-editor" role="dialog" aria-modal="true" aria-label={cap ? capLabel(cap) : id}>
    <header>
      <div class="title">
        <strong>{cap ? capLabel(cap) : id}</strong>
        <code>{id}</code>
      </div>
      <button class="ok-btn" onclick={() => editor.closeDialog()} aria-label={t('Close')}>✕</button>
    </header>

    <form onsubmit={submit}>
      <div class="ok-scroll body">
        {#if (cmd?.params.length ?? 0) === 0}
          <p class="empty">{t('This command takes no arguments.')}</p>
        {:else}
          {#each cmd?.params ?? [] as p (p.name)}
            <ParamField spec={p} value={args[p.name]} onchange={(v) => (args = { ...args, [p.name]: v })} />
          {/each}
        {/if}
        {#if cap}
          <p class="meta">
            {t('operates on')} <b>{cap.operand}</b> · {t('category')} <b>{cap.category}</b> · {t('returns')}
            <code>{cap.returns || 'void'}</code>
          </p>
        {/if}
      </div>
      <footer>
        <button type="button" class="ok-btn" onclick={() => editor.closeDialog()}>{t('Cancel')}</button>
        <button type="submit" class="ok-btn primary">{t('Apply')}</button>
      </footer>
    </form>
  </div>
</div>

<style>
  .backdrop {
    position: fixed;
    inset: 0;
    background: rgba(0, 0, 0, 0.28);
    display: grid;
    place-items: center;
    z-index: 200;
  }
  .dialog {
    width: min(440px, 92vw);
    max-height: 84vh;
    background: var(--ok-panel);
    border-radius: var(--ok-radius-lg);
    box-shadow: var(--ok-shadow-lg);
    display: flex;
    flex-direction: column;
    overflow: hidden;
  }
  header {
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 12px 14px;
    border-bottom: 1px solid var(--ok-border);
  }
  .title {
    display: flex;
    flex-direction: column;
    gap: 2px;
  }
  .title code {
    font-family: var(--ok-mono);
    font-size: 10px;
    color: var(--ok-text-3);
  }
  form {
    display: flex;
    flex-direction: column;
    min-height: 0;
  }
  .body {
    padding: 14px;
    overflow: auto;
  }
  .empty {
    color: var(--ok-text-2);
    margin: 4px 0 0;
  }
  .meta {
    margin-top: 10px;
    padding-top: 10px;
    border-top: 1px dashed var(--ok-border);
    color: var(--ok-text-3);
    font-size: 11px;
  }
  .meta code {
    font-family: var(--ok-mono);
  }
  footer {
    display: flex;
    justify-content: flex-end;
    gap: 8px;
    padding: 10px 14px;
    border-top: 1px solid var(--ok-border);
    background: var(--ok-panel-2);
  }
  .primary {
    background: var(--ok-accent);
    color: #fff;
    border-color: var(--ok-accent);
  }
  .primary:hover {
    background: var(--ok-accent-2);
    border-color: var(--ok-accent-2);
  }
</style>
