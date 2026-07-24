<script lang="ts">
  import { getEditor } from '../core/context.ts';
  const editor = getEditor();
</script>

<div class="toasts" role="status" aria-live="polite">
  {#each editor.toasts as t (t.id)}
    <div class="toast {t.kind}">{t.message}</div>
  {/each}
</div>

<style>
  .toasts {
    position: fixed;
    bottom: 40px;
    left: 50%;
    transform: translateX(-50%);
    display: flex;
    flex-direction: column;
    gap: 6px;
    z-index: 300;
    pointer-events: none;
  }
  .toast {
    padding: 8px 16px;
    border-radius: var(--ok-radius-lg);
    box-shadow: var(--ok-shadow-lg);
    font-size: 12px;
    color: #fff;
    max-width: 60vw;
  }
  .toast.info {
    background: #323130;
  }
  .toast.error {
    background: var(--ok-danger);
  }
</style>
