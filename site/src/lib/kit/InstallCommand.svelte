<script lang="ts">
  type Props = {
    /** npm package name to install. */
    pkg: string;
  };

  const { pkg }: Props = $props();

  const command = $derived(`npm i ${pkg}`);

  let status = $state<'idle' | 'copied' | 'failed'>('idle');
  let timer: ReturnType<typeof setTimeout> | undefined;

  async function copy(): Promise<void> {
    try {
      await navigator.clipboard.writeText(command);
      status = 'copied';
    } catch {
      // Clipboard access is denied on insecure origins and in some embedded views.
      status = 'failed';
    }
    clearTimeout(timer);
    timer = setTimeout(() => (status = 'idle'), 1600);
  }
</script>

<button type="button" class="install" onclick={copy} aria-label="Copy the install command: {command}">
  <span class="prompt" aria-hidden="true">$</span>
  <code>{command}</code>
  <span class="state" aria-live="polite">
    {#if status === 'copied'}Copied{:else if status === 'failed'}Press ⌘C{:else}Copy{/if}
  </span>
</button>

<style>
  .install {
    display: inline-flex;
    align-items: center;
    gap: 0.6rem;
    min-height: 44px;
    max-width: 100%;
    padding: 0 0.5rem 0 0.95rem;
    border: 1px solid var(--line-strong);
    border-radius: var(--radius);
    background: var(--wash);
    color: var(--ink);
    font-family: var(--mono);
    font-size: 0.88rem;
    cursor: pointer;
  }

  .install:hover {
    border-color: var(--ink-3);
  }

  .prompt {
    color: var(--accent-ink);
    font-weight: 600;
  }

  code {
    padding: 0;
    border: none;
    background: none;
    font-size: inherit;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
  }

  @media (max-width: 520px) {
    .install {
      width: 100%;
      justify-content: space-between;
      font-size: 0.8rem;
    }

    code {
      flex: 1;
      text-align: left;
    }
  }

  .state {
    flex: none;
    min-width: 4.2rem;
    padding: 0.3rem 0.5rem;
    border-radius: var(--radius-sm);
    background: var(--paper);
    border: 1px solid var(--line);
    color: var(--ink-2);
    font-family: var(--sans);
    font-size: 0.8rem;
    font-weight: 550;
    text-align: center;
  }
</style>
