<script lang="ts">
  import type { Snippet } from 'svelte';
  import { getEditor } from '../core/context.ts';

  // A Format pane section: PowerPoint's chevron header band over label-left /
  // control-right rows. Every section starts collapsed, as natively, and the
  // editor remembers which ones the user opened for the rest of the session.
  let { id, label, children }: { id: string; label: string; children: Snippet } = $props();
  const editor = getEditor();
  const stored = $derived(editor.formatPaneSections[id] ?? false);
  // The DOM opens on click but `toggle` reaches the store a task later, so the
  // store may briefly lag. Push it into the element only when this section's
  // own value changes: re-asserting it on every re-render (a language switch
  // re-renders the label, a neighbour's toggle replaces the record) would
  // close a section the user has just opened.
  const syncOpen = (node: HTMLDetailsElement) => {
    node.open = stored;
  };
</script>

<details class="pane-section" data-section={id} {@attach syncOpen}
  ontoggle={(event) => {
    const open = event.currentTarget.open;
    if (stored !== open) editor.formatPaneSections = { ...editor.formatPaneSections, [id]: open };
  }}>
  <summary>{label}</summary>
  <div class="pane-fields">{@render children()}</div>
</details>

<style>
  .pane-section { margin: 0 -10px; font-size: 12px; }
  /* Mac PowerPoint: 25 pt header pitch, a chevron (› closed, ⌄ open) on a band. */
  summary {
    display: flex;
    align-items: center;
    gap: 6px;
    box-sizing: border-box;
    min-height: 25px;
    padding: 0 8px;
    list-style: none;
    background: var(--ok-hover);
    border-bottom: 1px solid var(--ok-panel);
    cursor: pointer;
  }
  summary::-webkit-details-marker { display: none; }
  summary::before { content: '›'; display: inline-block; width: 10px; text-align: center; font-size: 14px; transition: transform 0.12s; }
  details[open] > summary::before { transform: rotate(90deg); }
  /* Radio options on a 20 pt pitch, a rule, then 30 pt rows with 26 pt
     controls ending 17 pt from the pane edge. */
  .pane-fields { display: flex; flex-direction: column; gap: 4px; padding: 10px 17px 10px 16px; }
  .pane-fields :global(.pane-row) { display: flex; align-items: center; justify-content: space-between; gap: 8px; min-height: 26px; }
  .pane-fields :global(.pane-row > select) { box-sizing: border-box; width: 112px; height: 26px; font-size: inherit; }
  .pane-fields :global(.pane-rule) { width: 100%; margin: 2px 0; border: none; border-top: 1px solid var(--ok-border); }
  .pane-fields :global(.pane-radios) { border: 0; margin: 0; padding: 0; display: flex; flex-direction: column; gap: 2px; }
  .pane-fields :global(.pane-radios label) { display: flex; align-items: center; gap: 6px; min-height: 18px; font-size: 12px; }
  .pane-fields :global(.pane-radios input) { margin: 0; accent-color: var(--ok-accent); }
  .pane-fields :global(.pane-check) { display: flex; align-items: center; gap: 6px; min-height: 26px; }
  .pane-fields :global(.pane-check input) { margin: 0; }
  .pane-fields :global(.pane-color .trigger) { width: 39px; height: 26px; }
</style>
