import { loadPresentation } from '@office-kit/pptx';
import { mergeDecks, sameDeck, type DeckConflict } from '../merge/deck-merge.ts';
import type { PendingProposal, ProposalOrigin, ProposeOptions, ProposeResult } from './proposal.ts';
import { t } from '../i18n/i18n.svelte.ts';
import type { EditorController } from './controller.svelte.ts';
import type { EditorDocument } from './document.svelte.ts';
import { reconcileSelection } from './shape-ref.ts';

/** Resolves once no drag or resize is in progress, so a whole gesture is never split. */
function idle(doc: EditorDocument): Promise<void> {
  if (!doc.liveEditing) return Promise.resolve();
  return new Promise((resolve) => {
    const stop = $effect.root(() => {
      $effect(() => {
        if (doc.liveEditing) return;
        queueMicrotask(stop);
        resolve();
      });
    });
  });
}

function ask(
  editor: EditorController,
  from: ProposalOrigin,
  conflicts: readonly DeckConflict[],
  edited: Uint8Array,
): Promise<'mine' | 'theirs'> {
  return new Promise((resolve, reject) => {
    editor.proposal = {
      from,
      conflicts,
      edited,
      choose(choice) {
        editor.proposal = null;
        resolve(choice);
      },
      abandon(reason) {
        editor.proposal = null;
        reject(reason);
      },
    };
  });
}

/**
 * Merges `edited` (made from `base`) into the open presentation, which may
 * have changed since. Without a `base` there is nothing to merge against, so
 * any difference is the user's choice.
 */
export async function propose(
  editor: EditorController,
  base: Uint8Array | null,
  edited: Uint8Array,
  { label, from = 'agent' }: ProposeOptions,
): Promise<ProposeResult> {
  const doc = editor.doc;
  const stepLabel = from === 'agent' ? t('Agent: {label}').replace('{label}', () => label) : label;
  // The source is what the host has saved: a deck that ends up as exactly that
  // version is not an unsaved change.
  const replace = async (bytes: Uint8Array, version?: number): Promise<boolean> => {
    const pres = await loadPresentation(bytes);
    await idle(doc);
    if (version !== undefined && version !== doc.version) return false;
    doc.transact(
      stepLabel,
      () => {
        doc.pres = pres;
        doc.select(reconcileSelection(pres, doc.selection));
      },
      from,
      from === 'agent' || !sameDeck(bytes, edited),
    );
    return true;
  };
  for (;;) {
    await idle(doc);
    const version = doc.version;
    const ours = await doc.toBytes();
    // The user edited while the deck was serialized: merge with that edit too.
    if (version !== doc.version) continue;
    const merged = base
      ? mergeDecks(base, ours, edited)
      : sameDeck(ours, edited)
        ? ({ ok: true, bytes: ours } as const)
        : ({ ok: false, conflicts: [] } as const);
    if (merged.ok) {
      if (sameDeck(merged.bytes, ours)) return { status: 'applied' };
      if (await replace(merged.bytes, version)) return { status: 'applied' };
      continue;
    }
    const { conflicts } = merged;
    if ((await ask(editor, from, conflicts, edited)) === 'mine')
      return { status: 'kept-mine', conflicts };
    await replace(edited);
    return { status: 'took-theirs', conflicts };
  }
}
