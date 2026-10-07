import type { DeckConflict } from '../merge/deck-merge.ts';

/**
 * Where a proposed version comes from: an agent that worked on a copy, or the
 * presentation's own source (the file changed elsewhere, or was rebuilt).
 */
export type ProposalOrigin = 'agent' | 'source';

/** How `EditorHandle.propose` ended. */
export type ProposeResult =
  /** The proposal merged with the open presentation (or changed nothing). */
  | { readonly status: 'applied' }
  /** It collided with the user's edits, and the user kept theirs; nothing of it was applied. */
  | { readonly status: 'kept-mine'; readonly conflicts: readonly DeckConflict[] }
  /** It collided with the user's edits, and the user replaced theirs with it (one undo step). */
  | { readonly status: 'took-theirs'; readonly conflicts: readonly DeckConflict[] };

export interface ProposeOptions {
  /** Names the undo step: "Agent: `label`" for an agent, `label` itself for the source. */
  readonly label: string;
  /** Defaults to `'agent'`. */
  readonly from?: ProposalOrigin;
}

/** A proposal that collided with the user's edits; the title bar asks which to keep. */
export interface PendingProposal {
  readonly from: ProposalOrigin;
  readonly conflicts: readonly DeckConflict[];
  /** The proposed version, offered for download before choosing. */
  readonly edited: Uint8Array;
  choose(choice: 'mine' | 'theirs'): void;
  /** Ends the proposal without a choice: the editor is going away. */
  abandon(reason: Error): void;
}
