// `@office-kit/pptx-editor/merge`: the three-way merge `EditorHandle.propose`
// uses, for hosts that merge presentations elsewhere (for example on a
// server). It has no DOM dependency and runs in Node as well as the browser.

export {
  comparablePart,
  describeConflict,
  mergeDecks,
  type DeckConflict,
  type DeckMerge,
} from './merge/deck-merge.ts';
