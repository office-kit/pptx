/**
 * Who made a change: the person using the editor; an agent, through
 * `EditorHandle.apply` or `propose`; or a newer version of the presentation's
 * source, through `propose` with `from: 'source'`.
 */
export type ChangeSource = 'user' | 'agent' | 'source';
