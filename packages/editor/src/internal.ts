// UNSTABLE, NOT PUBLIC API. `@office-kit/pptx-editor/internal` and
// `@office-kit/pptx-editor/internal/animation-player` exist only for
// `@office-kit/pptx-dev` while that host still needs more than `mountEditor`
// offers (its save status and conflict prompts, draft recovery, the preview
// frame's slide show and Agents pane). They change or disappear in any release,
// without notice, and are removed once the dev host uses the public API alone.

import DevEditor from './dev/DevEditor.svelte';
import { mountInPage } from './mount.ts';

/**
 * Mounts the development host's editor, which loads, saves and merges the deck
 * through the `pptx-dev` server's `/editor/*` endpoints.
 * @internal
 */
export function mountDevEditor(target: HTMLElement): void {
  mountInPage(target, DevEditor, {});
}
