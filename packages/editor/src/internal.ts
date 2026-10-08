// UNSTABLE, NOT PUBLIC API. `@office-kit/pptx-editor/internal` and
// `@office-kit/pptx-editor/internal/animation-player` exist only for
// `@office-kit/pptx-dev`. They change or disappear in any release, without
// notice. Everything else the dev host needs goes through `mountEditor`.

import { mountEditor, type EditorHandle, type EditorOptions } from './index.ts';
import { devHostOptions, type DevHostOptions } from './core/dev-host.ts';

export type { DevHostOptions } from './core/dev-host.ts';

/**
 * `mountEditor` for `pptx-dev`'s editor page. Inside the preview's frame
 * (`previewFrame`), Slide Show, Reading View, Rehearse Timings and the Agents
 * pane button talk to that page over `postMessage`, and the editor reports the
 * slide being edited to it. `feedback` adds Help ▸ Feedback, which hosts of the
 * public `mountEditor` never get.
 * @internal
 */
export function mountDevEditor(
  target: HTMLElement,
  options: EditorOptions,
  dev: DevHostOptions,
): EditorHandle {
  devHostOptions.set(options, dev);
  return mountEditor(target, options);
}
