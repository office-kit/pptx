// UNSTABLE, NOT PUBLIC API. `@office-kit/pptx-editor/internal` and
// `@office-kit/pptx-editor/internal/animation-player` exist only for
// `@office-kit/pptx-dev`. They change or disappear in any release, without
// notice. Everything else the dev host needs goes through `mountEditor`.

import { mountEditor, type EditorHandle, type EditorOptions } from './index.ts';
import { previewFrameOptions } from './core/preview-frame.ts';

/**
 * `mountEditor` for the frame inside the `pptx-dev` preview page. Slide Show,
 * Reading View, Rehearse Timings and the Agents pane button talk to that page
 * over `postMessage`, and the editor reports the slide being edited to it.
 * @internal
 */
export function mountPreviewFrameEditor(target: HTMLElement, options: EditorOptions): EditorHandle {
  previewFrameOptions.add(options);
  return mountEditor(target, options);
}
