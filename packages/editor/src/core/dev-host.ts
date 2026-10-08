/** What `@office-kit/pptx-editor/internal` adds to a mount for `@office-kit/pptx-dev`. */
export interface DevHostOptions {
  /**
   * The editor runs in the preview's frame, whose page presents slide shows
   * and hosts the Agents pane over `postMessage`.
   */
  readonly previewFrame?: boolean;
  /**
   * Shows Help ▸ Feedback, opening `url` in a new tab. A host that embeds the
   * editor owns its own support channel, so the public `mountEditor` shows no
   * feedback entry point and the editor bundle carries no feedback URL.
   */
  readonly feedback?: { readonly url: string };
}

/** The dev-host options of each internal mount, keyed by its `EditorOptions`. */
export const devHostOptions: WeakMap<object, DevHostOptions> = new WeakMap();
