import {
  editorModel,
  restoreEditorSession,
  readHistory,
  replayEdits,
  type EditorModel,
  type EditHistory,
} from './editor.ts';
import { join } from 'node:path';
import { pathToFileURL } from 'node:url';
import { compile, type Node } from '@office-kit/pptx-dsl';
import {
  getSlideSize,
  loadPresentation,
  savePresentation,
  validatePresentation,
} from '@office-kit/pptx';
import { renderPreview, type PreviewCache } from './preview-cache.ts';

export interface BuildResult {
  bytes: Uint8Array;
  editor: EditorModel;
  history: EditHistory;
  slides: string[];
  slideTexts: string[];
  aspectRatio: number;
  dependencies: string[];
  diagnostics: ReturnType<typeof validatePresentation>;
}
/** Evaluates trusted local TSX. This is code execution, not a sandbox. */
export async function buildDeck(
  directory: string,
  dependencies: string[],
  previous?: PreviewCache,
  entry?: string,
): Promise<{ result: BuildResult; cache: PreviewCache }> {
  const output = join(directory, 'deck.mjs');
  const module: { default?: Node } = await import(pathToFileURL(output).href);
  if (!module.default) throw new Error('The TSX file must default-export a Presentation.');
  const compiled = await compile(module.default);
  const presentation = await loadPresentation(await savePresentation(compiled));
  const history = entry
    ? await readHistory(entry)
    : { version: 1 as const, cursor: 0, entries: [] };
  replayEdits(presentation, history);
  const diagnostics = validatePresentation(presentation);
  const errors = diagnostics.filter((issue) => issue.severity === 'error');
  if (errors.length) throw new Error(`Invalid presentation: ${JSON.stringify(errors)}`);
  const bytes = await savePresentation(presentation);
  // Preview serialized output too, so persistence defects are visible during authoring.
  const saved = await loadPresentation(bytes);
  restoreEditorSession(saved, editorModel(presentation));
  const size = getSlideSize(saved);
  const { slides, slideTexts, cache } = renderPreview(saved, bytes, previous);
  return {
    cache,
    result: {
      bytes,
      editor: editorModel(saved),
      history,
      aspectRatio: size ? size.width / size.height : 16 / 9,
      slides,
      dependencies,
      slideTexts,
      diagnostics,
    },
  };
}
