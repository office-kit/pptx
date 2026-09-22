import { join } from 'node:path';
import { pathToFileURL } from 'node:url';
import { compile, type Node } from '@office-kit/pptx-dsl';
import {
  getSlideAnimations,
  getSlideSize,
  getSlides,
  getSlideNotes,
  isSlideHidden,
  getSlideTransition,
  loadPresentation,
  savePresentation,
  validatePresentation,
} from '@office-kit/pptx';
import { renderPreview, type PreviewCache } from './preview-cache.ts';

export interface BuildResult {
  bytes: Uint8Array;
  slides: string[];
  slideTexts: string[];
  notes: (string | null)[];
  hiddenSlides: boolean[];
  transitions: ReturnType<typeof getSlideTransition>[];
  /** What each slide animates, in click order — the preview's player reads this. */
  animations: ReturnType<typeof getSlideAnimations>[];
  aspectRatio: number;
  dependencies: string[];
  diagnostics: ReturnType<typeof validatePresentation>;
}
/** Evaluates trusted local TSX. This is code execution, not a sandbox. */
export async function buildDeck(
  directory: string,
  dependencies: string[],
  previous?: PreviewCache,
): Promise<{ result: BuildResult; cache: PreviewCache }> {
  const output = join(directory, 'deck.mjs');
  const module: { default?: Node } = await import(pathToFileURL(output).href);
  if (!module.default) throw new Error('The TSX file must default-export a Presentation.');
  const presentation = await compile(module.default);
  const diagnostics = validatePresentation(presentation);
  const errors = diagnostics.filter((issue) => issue.severity === 'error');
  if (errors.length) throw new Error(`Invalid presentation: ${JSON.stringify(errors)}`);
  const bytes = await savePresentation(presentation);
  // Preview serialized output too, so persistence defects are visible during authoring.
  return renderDeck(bytes, dependencies, previous);
}

export async function renderDeck(
  bytes: Uint8Array,
  dependencies: string[],
  previous?: PreviewCache,
): Promise<{ result: BuildResult; cache: PreviewCache }> {
  const saved = await loadPresentation(bytes);
  const diagnostics = validatePresentation(saved);
  const errors = diagnostics.filter((issue) => issue.severity === 'error');
  if (errors.length) throw new Error(`Invalid presentation: ${JSON.stringify(errors)}`);
  const size = getSlideSize(saved);
  const { slides, slideTexts, cache } = renderPreview(saved, bytes, previous);
  return {
    cache,
    result: {
      bytes,
      aspectRatio: size ? size.width / size.height : 16 / 9,
      slides,
      dependencies,
      slideTexts,
      notes: getSlides(saved).map(getSlideNotes),
      hiddenSlides: getSlides(saved).map(isSlideHidden),
      transitions: getSlides(saved).map(getSlideTransition),
      animations: getSlides(saved).map(getSlideAnimations),
      diagnostics,
    },
  };
}
