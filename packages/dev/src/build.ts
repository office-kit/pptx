import { join } from 'node:path';
import { pathToFileURL } from 'node:url';
import { compile, getShapeJsxSources, type JsxSource, type Node } from '@office-kit/pptx-dsl';
import {
  getGroupChildren,
  getShapeId,
  getSlideShapes,
  type PresentationData,
  type SlideShapeData,
  getSlideAnimations,
  getCustomShows,
  getSlideShowProperties,
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
import { getPreviewMedia, type PreviewMedia } from './media-manifest.ts';

export interface BuildResult {
  bytes: Uint8Array;
  slides: string[];
  slideTexts: string[];
  notes: (string | null)[];
  hiddenSlides: boolean[];
  transitions: ReturnType<typeof getSlideTransition>[];
  /** What each slide animates, in click order — the preview's player reads this. */
  animations: ReturnType<typeof getSlideAnimations>[];
  showProperties: ReturnType<typeof getSlideShowProperties>;
  customShows: { id: number; name: string; slideIndices: number[] }[];
  /** Media clips keyed by slide and shape id for HTML playback in preview mode. */
  media: PreviewMedia[];
  aspectRatio: number;
  dependencies: string[];
  /**
   * Slide index → shape id → the JSX elements being evaluated when the shape
   * was made, outermost first. Only decks built from TSX have it; a saved
   * editor deck has no source to point at.
   */
  shapeSources?: Record<number, Record<number, readonly JsxSource[]>>;
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
  const shapeSources = jsxSourcesBySlide(presentation);
  const bytes = await savePresentation(presentation);
  // Preview serialized output too, so persistence defects are visible during authoring.
  const rendered = await renderDeck(bytes, dependencies, previous);
  return { ...rendered, result: { ...rendered.result, shapeSources } };
}

// Shape ids survive saving, so sources read before `savePresentation` still
// name the shapes of the saved deck the editor loads.
function jsxSourcesBySlide(
  presentation: PresentationData,
): Record<number, Record<number, readonly JsxSource[]>> {
  const result: Record<number, Record<number, readonly JsxSource[]>> = {};
  getSlides(presentation).forEach((slide, index) => {
    const byId: Record<number, readonly JsxSource[]> = {};
    const visit = (shape: SlideShapeData) => {
      const sources = getShapeJsxSources(shape);
      if (sources) byId[getShapeId(shape)] = sources;
      for (const child of getGroupChildren(shape)) visit(child);
    };
    for (const shape of getSlideShapes(slide)) visit(shape);
    result[index] = byId;
  });
  return result;
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
  const deckSlides = getSlides(saved);
  const slideIndices = new Map(deckSlides.map((slide, index) => [slide, index]));
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
      showProperties: getSlideShowProperties(saved),
      customShows: getCustomShows(saved).map((show) => ({
        id: show.id,
        name: show.name,
        slideIndices: show.slides.map((slide) => slideIndices.get(slide)!),
      })),
      media: getPreviewMedia(saved),
      diagnostics,
    },
  };
}
