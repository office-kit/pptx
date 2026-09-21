import { createHash } from 'node:crypto';
import { unzipSync } from 'fflate';
import { getSlidePartName, getSlides, getSlideText, type PresentationData } from '@office-kit/pptx';
import { renderSlideToSvg } from '@office-kit/pptx-preview';

export interface PreviewCache {
  shared: string;
  slides: Record<string, { hash: string; svg: string; text: string }>;
}

/** Reuse SVG only when both slide XML and every shared rendering resource match. */
export function renderPreview(
  presentation: PresentationData,
  bytes: Uint8Array,
  previous?: PreviewCache,
) {
  const slides = getSlides(presentation);
  const names = slides.map((slide) => getSlidePartName(slide).replace(/^\//, ''));
  const slideParts = new Set(names);
  const parts = unzipSync(bytes);
  const hashes = new Map(
    Object.keys(parts)
      .sort()
      .map((name) => [name, createHash('sha256').update(parts[name]!).digest('hex')]),
  );
  // Relationships, themes, layouts, charts and media all invalidate the cache.
  // Document metadata changes on each save but cannot affect slide rendering.
  const shared = createHash('sha256')
    .update(
      JSON.stringify(
        [...hashes].filter(([name]) => !slideParts.has(name) && !name.startsWith('docProps/')),
      ),
    )
    .digest('hex');
  const cache: PreviewCache = { shared, slides: {} };
  const rendered = slides.map((slide, index) => {
    const name = names[index]!;
    const hash = hashes.get(name)!;
    const cached = previous?.shared === shared ? previous.slides[name] : undefined;
    const result =
      cached?.hash === hash
        ? cached
        : { hash, svg: renderSlideToSvg(presentation, slide), text: getSlideText(slide) };
    cache.slides[name] = result;
    return result;
  });
  return {
    slides: rendered.map((slide) => slide.svg),
    slideTexts: rendered.map((slide) => slide.text),
    cache,
  };
}
