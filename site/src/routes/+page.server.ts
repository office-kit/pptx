import { getSlides } from '@office-kit/pptx';
import { renderSlideToSvg } from '@office-kit/pptx-preview';
import { buildHeroDeck } from '$lib/examples/hero-deck';
import heroDeckSource from '$lib/examples/hero-deck.ts?raw';
import templateFillSource from '$lib/examples/template-fill.ts?raw';
import { highlight } from '$lib/server/highlight';
import type { PageServerLoad } from './$types';

// Each example opens with a comment addressed to maintainers; visitors only
// need the code, so a snippet starts where the code does.
const from = (source: string, marker: string): string => source.slice(source.indexOf(marker));

export const load: PageServerLoad = async () => {
  // Rendering at prerender time puts the slide in the static HTML: it paints
  // with the page, needs no client JS, and cannot drift from the code beside it.
  const deck = buildHeroDeck();
  const slideSvg = renderSlideToSvg(deck, getSlides(deck)[0]!);

  const [heroCode, templateCode] = await Promise.all([
    highlight(from(heroDeckSource, 'export function'), 'ts'),
    highlight(from(templateFillSource, 'import '), 'ts'),
  ]);

  return { slideSvg, heroCode, templateCode };
};
