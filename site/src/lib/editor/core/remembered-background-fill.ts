import { asColor, type PresentationData, type SlideData } from '@office-kit/pptx';
import { readSlideBackground } from './slide-background.ts';
import { patterns } from '../panels/patterns.ts';
import type { RememberedFill } from './remembered-fill.ts';

/** Mac PowerPoint retains background settings when switching between fill types. */
export function rememberBackgroundFill(
  pres: PresentationData,
  slide: SlideData,
  remembered: RememberedFill,
): void {
  const { fill, gradient, pattern } = readSlideBackground(pres, slide, { preserveTheme: true });
  if (fill.kind === 'solid')
    remembered.solid = { color: asColor(fill.color) ?? '#FFFFFF', opacity: fill.opacity };
  if (fill.kind === 'pattern' && pattern) {
    const preset = patterns.find(([value]) => value === pattern.preset)?.[0];
    const foreground = asColor(pattern.foreground);
    const background = asColor(pattern.background);
    remembered.pattern =
      preset && foreground && background ? { preset, foreground, background } : undefined;
  }
  if (fill.kind === 'gradient' && gradient)
    remembered.gradient = {
      ...gradient,
      stops: gradient.stops.map((stop) => {
        const color = asColor(stop.color);
        return {
          ...stop,
          color: color ?? asColor(stop.resolvedColor ?? '') ?? 'accent1',
          brightness: color ? stop.brightness : 0,
          colorTransforms: color ? stop.colorTransforms : undefined,
        };
      }),
    };
}
