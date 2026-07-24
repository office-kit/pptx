// `@office-kit/pptx-preview` — browser-safe entry.
//
// Renders a @office-kit/pptx slide model to an SVG string. This module pulls in NO
// Node built-ins (no `node:fs`, no resvg, no fontkit), so it bundles for the
// browser. For Node rasterization to PNG/RGBA, import `@office-kit/pptx-preview/node`.
//
// `renderSlideToSvg` is the canonical preview entry. By default it lays text
// out with `<foreignObject>` (the browser measures + wraps), which is correct
// in a browser but invisible to headless rasterizers. Pass
// `{ textLayout: 'svg', measureText }` to get pure `<text>` output that resvg
// (or any SVG rasterizer) can paint — this is what the `/node` entry does.

export { renderSlideSvg as renderSlideToSvg } from './render-slide.ts';

// Text-layout audit — overflow (はみ出し) and soft-wrap (段落ち) detection,
// measured with the same layout engine the preview renders with.
export { auditTextLayout } from './audit.ts';
export type { AuditTextLayoutOptions, TextAuditIssue } from './audit.ts';

export type {
  RenderSlideOptions,
  TextLayoutMode,
  TextMeasurer,
  FontSpec,
  MeasureResult,
} from './text-layout.ts';

// The substitute family names a custom `measureText` must key off, plus the
// default (heuristic) measurer and the family-substitution map. Exposed so a
// caller supplying their own measurer can resolve the same families the
// emitter writes into the SVG.
export {
  substituteFamily,
  defaultMeasurer,
  SANS,
  SERIF,
  ARIAL,
  TIMES,
  MONO,
} from './text-layout.ts';
