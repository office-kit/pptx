// Package introspection escape hatches.
import { getSlides, isSlideHidden } from './slide-query.ts';

import { partName } from '../../internal/opc/index.ts';
import type { OpcPackage } from '../../internal/parts/index.ts';
import { type ShapeKind } from '../../internal/presentationml/index.ts';
import { NS } from '../../internal/xml/index.ts';
import {
  INTERNAL_PACKAGE,
  type PresentationData,
  SHAPE_SNAPSHOT,
  SLIDE_DOCUMENT,
  SLIDE_PART_NAME,
  SLIDE_SHAPES,
  type SlideData,
} from '../_internal-symbols.ts';
import { getPresentationTheme } from './theme.ts';
import { getSlideLayouts } from './layouts.ts';
import { getSlideSections } from './sections.ts';

// ---------------------------------------------------------------------------
// Package introspection escape hatches.
//
// `_internalPackageOf` is the heavy escape hatch for hot-path power
// users; these two helpers cover the 80% case (just enumerate parts
// or read a single part's bytes) without exposing the OpcPackage
// class.

/**
 * Power-user escape hatch. Returns the underlying `OpcPackage`
 * backing `pres`. Use this when you need to manipulate parts /
 * rels directly. Most callers should use the typed helpers above
 * (`listPackageParts`, `readPackagePart`, `getMediaParts`, etc.).
 *
 * @internal — used by `@office-kit/pptx/node` to mount fs-backed helpers.
 */
export const _internalPackageOf = (pres: PresentationData): OpcPackage => pres[INTERNAL_PACKAGE];

/** One entry in the package's parts list. */
export interface PackagePartInfo {
  readonly name: string;
  readonly contentType: string;
  readonly byteLength: number;
}

/**
 * Enumerates every OPC part in the package. Useful for advanced
 * inspection (e.g. "what parts does this template carry?") without
 * dropping to `_internalPackageOf`.
 */
export const listPackageParts = (pres: PresentationData): ReadonlyArray<PackagePartInfo> =>
  pres[INTERNAL_PACKAGE].parts.map((p) => ({
    name: p.name,
    contentType: p.contentType,
    byteLength: p.data.byteLength,
  }));

/**
 * Reads a single OPC part's bytes by part name (e.g.
 * `'/ppt/slides/slide1.xml'`). Returns `null` when no such part
 * exists. The returned `Uint8Array` is a live view into the
 * package — DO NOT mutate it. Use this for read-only inspection
 * (e.g. parsing custom extension parts).
 */
export const readPackagePart = (pres: PresentationData, name: string): Uint8Array | null => {
  // OPC part names compare case-insensitively (ECMA-376 Part 2 §9.1.1), and
  // OpcPackage.getPart already honors that — match it so a caller passing
  // `/ppt/Media/Image1.PNG` resolves the same part the package stores.
  const target = name.toLowerCase();
  const part = pres[INTERNAL_PACKAGE].parts.find((p) => p.name.toLowerCase() === target);
  return part?.data ?? null;
};

/** A media (image / video / audio) part embedded in the package. */
export interface MediaPart {
  readonly name: string;
  readonly contentType: string;
  readonly data: Uint8Array;
}

/**
 * Returns the total size of the package's parts in bytes
 * (uncompressed). Useful for storage estimation, quota checks,
 * and "how big is this deck before save?" diagnostics. The
 * actual `savePresentation` output is typically smaller after
 * DEFLATE; this is an upper bound.
 */
export const getPackageSize = (pres: PresentationData): number => {
  let total = 0;
  for (const part of pres[INTERNAL_PACKAGE].parts) total += part.data.byteLength;
  return total;
};

/**
 * Returns every `/ppt/media/...` part in the package. Useful for
 * audit / export workflows — e.g. "extract every embedded image."
 */
export const getMediaParts = (pres: PresentationData): ReadonlyArray<MediaPart> => {
  const out: MediaPart[] = [];
  for (const p of pres[INTERNAL_PACKAGE].parts) {
    if (p.name.startsWith('/ppt/media/')) {
      out.push({ name: p.name, contentType: p.contentType, data: p.data });
    }
  }
  return out;
};

/**
 * Returns every media part NOT referenced by any rels in the
 * package — the set `compactPackage` would remove. Non-destructive;
 * the caller decides whether to delete.
 *
 * Useful for audit UIs that want to surface bloat before cleaning,
 * and for "is this asset still used?" checks.
 */
export const getOrphanMediaPartNames = (pres: PresentationData): ReadonlyArray<string> => {
  const pkg = pres[INTERNAL_PACKAGE];
  const referenced = new Set<string>();
  const resolve = (sourcePart: string, target: string): string => {
    if (target.startsWith('/')) return target;
    const dir = sourcePart.split('/').slice(0, -1);
    const segments: string[] = [];
    for (const seg of [...dir, ...target.split('/')]) {
      if (seg === '..') segments.pop();
      else if (seg !== '.' && seg.length > 0) segments.push(seg);
    }
    return `/${segments.join('/')}`;
  };
  for (const part of pkg.parts) {
    if (!part.name.endsWith('.rels')) continue;
    // /ppt/slides/_rels/slide1.xml.rels → /ppt/slides/slide1.xml
    const m = part.name.match(/^(.*)\/_rels\/(.+)\.rels$/);
    let sourceName: string;
    if (part.name === '/_rels/.rels') {
      sourceName = '/';
    } else if (m) {
      sourceName = `${m[1]}/${m[2]}`;
    } else {
      continue;
    }
    const sourceRels = sourceName === '/' ? pkg.rootRels() : pkg.getRels(sourceName as never);
    if (!sourceRels) continue;
    for (const rel of sourceRels.items) {
      if (rel.targetMode === 'External') continue;
      // Lowercase keys: OPC part names compare case-insensitively, so a rel
      // target whose case differs from the media part name still counts as a
      // reference (otherwise a referenced image is wrongly reported as orphan /
      // deleted by compactPackage).
      referenced.add(resolve(sourceName, rel.target).toLowerCase());
    }
  }
  const out: string[] = [];
  for (const part of pkg.parts) {
    if (!part.name.startsWith('/ppt/media/')) continue;
    if (!referenced.has(part.name.toLowerCase())) out.push(part.name);
  }
  return out;
};

/**
 * Returns every media part name the slide's rels reference
 * (typically `/ppt/media/imageN.ext`). Walks the slide's rels
 * graph and resolves each internal target. Useful for "which
 * media files does this slide depend on?" audits.
 */
export const getSlideMediaPartNames = (slide: SlideData): ReadonlyArray<string> => {
  const pkg = slide[INTERNAL_PACKAGE];
  const rels = pkg.getRels(slide[SLIDE_PART_NAME]);
  if (!rels) return [];
  const resolve = (sourcePart: string, target: string): string => {
    if (target.startsWith('/')) return target;
    const dir = sourcePart.split('/').slice(0, -1);
    const segments: string[] = [];
    for (const seg of [...dir, ...target.split('/')]) {
      if (seg === '..') segments.pop();
      else if (seg !== '.' && seg.length > 0) segments.push(seg);
    }
    return `/${segments.join('/')}`;
  };
  const out: string[] = [];
  const seen = new Set<string>();
  for (const rel of rels.items) {
    if (rel.targetMode === 'External') continue;
    const resolved = resolve(slide[SLIDE_PART_NAME], rel.target);
    if (!resolved.startsWith('/ppt/media/')) continue;
    if (seen.has(resolved)) continue;
    seen.add(resolved);
    out.push(resolved);
  }
  return out;
};

/**
 * Returns every slide that references the given media part name
 * (typically `/ppt/media/imageN.ext`). Walks each slide's rels and
 * checks whether any internal rel resolves to `mediaPartName`.
 *
 * Useful for image-audit workflows: "before I replace this image,
 * which slides will the change affect?"
 */
export const slidesUsingMediaPart = (
  pres: PresentationData,
  mediaPartName: string,
): ReadonlyArray<SlideData> => {
  const pkg = pres[INTERNAL_PACKAGE];
  const resolve = (sourcePart: string, target: string): string => {
    if (target.startsWith('/')) return target;
    const dir = sourcePart.split('/').slice(0, -1);
    const segments: string[] = [];
    for (const seg of [...dir, ...target.split('/')]) {
      if (seg === '..') segments.pop();
      else if (seg !== '.' && seg.length > 0) segments.push(seg);
    }
    return `/${segments.join('/')}`;
  };
  const out: SlideData[] = [];
  for (const slide of getSlides(pres)) {
    const rels = pkg.getRels(slide[SLIDE_PART_NAME]);
    if (!rels) continue;
    const hit = rels.items.some(
      (r) =>
        r.targetMode !== 'External' && resolve(slide[SLIDE_PART_NAME], r.target) === mediaPartName,
    );
    if (hit) out.push(slide);
  }
  return out;
};

/**
 * Removes media parts that no rels graph references. Returns the
 * list of removed part names. Useful after a sequence of slide
 * removals leaves orphan images behind.
 *
 * Only `/ppt/media/...` parts are considered. The check walks every
 * `.rels` part in the package and resolves each internal rel target
 * against its source part name to build the live media set.
 */
export const compactPackage = (
  pres: PresentationData,
): { readonly removed: ReadonlyArray<string> } => {
  const pkg = pres[INTERNAL_PACKAGE];
  const referenced = new Set<string>();

  const resolve = (sourcePart: string, target: string): string => {
    if (target.startsWith('/')) return target;
    const dir = sourcePart.split('/').slice(0, -1);
    const segments: string[] = [];
    for (const seg of [...dir, ...target.split('/')]) {
      if (seg === '..') segments.pop();
      else if (seg !== '.' && seg.length > 0) segments.push(seg);
    }
    return `/${segments.join('/')}`;
  };

  for (const part of pkg.parts) {
    if (!part.name.endsWith('.rels')) continue;
    // /ppt/slides/_rels/slide1.xml.rels → /ppt/slides/slide1.xml
    // /_rels/.rels                       → / (package root)
    let sourceName = part.name.replace('/_rels/', '/').replace(/\.rels$/, '');
    if (sourceName === '/' || sourceName === '') {
      // Root rels — `rel.target` is relative to the package root.
      // We don't need to consult pkg.getRels for it (the only thing it
      // points at that we care about is the presentation.xml, which
      // has its own rels we'll walk). Just parse the part data
      // directly for completeness.
      sourceName = '/';
    }
    const rels = sourceName === '/' ? null : pkg.getRels(partName(sourceName));
    if (!rels) continue;
    for (const rel of rels.items) {
      if (rel.targetMode === 'External') continue;
      // Lowercase keys: OPC part names compare case-insensitively, so a rel
      // target whose case differs from the media part name still counts as a
      // reference (otherwise a referenced image is wrongly reported as orphan /
      // deleted by compactPackage).
      referenced.add(resolve(sourceName, rel.target).toLowerCase());
    }
  }

  const removed: string[] = [];
  const orphans: string[] = [];
  for (const part of pkg.parts) {
    if (!part.name.startsWith('/ppt/media/')) continue;
    if (!referenced.has(part.name.toLowerCase())) orphans.push(part.name);
  }
  for (const name of orphans) {
    pkg.removePart(partName(name));
    removed.push(name);
  }
  return { removed };
};

/**
 * Replaces the bytes of a media part in place. Returns `true` if the
 * part was found and updated, `false` otherwise. The content type is
 * preserved.
 *
 * Useful for the "swap every instance of this logo" workflow — pick
 * the right `partName` via `getMediaParts` and call this once. Every
 * `<a:blip r:embed="…"/>` reference is unaffected because the rels
 * already point at this part name.
 */
export const setMediaPartBytes = (
  pres: PresentationData,
  partName: string,
  bytes: Uint8Array,
): boolean => {
  // Case-insensitive match, consistent with OPC part-name semantics and
  // OpcPackage.getPart (see readPackagePart).
  const target = partName.toLowerCase();
  const part = pres[INTERNAL_PACKAGE].parts.find((p) => p.name.toLowerCase() === target);
  if (!part) return false;
  part.data = bytes;
  return true;
};

/**
 * High-level snapshot of the presentation's structure. Useful as a
 * diagnostic checklist when debugging a template or generating audit
 * reports. The numbers reflect what's reachable through the typed
 * API on the current package state.
 */
export interface PresentationSummary {
  readonly slideCount: number;
  readonly hiddenSlideCount: number;
  readonly totalShapes: number;
  readonly shapesByKind: Readonly<Record<ShapeKind, number>>;
  readonly layoutCount: number;
  readonly sectionCount: number;
  readonly partCount: number;
  readonly hasCharts: boolean;
  readonly hasComments: boolean;
  readonly hasAnimations: boolean;
  readonly themeName: string | null;
}

const CHART_CONTENT_TYPE_FN = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml';
const COMMENTS_CONTENT_TYPE_FN =
  'application/vnd.openxmlformats-officedocument.presentationml.comments+xml';

export const getPresentationSummary = (pres: PresentationData): PresentationSummary => {
  const pkg = pres[INTERNAL_PACKAGE];
  const slides = getSlides(pres);
  let hiddenSlideCount = 0;
  let totalShapes = 0;
  const shapesByKind: Record<ShapeKind, number> = {
    shape: 0,
    picture: 0,
    group: 0,
    graphicFrame: 0,
    connector: 0,
  };
  let hasAnimations = false;
  for (const slide of slides) {
    if (isSlideHidden(slide)) hiddenSlideCount++;
    for (const s of slide[SLIDE_SHAPES]) {
      totalShapes++;
      shapesByKind[s[SHAPE_SNAPSHOT].kind]++;
    }
    // <p:timing> presence = at least one animation.
    if (
      !hasAnimations &&
      slide[SLIDE_DOCUMENT].root.children.some(
        (c) =>
          c.kind === 'element' && c.name.namespaceURI === NS.pml && c.name.localName === 'timing',
      )
    ) {
      hasAnimations = true;
    }
  }

  const hasCharts = pkg.parts.some((p) => p.contentType === CHART_CONTENT_TYPE_FN);
  const hasComments = pkg.parts.some((p) => p.contentType === COMMENTS_CONTENT_TYPE_FN);
  const theme = getPresentationTheme(pres);

  return {
    slideCount: slides.length,
    hiddenSlideCount,
    totalShapes,
    shapesByKind,
    layoutCount: getSlideLayouts(pres).length,
    sectionCount: getSlideSections(pres).length,
    partCount: pkg.parts.length,
    hasCharts,
    hasComments,
    hasAnimations,
    themeName: theme?.name ?? null,
  };
};
