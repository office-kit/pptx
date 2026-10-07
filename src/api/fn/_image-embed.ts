// Embeds raster pictures for fills that reference them by relationship id.

import type { ImageEmbedder } from '../../internal/drawingml/index.ts';
import {
  contentTypeForFormat,
  detectImageFormat,
  emptyRels,
  extensionForFormat,
  type ImageFormat,
  nextRelId,
  type PartName,
  partName,
  resolveTarget,
} from '../../internal/opc/index.ts';
import type { OpcPackage } from '../../internal/parts/index.ts';
import { REL_TYPES } from '../../internal/presentationml/index.ts';
import { setOpcDefault } from './_helpers.ts';

const MEDIA_IMAGE = /^\/ppt\/media\/image(\d+)\./;

/**
 * Adds `bytes` as the next `/ppt/media/imageN.<ext>` part and a relationship
 * from `owner` to it, returning the relationship id. `owner` lives one folder
 * below `/ppt` (a slide or notes slide), hence the `../media` target.
 */
export const embedImagePart = (
  pkg: OpcPackage,
  owner: PartName,
  bytes: Uint8Array,
  format: ImageFormat,
): string => {
  let next = 1;
  for (const part of pkg.parts) {
    const match = part.name.match(MEDIA_IMAGE);
    if (match?.[1] !== undefined) next = Math.max(next, Number.parseInt(match[1], 10) + 1);
  }
  const extension = extensionForFormat(format);
  const contentType = contentTypeForFormat(format);
  setOpcDefault(pkg, extension, contentType);
  pkg.addPart(partName(`/ppt/media/image${next}.${extension}`), contentType, bytes);
  const rels = pkg.getRels(owner) ?? emptyRels();
  const id = nextRelId(rels.items.map((rel) => rel.id));
  rels.items.push({
    id,
    type: REL_TYPES.image,
    target: `../media/image${next}.${extension}`,
    targetMode: 'Internal',
  });
  pkg.setRels(owner, rels);
  return id;
};

/**
 * Resolves an internal picture relationship of `owner` to its bytes, for the
 * run readers' picture fills. `null` for a missing or external target.
 */
export const relatedImageBytes =
  (pkg: OpcPackage, owner: PartName) =>
  (relationshipId: string): Uint8Array | null => {
    const rel = pkg.getRels(owner)?.items.find((item) => item.id === relationshipId);
    if (!rel || rel.targetMode === 'External') return null;
    const target = rel.target.startsWith('/')
      ? partName(rel.target)
      : resolveTarget(owner, rel.target);
    return pkg.getPart(target)?.data ?? null;
  };

/**
 * The picture embedder a text edit on `owner` hands to the run formatter. One
 * edit that applies the same bytes to many runs embeds them once.
 */
export const createImageEmbedder = (
  pkg: OpcPackage,
  owner: PartName,
  caller: string,
): ImageEmbedder => {
  const embedded = new Map<Uint8Array, string>();
  const rasterFormat = (bytes: Uint8Array): ImageFormat => {
    const format = detectImageFormat(bytes);
    // An SVG blip needs a raster fallback PowerPoint renders instead; a
    // picture text fill has nowhere to keep one.
    if (format === null || format === 'svg')
      throw new Error(`${caller}: picture fill bytes must be PNG, JPEG, GIF, BMP, TIFF or WebP`);
    return format;
  };
  return {
    check: (bytes) => {
      rasterFormat(bytes);
    },
    embed: (bytes) => {
      let id = embedded.get(bytes);
      if (id === undefined) {
        id = embedImagePart(pkg, owner, bytes, rasterFormat(bytes));
        embedded.set(bytes, id);
      }
      return id;
    },
  };
};
