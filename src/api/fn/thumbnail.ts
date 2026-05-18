// Document thumbnail.

import {
  type ImageFormat,
  type PartName,
  contentTypeForFormat,
  detectImageFormat,
  emptyRels,
  extensionForFormat,
  nextRelId,
  partName,
} from '../../internal/opc/index.ts';
import type { OpcPackage } from '../../internal/parts/index.ts';
import { REL_TYPES } from '../../internal/presentationml/index.ts';
import { INTERNAL_PACKAGE, type PresentationData } from '../_internal-symbols.ts';
import { setOpcDefault } from './_helpers.ts';

// ---------------------------------------------------------------------------
// Document thumbnail (`/docProps/thumbnail.jpeg` typically).

/**
 * The package's thumbnail image, when present. PowerPoint, the OS
 * file picker, and SharePoint preview decks via this. Format is
 * what's encoded in the thumbnail part — usually JPEG.
 */
export interface PresentationThumbnail {
  readonly format: ImageFormat;
  readonly bytes: Uint8Array;
}

const findThumbnailRel = (pkg: OpcPackage): { partName: PartName; rId: string } | null => {
  const rootRels = pkg.rootRels();
  if (!rootRels) return null;
  const rel = rootRels.items.find((r) => r.type === REL_TYPES.thumbnail);
  if (!rel) return null;
  const target = rel.target;
  const name = target.startsWith('/') ? partName(target) : partName(`/${target}`);
  return { partName: name, rId: rel.id };
};

/**
 * Returns the package's thumbnail bytes, plus the detected image
 * format. Returns `null` when the package has no thumbnail rel or
 * the rel target part is missing.
 *
 * The returned `bytes` is a live view into the thumbnail part —
 * treat it as read-only; copy if you need an independent buffer.
 */
export const getThumbnail = (pres: PresentationData): PresentationThumbnail | null => {
  const pkg = pres[INTERNAL_PACKAGE];
  const hit = findThumbnailRel(pkg);
  if (!hit) return null;
  const part = pkg.getPart(hit.partName);
  if (!part) return null;
  const format = detectImageFormat(part.data);
  if (format === null) return null;
  return { format, bytes: part.data };
};

/**
 * Replaces the package's thumbnail. Auto-detects the image format
 * from the bytes (pass `options.format` to override). Bootstraps the
 * thumbnail part + root rel + content-type default if the package
 * had no thumbnail; otherwise replaces the existing thumbnail in
 * place, switching its filename / extension if the format changed.
 */
export const setThumbnail = (
  pres: PresentationData,
  bytes: Uint8Array,
  options: { format?: ImageFormat } = {},
): void => {
  const format = options.format ?? detectImageFormat(bytes);
  if (format === null) {
    throw new Error('setThumbnail: could not detect image format. Pass options.format explicitly.');
  }
  const pkg = pres[INTERNAL_PACKAGE];
  const contentType = contentTypeForFormat(format);
  const extension = extensionForFormat(format);
  const desiredName = partName(`/docProps/thumbnail.${extension}`);

  const hit = findThumbnailRel(pkg);
  if (hit) {
    // Replace in place if the existing part is the same path; otherwise
    // remove the old one and add a new one with the right extension.
    if (hit.partName === desiredName) {
      const part = pkg.getPart(hit.partName);
      if (!part) throw new Error(`thumbnail rel points at missing part ${hit.partName}`);
      part.data = bytes;
      part.contentType = contentType;
      setOpcDefault(pkg, extension, contentType);
      return;
    }
    pkg.removePart(hit.partName);
    pkg.addPart(desiredName, contentType, bytes);
    setOpcDefault(pkg, extension, contentType);
    const rootRels = pkg.rootRels() ?? emptyRels();
    const existing = rootRels.items.find((r) => r.id === hit.rId);
    if (existing) {
      existing.target = `docProps/thumbnail.${extension}`;
    }
    pkg.setRootRels(rootRels);
    return;
  }

  // Bootstrap.
  setOpcDefault(pkg, extension, contentType);
  pkg.addPart(desiredName, contentType, bytes);
  const rootRels = pkg.rootRels() ?? emptyRels();
  const rId = nextRelId(rootRels.items.map((r) => r.id));
  rootRels.items.push({
    id: rId,
    type: REL_TYPES.thumbnail,
    target: `docProps/thumbnail.${extension}`,
    targetMode: 'Internal',
  });
  pkg.setRootRels(rootRels);
};

/**
 * Removes the package's thumbnail entirely (drops the rel + the
 * underlying part). No-op when the package has no thumbnail.
 */
export const removeThumbnail = (pres: PresentationData): void => {
  const pkg = pres[INTERNAL_PACKAGE];
  const hit = findThumbnailRel(pkg);
  if (!hit) return;
  pkg.removePart(hit.partName);
  const rootRels = pkg.rootRels();
  if (!rootRels) return;
  rootRels.items = rootRels.items.filter((r) => r.id !== hit.rId);
  pkg.setRootRels(rootRels);
};
