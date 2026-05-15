// Lightweight invariant validator.
//
// Scope: catch common authoring mistakes without depending on an XSD
// engine. Runs in Node and the browser, so it's safe to expose as a
// public free function in `fn.ts`. XSD-level validation lives in
// `test/lib/expect-schema-valid.ts` and stays Node-only (xmllint
// subprocess).
//
// What this catches today:
//
//   - `/ppt/presentation.xml` or its `.rels` part missing.
//   - A `<p:sldId>` whose `r:id` has no matching rel.
//   - A slide rel whose target part is absent from the package.
//   - A slide that has no `slideLayout` rel.
//   - A slide layout that has no `slideMaster` rel.
//   - Duplicate `<p:sldId>` `id` or `r:id` within `<p:sldIdLst>`.
//
// What this doesn't catch (yet):
//
//   - XSD-level structural problems (use the Layer-1 test path for that).
//   - Theme / colorMap consistency.
//   - Chart / xlsx round-trip integrity.
//   - PowerPoint-vs-spec quirks (handled by the planned quirks module).

import {
  type PartName,
  type Relationship,
  partName,
  resolveTarget,
} from '../opc/index.ts';
import type { OpcPackage } from '../parts/index.ts';
import { REL_TYPES } from '../presentationml/index.ts';
import {
  NS,
  allChildElements,
  firstChildElement,
  getAttrValue,
  parseXml,
  qname,
} from '../xml/index.ts';

const PRES_PART = partName('/ppt/presentation.xml');
const PRES_RELS = partName('/ppt/_rels/presentation.xml.rels');

const NAME_SLD_ID_LST = qname('p', 'sldIdLst', NS.pml);
const NAME_SLD_ID = qname('p', 'sldId', NS.pml);
const ATTR_ID = qname('', 'id', '');
const ATTR_R_ID = qname('r', 'id', NS.officeDocRels);

export type IssueSeverity = 'error' | 'warning';

export interface ValidationIssue {
  readonly severity: IssueSeverity;
  readonly message: string;
  readonly partName?: PartName;
}

const decoder = new TextDecoder();

const decode = (b: Uint8Array): string => decoder.decode(b);

const targetPartName = (sourcePart: PartName, rel: Relationship): PartName =>
  rel.target.startsWith('/') ? partName(rel.target) : resolveTarget(sourcePart, rel.target);

/**
 * Runs every invariant check in this module against `pkg`. Returns a
 * list of issues; an empty list means the deck passes every check. The
 * function never throws — callers decide how to surface issues (some
 * may treat warnings as informational, others may want a strict gate).
 */
export const validatePresentationPackage = (pkg: OpcPackage): ValidationIssue[] => {
  const issues: ValidationIssue[] = [];

  const presPart = pkg.getPart(PRES_PART);
  if (!presPart) {
    issues.push({
      severity: 'error',
      message: '/ppt/presentation.xml is missing',
      partName: PRES_PART,
    });
    return issues;
  }
  const presRels = pkg.getRels(PRES_PART);
  if (!presRels) {
    issues.push({
      severity: 'error',
      message: '/ppt/presentation.xml has no .rels part',
      partName: PRES_RELS,
    });
    return issues;
  }

  let presDocRoot: ReturnType<typeof parseXml>['root'];
  try {
    presDocRoot = parseXml(decode(presPart.data)).root;
  } catch (e) {
    issues.push({
      severity: 'error',
      message: `presentation.xml failed to parse: ${(e as Error).message}`,
      partName: PRES_PART,
    });
    return issues;
  }

  // Walk <p:sldIdLst> for duplicate id / r:id collisions and resolve
  // every slide reference back into a real part.
  const sldIdLst = firstChildElement(presDocRoot, NAME_SLD_ID_LST);
  const seenIds = new Set<string>();
  const seenRIds = new Set<string>();
  const slidePartNames: PartName[] = [];

  if (sldIdLst !== null) {
    for (const sldId of allChildElements(sldIdLst, NAME_SLD_ID)) {
      const idRaw = getAttrValue(sldId, ATTR_ID);
      const rIdRaw = getAttrValue(sldId, ATTR_R_ID);
      if (idRaw === null) {
        issues.push({
          severity: 'error',
          message: '<p:sldId> missing id attribute',
          partName: PRES_PART,
        });
        continue;
      }
      if (rIdRaw === null) {
        issues.push({
          severity: 'error',
          message: `<p:sldId id="${idRaw}"> missing r:id`,
          partName: PRES_PART,
        });
        continue;
      }
      if (seenIds.has(idRaw)) {
        issues.push({
          severity: 'error',
          message: `duplicate <p:sldId id="${idRaw}">`,
          partName: PRES_PART,
        });
      } else {
        seenIds.add(idRaw);
      }
      if (seenRIds.has(rIdRaw)) {
        issues.push({
          severity: 'error',
          message: `duplicate r:id="${rIdRaw}" in <p:sldIdLst>`,
          partName: PRES_PART,
        });
      } else {
        seenRIds.add(rIdRaw);
      }

      const rel = presRels.items.find((r) => r.id === rIdRaw);
      if (!rel) {
        issues.push({
          severity: 'error',
          message: `<p:sldId r:id="${rIdRaw}"> has no matching rel`,
          partName: PRES_PART,
        });
        continue;
      }
      const slidePartName = targetPartName(PRES_PART, rel);
      if (pkg.getPart(slidePartName) === null) {
        issues.push({
          severity: 'error',
          message: `slide rel target ${slidePartName} not in package`,
          partName: slidePartName,
        });
        continue;
      }
      slidePartNames.push(slidePartName);
    }
  }

  // Each slide must have a layout rel; each layout a master rel.
  for (const slideName of slidePartNames) {
    const slideRels = pkg.getRels(slideName);
    if (!slideRels) {
      issues.push({
        severity: 'warning',
        message: `slide ${slideName} has no .rels part`,
        partName: slideName,
      });
      continue;
    }
    const layoutRel = slideRels.items.find((r) => r.type === REL_TYPES.slideLayout);
    if (!layoutRel) {
      issues.push({
        severity: 'error',
        message: `slide ${slideName} is missing a slideLayout rel`,
        partName: slideName,
      });
      continue;
    }
    const layoutName = targetPartName(slideName, layoutRel);
    if (pkg.getPart(layoutName) === null) {
      issues.push({
        severity: 'error',
        message: `slide ${slideName} references missing layout ${layoutName}`,
        partName: slideName,
      });
      continue;
    }
    const layoutRels = pkg.getRels(layoutName);
    if (!layoutRels) {
      issues.push({
        severity: 'warning',
        message: `slide layout ${layoutName} has no .rels part`,
        partName: layoutName,
      });
      continue;
    }
    const masterRel = layoutRels.items.find((r) => r.type === REL_TYPES.slideMaster);
    if (!masterRel) {
      issues.push({
        severity: 'error',
        message: `slide layout ${layoutName} is missing a slideMaster rel`,
        partName: layoutName,
      });
      continue;
    }
    const masterName = targetPartName(layoutName, masterRel);
    if (pkg.getPart(masterName) === null) {
      issues.push({
        severity: 'error',
        message: `slide layout ${layoutName} references missing master ${masterName}`,
        partName: layoutName,
      });
    }

    // Dangling media / chart / hyperlink rels.
    for (const rel of slideRels.items) {
      if (rel.targetMode === 'External') continue;
      if (
        rel.type !== REL_TYPES.image &&
        rel.type !== REL_TYPES.chart &&
        rel.type !== REL_TYPES.notesSlide &&
        rel.type !== REL_TYPES.comments &&
        rel.type !== REL_TYPES.oleObject &&
        rel.type !== REL_TYPES.package
      ) {
        continue;
      }
      const targetName = targetPartName(slideName, rel);
      if (pkg.getPart(targetName) === null) {
        issues.push({
          severity: 'error',
          message: `slide ${slideName} has dangling ${rel.type.split('/').pop()} rel → ${targetName}`,
          partName: slideName,
        });
      }
    }
  }

  // Chart parts must resolve to their embedded xlsx workbooks.
  for (const part of pkg.parts) {
    if (
      part.contentType !==
      'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'
    ) {
      continue;
    }
    const chartRels = pkg.getRels(part.name);
    if (!chartRels) {
      issues.push({
        severity: 'warning',
        message: `chart part ${part.name} has no .rels`,
        partName: part.name,
      });
      continue;
    }
    const xlsxRel = chartRels.items.find((r) => r.type === REL_TYPES.package);
    if (xlsxRel) {
      const xlsxName = targetPartName(part.name, xlsxRel);
      if (pkg.getPart(xlsxName) === null) {
        issues.push({
          severity: 'error',
          message: `chart ${part.name} references missing embedded workbook ${xlsxName}`,
          partName: part.name,
        });
      }
    }
  }

  return issues;
};

