// Package validation.

import {
  type IssueSeverity,
  type ValidationIssue,
  validatePresentationPackage,
} from '../../internal/validator/index.ts';
import { INTERNAL_PACKAGE, type PresentationData } from '../_internal-symbols.ts';
// ---------------------------------------------------------------------------
// Validator.

export type { IssueSeverity, ValidationIssue };

/**
 * Runs a set of lightweight invariant checks on the package and
 * returns the list of issues found. An empty array means the deck
 * passes every check.
 *
 * Catches the common authoring mistakes — missing presentation.xml,
 * dangling slide rels, slides without a layout, etc. — without
 * depending on a heavyweight XSD engine, so it runs identically in
 * Node and the browser.
 *
 * Use it as a pre-save sanity check, especially after orchestrating
 * lots of mutations against the same package. Higher-fidelity XSD
 * validation lives in the test harness (Layer 1) and stays Node-only.
 */
export const validatePresentation = (pres: PresentationData): ReadonlyArray<ValidationIssue> =>
  validatePresentationPackage(pres[INTERNAL_PACKAGE]);
