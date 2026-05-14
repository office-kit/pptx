// Layer 1 of the testing strategy: XSD schema validation via `xmllint`.
//
// Strict, fast, and authoritative. If our emitted XML doesn't conform to
// the ECMA-376 schema, the test fails — no human judgment required.
//
// Why xmllint and not a JS validator: libxml2's schema engine handles the
// OOXML schema family's complex import graph correctly out of the box,
// while pure-JS validators tend to choke on cross-namespace `xs:include` /
// `xs:import`. xmllint is everywhere on macOS / Linux dev machines; CI
// installs it via `apt install libxml2-utils`.
//
// Tests that need this layer simply call `expectSchemaValid(xml, 'pml')`.
// Tests that need to run on machines without xmllint should guard with
// `isSchemaValidationAvailable()` and skip cleanly.

import { spawnSync } from 'node:child_process';
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { fileURLToPath } from 'node:url';
import { join } from 'node:path';

const SCHEMAS = {
  pml: 'ECMA-376/OfficeOpenXML-XMLSchema-Transitional/pml.xsd',
  dml: 'ECMA-376/OfficeOpenXML-XMLSchema-Transitional/dml-main.xsd',
  chart: 'ECMA-376/OfficeOpenXML-XMLSchema-Transitional/dml-chart.xsd',
  rels: 'ECMA-376/OpenPackagingConventions-XMLSchema/opc-relationships.xsd',
  contentTypes: 'ECMA-376/OpenPackagingConventions-XMLSchema/opc-contentTypes.xsd',
} as const;

export type SchemaKind = keyof typeof SCHEMAS;

// Repo-relative path of the ecma-376 submodule.
const ECMA_376_ROOT = fileURLToPath(new URL('../../references/ecma-376-5th/', import.meta.url));

const schemaPath = (kind: SchemaKind): string => join(ECMA_376_ROOT, SCHEMAS[kind]);

/**
 * Returns true if `xmllint` is on PATH. Tests that need schema validation
 * should skip when this is false rather than fail (developers without
 * libxml2 installed can still run the rest of the suite).
 */
export const isSchemaValidationAvailable = (): boolean => {
  const which = spawnSync('xmllint', ['--version'], { encoding: 'utf8', stdio: 'pipe' });
  return which.status === 0;
};

/**
 * Asserts `xml` validates against the ECMA-376 schema for `kind`. Throws
 * with the verbatim xmllint diagnostic on failure.
 */
export const expectSchemaValid = (xml: string, kind: SchemaKind): void => {
  if (!isSchemaValidationAvailable()) {
    throw new Error(
      'xmllint is not installed; install libxml2-utils to run schema-validation tests',
    );
  }
  const dir = mkdtempSync(join(tmpdir(), 'pptx-kit-schema-'));
  const file = join(dir, 'doc.xml');
  try {
    writeFileSync(file, xml, 'utf8');
    const r = spawnSync('xmllint', ['--noout', '--schema', schemaPath(kind), file], {
      encoding: 'utf8',
      stdio: 'pipe',
    });
    if (r.status !== 0) {
      const out = `${r.stdout}\n${r.stderr}`.trim();
      throw new Error(`schema validation failed (${kind}):\n${out}`);
    }
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
};
