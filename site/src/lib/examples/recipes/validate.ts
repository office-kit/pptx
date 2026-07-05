// Run the invariant checker before saving — useful in CI to catch
// dangling rels / missing parts / off-spec IDs.

import { type PresentationData, validatePresentation } from '@office-kit/pptx';

declare const pres: PresentationData;

const issues: ReturnType<typeof validatePresentation> = validatePresentation(pres);
for (const issue of issues) {
  console.error(`[${issue.severity}] ${issue.message}`);
}
if (issues.some((i) => i.severity === 'error')) {
  throw new Error(`@office-kit/pptx: ${issues.length} validation issue(s)`);
}
