import { loadPresentation } from '@office-kit/pptx';
import { auditTextLayout, type TextAuditIssue } from '@office-kit/pptx-preview';
import { buildFontkitMeasurer } from '@office-kit/pptx-preview/node';
import { buildDeck } from './index.ts';

/** The smallest size a footnote should use on a slide. */
export const DEFAULT_MIN_FONT_SIZE_PT = 10;
const OVERFLOW_TOLERANCE_PX = 1;

export interface CheckFinding {
  /** Errors are text leaving its shape; warnings need a look (spills into the insets, estimated widths, line breaks, small type). */
  severity: 'error' | 'warning';
  message: string;
}

/** Builds the deck and measures its text: overflow, split words, stranded last lines, small type. */
export async function checkDeck(
  entry: string,
  minFontSizePt = DEFAULT_MIN_FONT_SIZE_PT,
): Promise<CheckFinding[]> {
  const { bytes } = await buildDeck(entry);
  const issues = auditTextLayout(await loadPresentation(bytes), {
    measureText: buildFontkitMeasurer(),
    reportBadWraps: true,
    minFontSizePt,
  });
  return issues.map((issue) => ({
    severity: isError(issue) ? 'error' : 'warning',
    message: `Slide ${issue.slideIndex + 1}, ${issue.shapeName ? `"${issue.shapeName}"` : 'unnamed shape'}: ${describe(issue, minFontSizePt)}${issue.approximate ? ' (estimated width)' : ''}`,
  }));
}

// Only text leaving the shape fails: a spill into the insets is invisible. An
// estimated width can be off by about a line (a misplaced wrap, or a line's
// worth of glyph error), so beyond that the overflow is real too (Japanese text
// is always estimated).
function isError(issue: TextAuditIssue): boolean {
  if (issue.kind !== 'overflow-x' && issue.kind !== 'overflow-y') return false;
  return issue.shapeOverflowPx > (issue.approximate ? issue.lineHeightPx : OVERFLOW_TOLERANCE_PX);
}

function describe(issue: TextAuditIssue, minFontSizePt: number): string {
  switch (issue.kind) {
    case 'overflow-x':
      return issue.shapeOverflowPx > 0
        ? `text runs ${issue.shapeOverflowPx}px past the sides of its shape`
        : `text spills ${issue.overflowPx}px into the side margins of its box`;
    case 'overflow-y':
      return issue.shapeOverflowPx > 0
        ? `text runs ${issue.shapeOverflowPx}px past the top or bottom of its shape`
        : `text spills ${issue.overflowPx}px into the top or bottom margin of its box`;
    // Not requested by checkDeck; listed so the switch stays exhaustive.
    case 'soft-wrap':
      return `paragraph ${issue.paragraphIndex + 1} wraps onto ${issue.extraLines} more line(s)`;
    case 'word-break':
      return `paragraph ${issue.paragraphIndex + 1} breaks the word "${issue.text}" across lines`;
    case 'short-last-line':
      return `paragraph ${issue.paragraphIndex + 1} ends on a line of only "${issue.text}"`;
    case 'small-font':
      return `text at ${issue.sizePt}pt is below ${minFontSizePt}pt`;
  }
}
