import { loadPresentation } from '@office-kit/pptx';
import { auditTextLayout, type TextAuditIssue } from '@office-kit/pptx-preview';
import { buildFontkitMeasurer } from '@office-kit/pptx-preview/node';
import { buildDeck } from './index.ts';

/** The smallest size a footnote should use on a slide. */
export const DEFAULT_MIN_FONT_SIZE_PT = 10;

export interface CheckFinding {
  /** Errors break the slide; warnings are line breaks worth a look. */
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
    reportLineBreaks: true,
    minFontSizePt,
  });
  return issues.map((issue) => ({
    severity: issue.kind === 'word-break' || issue.kind === 'short-last-line' ? 'warning' : 'error',
    message: `Slide ${issue.slideIndex + 1}, ${issue.shapeName ? `"${issue.shapeName}"` : 'unnamed shape'}: ${describe(issue, minFontSizePt)}${issue.approximate ? ' (estimated width)' : ''}`,
  }));
}

function describe(issue: TextAuditIssue, minFontSizePt: number): string {
  switch (issue.kind) {
    case 'overflow-x':
      return `text runs ${issue.overflowPx}px past the sides of its box`;
    case 'overflow-y':
      return `text runs ${issue.overflowPx}px past the top or bottom of its box`;
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
