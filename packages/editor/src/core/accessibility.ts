import {
  getShapeDescription,
  getShapeId,
  getShapeKind,
  getShapeName,
  getSlides,
  getSlideShapes,
  getSlideTitle,
  type PresentationData,
} from '@office-kit/pptx';

export type AccessibilityIssue =
  | { kind: 'alt-text'; slide: number; shapeId: number; name: string }
  | { kind: 'slide-title'; slide: number };

// PowerPoint's Accessibility Checker flags pictures, charts, tables, media and
// groups without alternative text, and slides without a title. Text boxes and
// plain shapes read their own text, so they are not flagged.
const NEEDS_ALT_TEXT = new Set(['picture', 'graphicFrame', 'group']);

export function accessibilityIssues(pres: PresentationData): AccessibilityIssue[] {
  const issues: AccessibilityIssue[] = [];
  getSlides(pres).forEach((slide, index) => {
    if (!getSlideTitle(slide)?.trim()) issues.push({ kind: 'slide-title', slide: index });
    for (const shape of getSlideShapes(slide)) {
      if (NEEDS_ALT_TEXT.has(getShapeKind(shape)) && !getShapeDescription(shape)?.trim())
        issues.push({
          kind: 'alt-text',
          slide: index,
          shapeId: getShapeId(shape),
          name: getShapeName(shape),
        });
    }
  });
  return issues;
}
