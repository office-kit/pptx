import ts from 'typescript';

/** Longest `before` / `after` a text edit accepts, in UTF-16 code units. */
export const TEXT_EDIT_MAX_LENGTH = 12000;

export interface SourceFile {
  path: string;
  source: string;
}

export type TextEditPlan =
  | { ok: true; change: { path: string; before: string; after: string } }
  | { ok: false; reason: 'invalid' | 'not-unique-on-slide' | 'not-found' | 'ambiguous' };

/** A `getShapeJsxSources` entry, with `fileName` given as the matching `SourceFile.path`. */
export interface TextEditAnchor {
  path: string;
  /** 1-based, as in `JsxSource`. */
  lineNumber: number;
  /** 1-based, as in `JsxSource`. */
  columnNumber: number;
}

export interface TextEditRequest {
  /** Only these files are searched; pass the slide sources, not shared components. */
  files: readonly SourceFile[];
  /** The rendered text of the slide being edited. `before` must occur in it exactly once. */
  slideText: string;
  before: string;
  after: string;
  /** Searches only the JSX element that starts here. */
  anchor?: TextEditAnchor;
}

// A JSX text node renders the way JSX collapses it: each line trimmed where
// it meets a line break, blank lines dropped, the rest joined by one space.
function jsxText(node: ts.JsxText): string {
  return node.text
    .split(/\r\n|\n|\r/)
    .map((line, index, lines) => {
      let value = line.replaceAll('\t', ' ');
      if (index) value = value.trimStart();
      if (index < lines.length - 1) value = value.trimEnd();
      return value;
    })
    .filter(Boolean)
    .join(' ');
}

function anchoredElement(source: ts.SourceFile, anchor: TextEditAnchor): ts.Node | undefined {
  const lines = source.getLineStarts();
  const lineStart = lines[anchor.lineNumber - 1];
  // A stale anchor may point past the file or past the end of its line.
  if (lineStart === undefined || anchor.columnNumber < 1) return undefined;
  const position = lineStart + anchor.columnNumber - 1;
  if (position >= (lines[anchor.lineNumber] ?? source.end)) return undefined;
  let found: ts.Node | undefined;
  const visit = (node: ts.Node) => {
    if (found || node.getStart(source) > position || node.end <= position) return;
    if (
      node.getStart(source) === position &&
      (ts.isJsxElement(node) || ts.isJsxSelfClosingElement(node) || ts.isJsxFragment(node))
    )
      found = node;
    else ts.forEachChild(node, visit);
  };
  visit(source);
  return found;
}

/**
 * Plans replacing the source literal that renders `before` with `after`.
 * Only a JSX text or a string literal written out whole qualifies; computed
 * text and text written in more than one place are refused, so the caller
 * can fall back to a broader edit.
 */
export function planTextEdit(request: TextEditRequest): TextEditPlan {
  const { files, slideText, before, after, anchor } = request;
  if (
    !before.trim() ||
    before.length > TEXT_EDIT_MAX_LENGTH ||
    after.length > TEXT_EDIT_MAX_LENGTH ||
    after === before
  )
    return { ok: false, reason: 'invalid' };
  if (slideText.split(before).length !== 2) return { ok: false, reason: 'not-unique-on-slide' };
  const changes: { path: string; before: string; after: string }[] = [];
  for (const file of files) {
    if (anchor && file.path !== anchor.path) continue;
    const source = ts.createSourceFile(
      file.path,
      file.source,
      ts.ScriptTarget.Latest,
      true,
      /[jt]sx$/.test(file.path) ? ts.ScriptKind.TSX : ts.ScriptKind.TS,
    );
    const root = anchor ? anchoredElement(source, anchor) : source;
    if (!root) continue;
    const visit = (node: ts.Node) => {
      const text = ts.isJsxText(node)
        ? jsxText(node)
        : ts.isStringLiteral(node) || ts.isNoSubstitutionTemplateLiteral(node)
          ? node.text
          : undefined;
      if (text === before) {
        // Attribute strings use JSX entity rules, not JavaScript string escapes.
        const jsx = ts.isJsxText(node) || ts.isJsxAttribute(node.parent);
        const replacement = jsx ? '{' + JSON.stringify(after) + '}' : JSON.stringify(after);
        const start = ts.isJsxText(node) ? node.pos : node.getStart(source);
        changes.push({
          path: file.path,
          before: file.source,
          after: file.source.slice(0, start) + replacement + file.source.slice(node.end),
        });
      }
      ts.forEachChild(node, visit);
    };
    visit(root);
  }
  if (changes.length === 0) return { ok: false, reason: 'not-found' };
  if (changes.length > 1) return { ok: false, reason: 'ambiguous' };
  return { ok: true, change: changes[0]! };
}

export type TextEditVerification =
  | { ok: true }
  | { ok: false; reason: 'slide-count-changed' | 'other-slides-changed' | 'text-mismatch' };

export interface PreviewSnapshot {
  /** One rendering per slide (SVG); compared for equality only. */
  slides: readonly string[];
  slideTexts: readonly string[];
}

/** Checks that a rebuilt deck differs from the previous one only by `before` → `after` on `slide`. */
export function verifyTextEdit(request: {
  previous: PreviewSnapshot;
  next: PreviewSnapshot;
  slide: number;
  before: string;
  after: string;
}): TextEditVerification {
  const { previous, next, slide, before, after } = request;
  if (next.slides.length !== previous.slides.length)
    return { ok: false, reason: 'slide-count-changed' };
  if (next.slides.some((svg, index) => index !== slide && svg !== previous.slides[index]))
    return { ok: false, reason: 'other-slides-changed' };
  // A function replacement: `after` may contain `$&` and friends.
  if (next.slideTexts[slide] !== previous.slideTexts[slide]?.replace(before, () => after))
    return { ok: false, reason: 'text-mismatch' };
  return { ok: true };
}
