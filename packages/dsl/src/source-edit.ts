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

/** DSL elements that create a shape; only their attributes are edited in place. */
export const SHAPE_ELEMENTS = [
  'Text',
  'Shape',
  'Image',
  'Line',
  'Table',
  'Chart',
  'Media',
  'Group',
] as const;

/**
 * The OOXML wire unit behind each numeric prop. Geometry props are inches that
 * `inches()` rounds to whole EMU (`ST_Coordinate`); `rotation` is degrees the
 * core rounds to whole 60000ths of a degree (`ST_Angle`).
 */
export const PROP_WIRE_UNITS = {
  x: 914400,
  y: 914400,
  width: 914400,
  height: 914400,
  rotation: 60000,
} as const;
export type NumericProp = keyof typeof PROP_WIRE_UNITS;

export interface PropEdit {
  prop: NumericProp;
  /** The wire value the current source produces (whole EMU or 60000ths of a degree). */
  from: number;
  /** The wire value the edited deck has. */
  to: number;
}

export interface PropEditRequest {
  file: SourceFile;
  /** The element that created the shape: the innermost `getShapeJsxSources` entry. */
  anchor: TextEditAnchor;
  edits: readonly PropEdit[];
}

export type PropEditPlan =
  | { ok: true; change: { path: string; before: string; after: string } }
  | {
      ok: false;
      reason: 'invalid' | 'not-element' | 'not-shape-element' | 'spread' | 'not-numeric';
    };

// Six decimals of an inch and five of a degree always pin a wire value; the
// search returns the shortest decimal that rounds to it.
const MAX_DECIMALS = 10;
function wireDecimal(units: number, wire: number): string {
  for (let digits = 0; digits <= MAX_DECIMALS; digits++) {
    const text = String(Number((wire / units).toFixed(digits)));
    if (Math.round(Number(text) * units) === wire) return text;
  }
  return String(wire / units);
}
// A delta added to an expression shifts its wire value by exactly `delta`
// units once the sum is rounded, so the decimal only has to land within a
// tiny fraction of a unit rather than round to it.
function deltaDecimal(units: number, delta: number): string {
  for (let digits = 0; digits <= MAX_DECIMALS; digits++) {
    const text = String(Number((delta / units).toFixed(digits)));
    if (Math.abs(Number(text) * units - delta) < 1e-6) return text;
  }
  return String(delta / units);
}
// Degrees wrap, so a turn is written the short way round: 350° → 10° is −20°.
const FULL_TURN = 360 * 60000;
function wrappedDelta(prop: NumericProp, from: number, to: number): number {
  const delta = to - from;
  if (prop !== 'rotation') return delta;
  const turned = ((delta % FULL_TURN) + FULL_TURN) % FULL_TURN;
  return turned > FULL_TURN / 2 ? turned - FULL_TURN : turned;
}

function numericLiteral(node: ts.Expression): boolean {
  return (
    ts.isNumericLiteral(node) ||
    (ts.isPrefixUnaryExpression(node) &&
      node.operator === ts.SyntaxKind.MinusToken &&
      ts.isNumericLiteral(node.operand))
  );
}
// Expressions that bind tighter than `+`, so `expr + d` needs no parentheses.
function bindsTighterThanAdd(node: ts.Expression): boolean {
  if (ts.isBinaryExpression(node)) {
    const op = node.operatorToken.kind;
    return (
      op === ts.SyntaxKind.AsteriskToken ||
      op === ts.SyntaxKind.SlashToken ||
      op === ts.SyntaxKind.PercentToken ||
      op === ts.SyntaxKind.AsteriskAsteriskToken
    );
  }
  return !(
    ts.isConditionalExpression(node) ||
    ts.isArrowFunction(node) ||
    ts.isAsExpression(node) ||
    ts.isSatisfiesExpression(node) ||
    ts.isYieldExpression(node) ||
    ts.isAwaitExpression(node) ||
    ts.isPrefixUnaryExpression(node) ||
    ts.isTypeOfExpression(node) ||
    ts.isVoidExpression(node) ||
    ts.isDeleteExpression(node)
  );
}

function rewriteExpression(
  source: ts.SourceFile,
  expression: ts.Expression,
  edit: PropEdit,
): string {
  const units = PROP_WIRE_UNITS[edit.prop];
  if (numericLiteral(expression)) return wireDecimal(units, edit.to);
  const delta = wrappedDelta(edit.prop, edit.from, edit.to);
  // `expr ± n` from an earlier edit: fold into the same trailing literal.
  if (
    ts.isBinaryExpression(expression) &&
    (expression.operatorToken.kind === ts.SyntaxKind.PlusToken ||
      expression.operatorToken.kind === ts.SyntaxKind.MinusToken) &&
    ts.isNumericLiteral(expression.right)
  ) {
    const sign = expression.operatorToken.kind === ts.SyntaxKind.PlusToken ? 1 : -1;
    const previous = Math.round(sign * Number(expression.right.text) * units);
    const total = previous + delta;
    const left = expression.left.getText(source);
    if (total === 0) return left;
    return `${left} ${total > 0 ? '+' : '-'} ${deltaDecimal(units, Math.abs(total))}`;
  }
  const text = expression.getText(source);
  const base = bindsTighterThanAdd(expression) ? text : `(${text})`;
  return `${base} ${delta > 0 ? '+' : '-'} ${deltaDecimal(units, Math.abs(delta))}`;
}

/**
 * Plans writing edited numeric props back onto the JSX element that created a
 * shape. A literal is replaced by the shortest decimal that reproduces the
 * wire value; any other expression keeps its meaning and gains a delta
 * (`x={col * 2.5}` → `x={col * 2.5 + 0.3}`), folding into an earlier delta. A
 * missing prop is inserted. Only this element changes, never a shared value.
 */
export function planPropEdit(request: PropEditRequest): PropEditPlan {
  const { file, anchor, edits } = request;
  if (
    edits.length === 0 ||
    edits.some(
      (edit) =>
        !Object.hasOwn(PROP_WIRE_UNITS, edit.prop) ||
        !Number.isInteger(edit.from) ||
        !Number.isInteger(edit.to),
    ) ||
    new Set(edits.map((edit) => edit.prop)).size !== edits.length
  )
    return { ok: false, reason: 'invalid' };
  const source = ts.createSourceFile(
    file.path,
    file.source,
    ts.ScriptTarget.Latest,
    true,
    /[jt]sx$/.test(file.path) ? ts.ScriptKind.TSX : ts.ScriptKind.TS,
  );
  const element = anchoredElement(source, anchor);
  if (!element || !(ts.isJsxElement(element) || ts.isJsxSelfClosingElement(element)))
    return { ok: false, reason: 'not-element' };
  const opening = ts.isJsxElement(element) ? element.openingElement : element;
  if (
    !ts.isIdentifier(opening.tagName) ||
    !(SHAPE_ELEMENTS as readonly string[]).includes(opening.tagName.text)
  )
    return { ok: false, reason: 'not-shape-element' };
  const attributes = opening.attributes.properties;
  const replacements: { start: number; end: number; text: string }[] = [];
  const inserted: string[] = [];
  for (const edit of edits) {
    if (edit.from === edit.to) continue;
    const attribute = attributes.find(
      (item): item is ts.JsxAttribute =>
        ts.isJsxAttribute(item) && item.name.getText(source) === edit.prop,
    );
    if (!attribute) {
      // A spread could be what sets the prop now; an inserted attribute would
      // silently win over it or be overridden by it depending on order.
      if (attributes.some(ts.isJsxSpreadAttribute)) return { ok: false, reason: 'spread' };
      inserted.push(`${edit.prop}={${wireDecimal(PROP_WIRE_UNITS[edit.prop], edit.to)}}`);
      continue;
    }
    const initializer = attribute.initializer;
    if (!initializer || !ts.isJsxExpression(initializer) || !initializer.expression)
      return { ok: false, reason: 'not-numeric' };
    replacements.push({
      start: initializer.expression.getStart(source),
      end: initializer.expression.end,
      text: rewriteExpression(source, initializer.expression, edit),
    });
  }
  if (replacements.length === 0 && inserted.length === 0) return { ok: false, reason: 'invalid' };
  if (inserted.length > 0) {
    const at = attributes.end;
    replacements.push({ start: at, end: at, text: ' ' + inserted.join(' ') });
  }
  let after = file.source;
  for (const { start, end, text } of replacements.sort((a, b) => b.start - a.start))
    after = after.slice(0, start) + text + after.slice(end);
  return { ok: true, change: { path: file.path, before: file.source, after } };
}
