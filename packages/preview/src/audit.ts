// Text-layout audit: measures every shape's laid-out text against its box and
// reports where ink escapes (overflow) or paragraphs soft-wrap (段落ち).
//
// The measurement runs through the SAME pipeline the preview renders with —
// `resolveTextBodyModel` (effective bodyPr / autofit / paragraph model) into
// `layoutCore` (wrap + baseline placement) — so a reported overflow is exactly
// what the rasterized preview would paint outside the box.
//
// Browser-safe like text-layout.ts: the measurer is injected. The default
// heuristic measurer estimates widths per character (results are flagged
// `approximate`); pass `buildFontkitMeasurer()` from
// `@office-kit/pptx-preview/node` for glyph-accurate metrics.

import {
  getPresentationTheme,
  getShapeBoundsResolved,
  getShapeKind,
  getShapeName,
  getShapePlaceholderType,
  getShapeTextDirection,
  getSlides,
  getSlideShapes,
  type PresentationData,
  type PresentationTheme,
  type SlideShapeData,
} from '@office-kit/pptx';
import {
  buildSvgTextInput,
  EMU_PER_PX,
  resolveTextBodyModel,
  verticalLayoutOf,
} from './render-slide.ts';
import {
  defaultMeasurer,
  layoutCore,
  SANS,
  type ColumnLayout,
  type FontSpec,
  type TextMeasurer,
} from './text-layout.ts';

interface IssueBase {
  /** 0-based deck position of the slide the shape sits on. */
  readonly slideIndex: number;
  readonly shapeName: string | null;
  /** True when any run was measured by estimate rather than real glyph
   *  metrics (heuristic measurer, or a font without the needed glyphs).
   *  Treat borderline overflows as advisory when set. */
  readonly approximate: boolean;
}

export type TextAuditIssue =
  | (IssueBase & {
      readonly kind: 'overflow-x' | 'overflow-y';
      /** How far the text ink escapes the box, in px at 96 DPI. */
      readonly overflowPx: number;
    })
  | (IssueBase & {
      readonly kind: 'soft-wrap';
      readonly paragraphIndex: number;
      /** Lines beyond the paragraph's authored line count (1 + explicit
       *  breaks) — i.e. how many times the text wrapped on its own. */
      readonly extraLines: number;
    })
  | (IssueBase & {
      /** A wrap split a word: Latin / Greek / Cyrillic letters or digits, or a katakana run. */
      readonly kind: 'word-break';
      readonly paragraphIndex: number;
      /** The split word, with `|` where the line breaks. */
      readonly text: string;
    })
  | (IssueBase & {
      /** A wrapped paragraph whose last line holds only a character or two (泣き別れ). */
      readonly kind: 'short-last-line';
      readonly paragraphIndex: number;
      /** The text of that last line. */
      readonly text: string;
    })
  | (IssueBase & {
      readonly kind: 'small-font';
      /** The smallest font size drawn in the shape, after autofit shrinking. */
      readonly sizePt: number;
    });

export interface AuditTextLayoutOptions {
  /** Measurer for run widths / vertical metrics. Defaults to the heuristic
   *  measurer (browser-safe, approximate); pass the fontkit measurer from
   *  `@office-kit/pptx-preview/node` for real glyph metrics. */
  readonly measureText?: TextMeasurer;
  /** Overflow at or below this many px (96 DPI) is ignored. Default 1 —
   *  measurement and the reference desktop app disagree by sub-pixel amounts routinely. */
  readonly tolerancePx?: number;
  /** Also report paragraphs that wrap onto more lines than their explicit
   *  breaks author (段落ち). Off by default: wrapping is normal for body
   *  text, so this is opt-in for single-line-intent content like titles. */
  readonly reportSoftWraps?: boolean;
  /** Also report wraps that split a word (`word-break`) and wrapped
   *  paragraphs that end on a line of one or two characters
   *  (`short-last-line`). Horizontal text only. */
  readonly reportLineBreaks?: boolean;
  /** Report shapes drawing text below this size in points (`small-font`). */
  readonly minFontSizePt?: number;
}

const DEFAULT_TOLERANCE_PX = 1;
const PT_PER_PX = 72 / 96;
const SHORT_LAST_LINE_MAX_CHARS = 2;
const WORD_CHAR = /[\p{Script=Latin}\p{Script=Greek}\p{Script=Cyrillic}\p{Nd}]/u;
// U+30FC (the prolonged sound mark) is Script=Common but belongs to the katakana word.
const KATAKANA_CHAR = /[\p{Script=Katakana}\u30FC]/u;
const WORD_RUN =
  /[\p{Script=Latin}\p{Script=Greek}\p{Script=Cyrillic}\p{Nd}\p{Script=Katakana}\u30FC]*$/u;
const WORD_RUN_START =
  /^[\p{Script=Latin}\p{Script=Greek}\p{Script=Cyrillic}\p{Nd}\p{Script=Katakana}\u30FC]*/u;

const splitsWord = (before: string, after: string): boolean =>
  (WORD_CHAR.test(before) && WORD_CHAR.test(after)) ||
  (KATAKANA_CHAR.test(before) && KATAKANA_CHAR.test(after));

const round2 = (n: number): number => Math.round(n * 100) / 100;

// getSlideShapes already flattens group descendants into the list, so a
// plain filter reaches every auditable shape exactly once — recursing into
// groups on top of that would audit each child twice. Children are audited
// in their own (child) coordinate space, where box and text agree; a group's
// non-uniform scale stretches both equally in the rendered output, so
// overflow verdicts hold.
function* walkShapes(shapes: ReadonlyArray<SlideShapeData>): Generator<SlideShapeData> {
  for (const shape of shapes) {
    if (getShapeKind(shape) !== 'group') yield shape;
  }
}

// The audit hands the measurer the AUTHORED font name (falling back to the
// generic sans substitute), unlike the render path's `substituteFamily` —
// a fontkit measurer with user-registered fonts resolves it first and only
// then falls back to the same substitution map.
const passthroughFamily = (family: string | null): string => family ?? SANS;

interface ResolvedOptions {
  readonly tolerancePx: number;
  readonly reportSoftWraps: boolean;
  readonly reportLineBreaks: boolean;
  readonly minFontSizePt: number | undefined;
}

const auditShape = (
  pres: PresentationData,
  theme: PresentationTheme | null,
  shape: SlideShapeData,
  slideIndex: number,
  measure: TextMeasurer,
  options: ResolvedOptions,
  issues: TextAuditIssue[],
): void => {
  const { tolerancePx, reportSoftWraps, reportLineBreaks, minFontSizePt } = options;
  const kind = getShapeKind(shape);
  if (kind !== 'shape' && kind !== 'graphicFrame') return;
  const bounds = getShapeBoundsResolved(pres, shape);
  if (!bounds) return;
  const model = resolveTextBodyModel(
    pres,
    shape,
    { x: bounds.x as number, y: bounds.y as number, w: bounds.w as number, h: bounds.h as number },
    theme,
    getShapePlaceholderType(shape),
    measure,
    '#000000', // colors don't affect metrics
    false,
    passthroughFamily,
  );
  if (model === null) return;

  const vert = verticalLayoutOf(model.effectiveBody.vert ?? getShapeTextDirection(shape));
  const cols = model.effectiveBody.columns;
  const columns: ColumnLayout | null =
    vert === 'none' && cols && cols.count >= 2
      ? {
          count: cols.count,
          gapPx: cols.gapEmu !== undefined ? cols.gapEmu / EMU_PER_PX : 0,
        }
      : null;
  const rect = model.svgTextRect(vert);
  if (rect.w <= 0 || rect.h <= 0) return;

  // SVG semantics (see the render path's `svgScale`): only an authored
  // autofit shrinks; the heuristic factor is foreignObject-only.
  const autoFitScale = model.authoredAutofit ? model.autoFitScale : 1;
  const input = buildSvgTextInput({
    pres,
    shape,
    theme,
    paraData: model.paraData,
    numberLabels: model.numberLabels,
    autoFitScale,
    lineHeightScale: model.lineHeightScale,
    defaultPt: model.defaultPt,
    themeFace: model.themeFace,
    defaultColor: '#000000',
    anchor: model.anchor,
    anchorCentered: model.effectiveBody.anchorCentered ?? false,
    wrap: model.effectiveBody.wrap !== 'none',
    innerX: rect.x,
    innerY: rect.y,
    innerW: rect.w,
    innerH: rect.h,
    measure,
    vert,
    columns,
    resolveFamily: passthroughFamily,
  });
  const core = layoutCore(input, measure);

  // Mirror layoutCore's frame: the ±90° rotations lay out into a frame with
  // swapped extents re-centred on the box, so ink must be compared against
  // that frame, not the box itself.
  const rotated = vert === 'cw90' || vert === 'cw270';
  const frame = rotated
    ? {
        x: input.boxXpx + input.boxWpx / 2 - input.boxHpx / 2,
        y: input.boxYpx + input.boxHpx / 2 - input.boxWpx / 2,
        w: input.boxHpx,
        h: input.boxWpx,
      }
    : { x: input.boxXpx, y: input.boxYpx, w: input.boxWpx, h: input.boxHpx };

  // Ink extents over every placed line. Empty lines (blank paragraphs, lines
  // of spaces) paint nothing — they influence later lines' positions but are
  // not themselves visible overflow.
  let inkTop = Infinity;
  let inkBottom = -Infinity;
  let overflowX = 0;
  let anyInk = false;
  for (const p of core.placements) {
    const toks = [...p.line.tokens];
    while (toks.length > 0 && (toks[toks.length - 1]!.isSpace || toks[toks.length - 1]!.isBreak)) {
      toks.pop();
    }
    if (!toks.some((t) => !t.isSpace && !t.isBreak)) continue;
    anyInk = true;
    let lineW = 0;
    for (const t of toks) {
      if (!t.isBreak) lineW += t.width;
    }
    const left =
      p.line.textAnchor === 'middle'
        ? p.line.anchorX - lineW / 2
        : p.line.textAnchor === 'end'
          ? p.line.anchorX - lineW
          : p.line.anchorX;
    const over = Math.max(frame.x - (left + p.dx), left + p.dx + lineW - (frame.x + frame.w));
    if (over > overflowX) overflowX = over;
    inkTop = Math.min(inkTop, p.baselineY - p.line.ascent);
    inkBottom = Math.max(inkBottom, p.baselineY + p.line.descent);
  }
  if (!anyInk) return;

  const approximate = detectApproximate(input.paragraphs, measure);
  const slideRef = { slideIndex, shapeName: getShapeName(shape), approximate };

  if (overflowX > tolerancePx) {
    issues.push({ ...slideRef, kind: 'overflow-x', overflowPx: round2(overflowX) });
  }
  const overflowY = Math.max(frame.y - inkTop, inkBottom - (frame.y + frame.h));
  if (overflowY > tolerancePx) {
    issues.push({ ...slideRef, kind: 'overflow-y', overflowPx: round2(overflowY) });
  }

  // 段落ち — a paragraph occupying more lines than 1 + its explicit <a:br>
  // count wrapped on its own. Meaningless for `upright` (one glyph per line
  // by construction), so horizontal text only.
  if (reportSoftWraps && vert === 'none') {
    const lineCounts = new Map<number, number>();
    for (const p of core.placements) {
      lineCounts.set(p.line.paraIndex, (lineCounts.get(p.line.paraIndex) ?? 0) + 1);
    }
    input.paragraphs.forEach((para, pi) => {
      const breaks = para.pieces.reduce((n, pc) => n + (pc.isBreak ? 1 : 0), 0);
      const extraLines = (lineCounts.get(pi) ?? 0) - 1 - breaks;
      if (extraLines > 0) {
        issues.push({ ...slideRef, kind: 'soft-wrap', paragraphIndex: pi, extraLines });
      }
    });
  }
  if (reportLineBreaks && vert === 'none') {
    auditLineBreaks(input.paragraphs, core.placements, slideRef, issues);
  }
  if (minFontSizePt !== undefined) {
    let smallestPx = Infinity;
    for (const para of input.paragraphs) {
      for (const piece of para.pieces) {
        if (!piece.isBreak && piece.text.trim() !== '' && piece.sizePx < smallestPx) {
          smallestPx = piece.sizePx;
        }
      }
    }
    const sizePt = Math.round(smallestPx * PT_PER_PX * 10) / 10;
    if (sizePt < minFontSizePt) issues.push({ ...slideRef, kind: 'small-font', sizePt });
  }
};

// Lines carry no record of the explicit break that ended them, so each line is
// located in its paragraph's text (a break piece written as `\n`): what lies
// between two lines tells a hard break, a wrap at a space, and a wrap inside a word.
const auditLineBreaks = (
  paragraphs: ReturnType<typeof buildSvgTextInput>['paragraphs'],
  placements: ReturnType<typeof layoutCore>['placements'],
  slideRef: IssueBase,
  issues: TextAuditIssue[],
): void => {
  const linesByParagraph = new Map<number, string[]>();
  for (const { line } of placements) {
    const text = line.tokens.map((t) => (t.isTab ? '\t' : t.text)).join('');
    if (text.trim() === '') continue;
    const lines = linesByParagraph.get(line.paraIndex);
    if (lines) lines.push(text);
    else linesByParagraph.set(line.paraIndex, [text]);
  }
  for (const [paragraphIndex, lines] of linesByParagraph) {
    const source = paragraphs[paragraphIndex]!.pieces.map((p) => (p.isBreak ? '\n' : p.text)).join(
      '',
    );
    auditParagraphBreaks(source, lines, { ...slideRef, paragraphIndex }, issues);
  }
};

const auditParagraphBreaks = (
  source: string,
  lines: readonly string[],
  ref: IssueBase & { readonly paragraphIndex: number },
  issues: TextAuditIssue[],
): void => {
  const found: TextAuditIssue[] = [];
  let previousEnd = -1;
  let wrappedLines = 0;
  const closeSegment = (lastLine: string): void => {
    const visible = [...lastLine.replace(/\s/gu, '')];
    if (wrappedLines > 0 && visible.length <= SHORT_LAST_LINE_MAX_CHARS) {
      found.push({ ...ref, kind: 'short-last-line', text: lastLine.trim() });
    }
  };
  for (let i = 0; i < lines.length; i++) {
    const start = source.indexOf(lines[i]!, Math.max(previousEnd, 0));
    // A line the layout rewrote (e.g. caps) cannot be placed; judge nothing.
    if (start < 0) return;
    if (previousEnd >= 0) {
      const gap = source.slice(previousEnd, start);
      if (gap.includes('\n')) {
        closeSegment(lines[i - 1]!);
        wrappedLines = 0;
      } else {
        wrappedLines++;
        const before = /.$/u.exec(source.slice(0, start))?.[0] ?? '';
        const after = /^./u.exec(source.slice(start))?.[0] ?? '';
        if (gap === '' && splitsWord(before, after)) {
          const head = WORD_RUN.exec(source.slice(0, start))![0];
          const tail = WORD_RUN_START.exec(source.slice(start))![0];
          found.push({ ...ref, kind: 'word-break', text: `${head}|${tail}` });
        }
      }
    }
    previousEnd = start + lines[i]!.length;
  }
  closeSegment(lines[lines.length - 1]!);
  issues.push(...found);
};

// A shape's verdict is approximate when any of its runs was measured by
// estimate: the heuristic measurer (no vertical metrics) or a fontkit
// measurer that hit a missing glyph (`approximate` flag).
const detectApproximate = (
  paragraphs: ReturnType<typeof buildSvgTextInput>['paragraphs'],
  measure: TextMeasurer,
): boolean => {
  const seen = new Set<string>();
  for (const para of paragraphs) {
    for (const piece of para.pieces) {
      if (piece.isBreak || piece.text === '') continue;
      const key = `${piece.family}|${piece.sizePx}|${piece.bold}|${piece.italic}|${piece.kerning ?? true}|${piece.text}`;
      if (seen.has(key)) continue;
      seen.add(key);
      const spec: FontSpec = {
        family: piece.family,
        sizePx: piece.sizePx,
        bold: piece.bold,
        italic: piece.italic,
        letterSpacingPx: piece.letterSpacingPx,
        kerning: piece.kerning ?? true,
      };
      const r = measure(piece.text, spec);
      if (r.approximate === true || r.ascentPx === undefined) return true;
    }
  }
  return false;
};

/**
 * Measures every text body in the deck and reports text that escapes its box
 * (`overflow-x` / `overflow-y`) and, opt-in, paragraphs that soft-wrap
 * (`soft-wrap`, 段落ち), wraps that split a word or strand a line of one or
 * two characters (`word-break`, `short-last-line`), and text below a minimum
 * size (`small-font`).
 *
 * Table cell text is not audited (v1 covers shape text bodies, including
 * placeholders and shapes inside groups).
 */
export const auditTextLayout = (
  pres: PresentationData,
  options: AuditTextLayoutOptions = {},
): ReadonlyArray<TextAuditIssue> => {
  const measure = options.measureText ?? defaultMeasurer;
  const resolved: ResolvedOptions = {
    tolerancePx: options.tolerancePx ?? DEFAULT_TOLERANCE_PX,
    reportSoftWraps: options.reportSoftWraps ?? false,
    reportLineBreaks: options.reportLineBreaks ?? false,
    minFontSizePt: options.minFontSizePt,
  };
  const theme = getPresentationTheme(pres);
  const issues: TextAuditIssue[] = [];
  const slides = getSlides(pres);
  for (let slideIndex = 0; slideIndex < slides.length; slideIndex++) {
    for (const shape of walkShapes(getSlideShapes(slides[slideIndex]!))) {
      auditShape(pres, theme, shape, slideIndex, measure, resolved, issues);
    }
  }
  return issues;
};
