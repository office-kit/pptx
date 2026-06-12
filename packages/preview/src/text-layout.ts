// Pure-SVG text layout engine for the preview renderer.
//
// The playground's default text path lays text out with <foreignObject> + HTML
// and lets the browser wrap/position it. That cannot be rasterized without a
// browser, so the fidelity harness (and any future Node renderer) needs text as
// real SVG <text>/<tspan>. This module does that layout itself: measure runs,
// wrap to the inner box, position lines on their baselines, and emit SVG.
//
// It is PURE — no pptx-kit, no Node, no browser globals — so it runs in the
// SvelteKit bundle, in the Node harness, and in unit tests. The caller injects
// a `TextMeasurer`; in the browser that can wrap `ctx.measureText`, in Node it
// wraps fontkit (see site/fidelity/measure.ts). All geometry is in CSS px.

// ---------------------------------------------------------------------------
// Public injection contract.

/** What a measurer needs to size one run. Pixels at 96 DPI; the caller has
 *  already applied EMU→px, pt→px and the autofit fontScale. `family` is the
 *  resolved internal family name (see `substituteFamily`), not a CSS list. */
export interface FontSpec {
  readonly family: string;
  readonly sizePx: number;
  readonly bold: boolean;
  readonly italic: boolean;
  readonly letterSpacingPx: number;
}

/** Advance width of `text` in px, plus optional vertical metrics. A real
 *  measurer returns ascent/descent/lineGap so line height matches the font;
 *  the heuristic returns width only and the engine falls back to a ratio. */
export interface MeasureResult {
  readonly widthPx: number;
  readonly ascentPx?: number;
  readonly descentPx?: number;
  readonly lineGapPx?: number;
}

export type TextMeasurer = (text: string, spec: FontSpec) => MeasureResult;

export type TextLayoutMode = 'foreignObject' | 'svg';

export interface RenderSlideOptions {
  /** Measurer used by the pure-SVG text path. Required when `textLayout` is
   *  'svg'; ignored otherwise. */
  readonly measureText?: TextMeasurer;
  /** Which text path to use. Defaults to 'foreignObject' (the browser path)
   *  so existing callers are unaffected; the harness opts into 'svg'. */
  readonly textLayout?: TextLayoutMode;
}

// ---------------------------------------------------------------------------
// Font substitution. resvg matches a <text font-family> against the *internal*
// family name in the loaded TTF. We bundle metric-compatible open substitutes
// (the same ones LibreOffice renders with), so we must emit and measure under
// the substitute's exact internal name — "Liberation Sans" with a space, not
// "LiberationSans". This map is the single source of truth shared by the SVG
// emitter and the Node measurer (site/fidelity/measure.ts).

export const SANS = 'Carlito'; // Calibri / Aptos / generic sans substitute
export const SERIF = 'Caladea'; // Cambria substitute
export const ARIAL = 'Liberation Sans';
export const TIMES = 'Liberation Serif';
export const MONO = 'Liberation Mono';

export const substituteFamily = (family: string | null | undefined): string => {
  if (!family) return SANS;
  const f = family.trim().toLowerCase();
  if (f.startsWith('calibri') || f.startsWith('aptos')) return SANS;
  if (f.startsWith('cambria')) return SERIF;
  if (f === 'arial' || f === 'helvetica' || f === 'helvetica neue' || f.startsWith('arial '))
    return ARIAL;
  if (f === 'times new roman' || f === 'times' || f.startsWith('times ')) return TIMES;
  if (f === 'courier new' || f === 'courier' || f === 'consolas' || f.startsWith('courier '))
    return MONO;
  // Unknown family → generic sans substitute, matching DEFAULT_FONT's Calibri
  // lead and LibreOffice's behavior when the named face is unavailable.
  return SANS;
};

// CJK detection — lifted from render-slide.ts so the heuristic measurer
// matches today's autofit estimate exactly.
const isCjk = (cp: number): boolean =>
  (cp >= 0x3040 && cp <= 0x309f) ||
  (cp >= 0x30a0 && cp <= 0x30ff) ||
  (cp >= 0x4e00 && cp <= 0x9fff) ||
  (cp >= 0xac00 && cp <= 0xd7af);

// Mean glyph width as a fraction of size for a typical sans-serif — the same
// 0.55 PowerPoint's autofit estimator uses. Only used by the heuristic
// measurer (no real font metrics available).
const AVG_GLYPH_W_RATIO = 0.55;

export const defaultMeasurer: TextMeasurer = (text, spec) => {
  let w = 0;
  for (const ch of text) {
    const ratio = isCjk(ch.codePointAt(0) ?? 0) ? 1 : AVG_GLYPH_W_RATIO;
    w += spec.sizePx * ratio;
  }
  const n = [...text].length;
  if (n > 1) w += spec.letterSpacingPx * (n - 1); // tracking applies between glyphs
  return { widthPx: w };
};

// Vertical-metric fallback when the measurer returns width only. Calibri /
// Carlito sit near ascent 0.9em / descent 0.22em / gap 0.08em including the
// internal leading. The scored harness path always supplies real metrics, so
// this only drives the browser heuristic preview.
const FALLBACK_ASCENT = 0.9;
const FALLBACK_DESCENT = 0.22;
const FALLBACK_LINEGAP = 0.08;

// Vertical-centering calibration — see the use site. Fraction of a line's
// (ascent+descent) that center/bottom-anchored blocks sit LOWER than the
// win-metric line box predicts, fitted to LibreOffice/PowerPoint ground truth
// (top-anchored text needs no correction; only center/bottom do).
const CENTER_ANCHOR_DROP = 0.036;

// ---------------------------------------------------------------------------
// Engine input model. render-slide.ts resolves the OOXML cascade and hands the
// engine this already-normalized, px-native structure.

export interface PieceInput {
  readonly text: string;
  readonly family: string; // internal substituted name
  readonly sizePx: number; // post-autofit
  readonly bold: boolean;
  readonly italic: boolean;
  readonly letterSpacingPx: number;
  readonly fillHex: string;
  readonly underline: boolean;
  readonly strike: boolean;
  readonly superSub: 0 | 1 | -1; // 1 superscript, -1 subscript
  readonly href: string | null;
  readonly isBreak: boolean; // <a:br>
}

export interface BulletInput {
  readonly text: string;
  readonly family: string;
  readonly sizePx: number;
  readonly fillHex: string;
}

export interface ParaInput {
  readonly align: 'left' | 'center' | 'right' | 'justify';
  readonly marLpx: number;
  readonly marRpx: number;
  readonly firstIndentPx: number; // first-line indent (negative = hanging)
  readonly spcBefPx: number;
  readonly spcAftPx: number;
  readonly lineSpacing: { kind: 'pct'; value: number } | { kind: 'pts'; px: number } | null;
  readonly lineAdvanceScale: number; // lnSpcReduction → 1 - reduction
  readonly bullet: BulletInput | null;
  readonly pieces: readonly PieceInput[];
  readonly fallbackSizePx: number; // line height for an empty paragraph
}

/** Normalized text direction the engine understands. render-slide maps the
 *  OOXML `vert` token (`vert`/`eaVert`/`vert270`/…) onto one of these — the
 *  engine stays free of OOXML vocabulary. `cw90` rotates the horizontal layout
 *  90° clockwise (reading top-to-bottom, columns right-to-left); `cw270` is the
 *  opposite 270° rotation (columns left-to-right). */
export type VerticalLayout = 'none' | 'cw90' | 'cw270';

export interface ColumnLayout {
  readonly count: number; // PowerPoint clamps to >= 2 before this point
  readonly gapPx: number;
}

export interface TextBodyInput {
  readonly boxXpx: number;
  readonly boxYpx: number;
  readonly boxWpx: number;
  readonly boxHpx: number;
  readonly anchor: 'top' | 'center' | 'bottom';
  readonly wrap: boolean;
  readonly paragraphs: readonly ParaInput[];
  /** Vertical text direction; omitted / 'none' is the default horizontal flow. */
  readonly vert?: VerticalLayout;
  /** Multi-column body (`numCol`/`spcCol`); null / omitted is single column. */
  readonly columns?: ColumnLayout | null;
}

// ---------------------------------------------------------------------------
// Layout internals.

interface Token {
  readonly text: string;
  readonly piece: PieceInput;
  readonly isSpace: boolean;
  readonly isBreak: boolean;
  width: number;
}

interface Line {
  readonly tokens: Token[];
  ascent: number;
  descent: number;
  lineGap: number;
  topY: number;
  advance: number;
  anchorX: number;
  textAnchor: 'start' | 'middle' | 'end';
  bullet: { x: number; baselineDy: number; b: BulletInput } | null;
}

// The box (in local coords) that one horizontal layout pass fills. Horizontal
// text uses the shape box directly; vertical text uses a swapped-extent frame
// re-centred on the box (see layoutTextSvg).
interface Frame {
  readonly x: number;
  readonly y: number;
  readonly w: number;
  readonly h: number;
}

// A laid-out line ready to emit: its baseline Y plus an X shift (non-zero only
// for the second-and-later columns of a multi-column body).
interface Placement {
  readonly line: Line;
  readonly baselineY: number;
  readonly dx: number;
}

// Builds the line list for one column of a given content width. Closes over the
// paragraph model and measurer caches inside layoutTextSvg.
type LineBuilder = (contentLeft: number, contentRight: number) => { lines: Line[]; blockH: number };

const specOf = (piece: PieceInput): FontSpec => ({
  family: piece.family,
  sizePx: piece.sizePx,
  bold: piece.bold,
  italic: piece.italic,
  letterSpacingPx: piece.letterSpacingPx,
});

const bulletSpec = (b: BulletInput): FontSpec => ({
  family: b.family,
  sizePx: b.sizePx,
  bold: false,
  italic: false,
  letterSpacingPx: 0,
});

const escapeXml = (s: string): string =>
  s.replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c] ?? c);

const fmt = (n: number): string => {
  const r = Math.round(n * 100) / 100;
  return Object.is(r, -0) ? '0' : String(r);
};

// ---------------------------------------------------------------------------

export const layoutTextSvg = (input: TextBodyInput, measure: TextMeasurer): string => {
  const widthCache = new Map<string, number>();
  const metricCache = new Map<string, { a: number; d: number; g: number }>();
  const key = (text: string, s: FontSpec): string =>
    `${s.family}|${s.sizePx}|${s.bold}|${s.italic}|${s.letterSpacingPx}|${text}`;
  const mWidth = (text: string, s: FontSpec): number => {
    const k = key(text, s);
    let w = widthCache.get(k);
    if (w === undefined) {
      w = measure(text, s).widthPx;
      widthCache.set(k, w);
    }
    return w;
  };
  const mMetrics = (piece: PieceInput): { a: number; d: number; g: number } => {
    const s = specOf(piece);
    const k = `${s.family}|${s.sizePx}|${s.bold}|${s.italic}`;
    let m = metricCache.get(k);
    if (!m) {
      const r = measure('Mg', s);
      m =
        r.ascentPx !== undefined && r.descentPx !== undefined
          ? { a: r.ascentPx, d: r.descentPx, g: r.lineGapPx ?? 0 }
          : {
              a: piece.sizePx * FALLBACK_ASCENT,
              d: piece.sizePx * FALLBACK_DESCENT,
              g: piece.sizePx * FALLBACK_LINEGAP,
            };
      metricCache.set(k, m);
    }
    return m;
  };

  // One horizontal layout pass, parameterized on the content box's left/right
  // edges so the same wrap/measure code serves the shape box (horizontal), a
  // single column of a multi-column body, and the swapped frame of vertical
  // text. Closes over the measurer caches above.
  const buildLines: LineBuilder = (contentLeft, contentRight) => {
    const lines: Line[] = [];
    let cursorY = 0;

    for (const para of input.paragraphs) {
      cursorY += para.spcBefPx;
      const wrapLeft = contentLeft + para.marLpx;
      const wrapRight = contentRight - para.marRpx;
      const firstLeft = wrapLeft + para.firstIndentPx;
      const hasText = para.pieces.some((p) => !p.isBreak && p.text !== '');
      const bullet = para.bullet && hasText ? para.bullet : null;
      const bulletLead = bullet ? mWidth(`${bullet.text} `, bulletSpec(bullet)) : 0;

      // Tokenize: word / whitespace runs per piece, plus break markers. Pre-split
      // any single token wider than a full line into per-character tokens.
      const avail = Math.max(1, wrapRight - wrapLeft);
      const tokens: Token[] = [];
      for (const piece of para.pieces) {
        if (piece.isBreak) {
          tokens.push({ text: '', piece, isSpace: false, isBreak: true, width: 0 });
          continue;
        }
        for (const seg of piece.text.match(/\s+|\S+/g) ?? []) {
          const isSpace = /^\s+$/.test(seg);
          const w = mWidth(seg, specOf(piece));
          if (input.wrap && !isSpace && w > avail - bulletLead && [...seg].length > 1) {
            for (const ch of seg) {
              tokens.push({
                text: ch,
                piece,
                isSpace: false,
                isBreak: false,
                width: mWidth(ch, specOf(piece)),
              });
            }
          } else {
            tokens.push({ text: seg, piece, isSpace, isBreak: false, width: w });
          }
        }
      }

      const wrapped = wrapTokens(tokens, input.wrap, wrapRight - firstLeft - bulletLead, avail);
      const paraLines: Token[][] = wrapped.length > 0 ? wrapped : [[]];

      for (let li = 0; li < paraLines.length; li++) {
        const toks = paraLines[li]!;
        let ascent = 0;
        let descent = 0;
        let lineGap = 0;
        for (const t of toks) {
          if (t.isSpace || t.isBreak) continue;
          const m = mMetrics(t.piece);
          if (m.a > ascent) ascent = m.a;
          if (m.d > descent) descent = m.d;
          if (m.g > lineGap) lineGap = m.g;
        }
        if (ascent === 0) {
          ascent = para.fallbackSizePx * FALLBACK_ASCENT;
          descent = para.fallbackSizePx * FALLBACK_DESCENT;
          lineGap = para.fallbackSizePx * FALLBACK_LINEGAP;
        }
        const isFirst = li === 0;
        const lineLeft = (isFirst ? firstLeft : wrapLeft) + bulletLead;
        const line: Line = {
          tokens: toks,
          ascent,
          descent,
          lineGap,
          topY: cursorY,
          advance: 0,
          anchorX: lineLeft,
          textAnchor: 'start',
          bullet: null,
        };
        if (para.align === 'center') {
          line.textAnchor = 'middle';
          line.anchorX = (lineLeft + wrapRight) / 2;
        } else if (para.align === 'right') {
          line.textAnchor = 'end';
          line.anchorX = wrapRight;
        }
        if (isFirst && bullet) {
          line.bullet = { x: firstLeft, baselineDy: 0, b: bullet };
        }
        line.advance = lineAdvance(line, para);
        cursorY += line.advance;
        lines.push(line);
      }
      cursorY += para.spcAftPx;
    }
    return { lines, blockH: cursorY };
  };

  const vert = input.vert ?? 'none';
  const cx = input.boxXpx + input.boxWpx / 2;
  const cy = input.boxYpx + input.boxHpx / 2;
  // Vertical text lays out horizontally into a frame whose extents are swapped
  // and re-centred on the box, so a single rotate(±90) about the box centre
  // maps that horizontal layout exactly onto the shape. All wrap / anchor /
  // align math then stays identical to horizontal — only the wrapping transform
  // differs, which is what keeps the calibrated horizontal path byte-stable.
  const frame: Frame =
    vert === 'none'
      ? { x: input.boxXpx, y: input.boxYpx, w: input.boxWpx, h: input.boxHpx }
      : { x: cx - input.boxHpx / 2, y: cy - input.boxWpx / 2, w: input.boxHpx, h: input.boxWpx };

  // Columns apply to horizontal text only — a faithful rotated-column layout is
  // disproportionate and PowerPoint itself barely supports vert + numCol, so we
  // keep the rotation and drop numCol rather than emit a wrong combined layout.
  const columns = vert === 'none' ? (input.columns ?? null) : null;

  const placements =
    columns && columns.count >= 2
      ? placeColumns(frame, columns, input.anchor, buildLines)
      : placeSingle(frame, input.anchor, buildLines);

  const body = emitPlacements(placements);
  if (vert === 'none') return body;
  const deg = vert === 'cw90' ? 90 : 270;
  return `<g transform="rotate(${deg} ${fmt(cx)} ${fmt(cy)})">${body}</g>`;
};

// Vertical block offset for an anchor. Center / bottom get the empirical drop
// calibration (see site/fidelity/README.md): LibreOffice & PowerPoint sit
// center/bottom-anchored text ≈0.036 of a line's (ascent+descent) lower than
// the win-metric line box predicts; top-anchored text aligns exactly.
const anchorOffsetY = (
  frameY: number,
  frameH: number,
  blockH: number,
  anchor: 'top' | 'center' | 'bottom',
  firstLine: Line | undefined,
): number => {
  let offsetY = frameY;
  if (anchor === 'center') offsetY = frameY + (frameH - blockH) / 2;
  else if (anchor === 'bottom') offsetY = frameY + (frameH - blockH);
  if ((anchor === 'center' || anchor === 'bottom') && firstLine) {
    offsetY += CENTER_ANCHOR_DROP * (firstLine.ascent + firstLine.descent);
  }
  return offsetY;
};

const placeSingle = (
  frame: Frame,
  anchor: 'top' | 'center' | 'bottom',
  buildLines: LineBuilder,
): Placement[] => {
  const { lines, blockH } = buildLines(frame.x, frame.x + frame.w);
  const offsetY = anchorOffsetY(frame.y, frame.h, blockH, anchor, lines[0]);
  return lines.map((line) => ({
    line,
    baselineY: offsetY + line.topY + topPad(line) + line.ascent,
    dx: 0,
  }));
};

const placeColumns = (
  frame: Frame,
  columns: ColumnLayout,
  anchor: 'top' | 'center' | 'bottom',
  buildLines: LineBuilder,
): Placement[] => {
  const gap = columns.gapPx;
  const colW = Math.max(1, (frame.w - (columns.count - 1) * gap) / columns.count);
  const { lines } = buildLines(frame.x, frame.x + colW);
  // Sequential fill, matching PowerPoint (not CSS column-fill:balance): fill
  // column 1 down to the box height, then overflow into column 2, and so on. A
  // column only breaks once it already holds a line, so a single over-tall line
  // can't loop. With few lines everything stays in column 1.
  let col = 0;
  let colStartTopY = 0;
  const colHasLine: boolean[] = [];
  let tallest = 0;
  const placed: { line: Line; localTopY: number; col: number }[] = [];
  for (const line of lines) {
    const localBottom = line.topY - colStartTopY + line.advance;
    if (col < columns.count - 1 && localBottom > frame.h && colHasLine[col]) {
      col += 1;
      colStartTopY = line.topY;
    }
    const localTopY = line.topY - colStartTopY;
    placed.push({ line, localTopY, col });
    colHasLine[col] = true;
    if (localTopY + line.advance > tallest) tallest = localTopY + line.advance;
  }
  // Vertical anchor applies to the whole body via its tallest column block —
  // with sequential fill the filled columns reach the box height, so this
  // matches PowerPoint/LibreOffice anchoring the body as one unit.
  const offsetY = anchorOffsetY(frame.y, frame.h, tallest, anchor, lines[0]);
  return placed.map(({ line, localTopY, col: c }) => ({
    line,
    baselineY: offsetY + localTopY + topPad(line) + line.ascent,
    dx: c * (colW + gap),
  }));
};

const emitPlacements = (placements: Placement[]): string => {
  const parts: string[] = [];
  for (const { line, baselineY, dx } of placements) {
    if (line.bullet) {
      const b = line.bullet.b;
      parts.push(
        `<text x="${fmt(line.bullet.x + dx)}" y="${fmt(baselineY)}" font-family="${escapeXml(b.family)}" font-size="${fmt(b.sizePx)}" fill="${b.fillHex}" xml:space="preserve">${escapeXml(b.text)}</text>`,
      );
    }
    parts.push(emitLine(line, baselineY, dx));
  }
  return parts.join('');
};

const topPad = (line: Line): number => {
  // Half-leading: spread (advance - (ascent+descent)) above and below the
  // glyph box. For the natural advance this is lineGap/2.
  const lead = line.advance - (line.ascent + line.descent);
  return lead > 0 ? lead / 2 : 0;
};

const lineAdvance = (line: Line, para: ParaInput): number => {
  // Natural single-line height = ascent + descent + lineGap. The measurer
  // supplies the metric set the renderer (LibreOffice / GDI) actually uses:
  // win metrics (ascent = usWinAscent, lineGap 0) unless the font sets
  // USE_TYPO_METRICS, in which case typo metrics + typoLineGap. The baseline
  // sits at lineTop + ascent.
  const natural = line.ascent + line.descent + line.lineGap;
  let adv: number;
  if (para.lineSpacing?.kind === 'pct') adv = para.lineSpacing.value * natural;
  else if (para.lineSpacing?.kind === 'pts') adv = para.lineSpacing.px;
  else adv = natural;
  return adv * para.lineAdvanceScale;
};

const emitLine = (line: Line, baselineY: number, dx: number): string => {
  const toks = [...line.tokens];
  while (toks.length > 0 && (toks[toks.length - 1]!.isSpace || toks[toks.length - 1]!.isBreak)) {
    toks.pop();
  }
  const content = toks.filter((t) => !t.isBreak);
  if (content.length === 0) return '';
  const tspans = groupTokens(content)
    .map((g) => tspan(g))
    .join('');
  if (tspans === '') return '';
  return `<text x="${fmt(line.anchorX + dx)}" y="${fmt(baselineY)}" text-anchor="${line.textAnchor}" xml:space="preserve">${tspans}</text>`;
};

interface Group {
  text: string;
  piece: PieceInput;
}

const groupTokens = (toks: Token[]): Group[] => {
  const groups: Group[] = [];
  for (const t of toks) {
    if (t.isBreak) continue;
    const last = groups[groups.length - 1];
    if (last && samePiece(last.piece, t.piece)) last.text += t.text;
    else groups.push({ text: t.text, piece: t.piece });
  }
  return groups;
};

const samePiece = (a: PieceInput, b: PieceInput): boolean =>
  a.family === b.family &&
  a.sizePx === b.sizePx &&
  a.bold === b.bold &&
  a.italic === b.italic &&
  a.letterSpacingPx === b.letterSpacingPx &&
  a.fillHex === b.fillHex &&
  a.underline === b.underline &&
  a.strike === b.strike &&
  a.superSub === b.superSub &&
  a.href === b.href;

const tspan = (g: Group): string => {
  const p = g.piece;
  const sizePx = p.superSub !== 0 ? p.sizePx * 0.65 : p.sizePx;
  const attrs: string[] = [
    `font-family="${escapeXml(p.family)}"`,
    `font-size="${fmt(sizePx)}"`,
    `fill="${p.fillHex}"`,
  ];
  if (p.bold) attrs.push('font-weight="700"');
  if (p.italic) attrs.push('font-style="italic"');
  const deco: string[] = [];
  if (p.underline) deco.push('underline');
  if (p.strike) deco.push('line-through');
  if (deco.length) attrs.push(`text-decoration="${deco.join(' ')}"`);
  if (p.letterSpacingPx !== 0) attrs.push(`letter-spacing="${fmt(p.letterSpacingPx)}"`);
  if (p.superSub === 1) attrs.push(`baseline-shift="${fmt(p.sizePx * 0.33)}"`);
  else if (p.superSub === -1) attrs.push(`baseline-shift="${fmt(-p.sizePx * 0.16)}"`);
  return `<tspan ${attrs.join(' ')}>${escapeXml(g.text)}</tspan>`;
};

// ---------------------------------------------------------------------------
// Greedy first-fit line breaking. Over-long tokens are pre-split into chars by
// the caller, so here a token always fits on an empty line.

const wrapTokens = (
  tokens: Token[],
  wrap: boolean,
  firstAvail: number,
  avail: number,
): Token[][] => {
  const lines: Token[][] = [];
  let cur: Token[] = [];
  let lineW = 0; // width of every token pushed to cur (words + spaces)
  let trailingSpaceW = 0; // width of the trailing run of spaces
  let first = true;

  const trimTrailing = (): void => {
    while (cur.length > 0 && (cur[cur.length - 1]!.isSpace || cur[cur.length - 1]!.isBreak)) {
      cur.pop();
    }
  };
  const close = (): void => {
    trimTrailing();
    lines.push(cur);
    cur = [];
    lineW = 0;
    trailingSpaceW = 0;
    first = false;
  };

  for (const tok of tokens) {
    if (tok.isBreak) {
      cur.push(tok);
      close();
      continue;
    }
    if (tok.isSpace) {
      cur.push(tok);
      lineW += tok.width;
      trailingSpaceW += tok.width;
      continue;
    }
    const limit = first ? firstAvail : avail;
    const contentW = lineW - trailingSpaceW; // exclude trailing spaces from the fit test
    const hasContent = contentW > 0;
    if (wrap && hasContent && contentW + tok.width > limit + 0.5) {
      close();
      cur.push(tok);
      lineW = tok.width;
    } else {
      // Trailing spaces become committed inter-word spaces.
      cur.push(tok);
      lineW += tok.width;
    }
    trailingSpaceW = 0;
  }
  trimTrailing();
  if (cur.length > 0) lines.push(cur);
  return lines;
};
