// A deck's headline track and wording flags, for reviewing what a deck says.
//
// Flags are suspicions, not verdicts: the skill's references/review.md explains
// how to judge each one.

import type { ReadChartSpec } from '../../internal/chartml/index.ts';
import type { PresentationData, SlideData, SlideShapeData } from '../_internal-symbols.ts';
import { getShapeChartSpec } from './charts.ts';
import { getShapeRunFormatEffective } from './shape-paragraph.ts';
import { getShapeName, getShapeText } from './shape-read-base.ts';
import { getSlideShapes } from './shape-slide-read.ts';
import { getSlideNotes } from './slide-notes.ts';
import { getSlides } from './slide-query.ts';
import { getSlideTitle } from './slide-title.ts';
import { getTableCellText, getTableCells, isChartShape, isTableShape } from './tables.ts';

/** `talk` checks headlines against hand-made talk decks; `document` against written reports. */
export type PresentationTextReviewMode = 'talk' | 'document';

export interface PresentationTextReviewOptions {
  /** Adds the headline checks for this kind of deck. Omitted, neither set runs. */
  readonly mode?: PresentationTextReviewMode;
  /** Also lists every slide's text, chart labels and speaker notes. */
  readonly full?: boolean;
}

interface TextEntry {
  readonly shape: SlideShapeData | null;
  readonly text: string;
  readonly code: boolean;
  guessed: boolean;
}

interface SlideTextReview {
  readonly number: number;
  readonly headline: string | null;
  readonly guessed: boolean;
  readonly headlineRaw: string | null;
  readonly texts: ReadonlyArray<string>;
  readonly code: ReadonlyArray<string>;
  readonly chart: ReadonlyArray<string>;
  readonly notes: string;
}

type HeadlinedSlide = SlideTextReview & { readonly headline: string; readonly headlineRaw: string };

const HEADLINE_NAME = 'Headline';
// Longer text is body copy even when it is set large.
const GUESS_MAX_LENGTH = 80;
const MONOSPACE = /mono|courier|consolas|menlo|code|等幅/i;

/**
 * Returns a plain-text review of the deck: the headline track (one line per
 * slide, read from the shape named `Headline`, else the title placeholder, else
 * the largest short text), `[要確認]` markers left on slides and in notes, and
 * Japanese wording flags. `mode` adds talk or document headline checks; `full`
 * adds every slide's text. This is the output of the skill's `deck-text.mjs`.
 */
export const getPresentationTextReview = (
  pres: PresentationData,
  options: PresentationTextReviewOptions = {},
): string => {
  const slides = getSlides(pres).map((slide, index) => readSlide(pres, slide, index));
  const lines = headlineLines(slides);
  if (options.full) lines.push(...fullTextLines(slides));
  if (options.mode === 'talk') lines.push(...talkMetricLines(slides));
  if (options.mode === 'document') lines.push(...documentCheckLines(slides));
  lines.push(...openQuestionLines(slides), ...wordingFlagLines(slides));
  return lines.join('\n');
};

const readSlide = (pres: PresentationData, slide: SlideData, index: number): SlideTextReview => {
  const shapes: TextEntry[] = [];
  const chart: string[] = [];
  for (const shape of getSlideShapes(slide)) {
    if (isChartShape(shape)) {
      chart.push(...chartTexts(getShapeChartSpec(shape)));
      continue;
    }
    if (isTableShape(shape)) {
      for (const row of getTableCells(shape))
        shapes.push({
          shape: null,
          text: row.map((cell) => getTableCellText(cell)).join(' | '),
          code: false,
          guessed: false,
        });
      continue;
    }
    const text = getShapeText(shape).trim();
    if (text) shapes.push({ shape, text, code: isCode(pres, shape), guessed: false });
  }
  const headline =
    shapes.find(({ shape }) => shape !== null && getShapeName(shape) === HEADLINE_NAME) ??
    placeholderTitle(slide, shapes) ??
    largestText(pres, shapes);
  return {
    number: index + 1,
    headline: headline ? oneLine(headline.text) : null,
    guessed: headline?.guessed ?? false,
    headlineRaw: headline?.text ?? null,
    texts: shapes.filter((entry) => entry !== headline && !entry.code).map(({ text }) => text),
    code: shapes.filter((entry) => entry.code).map(({ text }) => text),
    chart,
    notes: getSlideNotes(slide) ?? '',
  };
};

const chartTexts = (spec: ReadChartSpec | null): string[] => {
  if (spec === null) return [];
  return [
    spec.title,
    spec.categoryAxisTitle,
    spec.valueAxisTitle,
    spec.secondaryValueAxis?.title,
    ...spec.categories,
    ...(spec.categoryGroupLevels ?? []).flat(),
    ...spec.series.flatMap((series) => [
      series.name,
      ...(series.pointDataLabels ?? []).map((labels) => labels?.text),
    ]),
  ].filter((text): text is string => typeof text === 'string' && text.trim() !== '');
};

// Code samples and prompt examples are quoted verbatim; wording flags there are noise.
const isCode = (pres: PresentationData, shape: SlideShapeData): boolean => {
  const format = getShapeRunFormatEffective(pres, shape, 0, 0);
  return MONOSPACE.test(`${format.font ?? ''} ${format.fontEastAsian ?? ''}`);
};

const placeholderTitle = (
  slide: SlideData,
  shapes: ReadonlyArray<TextEntry>,
): TextEntry | undefined => {
  const title = getSlideTitle(slide)?.trim();
  return title ? shapes.find((entry) => entry.text === title) : undefined;
};

// Decks made elsewhere rarely name their headline; the biggest short text is the best guess.
const largestText = (
  pres: PresentationData,
  shapes: ReadonlyArray<TextEntry>,
): TextEntry | undefined => {
  let best: TextEntry | undefined;
  let bestSize = 0;
  for (const entry of shapes) {
    if (entry.shape === null || entry.text.length > GUESS_MAX_LENGTH) continue;
    const size = getShapeRunFormatEffective(pres, entry.shape, 0, 0).size ?? 0;
    if (size > bestSize) [best, bestSize] = [entry, size];
  }
  if (best) best.guessed = true;
  return best;
};

const oneLine = (text: string): string => text.replace(/\s*[\n\v]\s*/g, ' ');

const label = (slide: SlideTextReview): string => `p${String(slide.number).padStart(2, '0')}`;

const hasHeadline = (slide: SlideTextReview): slide is HeadlinedSlide => slide.headline !== null;

const headlineLines = (all: ReadonlyArray<SlideTextReview>): string[] => {
  const lines = ['# Headline track'];
  const missing = all.filter((slide) => slide.headline === null).length;
  for (const slide of all)
    lines.push(`${label(slide)}${slide.guessed ? '?' : ' '} ${slide.headline ?? '(no headline)'}`);
  if (all.some((slide) => slide.guessed))
    lines.push(
      `\n"?" marks a guessed headline (largest text). Name the headline Text "${HEADLINE_NAME}".`,
    );
  if (missing > 0)
    lines.push(
      `\n${missing} slide(s) have no headline. Name the headline Text "${HEADLINE_NAME}" (or use a title placeholder).`,
    );
  return lines;
};

const fullTextLines = (all: ReadonlyArray<SlideTextReview>): string[] => {
  const lines = ['\n# Slide text'];
  for (const slide of all) {
    lines.push(`\n## ${label(slide)} ${slide.headline ?? ''}`);
    for (const text of slide.texts) lines.push(`- ${oneLine(text)}`);
    for (const text of slide.code) lines.push(`- [code] ${oneLine(text)}`);
    for (const text of slide.chart) lines.push(`- [chart] ${oneLine(text)}`);
    if (slide.notes) lines.push(`  notes: ${oneLine(slide.notes)}`);
  }
  return lines;
};

// Measured on hand-made Japanese conference decks (minorun365/minorun-marp-skill,
// slide-story). A deck far below these ranges reads as machine-uniform.
const TALK_RANGES = { exclamation: [7, 20], question: [10, 15], polite: [10, 12] } as const;
const POLITE_END = /(です|ます|ました|ません|でした|でしょう|ください)か?[。！？!?…]*$/;
const COLLOQUIAL_END = /(てる|ばいい|んです)[。！？!?]*$/;
const LEADING_CONNECTIVE = /^(しかも|実は|ちなみに|ただし|例えば|まずは|最後に|そして|さらに)/;
const COLON_HEADLINE = /[：:]/;
const MIN_RUN = 3;

const talkMetricLines = (all: ReadonlyArray<SlideTextReview>): string[] => {
  const headlines = all.filter(hasHeadline);
  if (headlines.length === 0) return [];
  const share = (test: (headline: string) => boolean): number =>
    Math.round((headlines.filter((slide) => test(slide.headline)).length / headlines.length) * 100);
  const shortest = headlines.reduce((a, b) => (a.headline.length <= b.headline.length ? a : b));
  const lines = [
    '\n# Talk headline mix (hand-made reference range in brackets)',
    `！ ${share((h) => /[！!]/.test(h))}% [${TALK_RANGES.exclamation.join('-')}%]`,
    `？ ${share((h) => /[？?]/.test(h))}% [${TALK_RANGES.question.join('-')}%]`,
    `polite form ${share((h) => POLITE_END.test(h))}% [${TALK_RANGES.polite.join('-')}%]`,
    `shortest ${shortest.headline.length} chars (${label(shortest)}) [3-6]; median ${median(headlines.map((s) => s.headline.length))}`,
  ];
  for (const slide of headlines)
    if (COLLOQUIAL_END.test(slide.headline))
      lines.push(`${label(slide)} colloquial ending: ${slide.headline}`);
  lines.push(
    ...runLines(headlines, 'colon headline', (h) => COLON_HEADLINE.test(h)),
    ...runLines(headlines, 'same leading connective', (h) => LEADING_CONNECTIVE.test(h)),
    ...runLines(headlines, 'comma-pause two-part headline', (h) =>
      /^[^、]{2,14}、[^、]{1,10}$/.test(h),
    ),
  );
  return lines;
};

const runLines = (
  slides: ReadonlyArray<HeadlinedSlide>,
  name: string,
  test: (headline: string) => boolean,
): string[] => {
  const lines: string[] = [];
  let run: HeadlinedSlide[] = [];
  const flush = (): void => {
    if (run.length >= MIN_RUN) lines.push(`${name} x${run.length}: ${run.map(label).join(', ')}`);
    run = [];
  };
  for (const slide of slides) {
    if (test(slide.headline)) run.push(slide);
    else flush();
  }
  flush();
  return lines;
};

// Full-width characters count 1, Latin letters and digits about half.
const LATIN_1_END = 0xff;
const HALF_WIDTH_KANA = [0xff61, 0xff9f] as const;
const displayWidth = (text: string): number => {
  let width = 0;
  for (const character of text) {
    const code = character.codePointAt(0)!;
    width +=
      code <= LATIN_1_END || (code >= HALF_WIDTH_KANA[0] && code <= HALF_WIDTH_KANA[1]) ? 0.5 : 1;
  }
  return width;
};

const median = (values: ReadonlyArray<number>): number | undefined => {
  const sorted = [...values].sort((a, b) => a - b);
  return sorted[Math.floor(sorted.length / 2)];
};

// Full-width characters per line in a 16:9 title box at ~28pt; carnot-tech/consulting-pptx-skill §2.1.
const DOCUMENT_TITLE_LINE = 40;
const DOCUMENT_TITLE_MAX_LINES = 2;
const ELEMENT_COUNT =
  /[0-9０-９一二三四五六七八九十]+ ?(つ|個|種類|点|段階?|ステップ|フェーズ|論点|柱)(の|で|が|を|に|$)/;

const documentCheckLines = (all: ReadonlyArray<SlideTextReview>): string[] => {
  const lines = ['\n# Document headline checks'];
  // The cover carries the talk title as given; it is not a headline to rewrite.
  for (const slide of all.slice(1).filter(hasHeadline)) {
    const headline = slide.headline;
    const issues: string[] = [];
    if (POLITE_END.test(headline)) issues.push('polite ending (use plain form)');
    const rawLines = slide.headlineRaw.split(/[\n\v]/);
    const widest = Math.max(...rawLines.map(displayWidth));
    if (rawLines.length > DOCUMENT_TITLE_MAX_LINES)
      issues.push(`${rawLines.length} lines (keep headlines to two)`);
    else if (widest > DOCUMENT_TITLE_LINE * DOCUMENT_TITLE_MAX_LINES)
      issues.push('longer than two lines');
    else if (widest > DOCUMENT_TITLE_LINE)
      issues.push('wraps without an explicit break (put one at a meaning boundary)');
    if (/^[^、。]{1,12}[：:]/.test(headline)) issues.push('label prefix before a colon');
    if (/^(この|その|ここまで|これら)/.test(headline)) issues.push('refers to another slide');
    if (ELEMENT_COUNT.test(headline))
      issues.push('element count in headline (state the content, or check it matches the body)');
    if (issues.length) lines.push(`${label(slide)} ${issues.join('; ')}: ${headline}`);
  }
  return lines;
};

// Each entry cites the rule in the skill's references/writing-ja.md. Order: highest confidence first.
const WORDING_PATTERNS: ReadonlyArray<readonly [string, RegExp]> = [
  // 効 inside 有効／効率／効果 is a different word.
  ['効く', /(?<![有無奏特即実発失])効(?:[くきけかこ]|いて|いた)/],
  ['dash', /[—―─]/],
  ['することができる', /することが(でき|可能)/],
  ['empty emphasis', /非常に|極めて|画期的|革新的|圧倒的|と言っても過言では|に他なら/],
  ['katakana padding', /ソリューション|シナジー|シームレス|エンドツーエンド|付加価値/],
  [
    'nominalized business verb',
    /(を|の)(活用|推進|実現|創出|担保|最適化|効率化|強化)(する|し|を|が|に)?/,
  ],
  ['stock opener', /昨今|近年|変化の激しい|結論から(言う|申し上げ)|以下の通り/],
  ['hedged conclusion', /と言えるでしょう|と言えます|のではないでしょうか|と考えられます/],
  [
    'announced count',
    /(理由|ポイント|軸|論点)は[0-9０-９一二三四五]+つ|見ていきましょう|深掘り|掘り下げ|言語化する|正面から/,
  ],
  ['empty adjective', /不可欠|核心的|鍵となる|多角的|包括的|肝となる/],
  [
    'translationese',
    /という観点から|にとって(重要|不可欠)|することによって|を持つ(こと|ため|。|$)/,
  ],
  ['stagy closer', /に尽き(る|ます)|これが.{1,12}の(形|正体)|そのもの(だ|です)|こそが/],
  ['showy metaphor verb', /届かない|拾ってくれ|地続き|一撃で|仕込[みむ]|たちが悪い|腑に落ち/],
  ['overclaimed effect', /これだけで|一気に|劇的に|驚くほど|目に見えて/],
  ['intensifying 切る', /(守り|止め|閉じ|抑え|塞ぎ|押し|言い)切[らりるれろっ]/],
  ['direction metaphor', /(側|方|向き)に倒|に倒(す|し|して|れる)(?![産])/],
  ['permission phrasing', /で構いません|でも構わない/],
  ['slide as 枚', /(次|前|この|全)の?枚|[0-9０-９]+枚目/],
];
// A Japanese label before a colon; ASCII labels are code, and source lines legitimately start with 出典.
const LABEL_COLON =
  /^\s*(?!出典|注)(?=[^\s：:、。[［]*[ぁ-んァ-ヶ一-龠])[^\s：:、。[［]{1,12}[：:]\s*\S/;
const FOREIGN_SCRIPT = /[가-힣Ѐ-ӿ]/;
const FLAG_EXCERPT_LENGTH = 80;

const wordingFlagLines = (all: ReadonlyArray<SlideTextReview>): string[] => {
  const lines = ['\n# Wording flags (judge each against references/writing-ja.md)'];
  for (const slide of all) {
    const onSlide = [
      slide.headline ?? '',
      ...[...slide.texts, ...slide.chart].flatMap((text) => text.split(/\n/)),
    ];
    const notes = slide.notes.split(/\n/);
    for (const [where, line] of [
      ...onSlide.map((line) => ['', line] as const),
      ...notes.map((line) => ['notes ', line] as const),
    ]) {
      const hits = flagsFor(line, where === '');
      if (hits.length === 0) continue;
      lines.push(
        `${label(slide)} ${where}[${hits.join(', ')}] ${oneLine(line).slice(0, FLAG_EXCERPT_LENGTH)}`,
      );
    }
  }
  if (lines.length === 1) lines.push('none');
  return lines;
};

// Notes are spoken prose, so the slide-layout check (label: description) does not apply.
const flagsFor = (line: string, onSlide: boolean): string[] => {
  if (!line.trim()) return [];
  const hits = WORDING_PATTERNS.filter(([, pattern]) => pattern.test(line)).map(([name]) => name);
  if (onSlide && LABEL_COLON.test(line)) hits.push('label: description');
  if (FOREIGN_SCRIPT.test(line)) hits.push('foreign script');
  return hits;
};

// Notes are listed apart: moving an open question there is one way to resolve it.
const OPEN_QUESTION = /[[［]要確認/;
const EXCERPT_LENGTH = 80;
const EXCERPT_LEAD = 20;

const openQuestionLines = (all: ReadonlyArray<SlideTextReview>): string[] => {
  const onSlides = all.flatMap((slide) =>
    [slide.headlineRaw ?? '', ...slide.texts, ...slide.code, ...slide.chart].map(
      (text) => [slide, text] as const,
    ),
  );
  return [
    '\n# [要確認] left on slides (fill each or move it to the notes before delivery)',
    ...matchLines(onSlides, ''),
    '\n# [要確認] in notes (report each to the user)',
    ...matchLines(
      all.map((slide) => [slide, slide.notes] as const),
      'notes ',
    ),
  ];
};

const matchLines = (
  entries: ReadonlyArray<readonly [SlideTextReview, string]>,
  where: string,
): string[] => {
  const lines: string[] = [];
  for (const [slide, text] of entries)
    for (const line of text.split(/\n/).map(oneLine)) {
      const at = line.search(OPEN_QUESTION);
      if (at === -1) continue;
      lines.push(`${label(slide)} ${where}${excerpt(line, at)}`);
    }
  if (lines.length === 0) lines.push('none');
  return lines;
};

// Keeps the marker visible when a long table row puts it past the cut.
const excerpt = (line: string, at: number): string => {
  if (line.length <= EXCERPT_LENGTH) return line;
  const start = Math.max(0, Math.min(at - EXCERPT_LEAD, line.length - EXCERPT_LENGTH));
  return `${start > 0 ? '…' : ''}${line.slice(start, start + EXCERPT_LENGTH)}${start + EXCERPT_LENGTH < line.length ? '…' : ''}`;
};
