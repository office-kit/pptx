#!/usr/bin/env node
// Prints a deck's headline track and flags wording for review.
// Usage (from the slide project): node <skill>/scripts/deck-text.mjs deck.pptx [--mode talk|document] [--full]
//
// Flags are suspicions, not verdicts: references/review.md explains how to judge each one.
import { readFile } from 'node:fs/promises';
import { existsSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { pathToFileURL } from 'node:url';
import { parseArgs } from 'node:util';

const USAGE = 'Usage: deck-text.mjs <deck.pptx> [--mode talk|document] [--full]';
const MODES = ['talk', 'document'];
const HEADLINE_NAME = 'Headline';
// Longer text is body copy even when it is set large.
const GUESS_MAX_LENGTH = 80;
const MONOSPACE = /mono|courier|consolas|menlo|code|等幅/i;

const { values, positionals } = parseArgs({
  allowPositionals: true,
  options: { mode: { type: 'string' }, full: { type: 'boolean', default: false } },
});
const [file] = positionals;
const mode = values.mode ?? null;
if (positionals.length !== 1 || (mode !== null && !MODES.includes(mode))) {
  console.error(USAGE);
  process.exit(2);
}
const full = values.full;

const pptx = await import(await resolveCore(process.cwd()));
const presentation = await pptx.loadPresentation(await readFile(file));

// The script ships with the skill, outside the slide project, so a bare import
// would resolve against the skill directory. Resolve the project's installed copy.
async function resolveCore(start) {
  for (let directory = start; ; directory = dirname(directory)) {
    const own = join(directory, 'package.json');
    if (existsSync(own) && JSON.parse(await readFile(own, 'utf8')).name === '@office-kit/pptx')
      return entryOf(directory);
    const installed = join(directory, 'node_modules', '@office-kit', 'pptx');
    if (existsSync(installed)) return entryOf(installed);
    if (dirname(directory) === directory)
      throw new Error('@office-kit/pptx is not installed; run this from the slide project.');
  }
}

async function entryOf(packageDirectory) {
  const manifest = JSON.parse(await readFile(join(packageDirectory, 'package.json'), 'utf8'));
  return pathToFileURL(join(packageDirectory, manifest.exports['.'].import)).href;
}

function readSlide(slide, index) {
  const shapes = [];
  for (const shape of pptx.getSlideShapes(slide)) {
    if (pptx.isTableShape(shape)) {
      for (const row of pptx.getTableCells(shape))
        shapes.push({ text: row.map((cell) => pptx.getTableCellText(cell)).join(' | ') });
      continue;
    }
    const text = pptx.getShapeText(shape).trim();
    if (text) shapes.push({ shape, text, code: isCode(shape) });
  }
  const headline =
    shapes.find(({ shape }) => shape && pptx.getShapeName(shape) === HEADLINE_NAME) ??
    placeholderTitle(slide, shapes) ??
    largestText(shapes);
  return {
    number: index + 1,
    headline: headline ? oneLine(headline.text) : null,
    guessed: headline?.guessed ?? false,
    headlineRaw: headline?.text ?? null,
    texts: shapes.filter((entry) => entry !== headline && !entry.code).map(({ text }) => text),
    code: shapes.filter((entry) => entry.code).map(({ text }) => text),
    notes: pptx.getSlideNotes(slide) ?? '',
  };
}

// Code samples and prompt examples are quoted verbatim; wording flags there are noise.
function isCode(shape) {
  const format = pptx.getShapeRunFormatEffective(presentation, shape, 0, 0);
  return MONOSPACE.test(`${format?.font ?? ''} ${format?.fontEastAsian ?? ''}`);
}

function placeholderTitle(slide, shapes) {
  const title = pptx.getSlideTitle(slide)?.trim();
  return title ? shapes.find((entry) => entry.text === title) : undefined;
}

// Decks made elsewhere rarely name their headline; the biggest short text is the best guess.
function largestText(shapes) {
  let best;
  let bestSize = 0;
  for (const entry of shapes) {
    if (!entry.shape || entry.text.length > GUESS_MAX_LENGTH) continue;
    const size = pptx.getShapeRunFormatEffective(presentation, entry.shape, 0, 0)?.size ?? 0;
    if (size > bestSize) [best, bestSize] = [entry, size];
  }
  if (best) best.guessed = true;
  return best;
}

function oneLine(text) {
  return text.replace(/\s*[\n\v]\s*/g, ' ');
}

function label(slide) {
  return `p${String(slide.number).padStart(2, '0')}`;
}

function printHeadlines(all) {
  console.log('# Headline track');
  const missing = all.filter((slide) => slide.headline === null).length;
  for (const slide of all)
    console.log(`${label(slide)}${slide.guessed ? '?' : ' '} ${slide.headline ?? '(no headline)'}`);
  if (all.some((slide) => slide.guessed))
    console.log(
      `\n"?" marks a guessed headline (largest text). Name the headline Text "${HEADLINE_NAME}".`,
    );
  if (missing > 0)
    console.log(
      `\n${missing} slide(s) have no headline. Name the headline Text "${HEADLINE_NAME}" (or use a title placeholder).`,
    );
}

function printFullText(all) {
  console.log('\n# Slide text');
  for (const slide of all) {
    console.log(`\n## ${label(slide)} ${slide.headline ?? ''}`);
    for (const text of slide.texts) console.log(`- ${oneLine(text)}`);
    for (const text of slide.code) console.log(`- [code] ${oneLine(text)}`);
    if (slide.notes) console.log(`  notes: ${oneLine(slide.notes)}`);
  }
}

// Measured on hand-made Japanese conference decks (minorun365/minorun-marp-skill,
// slide-story). A deck far below these ranges reads as machine-uniform.
const TALK_RANGES = { exclamation: [7, 20], question: [10, 15], polite: [10, 12] };
const POLITE_END = /(です|ます|ました|ません|でした|でしょう|ください)か?[。！？!?…]*$/;
const COLLOQUIAL_END = /(てる|ばいい|んです)[。！？!?]*$/;
const LEADING_CONNECTIVE = /^(しかも|実は|ちなみに|ただし|例えば|まずは|最後に|そして|さらに)/;
const COLON_HEADLINE = /[：:]/;

function printTalkMetrics(all) {
  const headlines = all.filter((slide) => slide.headline !== null);
  if (headlines.length === 0) return;
  const share = (test) =>
    Math.round((headlines.filter((slide) => test(slide.headline)).length / headlines.length) * 100);
  const shortest = headlines.reduce((a, b) => (a.headline.length <= b.headline.length ? a : b));
  console.log('\n# Talk headline mix (hand-made reference range in brackets)');
  console.log(`！ ${share((h) => /[！!]/.test(h))}% [${TALK_RANGES.exclamation.join('-')}%]`);
  console.log(`？ ${share((h) => /[？?]/.test(h))}% [${TALK_RANGES.question.join('-')}%]`);
  console.log(
    `polite form ${share((h) => POLITE_END.test(h))}% [${TALK_RANGES.polite.join('-')}%]`,
  );
  console.log(
    `shortest ${shortest.headline.length} chars (${label(shortest)}) [3-6]; median ${median(headlines.map((s) => s.headline.length))}`,
  );
  for (const slide of headlines)
    if (COLLOQUIAL_END.test(slide.headline))
      console.log(`${label(slide)} colloquial ending: ${slide.headline}`);
  reportRuns(headlines, 'colon headline', (h) => COLON_HEADLINE.test(h));
  reportRuns(headlines, 'same leading connective', (h) => LEADING_CONNECTIVE.test(h));
  reportRuns(headlines, 'comma-pause two-part headline', (h) =>
    /^[^、]{2,14}、[^、]{1,10}$/.test(h),
  );
}

function reportRuns(slides, name, test) {
  let run = [];
  const flush = () => {
    if (run.length >= 3) console.log(`${name} x${run.length}: ${run.map(label).join(', ')}`);
    run = [];
  };
  for (const slide of slides) {
    if (test(slide.headline)) run.push(slide);
    else flush();
  }
  flush();
}

// Full-width characters count 1, Latin letters and digits about half.
const LATIN_1_END = 0xff;
const HALF_WIDTH_KANA = [0xff61, 0xff9f];
function displayWidth(text) {
  let width = 0;
  for (const character of text) {
    const code = character.codePointAt(0);
    width +=
      code <= LATIN_1_END || (code >= HALF_WIDTH_KANA[0] && code <= HALF_WIDTH_KANA[1]) ? 0.5 : 1;
  }
  return width;
}

function median(values) {
  const sorted = [...values].sort((a, b) => a - b);
  return sorted[Math.floor(sorted.length / 2)];
}

// Full-width characters per line in a 16:9 title box at ~28pt; carnot-tech/consulting-pptx-skill §2.1.
const DOCUMENT_TITLE_LINE = 40;
const ELEMENT_COUNT =
  /[0-9０-９一二三四五六七八九十]+ ?(つ|個|種類|点|段階?|ステップ|フェーズ|論点|柱)(の|で|が|を|に|$)/;

function printDocumentChecks(all) {
  console.log('\n# Document headline checks');
  // The cover carries the talk title as given; it is not a headline to rewrite.
  for (const slide of all.slice(1)) {
    const headline = slide.headline;
    if (headline === null) continue;
    const issues = [];
    if (POLITE_END.test(headline)) issues.push('polite ending (use plain form)');
    const lines = slide.headlineRaw.split(/[\n\v]/);
    const widest = Math.max(...lines.map(displayWidth));
    if (lines.length > 2) issues.push(`${lines.length} lines (keep headlines to two)`);
    else if (widest > DOCUMENT_TITLE_LINE * 2) issues.push('longer than two lines');
    else if (widest > DOCUMENT_TITLE_LINE)
      issues.push('wraps without an explicit break (put one at a meaning boundary)');
    if (/^[^、。]{1,12}[：:]/.test(headline)) issues.push('label prefix before a colon');
    if (/^(この|その|ここまで|これら)/.test(headline)) issues.push('refers to another slide');
    if (ELEMENT_COUNT.test(headline))
      issues.push('element count in headline (state the content, or check it matches the body)');
    if (issues.length) console.log(`${label(slide)} ${issues.join('; ')}: ${headline}`);
  }
}

// Each entry cites the rule in references/writing-ja.md. Order: highest confidence first.
const WORDING_PATTERNS = [
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
const LABEL_COLON = /^\s*(?!出典|注)(?=[^\s：:、。]*[ぁ-んァ-ヶ一-龠])[^\s：:、。]{1,12}[：:]\s*\S/;
const FOREIGN_SCRIPT = /[가-힣Ѐ-ӿ]/;

function printWordingFlags(all) {
  console.log('\n# Wording flags (judge each against references/writing-ja.md)');
  let count = 0;
  for (const slide of all) {
    const lines = [slide.headline ?? '', ...slide.texts.flatMap((text) => text.split(/\n/))];
    const notes = slide.notes.split(/\n/);
    for (const [where, line] of [
      ...lines.map((line) => ['', line]),
      ...notes.map((line) => ['notes ', line]),
    ]) {
      const hits = flagsFor(line, where === '');
      if (hits.length === 0) continue;
      count++;
      console.log(`${label(slide)} ${where}[${hits.join(', ')}] ${oneLine(line).slice(0, 80)}`);
    }
  }
  if (count === 0) console.log('none');
}

// Notes are spoken prose, so the slide-layout check (label: description) does not apply.
function flagsFor(line, onSlide) {
  if (!line.trim()) return [];
  const hits = WORDING_PATTERNS.filter(([, pattern]) => pattern.test(line)).map(([name]) => name);
  if (onSlide && LABEL_COLON.test(line)) hits.push('label: description');
  if (FOREIGN_SCRIPT.test(line)) hits.push('foreign script');
  return hits;
}

const slides = pptx.getSlides(presentation).map(readSlide);
printHeadlines(slides);
if (full) printFullText(slides);
if (mode === 'talk') printTalkMetrics(slides);
if (mode === 'document') printDocumentChecks(slides);
printWordingFlags(slides);
