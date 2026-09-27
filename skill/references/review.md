# Reviewing the words and the story

Run this after the deck builds and the visual check passes, before delivering a new deck
or a substantial rewrite. For a small edit, run step 1 on the affected slides only.

Detection is mechanical; judgement is yours. The model that wrote the deck cannot see its
own habits, so the checks list suspicions and a separate reviewer reads the deck cold.

## 1. Extract the text and the flags

From the slide project, after `npm run build`:

```sh
node <skill-directory>/scripts/deck-text.mjs deck.pptx --mode talk      # or --mode document
node <skill-directory>/scripts/deck-text.mjs deck.pptx --full           # every text and note
```

The script prints the headline track, mode-specific checks and wording flags. It finds a
headline by the shape name `Headline`, so give each slide's headline `Text` the prop
`name="Headline"`. Without it the script falls back to a title placeholder or the largest
text and marks the guess with `?`.

## 2. Judge every flag

For each flag, decide "fix" or "keep (reason)". Keep a short ledger while doing so.

- A flag inside a code sample, a quoted log or a product's official copy is kept.
- A flagged word that the project uses as a defined term is kept.
- Otherwise rewrite using [writing-ja.md](writing-ja.md) for Japanese wording, then rerun the
  script to confirm the fix did not create a new flag.
- Do not rewrite in a new formula: replacing every 「〜ではなく」 with the same new pattern is
  the same problem again.

## 3. Read the headline track

- **Talk**: read the headlines as the speaker's sequence of discoveries. Check that every
  divider is answered by the next slide, nothing answers before it is asked, the mix of
  forms is within reach of the reference ranges, and no pattern repeats three times in a
  row ([story-talk.md](story-talk.md)).
- **Document**: read the headlines alone as one argument. Check that the answer comes first,
  each headline is an assertion, peers share a form and nothing repeats
  ([story-document.md](story-document.md)).

Fix the storyline before polishing sentences; wording cannot rescue a missing step.

## 4. Check facts across pages

The mistakes readers and reviewers find first are not awkward words but pages that
contradict each other. The author misses them because each page was right when written.
List, in the ledger:

- every quantity that appears on more than one page (counts of paths, layers, types; percentages;
  dates), with each page's value. One quantity has one value, and a headline count matches
  the items drawn on that slide;
- every mechanism explained on more than one page (where a check runs, what a marker is,
  what a store holds), with each page's wording. They must describe the same thing; if the
  design changed over time, say which version each page shows;
- every claim on the closing or summary slide, with the body page that supports it. A
  summary that says "no manual steps" while body pages describe manual ones is a
  contradiction, not a simplification;
- the central terms, with every variant used for each.

Resolve each conflict from the source material. When the source does not settle it, weaken
the claim and ask the user; do not pick the version that reads better.

## 5. Fresh-eye review

Give the exported deck to a new agent that shares none of your context: do not mention how
the deck was made, the skill, the modes or what you were proud of. Replace `{FILE}`, `{N}`
and `{KIND}` (「登壇で話すための資料」 or 「読んで判断するための資料」) and pass this prompt:

```text
あなたはプレゼン資料のレビュアーです。{KIND}（{N} ページ）を読み、作った本人には見えない
破綻を指摘してください。どう作られたかは知らされません。

ファイル {FILE} を最後まで読み、全ページを順に確認してください。一部だけ読んで推測しない
こと。発表者ノートがあれば、話す内容として参照してかまいません。

1. 日本語の言い回し: 日本語話者が読んで引っかかる表現をすべて挙げる。不自然な語尾、口語と
   文語の混在、主語や述語が抜けた圧縮語、そろっていない箇条書きの語尾、長い連体修飾、
   定義のない略語や造語、同じ概念の表記ゆれ、AI が書いたように感じる言い回し。
2. 論理展開: 見出しだけを上から読み、話が飛ぶ・戻る・同じことを二度言う・前提なしに結論が
   出る箇所を挙げる。
3. 破綻: 見出しの数と本文の数の食い違い、見出しと図の結論の食い違い、裏づけのない評価語
   （十分・問題ない・限定的）、既出ページの焼き直し、冒頭の予告と本編のずれ、同じ量が
   ページごとに違う値、同じ仕組みの説明がページごとに違う箇所、まとめのページと本編の
   矛盾、数値の出所の欠落。

各指摘は「ページ番号・該当箇所の引用・何が問題か・言い換え案」で書く。最後に、直すべき順に
上位 5 件を挙げる。盛らずに書くこと。無傷の資料はないので、欠点を最低 3 つ挙げること。
```

Tabulate the findings (page, finding verbatim, type, adopt / reject / hold, action or
reason). Decide each one; do not apply all of them. Show the user any finding that would
change meaning or structure and let them decide. Apply the adopted ones, rebuild, rerun
steps 1–4 and the visual check.

One round is the default. Report the number of findings, adopted changes and remaining
concerns in three lines, and ask before running another round.
